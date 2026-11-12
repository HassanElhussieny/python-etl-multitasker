"""Managed mode: re-processing, run folders, failures, crash recovery, honest run status."""

import json
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from sales_etl import managed as managed_module
from sales_etl import pipeline as pipeline_module
from sales_etl.config import PipelineConfig
from sales_etl.filestore import (
    FAILED,
    JOURNAL_NAME,
    LOCK_NAME,
    PROCESSED,
    RETAINED,
    Journal,
    MoveError,
    RunLock,
    RunLockedError,
)
from sales_etl.managed import (
    COMPLETED_WITH_PROBLEMS,
    FAILED_RUN,
    NO_INPUT_FILES,
    SUCCEEDED,
    run_managed,
)
from sales_etl.models import CheckStatus
from tests.conftest import R, S, Txn, write_erp_csv

REF = date(2026, 10, 18)  # the user's "yesterday"; deliberately different from the execution date
DAY = "2026-10-20"        # execution date used for folder names


class SimulatedCrash(BaseException):
    """Stands in for a power cut / killed process (not caught by normal handlers)."""


@pytest.fixture
def project(tmp_path):
    (tmp_path / "input").mkdir()
    return tmp_path


def run(project: Path, minute: int = 0, second: int = 0):
    config = PipelineConfig(input_dir=project / "input", output_dir=project / "output", reference_date=REF)
    return run_managed(config, now=datetime(2026, 10, 20, 10, minute, second))


def journal(project: Path) -> list[dict]:
    return Journal(project / "state" / JOURNAL_NAME).records()[0]


def events(project: Path) -> list[str]:
    return [r["event"] for r in journal(project)]


def names(folder: Path) -> list[str]:
    return sorted(p.name for p in folder.iterdir()) if folder.exists() else []


GOOD = [Txn("18-10-2026", S, "100", "1"), Txn("17-10-2026", S, "50", "2"),
        Txn("10-10-2026", R, "-20", "3"), Txn("19-10-2026", S, "999", "4")]


# --- success, folder layout ----------------------------------------------------------
def test_successful_file_is_archived_in_run_folder(project):
    src = write_erp_csv(project / "input" / "rep1.csv", GOOD)
    content = src.read_bytes()
    result = run(project)
    d = result.dispositions[0]
    assert (d.disposition, d.moved) == (PROCESSED, True)
    assert result.run_dir == project / "output" / f"Report_{DAY}"            # execution date, not REF
    assert result.processed_run_dir == project / "processed" / f"Run_{DAY}"
    assert names(project / "processed") == [f"Run_{DAY}"]                     # nothing in the root
    assert (project / "processed" / f"Run_{DAY}" / "rep1.csv").read_bytes() == content
    assert names(project / "input") == []
    m = result.pipeline.outcomes[0].metrics
    assert (m.sales_yesterday, m.sales_previous, m.returns_previous) == (Decimal("100"), Decimal("50"), Decimal("-20"))
    assert m.future_count == 1                                                 # 19/10 excluded
    assert result.status == SUCCEEDED and result.succeeded
    assert events(project) == ["RUN_STARTED", "COMMIT", "MOVED", "RUN_FINISHED"]


def test_manifest_matches_files_on_disk(project):
    write_erp_csv(project / "input" / "rep1.csv", GOOD)
    result = run(project)
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["status"] == SUCCEEDED and manifest["report_generated"] is True
    assert manifest["execution_date"] == DAY and manifest["reference_date"] == "2026-10-18"
    assert manifest["report_dir"] == str(project / "output" / f"Report_{DAY}")
    assert manifest["processed_run_dir"] == str(project / "processed" / f"Run_{DAY}")
    for entry in manifest["outputs"].values():
        assert entry["exists"] is True and Path(entry["path"]).is_file()
        assert Path(entry["path"]).parent == result.run_dir
    assert manifest["files"][0]["current_location"] == str(project / "processed" / f"Run_{DAY}" / "rep1.csv")


def test_multiple_runs_same_day_get_unique_folders(project):
    write_erp_csv(project / "input" / "a.csv", GOOD)
    first = run(project, minute=5)
    write_erp_csv(project / "input" / "a.csv", GOOD)  # same name again
    second = run(project, minute=7, second=9)
    write_erp_csv(project / "input" / "a.csv", GOOD)
    third = run(project, minute=7, second=9)          # identical clock -> numeric suffix
    assert [r.run_dir.name for r in (first, second, third)] == [
        f"Report_{DAY}", f"Report_{DAY}_100709", f"Report_{DAY}_100709_2"]
    assert names(project / "processed") == [f"Run_{DAY}", f"Run_{DAY}_100709", f"Run_{DAY}_100709_2"]
    for r in (first, second, third):
        assert r.pipeline.report_path.is_file() and r.manifest_path.is_file()
        assert names(r.processed_run_dir) == ["a.csv"]


def test_existing_report_and_archive_folders_are_never_overwritten(project):
    (project / "output" / f"Report_{DAY}").mkdir(parents=True)
    (project / "output" / f"Report_{DAY}" / "old.txt").write_text("keep")
    (project / "processed" / f"Run_{DAY}_100000").mkdir(parents=True)  # archive name already taken
    (project / "processed" / f"Run_{DAY}_100000" / "a.csv").write_text("keep")
    write_erp_csv(project / "input" / "a.csv", GOOD)
    result = run(project)
    assert result.run_dir.name == f"Report_{DAY}_100000_2"  # both folders free under this stamp
    assert result.processed_run_dir.name == f"Run_{DAY}_100000_2"
    assert (project / "output" / f"Report_{DAY}" / "old.txt").read_text() == "keep"
    assert (project / "processed" / f"Run_{DAY}_100000" / "a.csv").read_text() == "keep"


def test_historical_flat_processed_files_are_left_alone(project):
    (project / "processed").mkdir()
    (project / "processed" / "rep1.csv").write_bytes(b"historical file")
    write_erp_csv(project / "input" / "rep1.csv", GOOD)
    run(project)
    assert (project / "processed" / "rep1.csv").read_bytes() == b"historical file"
    assert names(project / "processed") == [f"Run_{DAY}", "rep1.csv"]


# --- no duplicate handling of any kind -----------------------------------------------
def test_previously_processed_content_is_processed_again(project):
    src = write_erp_csv(project / "input" / "rep1.csv", GOOD)
    content = src.read_bytes()
    first = run(project)
    (project / "input" / "rep1.csv").write_bytes(content)        # same name, same content
    (project / "input" / "rep1_copy.csv").write_bytes(content)   # renamed copy
    second = run(project, minute=1)
    assert {d.file_name: d.disposition for d in second.dispositions} == {
        "rep1.csv": PROCESSED, "rep1_copy.csv": PROCESSED}
    assert sorted(o.file_name for o in second.pipeline.outcomes) == ["rep1.csv", "rep1_copy.csv"]
    assert second.pipeline.grand_total.sales_yesterday == Decimal("200")  # counted again, by design
    assert second.notices == []                                  # no "already processed" notice
    assert not (project / "duplicates").exists()
    assert not (project / "state" / "fingerprints").exists()     # no fingerprints kept
    assert names(second.processed_run_dir) == ["rep1.csv", "rep1_copy.csv"]
    assert names(first.processed_run_dir) == ["rep1.csv"]
    assert second.status == SUCCEEDED


def test_identical_files_in_one_run_are_both_processed(project):
    src = write_erp_csv(project / "input" / "a.csv", GOOD)
    (project / "input" / "b.csv").write_bytes(src.read_bytes())
    result = run(project)
    assert [d.disposition for d in result.dispositions] == [PROCESSED, PROCESSED]
    assert result.pipeline.grand_total.sales_yesterday == Decimal("200")
    assert names(result.processed_run_dir) == ["a.csv", "b.csv"]


def test_journal_history_does_not_block_processing(project):
    """Existing history (incl. old DUPLICATE dispositions) is kept and never consulted to skip."""
    write_erp_csv(project / "input" / "rep1.csv", GOOD)
    run(project)
    history = (project / "state" / JOURNAL_NAME).read_text(encoding="utf-8")
    write_erp_csv(project / "input" / "rep1.csv", GOOD)
    second = run(project, minute=1)
    assert second.dispositions[0].disposition == PROCESSED
    assert (project / "state" / JOURNAL_NAME).read_text(encoding="utf-8").startswith(history)  # append-only


# --- failures --------------------------------------------------------------------------
def test_failed_files_go_to_failures_without_loss(project):
    (project / "input" / "empty.csv").write_bytes(b"")
    bad = write_erp_csv(project / "input" / "bad_rows.csv",
                        [Txn("18-10-2026", S, "100", "1"), Txn("31-02-2026", S, "5", "2")])
    bad_content = bad.read_bytes()
    write_erp_csv(project / "input" / "good.csv", GOOD)
    result = run(project)
    by_name = {d.file_name: d for d in result.dispositions}
    assert by_name["empty.csv"].disposition == FAILED and "EMPTY_FILE" in by_name["empty.csv"].reason
    assert by_name["bad_rows.csv"].disposition == FAILED and "NO_REJECTED_ROWS" in by_name["bad_rows.csv"].reason
    assert names(project / "failures") == ["bad_rows.csv", "empty.csv"]          # unchanged behaviour
    assert names(result.processed_run_dir) == ["good.csv"]                       # failed files never archived
    assert (project / "failures" / "bad_rows.csv").read_bytes() == bad_content
    assert result.status == COMPLETED_WITH_PROBLEMS and not result.succeeded


def test_failed_file_name_clash_never_overwrites(project):
    (project / "input" / "x.csv").write_bytes(b"")
    run(project)
    (project / "input" / "x.csv").write_bytes(b"")
    second = run(project, minute=1)
    assert second.dispositions[0].disposition == FAILED
    assert len(names(project / "failures")) == 2


# --- run status --------------------------------------------------------------------------
def test_empty_input_is_not_reported_as_success(project):
    result = run(project)
    assert result.status == NO_INPUT_FILES and not result.succeeded
    assert result.run_dir is None and names(project / "output") == []


def test_unverified_outputs_mean_failed_run_and_no_move(project, monkeypatch):
    write_erp_csv(project / "input" / "rep1.csv", GOOD)
    monkeypatch.setattr(managed_module, "verify_outputs", lambda result: ["audit JSON unreadable"])
    result = run(project)
    d = result.dispositions[0]
    assert d.disposition == RETAINED and not d.moved and "audit JSON unreadable" in d.reason
    assert names(project / "input") == ["rep1.csv"] and not result.processed_run_dir.exists()
    assert "COMMIT" not in events(project)
    assert result.status == FAILED_RUN
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["status"] == FAILED_RUN and manifest["report_generated"] is False


def test_report_write_error_gives_failed_manifest_and_keeps_file(project, monkeypatch):
    write_erp_csv(project / "input" / "rep1.csv", GOOD)

    def disk_full(*args, **kwargs):
        raise OSError("No space left on device")

    monkeypatch.setattr(pipeline_module, "write_report", disk_full)
    with pytest.raises(OSError):
        run(project)
    assert names(project / "input") == ["rep1.csv"]
    assert events(project) == ["RUN_STARTED", "RUN_FINISHED"] and journal(project)[-1]["status"] == FAILED_RUN
    manifest = json.loads((project / "output" / f"Report_{DAY}" / "run_manifest.json").read_text(encoding="utf-8"))
    assert manifest["status"] == FAILED_RUN and manifest["report_generated"] is False
    assert manifest["outputs"]["report"]["exists"] is False
    assert "No space left" in manifest["error"]
    monkeypatch.undo()
    assert run(project, minute=1).status == SUCCEEDED


# --- crash recovery ----------------------------------------------------------------------
def test_crash_before_commit_keeps_file_and_reprocesses(project, monkeypatch):
    write_erp_csv(project / "input" / "rep1.csv", GOOD)

    def crash(*args, **kwargs):
        raise SimulatedCrash()

    monkeypatch.setattr(pipeline_module, "write_audit_json", crash)
    with pytest.raises(SimulatedCrash):
        run(project)
    assert names(project / "input") == ["rep1.csv"] and names(project / "processed") == []
    assert events(project) == ["RUN_STARTED"]

    monkeypatch.undo()
    second = run(project, minute=1)
    assert [a.action for a in second.recovery] == ["RUN_VOIDED"]
    assert "VOID" in (project / "output" / f"Report_{DAY}" / "RUN_INTERRUPTED.txt").read_text(encoding="utf-8")
    assert second.status == SUCCEEDED and names(second.processed_run_dir) == ["rep1.csv"]


def test_file_left_in_input_after_interrupted_move_is_processed_again(project, monkeypatch):
    write_erp_csv(project / "input" / "rep1.csv", GOOD)

    def crash(*args, **kwargs):
        raise SimulatedCrash()

    monkeypatch.setattr(managed_module, "move_file", crash)
    with pytest.raises(SimulatedCrash):
        run(project)
    assert names(project / "input") == ["rep1.csv"] and events(project) == ["RUN_STARTED", "COMMIT"]

    monkeypatch.undo()
    second = run(project, minute=1)
    assert [a.action for a in second.recovery] == ["LEFT_IN_INPUT", "RUN_INTERRUPTED_AFTER_COMMIT"]
    assert second.pipeline is not None and second.dispositions[0].disposition == PROCESSED  # not skipped
    assert names(second.processed_run_dir) == ["rep1.csv"] and names(project / "input") == []
    assert "MOVE_ABANDONED" in events(project)  # old pending move closed as superseded
    third = run(project, minute=2)
    assert third.recovery == [] and third.status == NO_INPUT_FILES


def test_failed_move_leaves_file_and_next_run_processes_it(project, monkeypatch):
    write_erp_csv(project / "input" / "rep1.csv", GOOD)

    def locked(*args, **kwargs):
        raise MoveError("file is open in another program")

    monkeypatch.setattr(managed_module, "move_file", locked)
    first = run(project)
    assert not first.dispositions[0].moved
    assert first.status == COMPLETED_WITH_PROBLEMS  # report exists, but the file was not archived
    monkeypatch.undo()
    second = run(project, minute=1)
    assert second.dispositions[0].disposition == PROCESSED and second.status == SUCCEEDED
    assert names(project / "input") == []


def test_torn_journal_line_is_ignored_safely(project):
    write_erp_csv(project / "input" / "rep1.csv", GOOD)
    (project / "state").mkdir()
    (project / "state" / JOURNAL_NAME).write_text('{"event": "COMMIT", "run_id": "x", "fil', encoding="utf-8")
    result = run(project)
    assert any("interrupted write" in n for n in result.notices)
    assert result.dispositions[0].disposition == PROCESSED
    assert events(project)[-1] == "RUN_FINISHED"


def test_concurrent_run_is_refused(project):
    write_erp_csv(project / "input" / "rep1.csv", GOOD)
    with RunLock(project / "state" / LOCK_NAME):
        with pytest.raises(RunLockedError):
            run(project)
    assert names(project / "input") == ["rep1.csv"]


def test_changed_version_of_a_processed_file_is_just_processed(project):
    write_erp_csv(project / "input" / "rep1.csv", GOOD[:2])
    run(project)
    write_erp_csv(project / "input" / "rep1.csv", GOOD)
    second = run(project, minute=1)
    assert not any(c.check == "DUPLICATE_SCAN" for c in second.pipeline.checks)    # no duplicate check
    assert not hasattr(second.pipeline, "duplicates") and second.notices == []
    assert second.status == SUCCEEDED and second.pipeline.outcomes[0].metrics.sales_yesterday == Decimal("100")
