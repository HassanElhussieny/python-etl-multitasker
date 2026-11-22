"""The GUI's non-visual services: command building (accepted by the real CLI parser), lock probe, input scan,
history read-back."""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pytest

from gui.services import catalog
from gui.services.catalog import TASKS, ArgumentError, build_arguments, moves_files, task_for_folder
from gui.services.history import load_runs
from gui.services.lock import LockState, probe_lock
from gui.services.scanner import list_input_files, preflight
from gui.services.settings import AppSettings, load_settings, save_settings
from sales_etl.cli import build_parser
from sales_etl.filestore import RunLock
from tests.conftest import R, S, write_erp_csv
from tests.test_analysis import row

D16 = date(2026, 11, 16)
VALUES = {"1": {"ref_date": D16}, "2": {"start_date": date(2026, 11, 1), "end_date": D16},
          "5-1": {"ref_date": D16}, "5-2": {}, "5-3": {}}


@pytest.mark.parametrize("task", [t.key for t in TASKS])
@pytest.mark.parametrize("read_only", [True, False])
def test_every_task_builds_a_command_the_cli_accepts(gui_env, task, read_only):
    tmp, s = gui_env
    args = build_arguments(task, VALUES.get(task, {}), s, read_only=read_only)
    assert args[:2] == ["-m", "sales_etl"] and args[-1] == "--non-interactive"
    parsed = build_parser().parse_args(args[2:])                 # the real CLI parser: no SystemExit
    assert parsed.task == catalog.BY_KEY[task].task and parsed.input_dir == Path(s.input_dir)
    assert parsed.output_dir == Path(s.output_dir) and parsed.non_interactive
    if catalog.BY_KEY[task].read_only_only:
        assert not parsed.no_move and parsed.processed_dir is None    # option 3: no lifecycle at all
    elif read_only:
        assert parsed.no_move and parsed.processed_dir is None
    else:
        assert not parsed.no_move and parsed.state_dir == Path(s.state_dir)
        assert parsed.processed_dir == Path(s.processed_dir) and parsed.failures_dir == Path(s.failures_dir)
    assert moves_files(task, read_only) == (not read_only and not catalog.BY_KEY[task].read_only_only)


def test_exact_arguments_for_dated_tasks(gui_env):
    tmp, s = gui_env
    a1 = build_arguments("1", {"ref_date": D16}, s, read_only=True)
    assert a1[2:6] == ["--task", "1", "--ref-date", "16/11/2026"] and "--no-move" in a1
    a2 = build_arguments("2", VALUES["2"], s, read_only=True)
    assert a2[4:8] == ["--start-date", "01-11-2026", "--end-date", "16-11-2026"]
    a5 = build_arguments("5-1", {"ref_date": D16}, s, read_only=False)
    assert a5[2:8] == ["--task", "5", "--report-type", "1", "--ref-date", "16/11/2026"]
    assert a5[a5.index("--collection-dir") + 1] == s.collection_dir       # always explicit
    assert build_arguments("5-3", {}, s, read_only=False)[2:6] == ["--task", "5", "--report-type", "3"]


def test_missing_or_invalid_values_are_refused(gui_env):
    tmp, s = gui_env
    with pytest.raises(ArgumentError):
        build_arguments("1", {}, s, read_only=True)
    with pytest.raises(ArgumentError):
        build_arguments("2", {"start_date": D16, "end_date": date(2026, 11, 1)}, s, read_only=True)
    with pytest.raises(ArgumentError):
        build_arguments("5-3", {"sales_csv": str(tmp / "nope.csv")}, s, read_only=True)


def test_output_folders_map_to_their_task():
    assert task_for_folder("Report_2026-11-17").key == "1"
    assert task_for_folder("Task5_Receivables_2026-11-16").key == "5-3"
    assert task_for_folder("Task5_2026-11-17_092821").key in ("5-1", "5-2")
    assert task_for_folder("صافي_العملاء_الشهري_2026-11-01").key == "6"
    assert task_for_folder("run_ref-2026-10-01") is None


def test_settings_default_to_the_cli_folders_and_round_trip(gui_env):
    tmp, s = gui_env
    d = AppSettings(project_root=str(tmp))
    assert d.input_dir == str(tmp / "input") and d.state_dir == str(tmp / "state") and d.is_default_layout()
    path = save_settings(s)
    assert load_settings(path) == s


def test_lock_probe_never_takes_the_lock(gui_env):
    tmp, s = gui_env
    assert probe_lock(s.lock_file) is LockState.FREE and not s.lock_file.exists()   # not created
    with RunLock(s.lock_file):                                       # the engine's own lock
        assert probe_lock(s.lock_file) is LockState.BUSY
    assert probe_lock(s.lock_file) is LockState.FREE
    with RunLock(s.lock_file):                                       # the probe released it again
        pass


def test_preflight_lists_files_dates_and_warnings(gui_env):
    tmp, s = gui_env
    write_erp_csv(tmp / "input" / "year.csv", [row("12-01-2026", S, "10.00", 100001),
                                               row("17-11-2026", R, "-5.00", 100002)])
    write_erp_csv(tmp / "input" / "nov.csv", [row("03-11-2026", S, "1.00", 100003)])
    assert [f.name for f in list_input_files(tmp / "input")] == ["nov.csv", "year.csv"]
    p = preflight(tmp / "input", "1", {"ref_date": D16})
    by = {f.name: f for f in p.files}
    assert by["year.csv"].rows == 2 and by["year.csv"].first_date == date(2026, 1, 12)
    assert by["year.csv"].last_date == date(2026, 11, 17) and p.first_date == date(2026, 1, 12)
    assert any("year.csv" in w and "قبل بداية شهر" in w for w in p.warnings)     # the 17/11 case
    assert any("year.csv" in w and "بعد 16/11/2026" in w for w in p.warnings)
    assert not any("nov.csv" in w for w in p.warnings)
    assert sorted(x.name for x in (tmp / "input").iterdir()) == ["nov.csv", "year.csv"]   # nothing moved
    assert preflight(tmp / "empty_dir_missing", "3", {}).warnings                         # no files: warned


def test_history_reads_journals_manifests_and_interrupted_runs(gui_env):
    tmp, s = gui_env
    out, state = tmp / "output", tmp / "state"
    (out / "Customer_2026-11-20").mkdir()
    (out / "Customer_2026-11-20" / "r.xlsx").write_bytes(b"x")
    (out / "Customer_2026-11-20" / "~$r.xlsx").write_bytes(b"x")
    (out / "Report_2026-11-21").mkdir()
    (out / "Report_2026-11-21" / "run_manifest.json").write_text(json.dumps({"status": "COMPLETED_WITH_PROBLEMS"}))
    (out / "Report_2026-11-21" / "audit_ref-2026-11-20.json").write_text(json.dumps({"audit_status": "FAIL"}))
    (out / "Period_2026-11-22").mkdir()
    (out / "Period_2026-11-22" / "RUN_INTERRUPTED.txt").write_text("void")
    (out / "Representative_2026-11-22").mkdir()                      # no journal: a read-only run
    lines = [{"event": "RUN_STARTED", "run_id": "Customer_2026-11-20", "files": [{"file_name": "a.csv"}]},
             {"event": "COMMIT", "run_id": "Customer_2026-11-20",
              "files": [{"file_name": "a.csv", "disposition": "PROCESSED", "destination": "x"}]},
             {"event": "RUN_FINISHED", "run_id": "Customer_2026-11-20", "status": "SUCCEEDED"}]
    (state / "customer_journal.jsonl").write_text("\n".join(map(json.dumps, lines)) + "\n{torn", encoding="utf-8")
    runs = {r.folder.name: r for r in load_runs(out, state)}
    assert runs["Customer_2026-11-20"].status == "SUCCEEDED" and runs["Customer_2026-11-20"].source == "journal"
    assert runs["Customer_2026-11-20"].files[0]["disposition"] == "PROCESSED"
    assert [p.name for p in runs["Customer_2026-11-20"].reports] == ["r.xlsx"]
    assert runs["Report_2026-11-21"].status == "COMPLETED_WITH_PROBLEMS" and runs["Report_2026-11-21"].audit == "FAIL"
    assert runs["Period_2026-11-22"].status == "INTERRUPTED"
    assert runs["Representative_2026-11-22"].status == "READ_ONLY"
