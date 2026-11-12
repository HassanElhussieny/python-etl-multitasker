"""File lifecycle of tasks 3 (Customer) and 4 (Representative): move only after the report is saved."""

import json

import pytest

from sales_etl import grouped
from sales_etl.cli import EXIT_AUDIT_FAILED, EXIT_NO_INPUT, EXIT_OK, main
from tests.conftest import R, S, Txn, write_erp_csv

ROWS = [Txn("01-10-2026", S, "100.00", "1", "400100"), Txn("02-10-2026", R, "-10.00", "2", "400100")]
TASKS = pytest.mark.parametrize("task,name", [("3", "Customer"), ("4", "Representative")])


def _run(tmp_path, task, *extra):
    return main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--task", task, "--non-interactive", *extra])


@TASKS
def test_success_moves_files_to_own_processed_folder(tmp_path, task, name):
    a = write_erp_csv(tmp_path / "in" / "a.csv", ROWS)
    write_erp_csv(tmp_path / "in" / "b.csv", ROWS)
    original = a.read_bytes()
    assert _run(tmp_path, task) == EXIT_OK
    (report_dir,) = (tmp_path / "out").iterdir()
    (archive,) = (tmp_path / "processed").iterdir()
    assert report_dir.name.startswith(f"{name}_") and archive.name == report_dir.name   # own folder, same stamp
    assert sorted(p.name for p in archive.iterdir()) == ["a.csv", "b.csv"]
    assert (archive / "a.csv").read_bytes() == original
    assert list((tmp_path / "in").iterdir()) == []
    journal = tmp_path / "state" / f"{name.lower()}_journal.jsonl"
    events = [json.loads(line)["event"] for line in journal.read_text(encoding="utf-8").splitlines()]
    assert events == ["RUN_STARTED", "COMMIT", "MOVED", "MOVED", "RUN_FINISHED"]
    assert not (tmp_path / "state" / "processing_journal.jsonl").exists()             # Daily untouched
    assert not (tmp_path / "state" / "period_journal.jsonl").exists()                 # Period untouched


@TASKS
def test_failed_file_is_not_moved(tmp_path, task, name):
    write_erp_csv(tmp_path / "in" / "good.csv", ROWS)
    (tmp_path / "in" / "empty.csv").write_bytes(b"")                                  # unreadable
    write_erp_csv(tmp_path / "in" / "bad.csv", [*ROWS, Txn("xx-10-2026", S, "5.00", "9")])  # rejected row
    assert _run(tmp_path, task) == EXIT_AUDIT_FAILED
    assert sorted(p.name for p in (tmp_path / "in").iterdir()) == ["bad.csv", "empty.csv"]
    assert [p.name for p in (tmp_path / "processed").rglob("*.csv")] == ["good.csv"]
    assert not (tmp_path / "failures").exists() or not list((tmp_path / "failures").iterdir())


@TASKS
def test_report_generation_failure_moves_nothing(tmp_path, task, name, monkeypatch, capsys):
    a = write_erp_csv(tmp_path / "in" / "a.csv", ROWS)
    original = a.read_bytes()

    def broken(result, path):
        raise OSError("disk full")

    monkeypatch.setattr(grouped, "write_group_report", broken)
    assert _run(tmp_path, task) == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["a.csv"] and a.read_bytes() == original
    assert not (tmp_path / "processed").exists() or not list((tmp_path / "processed").rglob("*"))
    assert "No file was moved" in capsys.readouterr().err


@TASKS
def test_unverified_report_moves_nothing(tmp_path, task, name, monkeypatch):
    write_erp_csv(tmp_path / "in" / "a.csv", ROWS)
    monkeypatch.setattr(grouped, "verify_group_report", lambda result: ["report header mismatch"])
    assert _run(tmp_path, task) == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["a.csv"]


@TASKS
def test_unrelated_files_are_not_moved(tmp_path, task, name):
    write_erp_csv(tmp_path / "in" / "a.csv", ROWS)
    (tmp_path / "in" / "notes.txt").write_text("keep me", encoding="utf-8")
    (tmp_path / "in" / "book.xlsx").write_bytes(b"not a csv")
    (tmp_path / "in" / "~$lock.csv").write_bytes(b"temp")
    (tmp_path / "in" / "sub").mkdir()
    write_erp_csv(tmp_path / "in" / "sub" / "nested.csv", ROWS)
    assert _run(tmp_path, task) == EXIT_OK
    assert sorted(p.name for p in (tmp_path / "in").iterdir()) == ["book.xlsx", "notes.txt", "sub", "~$lock.csv"]
    assert (tmp_path / "in" / "sub" / "nested.csv").exists()
    assert [p.name for p in (tmp_path / "processed").rglob("*") if p.is_file()] == ["a.csv"]


@TASKS
def test_same_name_twice_never_overwrites(tmp_path, task, name):
    a = write_erp_csv(tmp_path / "in" / "a.csv", ROWS)
    content = a.read_bytes()
    assert _run(tmp_path, task) == EXIT_OK
    (tmp_path / "in" / "a.csv").write_bytes(content)
    assert _run(tmp_path, task) == EXIT_OK                       # no duplicate detection: counted again
    assert len(list((tmp_path / "out").iterdir())) == 2
    assert sorted(p.name for p in (tmp_path / "processed").rglob("*.csv")) == ["a.csv", "a.csv"]
    assert len(list((tmp_path / "processed").iterdir())) == 2


@TASKS
def test_empty_input(tmp_path, task, name):
    (tmp_path / "in").mkdir()
    assert _run(tmp_path, task) == EXIT_NO_INPUT
    assert not (tmp_path / "out").exists() and not (tmp_path / "processed").exists()
