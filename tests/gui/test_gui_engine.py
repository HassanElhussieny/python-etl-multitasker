"""End to end on temporary data: the desktop queue runs the REAL engine (``python -m sales_etl``) in a separate
process, the engine's lifecycle moves (or keeps) the files, and the CLI and the GUI work on the same folders."""

from __future__ import annotations

import hashlib
import json
from datetime import date

from gui.services.catalog import build_arguments
from gui.services.history import load_runs
from gui.services.jobs import Job, JobQueue, Status
from sales_etl.cli import main as cli_main
from tests.conftest import R, S, Txn, write_erp_csv
from tests.test_analysis import FIXED_CODES_NAME, write_fixed, write_main
from tests.gui.conftest import wait_until

NOV = ("01-11-2026", "16-11-2026")          # the ERP preamble period of the daily test file


def _txns():
    """Task 1 data with a correct ERP footer (the Daily task reconciles it)."""
    return [Txn("15-11-2026", S, "100.50", doc_no="1", customer="100001"),
            Txn("16-11-2026", S, "40.00", doc_no="2", customer="100002"),
            Txn("16-11-2026", R, "-20.25", doc_no="3", customer="100001")]


def _rows():
    from tests.test_analysis import row
    return [row("15-11-2026", S, "100.50", 100001), row("16-11-2026", S, "40.00", 100002),
            row("16-11-2026", R, "-20.25", 100001)]


def _sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def _run(qapp, s, task, values, read_only):
    q = JobQueue(s)
    job = q.add(Job(task, build_arguments(task, values, s, read_only=read_only), read_only))
    assert wait_until(qapp, lambda: job.status not in (Status.PENDING, Status.RUNNING), 180), job.text
    return job


def test_managed_customer_job_moves_the_file_after_verification(gui_env, qapp):
    tmp, s = gui_env
    src = write_erp_csv(tmp / "input" / "north.csv", _rows())
    sha = _sha(src)
    job = _run(qapp, s, "3", {}, read_only=False)
    assert job.status is Status.COMPLETED and job.exit_code == 0, job.text
    (run_dir,) = [p for p in (tmp / "output").iterdir()]
    assert run_dir.name.startswith("Customer_") and job.output_dirs == [str(run_dir)]
    assert job.reports and all(r.endswith(".xlsx") for r in job.reports)
    archived = tmp / "processed" / run_dir.name / "north.csv"
    assert archived.is_file() and _sha(archived) == sha and list((tmp / "input").iterdir()) == []
    assert (tmp / "state" / "customer_journal.jsonl").is_file()
    runs = load_runs(tmp / "output", tmp / "state")
    assert runs[0].status == "SUCCEEDED" and runs[0].files[0]["name"] == "north.csv"


def test_read_only_daily_job_never_moves_input(gui_env, qapp):
    tmp, s = gui_env
    src = write_erp_csv(tmp / "input" / "north.csv", _txns(), period=NOV)
    sha = _sha(src)
    job = _run(qapp, s, "1", {"ref_date": date(2026, 11, 16)}, read_only=True)
    assert job.status is Status.COMPLETED, job.text
    assert "مبيعات" in job.text and "�" not in job.text            # Arabic output decoded intact
    assert [p.name for p in (tmp / "input").iterdir()] == ["north.csv"] and _sha(src) == sha
    assert not any((tmp / "processed").iterdir()) and not any((tmp / "state").iterdir())
    assert any(r.endswith("sales_report_ref-2026-11-16.xlsx") for r in job.reports)
    assert "Reconciliation audit     : PASS" in job.text


def test_task5_option1_read_only_writes_four_reports(gui_env, qapp):
    tmp, s = gui_env
    write_erp_csv(tmp / "input" / "month.csv", _rows())
    write_fixed(tmp / "data" / FIXED_CODES_NAME, [100001, 555555])
    main_data = write_main(tmp / "data" / "main.xlsx")
    from tests.conftest import write_collection
    write_collection(tmp / "التحصيل" / "collection.xlsx", [(100001, 10, 20)])
    job = _run(qapp, s, "5-1", {"ref_date": date(2026, 11, 16), "main_data": str(main_data)}, read_only=True)
    assert job.status is Status.COMPLETED, job.text
    names = sorted(r.replace("\\", "/").split("/")[-1] for r in job.reports)
    assert len(names) == 4 and any(n.startswith("Customers_Not_Worked_") for n in names)
    assert [p.name for p in (tmp / "input").iterdir()] == ["month.csv"]


def test_engine_problem_is_shown_as_failed(gui_env, qapp):
    tmp, s = gui_env                                             # empty input: the engine reports no input
    job = _run(qapp, s, "3", {}, read_only=True)
    assert job.status is Status.FAILED and job.exit_code == 3 and "NO REPORT" in job.text


def test_cli_and_gui_share_the_same_folders(gui_env, qapp):
    tmp, s = gui_env
    write_erp_csv(tmp / "input" / "a.csv", _rows())
    code = cli_main(["--task", "4", "--input-dir", s.input_dir, "--output-dir", s.output_dir, "--processed-dir",
                     s.processed_dir, "--failures-dir", s.failures_dir, "--state-dir", s.state_dir,
                     "--non-interactive"])                         # the manual workflow, unchanged
    assert code == 0
    runs = load_runs(tmp / "output", tmp / "state")
    assert runs[0].folder.name.startswith("Representative_") and runs[0].status == "SUCCEEDED"
    write_erp_csv(tmp / "input" / "b.csv", _rows())
    job = _run(qapp, s, "4", {}, read_only=False)                # then the GUI on the same project
    assert job.status is Status.COMPLETED
    journal = (tmp / "state" / "representative_journal.jsonl").read_text(encoding="utf-8")
    assert sum(1 for line in journal.splitlines() if json.loads(line)["event"] == "RUN_FINISHED") == 2


def test_interrupting_a_managed_run_never_loses_a_file(gui_env, qapp):
    tmp, s = gui_env
    src = write_erp_csv(tmp / "input" / "north.csv", _rows())
    sha = _sha(src)
    q = JobQueue(s)
    job = q.add(Job("3", build_arguments("3", {}, s, read_only=False), False))
    assert wait_until(qapp, lambda: job.status is Status.RUNNING, 30)
    q.cancel(job.id)                                             # killed at once
    assert wait_until(qapp, lambda: job.finished is not None, 60)
    copies = list((tmp / "input").glob("north.csv")) + list((tmp / "processed").rglob("north.csv"))
    assert len(copies) == 1 and _sha(copies[0]) == sha           # in input or archived: never lost
    if copies[0].parent.name == "input":                         # the next run (recovery) handles it
        again = _run(qapp, s, "3", {}, read_only=False)
        assert again.status is Status.COMPLETED, again.text
        assert len(list((tmp / "processed").rglob("north.csv"))) == 1
