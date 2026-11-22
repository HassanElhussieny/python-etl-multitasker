"""The job queue with real processes: statuses, exit codes, parallel read-only jobs, managed jobs alone, the
engine's lock respected, cancellation."""

from __future__ import annotations

import sys

from gui.services.jobs import Job, JobQueue, Status
from gui.services.lock import LockState
from sales_etl.filestore import RunLock
from tests.gui.conftest import wait_until

PY = sys.executable


def fake(code: str, *, read_only: bool, task: str = "3") -> Job:
    """A job running a tiny Python script instead of the engine (task 3 = a task that may move files)."""
    return Job(task, ["-c", code], read_only, program=PY)


def test_exit_codes_become_statuses(gui_env, qapp):
    tmp, s = gui_env
    q = JobQueue(s)
    ok = q.add(fake("print('مرحبا')", read_only=True))
    bad = q.add(fake("import sys; sys.exit(1)", read_only=True))
    none = q.add(fake("import sys; sys.exit(3)", read_only=True))
    assert wait_until(qapp, lambda: all(j.status not in (Status.PENDING, Status.RUNNING) for j in (ok, bad, none)))
    assert ok.status is Status.COMPLETED and ok.exit_code == 0 and "مرحبا" in ok.text   # UTF-8 output
    assert bad.status is Status.FAILED and bad.exit_code == 1
    assert none.status is Status.FAILED and "لا توجد ملفات" in none.note
    assert ok.started and ok.finished and ok.elapsed() >= 0
    assert (tmp / "gui_jobs.jsonl").read_text(encoding="utf-8").count("\n") == 3          # job log kept


def test_failed_start_is_reported(gui_env, qapp):
    tmp, s = gui_env
    q = JobQueue(s)
    j = q.add(Job("3", ["-c", "1"], True, program=str(tmp / "no_python.exe")))
    assert wait_until(qapp, lambda: j.status is Status.FAILED, 30)
    assert "تعذّر تشغيل" in j.note


def test_read_only_jobs_run_in_parallel(gui_env, qapp):
    tmp, s = gui_env
    q = JobQueue(s)
    jobs = [q.add(fake("import time; time.sleep(2)", read_only=True)) for _ in range(2)]
    assert wait_until(qapp, lambda: all(j.status is Status.RUNNING for j in jobs), 20)
    assert wait_until(qapp, lambda: all(j.status is Status.COMPLETED for j in jobs))


def test_parallel_limit_is_respected(gui_env, qapp):
    tmp, s = gui_env
    s.max_parallel_readonly = 1
    q = JobQueue(s)
    a = q.add(fake("import time; time.sleep(1.5)", read_only=True))
    b = q.add(fake("print(1)", read_only=True))
    assert wait_until(qapp, lambda: a.status is Status.RUNNING, 20)
    assert b.status is Status.PENDING
    assert wait_until(qapp, lambda: b.status is Status.COMPLETED)
    assert b.started >= a.finished


def test_a_managed_job_runs_alone(gui_env, qapp):
    tmp, s = gui_env
    q = JobQueue(s)
    first = q.add(fake("import time; time.sleep(1.5)", read_only=True))
    managed = q.add(fake("import time; time.sleep(1.5)", read_only=False))
    after = q.add(fake("print('after')", read_only=True))
    assert managed.moves_files and not first.moves_files
    assert wait_until(qapp, lambda: first.status is Status.RUNNING, 20)
    assert managed.status is Status.PENDING and after.status is Status.PENDING      # FIFO, not beside it
    assert wait_until(qapp, lambda: managed.status is Status.RUNNING)
    assert after.status is Status.PENDING                                           # nothing beside a mover
    assert wait_until(qapp, lambda: after.status is Status.COMPLETED)
    assert managed.started >= first.finished and after.started >= managed.finished


def test_busy_lock_holds_jobs_back_until_released(gui_env, qapp):
    tmp, s = gui_env
    q = JobQueue(s)
    lock = RunLock(s.lock_file)
    lock.__enter__()                                   # e.g. a run started from the terminal
    try:
        assert q.lock_state() is LockState.BUSY
        managed = q.add(fake("print('m')", read_only=False))
        reader = q.add(fake("print('r')", read_only=True))
        wait_until(qapp, lambda: False, 3.5)
        assert managed.status is Status.PENDING and reader.status is Status.PENDING
        assert "القفل العام" in managed.note
    finally:
        lock.__exit__(None, None, None)
    assert wait_until(qapp, lambda: managed.status is Status.COMPLETED and reader.status is Status.COMPLETED)


def test_cancel_pending_and_running(gui_env, qapp):
    tmp, s = gui_env
    q = JobQueue(s)
    running = q.add(fake("import time; time.sleep(60)", read_only=False))
    pending = q.add(fake("print(1)", read_only=True))
    assert wait_until(qapp, lambda: running.status is Status.RUNNING, 20)
    q.cancel(pending.id)
    assert pending.status is Status.CANCELLED and pending.started is None
    q.cancel(running.id)
    assert wait_until(qapp, lambda: running.finished is not None, 20)
    assert running.status is Status.CANCELLED and running.elapsed() < 30
    assert "COMMIT" in running.note                                  # the recovery explanation for a mover
    q.clear_finished()
    assert q.jobs == []
