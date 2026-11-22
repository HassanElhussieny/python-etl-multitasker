"""Sidebar navigation: the right page every time, fast with a large output folder (the old refill of a few
hundred rows took 10-40 s per visit), nothing duplicated, no ETL work or file changes, responsive during a run."""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import pytest
from PySide6.QtCore import SIGNAL
from PySide6.QtWidgets import QTableWidget

from gui.app import create_application
from gui.services.jobs import Job, Status
from gui.windows.main_window import PAGES, MainWindow
from tests.gui.conftest import wait_until

RUNS = 160                                       # like the real output folder (142 runs, 291 reports)


@pytest.fixture
def big_window(gui_env, qapp):
    tmp, s = gui_env
    out, journal = tmp / "output", []
    for i in range(RUNS):
        folder = out / f"Customer_2026-10-{i % 28 + 1:02d}_{i:06d}"
        folder.mkdir()
        for name in ("customer_totals.xlsx", "extra.xlsx"):
            (folder / name).write_bytes(b"x" * 100)
        journal.append({"event": "RUN_FINISHED", "run_id": folder.name, "status": "SUCCEEDED"})
    (tmp / "state" / "customer_journal.jsonl").write_text("\n".join(map(json.dumps, journal)), encoding="utf-8")
    create_application([])
    w = MainWindow(s)
    w.show()
    qapp.processEvents()
    yield w, tmp, s
    w.close()


def _nav(qapp, w, key: str) -> float:
    i = [k for k, _ in PAGES].index(key)
    t = time.perf_counter()
    w.nav.setCurrentRow(i)
    qapp.processEvents()
    return time.perf_counter() - t


def _files(tmp: Path) -> list[tuple[str, int]]:
    return sorted((str(p), p.stat().st_mtime_ns) for p in tmp.rglob("*") if p.is_file())


def test_every_click_opens_the_right_page_quickly(big_window, qapp):
    w, tmp, s = big_window
    order = ["reports", "history", "monitor", "settings", "run"]
    first = {k: _nav(qapp, w, k) for k in order}
    assert w.pages["reports"].table.rowCount() == RUNS * 2 and w.pages["history"].table.rowCount() == RUNS
    timings = []
    for _ in range(4):                                       # repeated navigation
        for key in order:
            timings.append((key, _nav(qapp, w, key)))
            assert w.stack.currentWidget() is w.pages[key], key
            assert w.nav.currentItem().text() == dict(PAGES)[key]
    assert max(t for _, t in timings) < 0.5, timings         # unchanged data: no reload at all
    assert max(first.values()) < 3.0, first                  # first visit: one scan + one fill


def test_a_full_reload_stays_fast(big_window, qapp):
    w, tmp, s = big_window
    _nav(qapp, w, "reports")
    _nav(qapp, w, "history")
    for page in ("reports", "history"):                      # the table is already full: the slow case before
        t = time.perf_counter()
        w.pages[page].refresh()
        assert time.perf_counter() - t < 3.0
    (tmp / "output" / "Customer_2026-11-30_new").mkdir()     # a new run appears: the next visit reloads it
    (tmp / "output" / "Customer_2026-11-30_new" / "customer_totals.xlsx").write_bytes(b"x")
    _nav(qapp, w, "run")
    assert _nav(qapp, w, "reports") < 3.0
    assert w.pages["reports"].table.rowCount() == RUNS * 2 + 1


def test_navigation_accumulates_nothing_and_runs_nothing(big_window, qapp):
    w, tmp, s = big_window
    q = w.queue
    receivers = {sig: q.receivers(SIGNAL(sig)) for sig in ("job_changed(QString)", "job_added(QString)",
                                                           "output(QString,QString)")}
    tables = {k: len(p.findChildren(QTableWidget)) for k, p in w.pages.items()}
    widgets = {k: len(p.findChildren(object)) for k, p in w.pages.items()}
    before = _files(tmp)
    for _ in range(6):
        for key, _ in PAGES:
            _nav(qapp, w, key)
    assert {sig: q.receivers(SIGNAL(sig)) for sig in receivers} == receivers     # no duplicate connections
    assert {k: len(p.findChildren(QTableWidget)) for k, p in w.pages.items()} == tables
    assert {k: len(p.findChildren(object)) for k, p in w.pages.items()} == widgets   # no duplicate widgets
    assert list(w.pages) == [k for k, _ in PAGES]                                    # the same page objects
    assert q.jobs == []                                                              # no ETL run started
    assert _files(tmp) == before                                                     # nothing written


def test_responsive_while_a_task_runs(big_window, qapp):
    w, tmp, s = big_window
    job = w.queue.add(Job("3", ["-c", "import time; time.sleep(4)"], True, program=sys.executable))
    assert wait_until(qapp, lambda: job.status is Status.RUNNING, 20)
    timings = []
    while job.status is Status.RUNNING:
        for key, _ in PAGES:
            timings.append(_nav(qapp, w, key))
            assert w.stack.currentWidget() is w.pages[key]
        if len(timings) > 200:
            break
    assert wait_until(qapp, lambda: job.status is Status.COMPLETED, 30)
    assert timings and max(timings) < 0.5, max(timings)
    _nav(qapp, w, "monitor")
    assert w.pages["monitor"].table.rowCount() == 1                                  # the run is listed


def test_reports_stay_accessible(big_window, qapp, monkeypatch):
    w, tmp, s = big_window
    _nav(qapp, w, "reports")
    page = w.pages["reports"]
    page.table.setCurrentCell(0, 0)
    opened = []
    monkeypatch.setattr("gui.pages.reports.open_path", lambda path, parent=None: opened.append(Path(path)))
    page._open(False)
    assert opened and opened[0].suffix == ".xlsx" and opened[0].is_file()
    page.search.setText("extra")
    visible = [r for r in range(page.table.rowCount()) if not page.table.isRowHidden(r)]
    assert len(visible) == RUNS


def test_self_test_flag_is_routed_and_normal_start_is_not(monkeypatch):
    import gui.main as entry
    import gui.selftest as selftest
    calls = []
    monkeypatch.setattr(selftest, "run_selftest", lambda argv: calls.append(argv) or 0)
    assert entry.main(["--self-test", "r.json", "--sandbox", "x"]) == 0
    assert calls == [["--self-test", "r.json", "--sandbox", "x"]]
