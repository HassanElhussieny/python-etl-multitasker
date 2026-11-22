"""The simple workflow: choose task -> choose file(s) from any folder -> run -> download the report(s).
Everything on temporary data; the files are chosen outside the project folders."""

from __future__ import annotations

import hashlib
from datetime import date
from pathlib import Path

import pytest
from PySide6.QtCore import QDate
from PySide6.QtWidgets import QFileDialog, QMessageBox, QPushButton

from gui.app import create_application
from gui.services.catalog import TASKS
from gui.services.jobs import Status
from gui.services.staging import remove_staging, save_copy, stage_files, staging_root
from gui.windows.main_window import MainWindow
from tests.conftest import R, S, write_collection, write_erp_csv
from tests.gui.conftest import wait_until
from tests.test_analysis import FIXED_CODES_NAME, row, write_fixed, write_main


def _sha(p: Path) -> str:
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


@pytest.fixture
def window(gui_env, qapp):
    tmp, s = gui_env
    create_application([])
    w = MainWindow(s)
    w.show()
    qapp.processEvents()
    yield w, w.pages["run"], tmp, s
    w.close()


def _outside(tmp: Path) -> tuple[Path, Path]:
    """Two CSV files in two folders outside the project's input folder."""
    a = write_erp_csv(tmp / "elsewhere" / "D" / "north.csv", [row("15-11-2026", S, "100.50", 100001),
                                                              row("16-11-2026", R, "-20.25", 100001)])
    b = write_erp_csv(tmp / "other drive" / "south.csv", [row("10-11-2026", S, "40.00", 100002)])
    return a, b


def _wait_job(qapp, page):
    assert wait_until(qapp, lambda: page.job is not None and page.job.status not in (Status.PENDING, Status.RUNNING),
                      180), page.job.text if page.job else "no job"
    return page.job


# ------------------------------------------------------------------ 1. task selection
def test_selecting_each_task_shows_its_own_details(window):
    w, page, tmp, s = window
    assert w.stack.currentWidget() is page                         # the run screen is the first screen
    for i, t in enumerate(TASKS):
        page.task_list.setCurrentRow(i)
        assert page.title.text() == t.title and page.description.text() == t.description
        assert set(page.fields) == {p.key for p in t.visible_params}
        assert page.params_card.isVisibleTo(page) == bool(t.visible_params)
        assert page.task is t
    page.task_list.setCurrentRow(6)                                # Task 5 option 3: the reported mismatch
    assert page.title.text() == TASKS[6].title and "Receivables" in page.title.text()
    assert page.fields == {} and "Excel" in page.file_hint.text()
    page.task_list.setCurrentRow(1)                                # Period: two dates only
    assert set(page.fields) == {"start_date", "end_date"}
    page.task_list.setCurrentRow(4)                                # Task 5 option 1: the date only
    assert set(page.fields) == {"ref_date"} and page.fields["ref_date"].text().count("/") == 2


def test_no_queue_or_mode_controls_on_the_main_screen(window):
    w, page, tmp, s = window
    texts = {b.text() for b in page.findChildren(QPushButton)}
    assert {"اختيار الملف", "تشغيل التقرير"} <= texts
    assert not any("المحددة" in t for t in texts)                 # no multi-task add
    assert "tasks" not in w.pages and "dashboard" not in w.pages
    assert all(page.task_list.item(i).checkState().value == 0 for i in range(page.task_list.count()))


# ------------------------------------------------------------------ 2-6. files from anywhere, run, originals safe
def test_run_with_multiple_files_from_outside_the_project(window, qapp):
    w, page, tmp, s = window
    a, b = _outside(tmp)
    decoy = write_erp_csv(tmp / "input" / "untouched.csv", [row("01-11-2026", S, "1.00", 100009)])
    shas = {p: _sha(p) for p in (a, b, decoy)}
    page.task_list.setCurrentRow(2)                                # Customer
    page.add_files([a, b])
    assert page.file_list.rowCount() == 2 and page.file_list.item(0, 0).text() == "north.csv"
    assert page.file_list.item(0, 5).text() == str(a.parent)       # the folder is shown too
    assert page.file_list.item(0, 1).text() == "2" and page.file_list.item(0, 2).text() == "15/11/2026"
    job = page.run()
    assert job is not None and "--no-move" in job.arguments
    staged = Path(job.staging_dir)
    assert staged.parent == staging_root() and job.arguments[job.arguments.index("--input-dir") + 1] == str(staged)
    job = _wait_job(qapp, page)
    assert job.status is Status.COMPLETED, job.text
    assert {p: _sha(p) for p in (a, b, decoy)} == shas              # originals and the input folder untouched
    assert a.is_file() and b.is_file() and decoy.is_file()
    assert not staged.exists()                                     # staged copies removed after the run
    assert not any((tmp / "processed").iterdir()) and not any((tmp / "state").iterdir())
    assert job.reports and all(Path(r).parent.parent == tmp / "output" for r in job.reports)
    from openpyxl import load_workbook
    cells = {str(c) for r in job.reports for ws in load_workbook(r, read_only=True).worksheets
             for row_ in ws.iter_rows(values_only=True) for c in row_ if c is not None}
    assert {"100001", "100002"} <= cells                          # both files read (one customer each)
    downloads = [b_ for b_ in page.results.findChildren(QPushButton) if b_.text() == "تحميل التقرير"]
    assert len(downloads) == len(job.reports) and page.results.isVisibleTo(page)
    assert "بنجاح" in page.status.text()


def test_remove_and_replace_selected_files(window):
    w, page, tmp, s = window
    a, b = _outside(tmp)
    page.add_files([a, b, a])                                      # the same file once
    assert page.files == [a, b]
    page.file_list.setCurrentCell(0, 0)                            # what a mouse click on the row does
    page.remove_selected()
    assert page.files == [b]
    page.clear_files()
    assert page.files == [] and page.file_list.rowCount() == 0
    assert a.is_file() and b.is_file()                             # removing from the list never deletes


def test_file_picker_offers_the_right_types_from_any_folder(window, monkeypatch):
    w, page, tmp, s = window
    a, _ = _outside(tmp)
    seen = {}

    def fake(parent, title, start, flt):
        seen["filter"], seen["title"] = flt, title
        return [str(a)], flt

    monkeypatch.setattr(QFileDialog, "getOpenFileNames", staticmethod(fake))
    page.task_list.setCurrentRow(0)
    page.pick_files()
    assert seen["title"] == "اختيار الملف" and "*.csv" in seen["filter"] and "xlsx" not in seen["filter"]
    assert page.files == [a]
    page.task_list.setCurrentRow(6)
    page.pick_files()
    assert "*.xlsx" in seen["filter"]


def test_wrong_files_or_no_files_do_not_run(window, monkeypatch):
    w, page, tmp, s = window
    msgs = []
    monkeypatch.setattr(QMessageBox, "information", staticmethod(lambda *a, **k: msgs.append(a[2])))
    monkeypatch.setattr(QMessageBox, "warning", staticmethod(lambda *a, **k: msgs.append(a[2])))
    page.task_list.setCurrentRow(2)
    assert page.run() is None and msgs                             # nothing chosen
    book = tmp / "x" / "book.xlsx"
    book.parent.mkdir()
    book.write_bytes(b"x")
    page.add_files([book])
    assert page.run() is None and "book.xlsx" in msgs[-1]          # an Excel file for a CSV task
    assert not staging_root().exists() or not any(staging_root().iterdir())


def test_daily_task_with_its_date(window, qapp):
    w, page, tmp, s = window
    a, b = _outside(tmp)
    page.task_list.setCurrentRow(0)
    page.fields["ref_date"].setDate(QDate(2026, 11, 16))
    page.add_files([a])
    job = page.run()
    assert job.arguments[job.arguments.index("--ref-date") + 1] == "16/11/2026"
    job = _wait_job(qapp, page)
    assert job.exit_code in (0, 1)                                 # the engine's own verdict, shown as it is
    assert any(r.endswith("sales_report_ref-2026-11-16.xlsx") for r in job.reports)


# ------------------------------------------------------------------ 7-8. several reports, download (Save As)
def test_task5_option1_all_four_reports_downloadable(window, qapp, monkeypatch):
    w, page, tmp, s = window
    a, _ = _outside(tmp)
    write_fixed(tmp / "data" / FIXED_CODES_NAME, [100001, 555555])
    s.main_data = str(write_main(tmp / "data" / "main.xlsx"))
    write_collection(tmp / "التحصيل" / "collection.xlsx", [(100001, 10, 20)])
    page.task_list.setCurrentRow(4)
    page.fields["ref_date"].setDate(QDate(2026, 11, 16))
    page.add_files([a])
    page.run()
    job = _wait_job(qapp, page)
    assert job.status is Status.COMPLETED, job.text
    assert len(job.reports) == 4
    buttons = [b for b in page.results.findChildren(QPushButton) if b.text() == "تحميل التقرير"]
    assert len(buttons) == 4
    target = tmp / "my downloads"
    asked = []

    def save_as(parent, title, start, flt):
        asked.append((title, Path(start).name))
        return str(target / f"copy_{Path(start).name}"), flt

    monkeypatch.setattr(QFileDialog, "getSaveFileName", staticmethod(save_as))
    for b in buttons:
        b.click()                                                  # each report: its own Save As
    assert [t for t, _ in asked] == ["تحميل التقرير"] * 4
    for report in job.reports:
        r = Path(report)
        copy = target / f"copy_{r.name}"
        assert copy.is_file() and _sha(copy) == _sha(r)            # copied ...
        assert r.is_file() and r.parent.parent == tmp / "output"   # ... and still in output


def test_receivables_option_with_files_from_different_folders(window, qapp):
    from tests.test_receivables import _setup
    w, page, tmp, s = window
    main_path, sales, recv = _setup(tmp / "src")                   # sales and receivables in two folders
    s.main_data = str(main_path)
    shas = (_sha(sales), _sha(recv))
    page.task_list.setCurrentRow(6)
    page.add_files([sales, recv])
    page.run()
    job = _wait_job(qapp, page)
    assert job.status is Status.COMPLETED, job.text
    assert len(job.reports) == 1 and "Representatives_Receivables_Update" in job.reports[0]
    assert (_sha(sales), _sha(recv)) == shas


def test_save_copy_never_touches_the_report(tmp_path):
    report = tmp_path / "out" / "r.xlsx"
    report.parent.mkdir()
    report.write_bytes(b"report")
    dest = save_copy(report, tmp_path / "dl" / "r.xlsx")
    assert dest.read_bytes() == b"report" and report.read_bytes() == b"report"
    with pytest.raises(ValueError):
        save_copy(report, report)


def test_staging_only_ever_deletes_its_own_folders(gui_env):
    tmp, s = gui_env
    src = write_erp_csv(tmp / "u" / "a.csv", [row("01-11-2026", S, "1.00", 1)])
    twin = write_erp_csv(tmp / "v" / "a.csv", [row("02-11-2026", S, "2.00", 2)])
    run = stage_files([src, twin])
    assert [f.staged.name for f in run.files] == ["a.csv", "a (2).csv"] and len(run.renamed) == 1
    assert not remove_staging(tmp / "u") and src.is_file()        # outside the staging root: refused
    assert remove_staging(run.folder) and not run.folder.exists() and src.is_file() and twin.is_file()


# ------------------------------------------------------------------ «تحميل كل التقارير» (separate .xlsx files)
def _choose_folder(monkeypatch, folder: Path | None, asked: list):
    def fake(parent, title, start):
        asked.append(title)
        return str(folder) if folder else ""
    monkeypatch.setattr(QFileDialog, "getExistingDirectory", staticmethod(fake))


def _no_archives(folder: Path) -> None:
    assert not [p for p in folder.rglob("*") if p.suffix.lower() in (".zip", ".7z", ".rar", ".part")]


def test_download_all_one_report(window, qapp, monkeypatch):
    w, page, tmp, s = window
    a, b = _outside(tmp)
    page.task_list.setCurrentRow(2)                                # Customer: one report
    page.add_files([a, b])
    page.run()
    job = _wait_job(qapp, page)
    assert job.status is Status.COMPLETED and len(job.reports) == 1
    assert page.all_btn.isEnabled() and page.all_btn.text() == "تحميل كل التقارير"
    before = {r: _sha(r) for r in job.reports}
    asked = []
    _choose_folder(monkeypatch, tmp / "my folder", asked)
    page.all_btn.click()
    assert len(asked) == 1 and "اختر مجلد" in asked[0]
    files = sorted((tmp / "my folder").iterdir())
    assert [f.name for f in files] == [Path(r).name for r in job.reports]
    assert _sha(files[0]) == before[job.reports[0]]
    assert {r: _sha(r) for r in job.reports} == before             # the reports in output untouched
    assert "تم تحميل جميع التقارير بنجاح — 1 ملفات Excel" in page.status.text()
    _no_archives(tmp / "my folder")


def test_download_all_task5_option1_four_separate_workbooks(window, qapp, monkeypatch):
    from openpyxl import load_workbook
    w, page, tmp, s = window
    a, _ = _outside(tmp)
    (tmp / "output" / "Task5_2020-01-01").mkdir()                  # an older run: never included
    (tmp / "output" / "Task5_2020-01-01" / "old.xlsx").write_bytes(b"old")
    write_fixed(tmp / "data" / FIXED_CODES_NAME, [100001, 555555])
    s.main_data = str(write_main(tmp / "data" / "main.xlsx"))
    write_collection(tmp / "التحصيل" / "collection.xlsx", [(100001, 10, 20)])
    page.task_list.setCurrentRow(4)
    page.fields["ref_date"].setDate(QDate(2026, 11, 16))
    page.add_files([a])
    page.run()
    job = _wait_job(qapp, page)
    assert job.status is Status.COMPLETED, job.text
    assert len(job.reports) == 4
    before = {r: _sha(r) for r in job.reports}
    dest = tmp / "تقارير اليوم"
    asked = []
    _choose_folder(monkeypatch, dest, asked)
    assert page.download_all() == dest
    files = sorted(dest.iterdir())
    assert sorted(f.name for f in files) == sorted(Path(r).name for r in job.reports)   # 4 separate files
    assert len(files) == 4 and all(f.suffix == ".xlsx" for f in files)
    assert "old.xlsx" not in [f.name for f in files]
    for f in files:
        wb = load_workbook(f)                                      # each opens on its own
        assert wb.sheetnames
        wb.close()
    assert {f.name: _sha(f) for f in files} == {Path(r).name: h for r, h in before.items()}   # identical copies
    assert {r: _sha(r) for r in job.reports} == before             # originals unchanged
    assert "4 ملفات Excel" in page.status.text()
    _no_archives(dest)
    one = []                                                       # the individual buttons still work beside it

    def save_as(parent, title, start, flt):
        one.append(title)
        return str(tmp / "single" / "one.xlsx"), flt

    monkeypatch.setattr(QFileDialog, "getSaveFileName", staticmethod(save_as))
    [b for b in page.results.findChildren(QPushButton) if b.text() == "تحميل التقرير"][0].click()
    assert (tmp / "single" / "one.xlsx").is_file() and one == ["تحميل التقرير"]


def test_download_all_never_overwrites_existing_files(window, qapp, monkeypatch):
    w, page, tmp, s = window
    a, _ = _outside(tmp)
    page.task_list.setCurrentRow(3)                                # Representative
    page.add_files([a])
    page.run()
    job = _wait_job(qapp, page)
    name = Path(job.reports[0]).name
    dest = tmp / "dest"
    dest.mkdir()
    (dest / name).write_bytes(b"MY OWN FILE")                      # already in the folder
    _choose_folder(monkeypatch, dest, [])
    page.download_all()
    assert (dest / name).read_bytes() == b"MY OWN FILE"            # never overwritten
    stem, suffix = Path(name).stem, Path(name).suffix
    assert _sha(dest / f"{stem} (2){suffix}") == _sha(job.reports[0])
    assert "حُفظ باسم" in page.status.text()
    page.download_all()                                            # again: (3), still nothing overwritten
    assert (dest / f"{stem} (3){suffix}").is_file() and (dest / name).read_bytes() == b"MY OWN FILE"


def test_download_all_cancelled_does_nothing(window, qapp, monkeypatch):
    w, page, tmp, s = window
    a, _ = _outside(tmp)
    page.task_list.setCurrentRow(3)
    page.add_files([a])
    page.run()
    _wait_job(qapp, page)
    status_before = page.status.text()
    files_before = sorted(p for p in tmp.rglob("*") if p.is_file() and "gui_jobs" not in p.name)
    asked = []
    _choose_folder(monkeypatch, None, asked)                       # the user presses Cancel
    assert page.download_all() is None and len(asked) == 1
    assert page.status.text() == status_before                     # no false success message
    assert sorted(p for p in tmp.rglob("*") if p.is_file() and "gui_jobs" not in p.name) == files_before


def test_download_all_reports_failures_without_claiming_success(window, monkeypatch):
    from gui.services.jobs import Job
    w, page, tmp, s = window
    good = tmp / "output" / "Run_x" / "good.xlsx"
    good.parent.mkdir()
    good.write_bytes(b"ok")
    job = Job("3", ["-c", "1"], True)
    job.status, job.reports = Status.COMPLETED, [str(good), str(tmp / "output" / "Run_x" / "gone.xlsx")]
    page.job = job
    page.show_results(job)
    warnings = []
    monkeypatch.setattr(QMessageBox, "warning", staticmethod(lambda *a, **k: warnings.append(a[2])))
    _choose_folder(monkeypatch, tmp / "dest", [])
    page.download_all()
    assert (tmp / "dest" / "good.xlsx").read_bytes() == b"ok"
    assert warnings and "gone.xlsx" in warnings[0]                 # the failed file is named
    assert "بنجاح" not in page.status.text() and "1 من 2" in page.status.text()


def test_download_all_disabled_without_reports(window, monkeypatch):
    from gui.services.jobs import Job
    w, page, tmp, s = window
    job = Job("3", ["-c", "1"], True)                              # a finished run that wrote no report
    job.status, job.note = Status.FAILED, "لا توجد ملفات إدخال"
    page.job = job
    page.show_results(job)
    assert not page.all_btn.isEnabled() and page.results.isVisibleTo(page)
    msgs, dialogs = [], []
    monkeypatch.setattr(QMessageBox, "information", staticmethod(lambda *a, **k: msgs.append(a[2])))
    _choose_folder(monkeypatch, tmp / "never", dialogs)
    assert page.download_all() is None and msgs and not dialogs    # a clear message, no folder dialog
    assert not (tmp / "never").exists()


def test_copy_reports_to_folder_names_conflicts_and_failures(tmp_path):
    from gui.services.staging import copy_reports_to_folder
    a = tmp_path / "r1" / "report.xlsx"
    b = tmp_path / "r2" / "report.xlsx"                            # the same name twice in one run
    for p, data in ((a, b"A"), (b, b"B")):
        p.parent.mkdir()
        p.write_bytes(data)
    dest = tmp_path / "dest"
    res = copy_reports_to_folder([a, b, tmp_path / "missing.xlsx"], dest)
    assert [p.name for p in res.copied] == ["report.xlsx", "report (2).xlsx"]
    assert (dest / "report.xlsx").read_bytes() == b"A" and (dest / "report (2).xlsx").read_bytes() == b"B"
    assert res.renamed == [("report.xlsx", "report (2).xlsx")] and res.failed[0][0].endswith("missing.xlsx")
    assert a.read_bytes() == b"A" and b.read_bytes() == b"B"       # originals untouched
    assert sorted(p.name for p in dest.iterdir()) == ["report (2).xlsx", "report.xlsx"]   # no .part left over
