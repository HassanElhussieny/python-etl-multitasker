"""No task detects, skips or removes duplicates: repeated files and repeated rows are always counted."""

import pytest
from openpyxl import Workbook, load_workbook

from sales_etl.cli import EXIT_OK, main
from tests.conftest import R, S, Txn, write_erp_csv

# The first two rows are identical (same document number): both must be counted. Net per file: 190.
ROWS = [Txn("08-10-2026", S, "100.00", "1", "400100"), Txn("08-10-2026", S, "100.00", "1", "400100"),
        Txn("07-10-2026", R, "-10.00", "2", "400100")]


def _net_task1(path):    # Grand Total row, الصافي column
    return list(load_workbook(path, data_only=True)["التقرير"].iter_rows(values_only=True))[-1][2]


def _net_period(path):   # Grand Total row, last column
    return list(load_workbook(path, data_only=True)["التقرير"].iter_rows(values_only=True))[-1][-1]


def _net_grouped(path):  # no total row: sum of الصافي per code
    return sum(r[-1] for r in list(load_workbook(path, data_only=True)["التقرير"].iter_rows(values_only=True))[1:])


def _net_task5(path):    # الإجمالي العام of the sector sheet, الصافي column
    rows = list(load_workbook(path, data_only=True)["تقرير القطاعات"].iter_rows(values_only=True))
    return rows[-1][rows[1].index("الصافي")]                     # row 1: the report date, row 2: the header


TASKS = {
    "1 Daily": (["--ref-date", "08/10/2026"], "Report_*/sales_report_*.xlsx", "Run_", _net_task1),
    "2 Period": (["--start-date", "01-10-2026", "--end-date", "12-10-2026"], "Period_*/*.xlsx", "Period_",
                 _net_period),
    "3 Customer": (["--task", "3"], "Customer_*/*.xlsx", "Customer_", _net_grouped),
    "4 Representative": (["--task", "4"], "Representative_*/*.xlsx", "Representative_", _net_grouped),
    "5 Analysis": (["--task", "5", "--report-type", "2"], "Task5_*/تقرير تحليل المبيعات والمرتجعات.xlsx", "Task5_", _net_task5),
}


def _run(tmp_path, task):
    args, *_ = TASKS[task]
    if task.startswith("5"):
        main_data = tmp_path / "main.xlsx"
        if not main_data.exists():
            wb = Workbook()
            wb.active.append(["كود العميل", "اسم العميل", "قطاعات العملاء", "المندوب"])
            wb.active.append([400100, "عميل", "الجملة", "الاداره"])
            wb.save(main_data)
            fixed = Workbook()
            fixed.active.append([400100])
            fixed.save(tmp_path / "fixed code.xlsx")
        args = [*args, "--main-data", str(main_data)]
    return main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"),
                 "--non-interactive", *args])


def _nets(tmp_path, task):
    _, pattern, _, net = TASKS[task]
    return [net(p) for p in sorted((tmp_path / "out").glob(pattern), key=lambda p: p.parent.name)]


@pytest.mark.parametrize("task", TASKS)
def test_identical_files_with_different_names_are_both_counted(tmp_path, task):
    a = write_erp_csv(tmp_path / "in" / "a.csv", ROWS)
    (tmp_path / "in" / "copy_of_a.csv").write_bytes(a.read_bytes())
    assert _run(tmp_path, task) == EXIT_OK
    assert _nets(tmp_path, task) == [380]                         # 2 files x 190, repeated row included
    archived = sorted(p.name for p in (tmp_path / "processed").rglob("*.csv"))
    assert archived == ["a.csv", "copy_of_a.csv"] and list((tmp_path / "in").iterdir()) == []
    assert not (tmp_path / "duplicates").exists()
    assert not (tmp_path / "state" / "fingerprints").exists()


@pytest.mark.parametrize("task", TASKS)
def test_same_file_is_processed_again_after_a_successful_run(tmp_path, task):
    content = write_erp_csv(tmp_path / "in" / "a.csv", ROWS).read_bytes()
    assert _run(tmp_path, task) == EXIT_OK
    (tmp_path / "in" / "a.csv").write_bytes(content)              # supplied again, same name and content
    assert _run(tmp_path, task) == EXIT_OK
    assert _nets(tmp_path, task) == [190, 190]                    # not skipped the second time
    _, _, prefix, _ = TASKS[task]
    runs = sorted((tmp_path / "processed").glob(f"{prefix}*"))
    assert len(runs) == 2 and all([p.name for p in r.iterdir()] == ["a.csv"] for r in runs)
    assert list((tmp_path / "in").iterdir()) == []
