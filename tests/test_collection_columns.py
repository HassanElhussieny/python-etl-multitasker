"""Task 5 option 1, report 3 (Fixed_Customers_Report): المستحق 7-10 / المستحق الجديد from the current workbook of
the التحصيل folder - matching by customer code, columns F and L, the main sheet and every المسؤول sheet, the totals,
the errors that stop the step, re-reading on every run, and option 2 / reports 1-2 left as they were.

Every collection workbook here is written into the test's own temporary folder (never the project's التحصيل)."""

import hashlib
from datetime import datetime
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook

from sales_etl import analysis
from sales_etl import fixed_customers as fc
from sales_etl.analysis import FIXED_CODES_NAME, GRAND_TOTAL, MODE_ALL, MODE_YESTERDAY, load_fixed_codes, \
    load_main_data, run_analysis
from sales_etl.cli import EXIT_AUDIT_FAILED, EXIT_OK, EXIT_USAGE, main
from sales_etl.fixed_customers import load_fixed_customers
from tests.conftest import COLLECTION_HEADER, write_collection, write_erp_csv
from tests.test_analysis import _config, write_main
from tests.test_fixed_customers import CUST_MAIN_ROWS, N16, RESP_CSV, RESP_FIXED, _write_fixed_list

D = Decimal
COLL_ROWS = [  # code (column B), المستحق 7-10 (F), المستحق الجديد (L)
    (600010, 1000.5, 200),
    (600020, 0, None),            # in the file, L empty: stays empty (never an invented 0)
    (600030, -50, 75.25),         # a customer outside the 224 list
    ("0600070", 3, 4),            # a code written with a leading zero: the same customer 600070
    (999111, 5, 6),               # not in Main Data (غير مسجل)
    (888888, 7, 8),               # in the file only: no report row is added for it
]
MAIN_ROWS = [*CUST_MAIN_ROWS, (600060, "x", "الجملة", "اغلاق"), (600070, "y", "الجملة", "مندوب تجريبي 10")]
# report rows (fixed list first, then the others by code) -> (المستحق 7-10, المستحق الجديد) as saved
WANT = {600010: [1000.5, 200], 600005: [None, None], 600020: [0, None], 600060: [None, None],
        600030: [-50, 75.25], 600040: [None, None], 600070: [3, 4], 999111: [5, 6]}


def _coll_run(tmp_path, rows=COLL_ROWS, mode=MODE_YESTERDAY, folder="التحصيل", name="متابعة_التحصيل_15-11.xlsx",
              write=True):
    if write:
        write_collection(tmp_path / "التحصيل" / name, rows)
    write_erp_csv(tmp_path / "input" / "c.csv", RESP_CSV)
    fixed_path = _write_fixed_list(tmp_path / "data" / FIXED_CODES_NAME, rows=RESP_FIXED)
    return run_analysis(_config(tmp_path), mode, N16 if mode == MODE_YESTERDAY else None,
                        load_main_data(write_main(tmp_path / "data" / "main.xlsx", MAIN_ROWS)),
                        fixed_codes=load_fixed_codes(fixed_path), fixed_customers=load_fixed_customers(fixed_path),
                        now=datetime(2026, 11, 16, 13, 0), collection_dir=tmp_path / folder)


def _saved(path):
    wb = load_workbook(path)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


# ------------------------------------------------------------------ the columns
def test_collection_columns_on_the_main_sheet(tmp_path):
    result = _coll_run(tmp_path)
    assert analysis.verify_task5(result) == []
    sheets = _saved(result.customers_report_path)
    main_sheet = sheets[fc.SHEET_CUSTOMERS]
    header = main_sheet[1]
    assert header[:9] == list(fc.customers_columns(result))[:9]                  # the 9 columns, unchanged
    assert header[9:] == ["المستحق 7-10", "المستحق الجديد"] and len(header) == len(set(header)) == 11
    body = main_sheet[2:-3]
    *detail, total = body
    assert {r[0]: r[9:11] for r in detail} == WANT                               # by customer code (F, L)
    assert [r[0] for r in detail] == [600010, 600005, 600020, 600060, 600030, 600040, 600070, 999111]
    assert 888888 not in [r[0] for r in detail]                                  # never a row from the file
    assert total[0] == GRAND_TOTAL and total[9:11] == [958.5, 285.25]
    assert main_sheet[-2][9:11] == ["إجمالي المستحق 7-10", "إجمالي المستحق الجديد"]
    assert main_sheet[-1][9:11] == [958.5, 285.25]
    ws = load_workbook(result.customers_report_path)[fc.SHEET_CUSTOMERS]
    assert ws.cell(row=3, column=10).number_format == analysis.MONEY_FORMAT
    assert ws.print_area.endswith(f"$K${len(main_sheet)}") and ws.auto_filter.ref == f"A2:K{len(detail) + 2}"


def test_collection_columns_on_every_responsible_sheet(tmp_path):
    result = _coll_run(tmp_path)
    sheets = _saved(result.customers_report_path)
    main_rows = {r[0]: r for r in sheets[fc.SHEET_CUSTOMERS][2:-4]}
    people = [n for n in sheets if n != fc.SHEET_CUSTOMERS]
    assert len(people) == len(set(people)) and len(sheets) == 1 + len(people)    # no duplicate sheets
    seen = []
    for name in people:
        sheet = sheets[name]
        assert sheet[1] == sheets[fc.SHEET_CUSTOMERS][1]                          # the same 11 columns
        rows = sheet[2:-4]
        for r in rows:
            assert r == main_rows[r[0]]                                           # same values as the main sheet
            seen.append(r[0])
        block = sheet[-1]
        assert block[9:11] == [sum((r[i] or 0) for r in rows) for i in (9, 10)]
    assert sorted(seen) == sorted(main_rows)                                     # every customer exactly once
    assert sheets["مندوب تجريبي 02"][-1][9:11] == [1000.5, 200]
    assert sheets["مندوب تجريبى 10"][-1][9:11] == [3, 4]
    assert sheets["اغلاق"][-1][9:11] == [0, 0]                                    # its customer not in the file


def test_sales_and_returns_unchanged_by_the_collection(tmp_path):
    plain = _coll_run(tmp_path / "a", write=False, folder="none")                # without: the step stops
    assert plain.customers_report_path is None
    plain = run_analysis(_config(tmp_path / "a"), MODE_YESTERDAY, N16, plain.main_data,
                         fixed_codes=plain.fixed_codes, fixed_customers=plain.fixed_customers,
                         now=datetime(2026, 11, 16, 14, 0))                     # no collection folder given
    full = _coll_run(tmp_path / "b")
    assert len(fc.customers_columns(plain)) == 9 and len(fc.customers_columns(full)) == 11
    assert [r[:9] for r in fc.customers_table(full)] == fc.customers_table(plain)
    a, b = _saved(plain.reps_report_path), _saved(full.reps_report_path)          # report 2 unchanged
    assert a.keys() == b.keys() and all(a[n][1:] == b[n][1:] for n in a)
    assert _saved(plain.report_path) == _saved(full.report_path)                   # report 1 unchanged


def test_collection_stats(tmp_path):
    result = _coll_run(tmp_path)
    assert fc.collection_stats(result) == {"customers": 8, "matched": 5, "unmatched": 3, "file_customers": 6,
                                           "file_only": 1, "duplicates": 0, "rows_without_code": 2}
    assert result.collection.file.sheet == "القديم والجديد" and result.collection.file.header_row == 3


def test_option2_never_has_the_collection_columns(tmp_path):
    result = _coll_run(tmp_path, mode=MODE_ALL)
    assert analysis.verify_task5(result) == [] and result.collection is None
    header = _saved(result.customers_report_path)[fc.SHEET_CUSTOMERS][1]
    assert len(header) == 9 and "المستحق 7-10" not in header


# ------------------------------------------------------------------ the folder and the workbook
def _lock(tmp_path):
    (tmp_path / "التحصيل" / "~$متابعة.xlsx").write_bytes(b"lock")              # Excel's lock file: ignored


@pytest.mark.parametrize("prepare, message", [
    (lambda t: None, "does not exist"),
    (lambda t: (t / "التحصيل").mkdir(), "has no Excel workbook"),
    (lambda t: (write_collection(t / "التحصيل" / "a.xlsx", COLL_ROWS), write_collection(t / "التحصيل" / "b.xlsx",
                                                                                         COLL_ROWS)), "2 workbooks"),
    (lambda t: ((t / "التحصيل").mkdir(), (t / "التحصيل" / "x.xlsx").write_text("not excel")), "cannot open"),
    (lambda t: write_collection(t / "التحصيل" / "a.xlsx", COLL_ROWS, header=COLLECTION_HEADER[:11]),
     "no sheet has the columns"),                                                # المستحق الجديد missing
    (lambda t: write_collection(t / "التحصيل" / "a.xlsx", COLL_ROWS,
                                header=[*COLLECTION_HEADER[:11], "x", "المستحق الجديد"]), "expected column L"),
    (lambda t: write_collection(t / "التحصيل" / "a.xlsx", [*COLL_ROWS, (600010, 1, 1)]), "600010 (rows 4, 10)"),
])
def test_collection_problem_stops_report_3_only(tmp_path, prepare, message):
    prepare(tmp_path)
    result = _coll_run(tmp_path, write=False)
    assert result.customers_report_path is None and result.collection is None
    assert message in result.customers_error and "reports 1-2 are not affected" in result.customers_error
    assert result.report_path.is_file() and result.reps_report_path.is_file()  # reports 1-2 written
    problems = analysis.verify_task5(result)
    assert result.report_problems == [] and problems == result.customers_problems == [result.customers_error]
    assert not list(result.run_dir.glob("Fixed_Customers_Report_*"))


def test_lock_file_is_not_a_second_workbook(tmp_path):
    write_collection(tmp_path / "التحصيل" / "a.xlsx", COLL_ROWS)
    _lock(tmp_path)
    result = _coll_run(tmp_path, write=False)
    assert analysis.verify_task5(result) == [] and result.collection.path.name == "a.xlsx"


def test_collection_source_never_written_and_read_every_run(tmp_path):
    path = write_collection(tmp_path / "التحصيل" / "first_14-11.xlsx", COLL_ROWS)
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    first = _coll_run(tmp_path, write=False)
    assert analysis.verify_task5(first) == [] and hashlib.sha256(path.read_bytes()).hexdigest() == before
    again = run_analysis(_config(tmp_path), MODE_YESTERDAY, N16, first.main_data, fixed_codes=first.fixed_codes,
                         fixed_customers=first.fixed_customers, now=datetime(2026, 11, 16, 13, 5),
                         collection_dir=tmp_path / "التحصيل")                   # repeated run: the same values
    assert analysis.verify_task5(again) == []
    assert fc.customers_table(again) == fc.customers_table(first)
    # the workbook replaced by a newer one (another name, other values): the next run reads it, no cache
    path.unlink()
    write_collection(tmp_path / "التحصيل" / "second_15-11.xlsx", [(600010, 1, 2), (600070, 9, 10)])
    third = run_analysis(_config(tmp_path), MODE_YESTERDAY, N16, first.main_data, fixed_codes=first.fixed_codes,
                         fixed_customers=first.fixed_customers, now=datetime(2026, 11, 16, 13, 10),
                         collection_dir=tmp_path / "التحصيل")
    assert analysis.verify_task5(third) == [] and third.collection.path.name == "second_15-11.xlsx"
    rows = {r[0]: r[9:11] for r in _saved(third.customers_report_path)[fc.SHEET_CUSTOMERS][2:-4]}
    assert rows[600010] == [1, 2] and rows[600070] == [9, 10] and rows[600030] == [None, None]


# ------------------------------------------------------------------ the checks catch wrong values
def test_wrong_saved_collection_value_is_detected(tmp_path):
    result = _coll_run(tmp_path)
    wb = load_workbook(result.customers_report_path)
    wb[fc.SHEET_CUSTOMERS].cell(row=3, column=10).value = 1.0                     # 600010's المستحق 7-10
    wb.save(result.customers_report_path)
    problems = fc.verify_customers_report(result)
    assert any("row 3 المستحق 7-10 1" in p for p in problems)


def test_collection_value_differing_from_the_source_is_detected(tmp_path, monkeypatch):
    result = _coll_run(tmp_path)
    original = fc.customers_table

    def wrong(r):
        rows = original(r)
        rows[0][10] = D("999")                                                    # 600010's المستحق الجديد
        return rows

    monkeypatch.setattr(fc, "customers_table", wrong)
    assert any("customer 600010 المستحق 7-10 / المستحق الجديد" in p for p in fc.reconcile_customers(result))


def test_collection_changed_during_the_run_is_detected(tmp_path):
    result = _coll_run(tmp_path)
    write_collection(result.collection.path, [(600010, 1, 2)])
    assert any("changed during the run" in p for p in fc.reconcile_customers(result))


# ------------------------------------------------------------------ through the command (option 1)
def _cli(tmp_path, *extra, report="1"):
    write_erp_csv(tmp_path / "input" / "c.csv", RESP_CSV)
    _write_fixed_list(tmp_path / "data" / FIXED_CODES_NAME, rows=RESP_FIXED)
    write_main(tmp_path / "data" / "main.xlsx", MAIN_ROWS)
    date_args = ["--ref-date", "16/11/2026"] if report == "1" else []
    return main(["--input-dir", str(tmp_path / "input"), "--output-dir", str(tmp_path / "out"), "--task", "5",
                 "--report-type", report, *date_args, "--main-data", str(tmp_path / "data" / "main.xlsx"),
                 "--non-interactive", *extra])


def test_command_option1_reads_the_collection_folder_and_moves_files(tmp_path, capsys):
    path = write_collection(tmp_path / "التحصيل" / "متابعة.xlsx", COLL_ROWS)    # the default (redirected) folder
    sha = hashlib.sha256(path.read_bytes()).hexdigest()
    assert _cli(tmp_path) == EXIT_OK
    out = capsys.readouterr().out
    assert "Collection workbook" in out and "5 of 8 report customers matched by code, 3 not in the file" in out
    (report,) = (tmp_path / "out").glob("Task5_*/Fixed_Customers_Report_*.xlsx")
    assert _saved(report)[fc.SHEET_CUSTOMERS][1][9:] == ["المستحق 7-10", "المستحق الجديد"]
    assert list((tmp_path / "input").iterdir()) == []                            # moved after verification
    assert [p.name for p in (tmp_path / "التحصيل").iterdir()] == ["متابعة.xlsx"]  # never moved or written
    assert hashlib.sha256(path.read_bytes()).hexdigest() == sha


def test_command_collection_problem_moves_nothing(tmp_path, capsys):
    other = tmp_path / "elsewhere"
    assert _cli(tmp_path, "--collection-dir", str(other)) == EXIT_AUDIT_FAILED
    out = capsys.readouterr().out
    assert "Customer report check    : FAIL" in out and "collection folder does not exist" in out
    assert [p.name for p in (tmp_path / "input").iterdir()] == ["c.csv"]         # nothing moved


def test_command_collection_dir_belongs_to_option1(tmp_path):
    assert _cli(tmp_path, "--collection-dir", str(tmp_path), report="2") == EXIT_USAGE
    assert _cli(tmp_path, "--no-move", report="2") == EXIT_OK                    # option 2: no folder needed


# ------------------------------------------------------------------ the total row of every المسؤول sheet
def _independent_total(rows):
    """The eight values recounted with plain loops from the saved customer rows."""
    num = [v for r in rows for v in [r[2]] if isinstance(v, (int, float))]
    return [None, GRAND_TOTAL, sum(D(str(v)) for v in num), None,
            sum(D(str(r[4])) for r in rows), sum(D(str(r[5])) for r in rows), sum(D(str(r[6])) for r in rows),
            len({r[0] for r in rows if r[7] == "نعم"}), len({r[0] for r in rows if r[8] == "نعم"}),
            sum(D(str(r[9])) for r in rows if r[9] is not None), sum(D(str(r[10])) for r in rows if r[10] is not None)]


def _as_decimal(row):
    return [D(str(v)) if isinstance(v, float) or (isinstance(v, int) and i not in (7, 8)) else v
            for i, v in enumerate(row)]


def test_every_person_sheet_has_one_total_row_below_its_customers(tmp_path):
    result = _coll_run(tmp_path)
    assert analysis.verify_task5(result) == []
    sheets = _saved(result.customers_report_path)
    main_sheet = sheets[fc.SHEET_CUSTOMERS]
    people = [n for n in sheets if n != fc.SHEET_CUSTOMERS]
    for name in people:
        sheet = sheets[name]
        rows, total = sheet[2:-4], sheet[-4]
        assert sum(1 for r in sheet if GRAND_TOTAL in (r[0], r[1])) == 2, name   # this row + the block below
        assert sheet[-3] == [None] * 11 and sheet[-1][0] == GRAND_TOTAL           # the block stays as it was
        assert all(r[1] != GRAND_TOTAL for r in rows)
        assert _as_decimal(total) == _independent_total(rows), name
    dalia = sheets["مندوب تجريبي 02"][-4]
    # 600010 (C 10, sold 100 / -10) + 600020 (C 5): both inside, F 1000.5 + 0, L 200 + empty
    assert dalia == [None, GRAND_TOTAL, 15, None, 100, -10, 90, 2, 0, 1000.5, 200]
    shimi = sheets["مندوب تجريبى 10"][-4]                                           # 600005 inside + 600070 outside
    assert shimi[2] == 0 and shimi[7:9] == [1, 1] and shimi[9:11] == [3, 4]
    assert sheets["غير مسجل"][-4][2] == 0 and sheets["غير مسجل"][-4][7:9] == [0, 1]
    # the main sheet: its own total row exactly as before (no الشهر الأعلى قيمة sum, no counts)
    assert main_sheet[-4][:9] == [GRAND_TOTAL, None, None, None, 155, -22, 133, None, None]
    # the person totals add up to the main sheet's
    totals = [sheets[n][-4] for n in people]
    for i in (4, 5, 6, 9, 10):
        assert sum(D(str(t[i])) for t in totals) == D(str(main_sheet[-4][i])), i
    ws = load_workbook(result.customers_report_path)["مندوب تجريبي 02"]
    line = len(sheets["مندوب تجريبي 02"]) - 3
    assert ws.cell(row=line, column=5).number_format == analysis.MONEY_FORMAT
    assert ws.cell(row=line, column=8).number_format == analysis.COUNT_FORMAT
    assert ws.cell(row=line, column=1).fill.fgColor.rgb == ws.cell(row=line, column=5).fill.fgColor.rgb


def test_person_total_row_in_option2_without_collection_columns(tmp_path):
    result = _coll_run(tmp_path, mode=MODE_ALL)
    assert analysis.verify_task5(result) == []
    sheets = _saved(result.customers_report_path)
    for name in (n for n in sheets if n != fc.SHEET_CUSTOMERS):
        total = sheets[name][-4]
        assert len(total) == 9 and total[1] == GRAND_TOTAL and total[0] is None and total[3] is None, name


def test_wrong_person_total_row_is_detected(tmp_path):
    result = _coll_run(tmp_path)
    wb = load_workbook(result.customers_report_path)
    ws = wb["مندوب تجريبي 02"]
    line = ws.max_row - 3
    ws.cell(row=line, column=8).value = 5                                         # عملاء داخل: 2 -> 5
    ws.cell(row=line, column=3).value = 99                                        # الشهر الأعلى قيمة: 15 -> 99
    wb.save(result.customers_report_path)
    problems = fc.verify_customers_report(result)
    assert any("مندوب تجريبي 02" in p and "saved total row" in p for p in problems)
