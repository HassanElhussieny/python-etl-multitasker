"""Task 5 option 1, report 4 (Customers_Not_Worked_<Month>_<year>.xlsx): the customers of report 3 whose sales of
the month are 0 (also with returns), one sheet per المسؤول (report 3's sheets), one الإجمالي العام row each.

Every input here is written into the test's own temporary folder."""

import hashlib
from datetime import datetime
from decimal import Decimal

from openpyxl import load_workbook

from sales_etl import analysis
from sales_etl import fixed_customers as fc
from sales_etl import not_worked as nw
from sales_etl.analysis import FIXED_CODES_NAME, GRAND_TOTAL, MODE_ALL, MODE_YESTERDAY, load_fixed_codes, \
    load_main_data, run_analysis
from sales_etl.cli import EXIT_OK, main
from sales_etl.fixed_customers import load_fixed_customers
from tests.conftest import R, S, write_collection, write_erp_csv
from tests.test_analysis import _config, row, write_main
from tests.test_collection_columns import COLL_ROWS, MAIN_ROWS
from tests.test_fixed_customers import N16, RESP_CSV, RESP_FIXED, _write_fixed_list

D = Decimal
NAME = "Customers_Not_Worked_November_2026.xlsx"
CSV = [*RESP_CSV,
       row("09-11-2026", R, "-0.55", 600040, doc="40"),          # 600040: returns only, -7.00 - 0.55
       row("02-11-2026", S, "0.00", 600020, doc="41")]           # a zero sale: still no sales
# 600005 (fixed, no November activity) has receivables: a zero-sales customer with dues (decimal, negative)
COLL = [*COLL_ROWS, (600005, 12.34, -5.5)]
# report 3 (16/11/2026): 600010 sold 100 / -10, 600005 nothing, 600020 0.00, 600060 sold 12, 600030 sold 30 / -5,
# 600040 returns -7.55 only, 600070 sold 4, 999111 sold 9  ->  report 4: 600005, 600020, 600040
SHOWN = {600005, 600020, 600040}


def _run(tmp_path, mode=MODE_YESTERDAY, csv=CSV, coll=COLL, collection=True):
    if collection:
        write_collection(tmp_path / "التحصيل" / "متابعة.xlsx", coll)
    write_erp_csv(tmp_path / "input" / "c.csv", csv)
    fixed_path = _write_fixed_list(tmp_path / "data" / FIXED_CODES_NAME, rows=RESP_FIXED)
    return run_analysis(_config(tmp_path), mode, N16 if mode == MODE_YESTERDAY else None,
                        load_main_data(write_main(tmp_path / "data" / "main.xlsx", MAIN_ROWS)),
                        fixed_codes=load_fixed_codes(fixed_path), fixed_customers=load_fixed_customers(fixed_path),
                        now=datetime(2026, 11, 17, 9, 0), collection_dir=tmp_path / "التحصيل")


def _saved(path):
    wb = load_workbook(path)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


def _rows(sheets):
    return {name: (rows[2:-1], rows[-1]) for name, rows in sheets.items()}


# ------------------------------------------------------------------ the workbook
def test_four_reports_beside_each_other(tmp_path):
    result = _run(tmp_path)
    assert analysis.verify_task5(result) == []
    assert result.not_worked_path.name == NAME == nw.workbook_name(result)
    assert sorted(p.name for p in result.run_dir.iterdir()) == sorted([
        analysis.WORKBOOK_NAME, result.reps_report_path.name, result.customers_report_path.name, NAME])


def test_only_customers_with_zero_sales_are_shown(tmp_path):
    result = _run(tmp_path)
    shown = [r for rows, _ in _rows(_saved(result.not_worked_path)).values() for r in rows]
    assert {r[0] for r in shown} == SHOWN                                        # sold customers never shown
    assert all(r[4] == 0 for r in shown)
    by = {r[0]: r for r in shown}
    assert by[600005][4:7] == [0, 0, 0]                                          # no sales, no returns
    assert by[600040][4:7] == [0, -7.55, -7.55]                                  # no sales, returns: shown
    assert by[600020][4:7] == [0, 0, 0]                                          # a 0.00 sale is no sale
    report3 = {r[0]: r for r in _saved(result.customers_report_path)[fc.SHEET_CUSTOMERS][2:-4]}
    for code, r in by.items():
        assert r == report3[code]                                                # the same 11 values as report 3
    assert by[600005][9:11] == [12.34, -5.5]                                     # zero sales with dues
    assert by[600020][9:11] == [0, None] and by[600040][9:11] == [None, None]   # empty stays empty
    assert by[600005][7:9] == [fc.CUSTOMER_MARK, None] and by[600040][7:9] == [None, fc.CUSTOMER_MARK]


def test_sheets_are_report3_people_and_customers_on_their_own_sheet(tmp_path):
    result = _run(tmp_path)
    sheets = _saved(result.not_worked_path)
    people3 = [n for n in _saved(result.customers_report_path) if n != fc.SHEET_CUSTOMERS]
    assert list(sheets) == people3                                               # same sheets, same order
    rows = _rows(sheets)
    assert [r[0] for r in rows["مندوب تجريبى 10"][0]] == [600005]
    assert [r[0] for r in rows["مندوب تجريبي 02"][0]] == [600020]           # 600010 sold: not here
    assert [r[0] for r in rows["غير محدد"][0]] == [600040]
    for name in ("اغلاق", "مندوب تجريبي 11", "غير مسجل"):                   # no customer without sales
        assert rows[name][0] == [] and len(sheets[name]) == 3
        assert rows[name][1][:9] == [None, GRAND_TOTAL, 0, None, 0, 0, 0, 0, 0]
    codes = [r[0] for detail, _ in rows.values() for r in detail]
    assert len(codes) == len(set(codes)) == 3                                    # every code once
    for name, (detail, _) in rows.items():
        assert all(analysis.canonical(r[3] or "غير محدد") == analysis.canonical(name) for r in detail), name


def test_one_total_row_per_sheet_recounted(tmp_path):
    result = _run(tmp_path)
    sheets = _saved(result.not_worked_path)
    for name, rows in sheets.items():
        detail, total = rows[2:-1], rows[-1]
        assert sum(1 for r in rows if GRAND_TOTAL in r[:2]) == 1, name           # exactly one, the last row
        num = [D(str(r[2])) for r in detail if isinstance(r[2], (int, float))]
        want = [None, GRAND_TOTAL, sum(num, D(0)), None,
                *[sum((D(str(r[i])) for r in detail), D(0)) for i in (4, 5, 6)],
                len({r[0] for r in detail if r[7] == "نعم"}), len({r[0] for r in detail if r[8] == "نعم"}),
                *[sum((D(str(r[i])) for r in detail if r[i] is not None), D(0)) for i in (9, 10)]]
        assert [D(str(v)) if isinstance(v, float) or (isinstance(v, int) and i not in (7, 8)) else v
                for i, v in enumerate(total)] == want, name
    assert sheets["مندوب تجريبى 10"][-1] == [None, GRAND_TOTAL, 0, None, 0, 0, 0, 1, 0, 12.34, -5.5]
    assert sheets["مندوب تجريبي 02"][-1] == [None, GRAND_TOTAL, 5, None, 0, 0, 0, 1, 0, 0, 0]
    assert sheets["غير محدد"][-1] == [None, GRAND_TOTAL, 0, None, 0, -7.55, -7.55, 0, 1, 0, 0]
    ws = load_workbook(result.not_worked_path)["مندوب تجريبي 02"]
    assert ws.cell(row=4, column=3).number_format == analysis.MONEY_FORMAT
    assert ws.cell(row=4, column=8).number_format == analysis.COUNT_FORMAT
    assert ws.print_area.endswith("$K$4") and ws.auto_filter.ref == "A2:K3" and ws.sheet_view.rightToLeft


def test_reports_1_to_3_unchanged_by_report_4(tmp_path, monkeypatch):
    with_4 = _run(tmp_path / "a")
    monkeypatch.setattr(nw, "build_not_worked_report", lambda result, run_dir: None)
    without = _run(tmp_path / "b")
    assert without.not_worked_path is None
    for a, b in ((with_4.report_path, without.report_path), (with_4.reps_report_path, without.reps_report_path),
                 (with_4.customers_report_path, without.customers_report_path)):
        assert _saved(a) == _saved(b), a.name


def test_option2_has_no_report_4(tmp_path):
    result = _run(tmp_path, mode=MODE_ALL)
    assert analysis.verify_task5(result) == [] and result.not_worked_path is None
    assert not list(result.run_dir.glob("Customers_Not_Worked_*"))


def test_collection_problem_means_no_report_3_and_no_report_4(tmp_path):
    result = _run(tmp_path, collection=False)                                    # no التحصيل folder
    assert result.customers_report_path is None and result.not_worked_path is None
    assert analysis.verify_task5(result) == [result.customers_error]


def test_wrong_saved_report_4_is_detected(tmp_path):
    result = _run(tmp_path)
    wb = load_workbook(result.not_worked_path)
    ws = wb["غير محدد"]
    ws.cell(row=3, column=5).value = 3.0                                         # sales 0 -> 3
    ws.cell(row=4, column=6).value = -1.0                                        # the total's returns
    wb.save(result.not_worked_path)
    problems = nw.verify_not_worked_report(result)
    assert any("has sales 3" in p for p in problems)
    assert any("saved الإجمالي العام" in p for p in problems)


def test_wrong_selection_is_detected_by_the_recount(tmp_path, monkeypatch):
    result = _run(tmp_path)
    original = nw.sheets_table
    monkeypatch.setattr(nw, "sheets_table", lambda r: [(label, mine[:-1] if mine else mine)
                                                       for label, mine in original(r)])   # a customer dropped
    problems = nw.reconcile_not_worked(result, {600005, 600020})
    assert any("are missing" in p for p in problems)


def test_command_option1_writes_four_reports(tmp_path, capsys):
    path = write_collection(tmp_path / "التحصيل" / "متابعة.xlsx", COLL)
    sha = hashlib.sha256(path.read_bytes()).hexdigest()
    write_erp_csv(tmp_path / "input" / "c.csv", CSV)
    _write_fixed_list(tmp_path / "data" / FIXED_CODES_NAME, rows=RESP_FIXED)
    write_main(tmp_path / "data" / "main.xlsx", MAIN_ROWS)
    assert main(["--input-dir", str(tmp_path / "input"), "--output-dir", str(tmp_path / "out"), "--task", "5",
                 "--ref-date", "16/11/2026", "--main-data", str(tmp_path / "data" / "main.xlsx"),
                 "--non-interactive"]) == EXIT_OK
    out = capsys.readouterr().out
    assert "Not-worked report" in out and "3 customers with sales = 0 (1 with returns, 2 without)" in out
    (run_dir,) = (tmp_path / "out").iterdir()
    assert len(list(run_dir.glob("*.xlsx"))) == 4 and (run_dir / NAME).is_file()
    assert list((tmp_path / "input").iterdir()) == [] and hashlib.sha256(path.read_bytes()).hexdigest() == sha
