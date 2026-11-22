"""Task 5, report 3 (Fixed_Customers_Report): its own tests - the module ``sales_etl.fixed_customers``, its build
step, its separate check and reconciliation, the main sheet and the sheets per المسؤول."""

from datetime import date, datetime
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from sales_etl import analysis
from sales_etl import fixed_customers as fc
from sales_etl.analysis import FIXED_CODES_NAME, GRAND_TOTAL, MODE_ALL, MODE_YESTERDAY, MainDataError, load_fixed_codes, \
    load_main_data, run_analysis
from sales_etl.cli import EXIT_AUDIT_FAILED, EXIT_OK
from sales_etl.fixed_customers import CUSTOMER_MARK, load_fixed_customers
from tests.conftest import R, S, write_erp_csv
from tests.test_analysis import RUN_AT, _config, _content_width, _raw, _run, _sha, row, setup, write_main  # noqa: F401

D = Decimal


# ------------------------------------------------------------------ third workbook: Fixed_Customers_Report

IN_COLS, OUT_COLS = [CUSTOMER_MARK, None], [None, CUSTOMER_MARK]   # عملاء داخل / خارج قائمة الـ224, as saved

N16 = date(2026, 11, 16)
FIXED_ROWS = [  # file order on purpose (not sorted); D may differ from Main Data and is kept as written
    (600010, "  عميل ثابت أ", 500.25, "مندوب تجريبي 02"),
    (600005, "عميل ثابت ب", 0, "اغلاق"),
    (600020, "عميل ثابت ج", 1234, "مندوب تجريبي 08"),
]
CUST_MAIN_ROWS = [(600010, "م أ", "الجملة", "مندوب تجريبي 01"), (600005, "م ب", "الجملة", "مندوب تجريبي 01"),
                  (600020, "م ج", "مستشفيات", "مندوب تجريبي 08"), (600030, "م د", "الجملة", "مندوب تجريبي 03"),
                  (600040, "م هـ", "الجملة", None), (600050, "م و", "الجملة", "مندوب تجريبي 01")]
CUST_CSV = [
    row("16-11-2026", S, "60.00", 600010, "اسم ملف أ", doc="1"), row("16-11-2026", S, "40.00", 600010, doc="1"),
    row("05-11-2026", R, "-10.00", 600010, doc="2"),
    row("31-10-2026", S, "999.00", 600010, doc="3"),            # October: never in November
    row("17-11-2026", S, "50.00", 600010, doc="4"),             # after the report date
    row("20-10-2026", S, "70.00", 600005, doc="5"),             # 600005: nothing in November -> zeros
    row("01-11-2026", S, "30.00", 600030, "اسم ملف د", doc="6"), row("15-11-2026", R, "-5.00", 600030, doc="7"),
    row("10-11-2026", R, "-7.00", 600040, "اسم ملف خامس", doc="8"),   # returns only, blank Main Data representative
    row("12-11-2026", S, "9.00", 999111, "زائر", doc="9"),          # not in Main Data
    row("12-10-2026", S, "11.00", 600050, doc="10"),               # outside the list, October only: not shown
]


def _write_fixed_list(path, rows=FIXED_ROWS, header=("كود العميل", "اسم العميل", "الشهر الاعلي قيمه ", "المسئول")):
    wb = Workbook()
    wb.active.append(list(header))
    for r in rows:
        wb.active.append(list(r))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _cust_report(tmp_path, ref=N16, mode=MODE_YESTERDAY, rows=CUST_CSV):
    write_erp_csv(tmp_path / "input" / "c.csv", rows)
    fixed_path = _write_fixed_list(tmp_path / "data" / FIXED_CODES_NAME)
    result = run_analysis(_config(tmp_path), mode, ref if mode == MODE_YESTERDAY else None,
                          load_main_data(write_main(tmp_path / "data" / "main.xlsx", CUST_MAIN_ROWS)),
                          fixed_codes=load_fixed_codes(fixed_path), fixed_customers=load_fixed_customers(fixed_path),
                          now=datetime(2026, 11, 16, 13, 0))
    assert analysis.verify_task5(result) == []                          # incl. the third workbook
    ws = load_workbook(result.customers_report_path)[fc.SHEET_CUSTOMERS]
    saved = [list(r) for r in ws.iter_rows(values_only=True)]
    assert saved[-3] == [None] * 9 and saved[-1][0] == GRAND_TOTAL                # the الإجمالي العام block below
    return result, saved[0][0], saved[1], saved[2:-3]                            # customers + total row


def test_customers_report_is_written_beside_the_other_two(tmp_path):
    result, title, header, rows = _cust_report(tmp_path)
    assert result.customers_report_path.parent == result.report_path.parent == result.reps_report_path.parent
    assert result.customers_report_path.name == "Fixed_Customers_Report_2026-11-16.xlsx"      # the run date
    assert header == ["كود العميل", "اسم العميل", "الشهر الأعلى قيمة", "اسم المسؤول عن العميل",
                      "مبيعات شهر نوفمبر 2026", "مرتجعات شهر نوفمبر 2026", "صافي شهر نوفمبر 2026",
                      "عملاء داخل قائمة الـ224", "عملاء خارج قائمة الـ224"]
    assert title == "تقرير عملاء قائمة الـ224  |  الفترة من 01/11/2026 إلى 16/11/2026"


def test_customers_report_rows_amounts_and_classification(tmp_path):
    _, _, _, rows = _cust_report(tmp_path)
    body, total = rows[:-1], rows[-1]
    assert [r[0] for r in body[:3]] == [600010, 600005, 600020]                  # 1, 11: all, in the file's order
    assert body[0] == [600010, "  عميل ثابت أ", 500.25, "مندوب تجريبي 02", 100, -10, 90, *IN_COLS]
    assert body[1] == [600005, "عميل ثابت ب", 0, "اغلاق", 0, 0, 0, *IN_COLS]        # 10: no activity
    assert body[2] == [600020, "عميل ثابت ج", 1234, "مندوب تجريبي 08", 0, 0, 0, *IN_COLS]
    assert body[3:] == [                                                         # 3: outside, by code
        [600030, "اسم ملف د", None, "مندوب تجريبي 03", 30, -5, 25, *OUT_COLS],
        [600040, "اسم ملف خامس", None, "غير محدد", 0, -7, -7, *OUT_COLS],      # 12: Main Data, blank
        [999111, "زائر", None, "غير مسجل", 9, 0, 9, *OUT_COLS]]              # not in Main Data
    codes = [r[0] for r in body]
    assert len(codes) == len(set(codes)) and 600050 not in codes                 # 2: once each; October only: out
    assert all(r[7:] == IN_COLS for r in body[:3]) and all(r[7:] == OUT_COLS for r in body[3:])  # 4, 5
    assert total == [GRAND_TOTAL, None, None, None, 139, -22, 117, None, None]
    assert all(r[6] == r[4] + r[5] for r in rows)                                # 8: net = sales + returns


def test_customers_report_matches_the_raw_csv(tmp_path):
    """6, 7, 9: sales / returns per code = a plain sum of the raw CSV rows of 01/11 - 16/11 (October and 17/11 out)."""
    result, _, _, rows = _cust_report(tmp_path)
    want = {}
    for r in CUST_CSV:
        d = date(int(r[0][6:]), int(r[0][3:5]), int(r[0][:2]))
        if date(2026, 11, 1) <= d <= N16:
            s = want.setdefault(int(r[5]), [D(0), D(0)])
            s[0 if r[1] == S else 1] += D(r[19].replace(",", ""))
    got = {r[0]: [D(str(r[4])), D(str(r[5]))] for r in rows[:-1]}
    assert {k: v for k, v in got.items() if any(v)} == want
    assert fc.reconcile_customers(result) == []


def test_customers_report_period_follows_the_reference_month(tmp_path):
    rows = [row("30-11-2026", S, "5.00", 600010, doc="1"), row("02-12-2026", S, "7.00", 600010, doc="2"),
            row("05-12-2026", R, "-1.00", 600010, doc="3"), row("06-12-2026", S, "100.00", 600010, doc="4"),
            row("31-12-2026", S, "3.00", 600030, doc="5"), row("20-01-2027", S, "4.00", 600010, doc="6")]
    _, title, header, got = _cust_report(tmp_path / "dec", ref=date(2026, 12, 5), rows=rows)
    assert title.endswith("الفترة من 01/12/2026 إلى 05/12/2026") and header[4] == "مبيعات شهر ديسمبر 2026"
    assert got[0][4:7] == [7, -1, 6] and 600030 not in [r[0] for r in got]       # 30/11 and 06/12 out
    _, title, header, got = _cust_report(tmp_path / "jan", ref=date(2027, 1, 20), rows=rows)
    assert title.endswith("الفترة من 01/01/2027 إلى 20/01/2027") and header[6] == "صافي شهر يناير 2027"
    assert got[0][4:7] == [4, 0, 4] and got[-1][4:7] == [4, 0, 4]


def test_customers_report_all_transactions_uses_the_latest_month(tmp_path):
    result, title, header, rows = _cust_report(tmp_path, mode=MODE_ALL)
    assert "كل الحركات: أحدث شهر في ملفات الإدخال نوفمبر 2026" in title and header[4] == "مبيعات شهر نوفمبر 2026"
    assert rows[0][4:7] == [150, -10, 140]                                       # 17/11 counted (no end), October not
    assert rows[-1][4:7] == [189, -22, 167]
    more = [*CUST_CSV, row("03-12-2026", S, "8.00", 600020, doc="20")]          # a December row: the latest month
    _, title, header, rows = _cust_report(tmp_path / "dec", mode=MODE_ALL, rows=more)
    assert "ديسمبر 2026" in title and header[5] == "مرتجعات شهر ديسمبر 2026"
    assert [r[4:7] for r in rows[:3]] == [[0, 0, 0], [0, 0, 0], [8, 0, 8]] and len(rows) == 4   # only December


def test_customers_report_changes_neither_existing_workbook(tmp_path):
    """The same run without and with the third workbook: both existing workbooks identical, cell by cell."""
    write_erp_csv(tmp_path / "a" / "input" / "c.csv", CUST_CSV)
    write_erp_csv(tmp_path / "b" / "input" / "c.csv", CUST_CSV)
    fixed_path = _write_fixed_list(tmp_path / "fixed" / FIXED_CODES_NAME)
    main_data = load_main_data(write_main(tmp_path / "main.xlsx", CUST_MAIN_ROWS))
    fixed = load_fixed_codes(fixed_path)
    plain = run_analysis(_config(tmp_path / "a"), MODE_YESTERDAY, N16, main_data, fixed_codes=fixed, now=RUN_AT)
    full = run_analysis(_config(tmp_path / "b"), MODE_YESTERDAY, N16, main_data, fixed_codes=fixed, now=RUN_AT,
                        fixed_customers=load_fixed_customers(fixed_path))
    assert plain.customers_report_path is None and full.customers_report_path.is_file()
    assert _raw(plain.report_path) == _raw(full.report_path)
    assert _raw(plain.reps_report_path) == _raw(full.reps_report_path)


def test_fixed_code_file_is_read_exactly_and_never_written(tmp_path):
    path = _write_fixed_list(tmp_path / FIXED_CODES_NAME)
    before = _sha(path)
    customers = load_fixed_customers(path)
    assert [(c.code, c.name, c.top_month, c.responsible) for c in customers] == FIXED_ROWS
    assert [c.key for c in customers] == ["600010", "600005", "600020"] and _sha(path) == before
    with pytest.raises(MainDataError, match="columns A-D must be"):
        load_fixed_customers(_write_fixed_list(tmp_path / "bad" / FIXED_CODES_NAME,
                                               header=("كود", "الاسم", "القيمة", "المندوب")))
    with pytest.raises(MainDataError, match="appears twice"):
        load_fixed_customers(_write_fixed_list(tmp_path / "dup" / FIXED_CODES_NAME, rows=[*FIXED_ROWS, FIXED_ROWS[0]]))


def test_wrong_customers_report_is_detected(tmp_path, monkeypatch):
    result, _, _, _ = _cust_report(tmp_path)
    original = fc.customers_table

    def patched(change):
        monkeypatch.setattr(fc, "customers_table", lambda r: change(original(r)))
        return fc.reconcile_customers(result)

    assert any("customer 600030 sales / returns" in p for p in
               patched(lambda t: [[*r[:4], r[4] + 1, *r[5:]] if r[0] == 600030 else r for r in t]))
    assert any("not all there in the order" in p for p in patched(lambda t: [t[1], t[0], *t[2:]]))
    assert any("appears more than once" in p for p in patched(lambda t: [*t[:-1], t[3], t[-1]]))
    assert any("customer 600030 is classified" in p for p in
               patched(lambda t: [[*r[:7], CUSTOMER_MARK, ""] if r[0] == 600030 else r for r in t]))
    assert any("المسؤول" in p and "Main Data" in p for p in
               patched(lambda t: [[*r[:3], "مندوب تجريبي 01", *r[4:]] if r[0] == 600040 else r for r in t]))
    assert any("differ from the customers with activity" in p for p in patched(lambda t: [*t[:3], *t[4:]]))
    monkeypatch.setattr(fc, "customers_table", original)
    wb = load_workbook(result.customers_report_path)
    wb[fc.SHEET_CUSTOMERS].cell(row=3, column=5).value = 1.0              # a saved sales cell
    wb.save(result.customers_report_path)
    problems = analysis.verify_task5(result)
    assert any("Fixed_Customers_Report_2026-11-16.xlsx / تقرير عملاء الـ224: row 3" in p for p in problems)
    assert any("net is not sales + returns" in p for p in problems)


def test_customers_report_layout_and_print(tmp_path):
    result, _, _, rows = _cust_report(tmp_path)
    ws = load_workbook(result.customers_report_path)[fc.SHEET_CUSTOMERS]
    last = ws.max_row                                                            # the block's last row
    assert ws.sheet_view.rightToLeft and ws.freeze_panes == "A3" and ws.auto_filter.ref == f"A2:I{len(rows) + 1}"
    assert int(ws.page_setup.paperSize) == 9 and ws.page_setup.orientation == "landscape"
    assert ws.page_setup.fitToWidth == 1 and ws.print_title_rows == "$1:$2"
    assert ws.print_area.endswith(f"$A$1:$I${last}")
    assert [ws.cell(row=3, column=c).number_format for c in (3, 5, 6, 7)] == ["#,##0.00"] * 4
    total_row = len(rows) + 2
    assert ws.cell(row=total_row, column=5).number_format == "#,##0.00"
    assert ws.cell(row=total_row, column=1).value == ws.cell(row=last, column=1).value == GRAND_TOTAL
    header = ws.cell(row=2, column=1)
    assert header.font.name == "Segoe UI" and header.font.bold and header.fill.fgColor.rgb.endswith("BFBFBF")
    assert ws.cell(row=total_row, column=1).border.top.style == "double"
    for c in range(1, 10):                                                       # the table (the block wraps)
        assert ws.column_dimensions[get_column_letter(c)].width >= _content_width(ws, c, total_row)


def test_cli_writes_the_customers_report_each_run(tmp_path, setup, capsys):
    assert _run(tmp_path) == EXIT_OK
    out = capsys.readouterr().out
    (path,) = (tmp_path / "out").rglob("Fixed_Customers_Report_*.xlsx")
    assert "Customer report          :" in out and "3 fixed customers +" in out
    rows = [list(r) for r in load_workbook(path).active.iter_rows(min_row=3, values_only=True)]
    assert [r[0] for r in rows[:3]] == [100001, 555555, 100003]                  # fixed code.xlsx order



# ------------------------------------------------------------------ Fixed_Customers_Report: one sheet per المسؤول
def _cust_sheets(result):
    wb = load_workbook(result.customers_report_path)
    return wb, {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


BLOCK_HEADER = [None, "إجمالي العملاء", "عملاء داخل الـ224", "عملاء خارج الـ224", "إجمالي المبيعات",
                "إجمالي المرتجعات", "إجمالي الصافي", "عملاء داخل الـ224 حققوا مبيعات", "عملاء خارج الـ224 حققوا مبيعات"]
RESP_FIXED = [  # a person twice, the ى spelling of مندوب تجريبي 10, a person outside the given list (اغلاق)
    (600010, "أ", 10, "مندوب تجريبي 02"), (600005, "ب", 0, "مندوب تجريبى 10"), (600020, "ج", 5, "مندوب تجريبي 02"),
    (600060, "د", 7, "اغلاق"),
]
RESP_CSV = [*CUST_CSV, row("03-11-2026", S, "12.00", 600060, doc="30"),
            row("04-11-2026", S, "4.00", 600070, "اسم ملف ز", doc="31")]           # 600070: Main Data مندوب تجريبي 10


def _resp_run(tmp_path, mode=MODE_YESTERDAY):
    write_erp_csv(tmp_path / "input" / "c.csv", RESP_CSV)
    fixed_path = _write_fixed_list(tmp_path / "data" / FIXED_CODES_NAME, rows=RESP_FIXED)
    main_rows = [*CUST_MAIN_ROWS, (600060, "x", "الجملة", "اغلاق"), (600070, "y", "الجملة", "مندوب تجريبي 10")]
    result = run_analysis(_config(tmp_path), mode, N16 if mode == MODE_YESTERDAY else None,
                          load_main_data(write_main(tmp_path / "data" / "main.xlsx", main_rows)),
                          fixed_codes=load_fixed_codes(fixed_path), fixed_customers=load_fixed_customers(fixed_path),
                          now=datetime(2026, 11, 16, 13, 0))
    assert analysis.verify_task5(result) == []
    return result


def test_one_sheet_per_responsible_person(tmp_path):
    result = _resp_run(tmp_path)
    wb, sheets = _cust_sheets(result)
    names = wb.sheetnames
    assert names[0] == fc.SHEET_CUSTOMERS
    assert names[1:12] == list(fc.CUSTOMER_RESPONSIBLES)                    # 1: the 11 people, in order
    assert names[12:] == ["اغلاق", "مندوب تجريبي 03", "غير محدد", "غير مسجل"]      # the others: never dropped
    assert len(names) == len({n.casefold() for n in names})                      # 2: no duplicate sheet
    assert all(len(n) <= 31 and not any(ch in n for ch in "[]:*?/\\") for n in names)
    shimi = sheets["مندوب تجريبى 10"][2:-4]
    assert [r[0] for r in shimi] == [600005, 600070]                             # ى / ي: one person, one sheet
    assert [r[3] for r in shimi] == ["مندوب تجريبى 10", "مندوب تجريبي 10"]              # each name as written
    assert sheets["مندوب تجريبي 11"][2:] == [[None, GRAND_TOTAL, 0, None, 0, 0, 0, 0, 0],   # its total row
                                               [None] * 9, ["مندوب تجريبي 11", *BLOCK_HEADER[1:]],
                                               [GRAND_TOTAL, 0, 0, 0, 0, 0, 0, 0, 0]]   # listed, no customer


def test_every_customer_on_exactly_one_responsible_sheet(tmp_path):
    result = _resp_run(tmp_path)
    _, sheets = _cust_sheets(result)
    main = sheets[fc.SHEET_CUSTOMERS][2:-4]                                # customers only
    split = [tuple(r) for name, rows in sheets.items() if name != fc.SHEET_CUSTOMERS for r in rows[2:-4]]
    assert sorted(split, key=str) == sorted(map(tuple, main), key=str)          # 3, 4: none missing, none twice
    codes = [r[0] for r in split]
    assert len(codes) == len(set(codes))
    for name, rows in sheets.items():                                           # never moved to another person
        if name != fc.SHEET_CUSTOMERS:
            assert {analysis.canonical(r[3]) for r in rows[2:-4]} <= {analysis.canonical(name)}
    assert [r[0] for r in sheets["مندوب تجريبي 02"][2:-4]] == [600010, 600020]   # column D as written


def test_responsible_totals_add_up_to_the_main_block(tmp_path):
    result = _resp_run(tmp_path)
    _, sheets = _cust_sheets(result)
    main = sheets[fc.SHEET_CUSTOMERS]
    block = main[-1]
    assert main[-2] == BLOCK_HEADER and main[-3] == [None] * 9
    customers = main[2:-4]
    inside = [r for r in customers if r[7] == CUSTOMER_MARK]
    outside = [r for r in customers if r[8] == CUSTOMER_MARK]
    assert block[:4] == [GRAND_TOTAL, len(inside) + len(outside), len(inside), len(outside)] == [GRAND_TOTAL, 8, 4, 4]
    assert block[4:7] == main[-4][4:7] == [155, -22, 133]                        # = the main total row
    parts = [rows[-1] for name, rows in sheets.items() if name != fc.SHEET_CUSTOMERS]
    for i in range(1, 9):                                                        # 5, 6, 8 (+ who sold)
        assert sum(p[i] for p in parts) == block[i], i
    for name, rows in sheets.items():                                           # 7, 9: each sheet's own block
        detail = rows[2:-4]                                    # above the total row
        last = rows[-1]
        assert last[1] == len({r[0] for r in detail}) == last[2] + last[3], name
        assert last[2] == len({r[0] for r in detail if r[7] == CUSTOMER_MARK})
        assert last[7] == len({r[0] for r in detail if r[7] == CUSTOMER_MARK and r[4] > 0})
        assert last[8] == len({r[0] for r in detail if r[8] == CUSTOMER_MARK and r[4] > 0})
        assert last[4:7] == [sum(r[4] for r in detail), sum(r[5] for r in detail), sum(r[6] for r in detail)]
        assert last[6] == last[4] + last[5]
    assert sheets["مندوب تجريبي 03"][-1] == [GRAND_TOTAL, 1, 0, 1, 30, -5, 25, 0, 1]   # returns keep their sign


def test_customer_count_is_distinct_codes():
    rows = [[1, "a", 0, "x", D("5"), D("0"), D("5"), CUSTOMER_MARK, ""],
            [1, "a", 0, "x", D("3"), D("-1"), D("2"), CUSTOMER_MARK, ""],         # the same code twice
            ["0002", "b", 0, "x", D("0"), D("0"), D("0"), "", CUSTOMER_MARK],
            [2, "b", 0, "x", D("1"), D("0"), D("1"), "", CUSTOMER_MARK]]        # 0002 = 2
    assert fc.customers_summary(rows) == [GRAND_TOTAL, 2, 1, 1, D("9"), D("-1"), D("8"), 1, 1]


def test_responsible_sheet_layout_and_print(tmp_path):
    result = _resp_run(tmp_path)
    wb, sheets = _cust_sheets(result)
    for ws in wb.worksheets:
        rows = sheets[ws.title]
        customers = len(rows) - 6                                            # title, header, total, block
        assert ws.sheet_view.rightToLeft and ws.freeze_panes == "A3", ws.title
        assert ws.auto_filter.ref == f"A2:I{customers + 2}", ws.title            # the customer rows only
        assert int(ws.page_setup.paperSize) == 9 and ws.page_setup.orientation == "landscape"
        assert ws.page_setup.fitToWidth == 1 and ws.print_title_rows == "$1:$2"
        assert ws.print_area.endswith(f"$A$1:$I${ws.max_row}")                   # the block is printed
        assert ws.cell(row=ws.max_row, column=5).number_format == "#,##0.00"
        assert ws.cell(row=2, column=1).fill.fgColor.rgb.endswith("BFBFBF")
        assert ws.column_dimensions["A"].width >= len(GRAND_TOTAL) + 1              # the block label not cut
        if ws.title != fc.SHEET_CUSTOMERS:
            title = ws.cell(row=1, column=1).value
            assert title.startswith("تقرير عملاء قائمة الـ224  |  الفترة من 01/11/2026 إلى 16/11/2026")
            assert title.endswith(f"المسؤول: {ws.title}")                         # the full name


def test_responsible_sheets_in_all_transactions_mode_keep_the_period(tmp_path):
    result = _resp_run(tmp_path, mode=MODE_ALL)
    _, sheets = _cust_sheets(result)
    dalia = sheets["مندوب تجريبي 02"]
    assert "كل الحركات: أحدث شهر في ملفات الإدخال نوفمبر 2026" in dalia[0][0]       # 11: the dynamic period
    assert dalia[1][4] == "مبيعات شهر نوفمبر 2026" and dalia[-1][4] == 150         # 17/11 counted: no end date


def test_main_customer_sheet_is_unchanged_above_the_block(tmp_path):
    """The rows the main sheet had before the per-person sheets: identical, the block only added below."""
    _, _, _, rows = _cust_report(tmp_path)
    assert rows == [[600010, "  عميل ثابت أ", 500.25, "مندوب تجريبي 02", 100, -10, 90, *IN_COLS],
                    [600005, "عميل ثابت ب", 0, "اغلاق", 0, 0, 0, *IN_COLS],
                    [600020, "عميل ثابت ج", 1234, "مندوب تجريبي 08", 0, 0, 0, *IN_COLS],
                    [600030, "اسم ملف د", None, "مندوب تجريبي 03", 30, -5, 25, *OUT_COLS],
                    [600040, "اسم ملف خامس", None, "غير محدد", 0, -7, -7, *OUT_COLS],
                    [999111, "زائر", None, "غير مسجل", 9, 0, 9, *OUT_COLS],
                    [GRAND_TOTAL, None, None, None, 139, -22, 117, None, None]]


def test_wrong_responsible_split_is_detected(tmp_path, monkeypatch):
    result = _resp_run(tmp_path)
    original = fc.customers_by_responsible

    def drop(detail):                                                            # a customer lost from its sheet
        return [(label, rows[1:] if label == "مندوب تجريبي 02" else rows) for label, rows in original(detail)]
    monkeypatch.setattr(fc, "customers_by_responsible", drop)
    problems = fc.reconcile_customers(result)
    assert any("the المسؤول sheets miss customers ['600010']" in p for p in problems)
    assert any("do not add up to الإجمالي العام" in p for p in problems)

    def twice(detail):                                                           # a customer on two sheets
        groups = original(detail)
        last = next(rows for _, rows in reversed(groups) if rows)
        return [(groups[0][0], [*groups[0][1], last[0]]), *groups[1:]]
    monkeypatch.setattr(fc, "customers_by_responsible", twice)
    assert any("is on the sheets of" in p for p in fc.reconcile_customers(result))
    monkeypatch.setattr(fc, "customers_by_responsible", original)
    wb = load_workbook(result.customers_report_path)
    ws = wb["مندوب تجريبي 02"]
    ws.cell(row=ws.max_row, column=2).value = 5                                  # a saved customer count
    wb.save(result.customers_report_path)
    assert any("مندوب تجريبي 02" in p and "block" in p for p in analysis.verify_task5(result))



# ------------------------------------------------------------------ Fixed_Customers_Report: who sold (sales > 0)
SOLD_CSV = [
    row("02-11-2026", S, "10.00", 600010, doc="1"), row("03-11-2026", R, "-50.00", 600010, doc="2"),  # net < 0: sold
    row("04-11-2026", R, "-5.00", 600005, doc="3"),                                                 # returns only
    row("05-11-2026", S, "0.00", 600020, doc="4"),                                                  # a zero invoice
    row("06-11-2026", S, "5.00", 600030, doc="5"), row("07-11-2026", S, "6.00", 600030, doc="6"),    # two invoices
    row("08-11-2026", R, "-7.00", 600040, doc="7"),                                                 # returns only
    row("09-11-2026", S, "0.00", 999111, "زائر", doc="8"),                                          # outside, zero sale
]


def test_each_customer_has_exactly_one_list_column(tmp_path):
    result, _, header, rows = _cust_report(tmp_path, rows=SOLD_CSV)
    assert header[7:] == ["عملاء داخل قائمة الـ224", "عملاء خارج قائمة الـ224"]
    body = rows[:-1]
    assert all((r[7] == CUSTOMER_MARK) != (r[8] == CUSTOMER_MARK) for r in body)          # 1: never both / neither
    assert [r[7:] for r in body] == [IN_COLS] * 3 + [OUT_COLS] * 3                     # the same classification
    assert {r[7] for r in body} <= {CUSTOMER_MARK, None} and {r[8] for r in body} <= {CUSTOMER_MARK, None}


def test_customers_who_sold_are_distinct_codes_with_sales_above_zero(tmp_path):
    """2-7: 600010 sold (sales 10 although its net is -40), 600005 returns only, 600020 a zero invoice: 1 of 3
    inside; 600030 sold with two invoices (once), 600040 returns only, 999111 a zero sale: 1 of 3 outside."""
    result, _, _, rows = _cust_report(tmp_path, rows=SOLD_CSV)
    ws = load_workbook(result.customers_report_path)[fc.SHEET_CUSTOMERS]
    block = [c.value for c in ws[ws.max_row]]
    assert block == [GRAND_TOTAL, 6, 3, 3, 21, -62, -41, 1, 1]
    assert block[1] == block[2] + block[3]                                        # 2: inside + outside = all
    assert [r[4:7] for r in rows[:-1]] == [[10, -50, -40], [0, -5, -5], [0, 0, 0], [11, 0, 11], [0, -7, -7],
                                          [0, 0, 0]]                              # amounts as they were
    assert [c.value for c in ws[ws.max_row - 1]][7:] == ["عملاء داخل الـ224 حققوا مبيعات",
                                                          "عملاء خارج الـ224 حققوا مبيعات"]
    assert fc.reconcile_customers(result) == []


def test_who_sold_per_person_adds_up_to_the_main_sheet(tmp_path):
    result = _resp_run(tmp_path)
    _, sheets = _cust_sheets(result)
    main_block = sheets[fc.SHEET_CUSTOMERS][-1]
    parts = [rows[-1] for name, rows in sheets.items() if name != fc.SHEET_CUSTOMERS]
    assert [sum(p[7] for p in parts), sum(p[8] for p in parts)] == main_block[7:9]          # 8
    assert main_block[1:4] == [8, 4, 4] and main_block[7:9] == [2, 3]   # listed and in the report, not who sold
    for name, rows in sheets.items():
        detail = rows[2:-4]                                    # above the total row
        assert rows[-1][7] <= rows[-1][2] and rows[-1][8] <= rows[-1][3], name   # never more than listed


def test_wrong_who_sold_count_is_detected(tmp_path, monkeypatch):
    result, _, _, _ = _cust_report(tmp_path, rows=SOLD_CSV)
    original = fc.customers_summary
    monkeypatch.setattr(fc, "customers_summary", lambda rows, *a: [*original(rows, *a)[:7], 3, *original(rows, *a)[8:]])
    assert any("الإجمالي العام block" in p for p in fc.reconcile_customers(result))
    monkeypatch.setattr(fc, "customers_summary", original)
    wb = load_workbook(result.customers_report_path)
    ws = wb[fc.SHEET_CUSTOMERS]
    ws.cell(row=ws.max_row, column=9).value = 3                                  # a saved count of who sold
    ws.cell(row=3, column=9).value = CUSTOMER_MARK                               # a customer marked twice
    wb.save(result.customers_report_path)
    problems = analysis.verify_task5(result)
    assert any("marked in both or neither list column" in p for p in problems)
    assert any("saved الإجمالي العام block" in p for p in problems)



# ------------------------------------------------------------------ report 3 as its own step: name cell, own check
def test_responsible_name_above_the_grand_total_of_every_person_sheet(tmp_path):
    """1, 2: the cell right above الإجمالي العام (first column of the block's header row) holds the sheet's own
    المسؤول - found from the sheet itself, whatever its number of customers; the main sheet keeps it empty."""
    result = _resp_run(tmp_path)
    wb = load_workbook(result.customers_report_path)
    people = [ws for ws in wb.worksheets if ws.title != fc.SHEET_CUSTOMERS]
    seen = []
    for ws in people:
        total = max(r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == GRAND_TOTAL)
        label = ws.cell(row=total - 1, column=1)
        assert total == ws.max_row and label.value == ws.title, ws.title          # the full name (short here)
        assert ws.cell(row=1, column=1).value.endswith(f"المسؤول: {label.value}")
        assert label.font.bold and label.fill.fgColor.rgb.endswith("BFBFBF")     # the block header's style
        assert [ws.cell(row=total - 1, column=c).value for c in range(2, 10)] == BLOCK_HEADER[1:]
        others = {p.title for p in people} - {ws.title}
        assert not any(isinstance(c.value, str) and c.value in others for row_ in ws.iter_rows() for c in row_)
        seen.append(label.value)
    assert len(seen) == len(set(seen)) == len(people)
    main = wb[fc.SHEET_CUSTOMERS]
    assert main.cell(row=main.max_row - 1, column=1).value is None                # the main sheet: unchanged


def test_name_cell_changes_no_total(tmp_path):
    """3: the block values are those of the customer rows, exactly as before the name was added."""
    result = _resp_run(tmp_path)
    _, sheets = _cust_sheets(result)
    for name, rows in sheets.items():
        detail = rows[2:-4]                                    # above the total row
        assert rows[-1] == [GRAND_TOTAL, *fc.customers_summary([[*r[:7], r[7] or "", r[8] or ""] for r in detail])[1:]]


def test_report_three_is_checked_on_its_own(tmp_path):
    """4: report 3 has its own check; a broken report 3 never shows as a problem of reports 1-2, and vice versa."""
    result = _resp_run(tmp_path)
    assert fc.verify_customers_report(result) == [] and analysis.verify_analysis_report(result) == []
    wb = load_workbook(result.customers_report_path)
    wb[fc.SHEET_CUSTOMERS].cell(row=3, column=5).value = 1.0
    wb.save(result.customers_report_path)
    assert analysis.verify_analysis_report(result) == []                          # reports 1-2: still fine
    problems = analysis.verify_task5(result)
    assert result.report_problems == [] and result.customers_problems and problems == result.customers_problems
    assert all("Fixed_Customers_Report" in p or "تقرير عملاء الـ224" in p for p in result.customers_problems)


def test_a_failing_report_three_leaves_reports_one_and_two_untouched(tmp_path, monkeypatch):
    def broken(result, path):
        raise OSError("disk full")
    monkeypatch.setattr(fc, "write_customers_report", broken)
    write_erp_csv(tmp_path / "input" / "c.csv", CUST_CSV)
    fixed_path = _write_fixed_list(tmp_path / "data" / FIXED_CODES_NAME)
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, N16,
                          load_main_data(write_main(tmp_path / "data" / "main.xlsx", CUST_MAIN_ROWS)),
                          fixed_codes=load_fixed_codes(fixed_path), fixed_customers=load_fixed_customers(fixed_path))
    assert result.report_path.is_file() and result.reps_report_path.is_file()  # reports 1-2 written
    assert result.customers_report_path is None and "disk full" in result.customers_error
    assert analysis.verify_analysis_report(result) == []
    assert analysis.verify_task5(result) == [result.customers_error] == result.customers_problems


def test_cli_shows_report_three_pass_and_fail_separately(tmp_path, setup, capsys, monkeypatch):
    assert _run(tmp_path) == EXIT_OK
    out = capsys.readouterr().out
    assert "Reconciliation           : PASS" in out and "Customer report check    : PASS" in out
    assert out.index("Reconciliation") < out.index("Customer report check")       # report 3 after reports 1-2
    (tmp_path / "input").mkdir(exist_ok=True)
    for p in (tmp_path / "processed").rglob("*.csv"):                              # the same file again
        p.rename(tmp_path / "input" / p.name)
    monkeypatch.setattr(fc, "verify_customers_report", lambda result: ["report 3 broken on purpose"])
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    out = capsys.readouterr().out
    assert "Reconciliation           : PASS" in out                               # reports 1-2 verified
    assert "Customer report check    : FAIL" in out and "FAIL  report 3 broken on purpose" in out
    assert [p.name for p in (tmp_path / "input").iterdir()] == ["month.csv"]      # nothing moved
