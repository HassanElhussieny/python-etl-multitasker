"""Task 5: Sector & Representative Sales Analysis (Main Data mapping, four sheets incl. Master Data, lifecycle)."""

import hashlib
import json
from dataclasses import replace
from datetime import date, datetime
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from sales_etl import analysis
from sales_etl.analysis import (
    ALL_COLUMNS,
    ALL_COUNT_COLUMNS,
    COVER_COLUMNS,
    FIXED_CODES_NAME,
    FIXED_COLUMNS,
    MASTER_COLUMNS,
    GRAND_TOTAL,
    MISSING_COLUMNS,
    MODE_ALL,
    MODE_YESTERDAY,
    REPORT_REPRESENTATIVES,
    REPRESENTATIVES,
    RETURNS_ONLY_COLUMN,
    SECTORS,
    SHEET_COVER,
    SHEET_MASTER,
    SHEET_MISSING,
    SHEET_REPS,
    SHEET_SECTORS,
    TOTAL_CUSTOMERS_COLUMN,
    UNREGISTERED,
    WORKBOOK_NAME,
    YESTERDAY_COLUMNS,
    YESTERDAY_COUNT_COLUMNS,
    MainDataError,
    customer_key,
    find_main_data,
    load_fixed_codes,
    load_main_data,
    run_analysis,
)
from sales_etl.cli import EXIT_AUDIT_FAILED, EXIT_NO_INPUT, EXIT_OK, EXIT_USAGE, main
from sales_etl.config import PipelineConfig
from tests.conftest import S, R, alex8_rows, write_erp_csv

D = Decimal
WRONG_REP = ("99", "مندوب من الملف")   # CSV columns I/J: must never be used by task 5

MAIN = [  # A code, B name, C sector, D representative
    (100001, "عميل أ", "الجملة ", "سامح السيد محمود احم"),                 # trailing space as in the real file
    (100002, "عميل ب", "الجملة", "داليا صلاح الدين فكر"),                  # same sector, other rep
    (100003, "عميل ج", "صيدليات وسلاسل صيدليات", "سامح السيد محمود احم"),  # same rep, other sector
    (100004, "عميل د", "مستشفيات", "هند محمد رضا"),
]


def row(day, kind, net, customer, name="", doc="1"):
    """ERP row (21 cells): A date, B type, F code, G name, I/J representative (ignored), T صافى."""
    return [day, kind, doc, "93", "", str(customer), name or f"اسم {customer}", "", *WRONG_REP, net,
            *["0"] * 8, net, ""]


ROWS = [
    row("30-10-2026", S, "100.10", 100001, doc="1"),              # yesterday sales     الجملة / سامح
    row("29-10-2026", S, "200.20", 100002, doc="2"),              # previous sales      الجملة / داليا
    row("30-10-2026", R, "-30.00", 100003, doc="3"),              # yesterday returns   صيدليات / سامح
    row("01-10-2026", R, "-5.05", 100001, doc="4"),               # previous returns    الجملة / سامح
    row("30-10-2026", S, "50.00", 555555, "زائر جديد", doc="5"),  # unregistered
    row("15-10-2026", R, "-7.25", 555555, "زائر جديد", doc="6"),  # unregistered, same code
    row("10-10-2026", R, "2.00", 100004, doc="7"),                # positive return: sign kept
    row("31-10-2026", S, "999.99", 100001, doc="8"),              # after the date: excluded in type 1
    row("20-10-2026", "تسوية", "1000.00", 100001, doc="9"),        # unknown type: never counted
    row("12-10-2026", S, "0.01", 666666, "عميل مؤقت", doc="10"),  # second unregistered code
]
REF = date(2026, 10, 30)


def write_main(path, rows=MAIN):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet2"
    ws.append(["كود العميل", "اسم العميل", "قطاعات العملاء", "المندوب"])
    for r in rows:
        ws.append(list(r))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _config(tmp_path):
    return PipelineConfig(input_dir=tmp_path / "input", output_dir=tmp_path / "output", reference_date=REF)


def _num(v):
    return Decimal(str(v)).quantize(Decimal("0.01")) if isinstance(v, (int, float)) else v


def _sheets(path):
    """Sheet -> header + rows as saved, without the report-date line (row 1), the first column م, and - on
    Master Data-style sheets - the الإجمالي العام row under the transactions (these have their own tests)."""
    wb = load_workbook(path, data_only=True)
    sheets = {}
    for ws in wb.worksheets:
        rows = [[_num(v) for v in r] for r in ws.iter_rows(values_only=True)]
        if ws.title == "Pivot for Last Day":                                      # own layout: no م, own title
            sheets[ws.title] = rows[1:]
            continue
        assert str(rows[0][0]).startswith("تاريخ التقرير: ") and rows[1][0] == "م", (ws.title, rows[:2])
        rows = [r[1:] for r in rows[1:]]
        if rows[0][0] == "التاريخ":                                               # Master Data / representative
            assert rows[-1][0] == GRAND_TOTAL, ws.title
            rows = rows[:-1]
        sheets[ws.title] = rows
    return sheets


def _d(*values):
    return [D(v).quantize(D("0.01")) for v in values]


def _money_width(rows):
    yesterday = len(rows[0]) == 1 + len(YESTERDAY_COLUMNS) + len(YESTERDAY_COUNT_COLUMNS)
    return len(YESTERDAY_COLUMNS) if yesterday else len(ALL_COLUMNS)


def _dated(*columns, ref=REF):
    """The headers as printed for the reference date ``ref`` (أمس / الأيام السابقة carry the dates)."""
    return list(analysis.date_headers(columns, ref))


def _by_label(rows):
    """Label -> money columns (the distinct-customer columns that follow are read by _counts)."""
    n = _money_width(rows)
    return {r[0]: r[1:1 + n] for r in rows[1:]}


def _counts(rows):
    """Label -> distinct-customer columns up to إجمالي العملاء المميزين, as integers (the split: _split)."""
    n, end = _money_width(rows), rows[0].index(TOTAL_CUSTOMERS_COLUMN) + 1
    return {r[0]: [int(v) for v in r[1 + n:end]] for r in rows[1:]}


def _split(rows):
    """Label -> [عدد العملاء من الـ224 كود, عدد العملاء خارج الـ224 كود] (right after the total)."""
    i = rows[0].index(TOTAL_CUSTOMERS_COLUMN) + 1
    assert tuple(rows[0][i:i + 2]) == FIXED_COLUMNS
    return {r[0]: [int(v) for v in r[i:i + 2]] for r in rows[1:]}


def _returns_only(rows):
    """Label -> عدد عملاء المرتجعات فقط (the last column)."""
    assert rows[0][-1] == RETURNS_ONLY_COLUMN
    return {r[0]: int(r[-1]) for r in rows[1:]}


def write_fixed(path, codes):
    wb = Workbook()
    for code in codes:
        wb.active.append([code])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


@pytest.fixture
def setup(tmp_path):
    write_erp_csv(tmp_path / "input" / "month.csv", ROWS)
    write_fixed(tmp_path / "data" / FIXED_CODES_NAME, [100001, 555555, 100003])
    return load_main_data(write_main(tmp_path / "data" / "main.xlsx"))


# ------------------------------------------------------------------ calculations
def test_yesterday_and_previous_days(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    assert result.report_path.name == WORKBOOK_NAME == "تقرير تحليل المبيعات والمرتجعات.xlsx"
    sheets = _sheets(result.report_path)
    sectors, reps, missing = (sheets[n] for n in (SHEET_SECTORS, SHEET_REPS, SHEET_MISSING))
    assert sectors[0] == _dated("قطاع العملاء", *YESTERDAY_COLUMNS, *YESTERDAY_COUNT_COLUMNS)
    s = _by_label(sectors)
    assert s["الجملة"] == _d("100.10", "200.20", "300.30", "0", "-5.05", "-5.05", "295.25")
    assert s["صيدليات وسلاسل صيدليات"] == _d("0", "0", "0", "-30.00", "0", "-30.00", "-30.00")
    assert s["مستشفيات"] == _d("0", "0", "0", "0", "2.00", "2.00", "2.00")
    assert s[UNREGISTERED] == _d("50.00", "0.01", "50.01", "0", "-7.25", "-7.25", "42.76")
    assert s["شركات شحن"] == _d(*["0"] * 7)
    assert s[GRAND_TOTAL] == _d("150.10", "200.21", "350.31", "-30.00", "-10.30", "-40.30", "310.01")
    r = _by_label(reps)
    assert reps[0] == _dated("المندوب", *YESTERDAY_COLUMNS, *YESTERDAY_COUNT_COLUMNS)
    assert r["سامح السيد محمود احم"] == _d("100.10", "0", "100.10", "-30.00", "-5.05", "-35.05", "65.05")
    assert r["داليا صلاح الدين فكر"] == _d("0", "200.20", "200.20", "0", "0", "0", "200.20")
    assert r["هند محمد رضا"] == _d("0", "0", "0", "0", "2.00", "2.00", "2.00")
    assert r[UNREGISTERED] == s[UNREGISTERED] and r[GRAND_TOTAL] == s[GRAND_TOTAL]
    assert "مندوب من الملف" not in r and "99" not in r                     # CSV I/J never used
    assert missing[0] == list(MISSING_COLUMNS)
    assert missing[1:] == [
        ["555555", "زائر جديد", *_d("50.00", "-7.25", "42.75")],
        ["666666", "عميل مؤقت", *_d("0.01", "0", "0.01")],
        [GRAND_TOTAL, None, *_d("50.01", "-7.25", "42.76")],
    ]
    assert result.after_date_rows == 1 and result.other_type_rows == 1
    assert analysis.verify_analysis_report(result) == []


def test_all_transactions(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_ALL, None, setup)
    sheets = _sheets(result.report_path)
    s, r = _by_label(sheets[SHEET_SECTORS]), _by_label(sheets[SHEET_REPS])
    assert sheets[SHEET_SECTORS][0] == ["قطاع العملاء", *ALL_COLUMNS, *ALL_COUNT_COLUMNS]
    assert sheets[SHEET_REPS][0] == ["المندوب", *ALL_COLUMNS, *ALL_COUNT_COLUMNS]
    assert s["الجملة"] == _d("1300.29", "-5.05", "1295.24")                  # 31-10 row now included
    assert s[GRAND_TOTAL] == r[GRAND_TOTAL] == _d("1350.30", "-40.30", "1310.00")
    assert r["سامح السيد محمود احم"] == _d("1100.09", "-35.05", "1065.04")
    assert s[UNREGISTERED] == r[UNREGISTERED] == _d("50.01", "-7.25", "42.76")
    assert analysis.verify_analysis_report(result) == []


def test_decimal_precision(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row("30-10-2026", S, "0.10", 100001, doc=str(i)) for i in range(10)])
    result = run_analysis(_config(tmp_path), MODE_ALL, None, load_main_data(write_main(tmp_path / "main.xlsx")))
    assert result.sectors["الجملة"].sales == D("1.00") and result.direct.net == D("1.00")


def test_real_alex8_content_reconciles_with_erp_footer(tmp_path):
    """All 235 real transactions are unregistered here: the grand net must equal the ERP footer."""
    write_erp_csv(tmp_path / "input" / "alex8.csv", alex8_rows(), footer_net=D("1105646.62"), footer_count=235)
    result = run_analysis(_config(tmp_path), MODE_ALL, None, load_main_data(write_main(tmp_path / "main.xlsx")))
    assert result.direct.net == D("1105646.62") and analysis.verify_analysis_report(result) == []


# ------------------------------------------------------------------ distinct customer counts
def test_customer_counts_yesterday_and_previous_days(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    sheets = _sheets(result.report_path)
    s, r = _counts(sheets[SHEET_SECTORS]), _counts(sheets[SHEET_REPS])
    # [عدد عملاء أمس, عدد عملاء الأيام السابقة, إجمالي العملاء المميزين]: only customers with مبيعات
    assert s["الجملة"] == [1, 1, 2]               # 100001 sale yesterday (its earlier row is a return), 100002
    assert s["صيدليات وسلاسل صيدليات"] == [0, 0, 0]  # 100003 has only a return
    assert s["مستشفيات"] == [0, 0, 0]               # 100004 has only a return
    assert s[UNREGISTERED] == [1, 1, 2]           # 555555 sale yesterday (return before), 666666 sale before
    assert s["شركات شحن"] == [0, 0, 0]
    assert r["سامح السيد محمود احم"] == [1, 0, 1]  # 100001 only: 100003 has only a return
    assert r["داليا صلاح الدين فكر"] == [0, 1, 1] and r["هند محمد رضا"] == [0, 0, 0]
    assert r[UNREGISTERED] == [1, 1, 2]
    assert s[GRAND_TOTAL] == r[GRAND_TOTAL] == [2, 2, 4]   # distinct over the whole report
    assert analysis.verify_analysis_report(result) == []


def test_customer_counts_all_transactions(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_ALL, None, setup)
    sheets = _sheets(result.report_path)
    s, r = _counts(sheets[SHEET_SECTORS]), _counts(sheets[SHEET_REPS])
    assert s["الجملة"] == [2]                     # 100001 has 2 sales and a return here: counted once
    assert s["صيدليات وسلاسل صيدليات"] == [0] and s["مستشفيات"] == [0] and s[UNREGISTERED] == [2]  # returns only
    assert r["سامح السيد محمود احم"] == [1] and r["داليا صلاح الدين فكر"] == [1] and r[UNREGISTERED] == [2]
    assert s[GRAND_TOTAL] == r[GRAND_TOTAL] == [4]
    assert analysis.verify_analysis_report(result) == []


def test_many_transactions_of_one_customer_count_once(tmp_path):
    rows = [row("30-10-2026", S, "1.00", 100001, doc=str(i)) for i in range(5)]               # 5 x yesterday
    rows += [row(f"2{i}-10-2026", S, "1.00", 100001, doc=str(10 + i)) for i in range(10)]    # 10 x before
    rows += [row(f"2{i}-10-2026", R, "-1.00", 100001, doc=str(30 + i)) for i in range(5)]    # + returns
    write_erp_csv(tmp_path / "input" / "m.csv", rows)
    main_data = load_main_data(write_main(tmp_path / "main.xlsx"))
    yesterday = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, main_data)
    sheets = _sheets(yesterday.report_path)
    assert _counts(sheets[SHEET_SECTORS])["الجملة"] == [1, 1, 1]
    assert _counts(sheets[SHEET_REPS])["سامح السيد محمود احم"] == [1, 1, 1]
    assert _counts(sheets[SHEET_SECTORS])[GRAND_TOTAL] == [1, 1, 1]
    assert analysis.verify_analysis_report(yesterday) == []
    every = run_analysis(_config(tmp_path), MODE_ALL, None, main_data)
    assert _counts(_sheets(every.report_path)[SHEET_REPS])["سامح السيد محمود احم"] == [1]


def test_overlap_between_yesterday_and_previous_days_is_not_added(tmp_path):
    """20 customers yesterday, 80 before, 10 in both: 90 distinct, not 100."""
    rows = [row("30-10-2026", S, "1.00", 200000 + i, doc=str(i)) for i in range(20)]
    rows += [row("20-10-2026", S, "1.00", 200010 + i, doc=str(100 + i)) for i in range(80)]
    write_erp_csv(tmp_path / "input" / "m.csv", rows)
    main_data = load_main_data(write_main(tmp_path / "main.xlsx",
                                          [(200000 + i, "x", "الجملة", "الاداره") for i in range(90)]))
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, main_data)
    sheets = _sheets(result.report_path)
    assert _counts(sheets[SHEET_SECTORS])["الجملة"] == [20, 80, 90]
    assert _counts(sheets[SHEET_REPS])["الاداره"] == [20, 80, 90]
    assert _counts(sheets[SHEET_SECTORS])[GRAND_TOTAL] == [20, 80, 90]
    assert analysis.verify_analysis_report(result) == []


def test_customer_counts_ignore_excluded_rows_blank_codes_and_names(tmp_path):
    rows = [
        row("30-10-2026", S, "1.00", 100001, "اسم أول", doc="1"),
        row("29-10-2026", S, "1.00", "0100001", "اسم مختلف", doc="2"),   # same code, other spelling and name
        row("29-10-2026", S, "1.00", 100002, "اسم أول", doc="3"),        # same name, other code: 2 customers
        row("31-10-2026", S, "9.00", 100003, doc="4"),                   # after the date (type 1 only)
        row("20-10-2026", "تسوية", "9.00", 100004, doc="5"),             # unknown type: never a customer
        row("20-10-2026", S, "3.00", "", "بدون كود", doc="6"),           # blank code: amounts kept, not counted
    ]
    write_erp_csv(tmp_path / "input" / "m.csv", rows)
    main_data = load_main_data(write_main(tmp_path / "main.xlsx"))
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, main_data)
    sheets = _sheets(result.report_path)
    s = _counts(sheets[SHEET_SECTORS])
    assert s["الجملة"] == [1, 2, 2] and s["صيدليات وسلاسل صيدليات"] == [0, 0, 0] and s["مستشفيات"] == [0, 0, 0]
    assert s[UNREGISTERED] == [0, 0, 0]                                    # the blank code is not a customer
    assert _by_label(sheets[SHEET_SECTORS])[UNREGISTERED] == _d("0", "3.00", "3.00", "0", "0", "0", "3.00")
    assert s[GRAND_TOTAL] == [1, 2, 2]
    assert analysis.verify_analysis_report(result) == []
    every = _counts(_sheets(run_analysis(_config(tmp_path), MODE_ALL, None, main_data).report_path)[SHEET_SECTORS])
    assert every["صيدليات وسلاسل صيدليات"] == [1] and every[GRAND_TOTAL] == [3]   # 31-10 row counted in type 2


def test_grand_total_count_is_independent_of_the_rows(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    assert analysis.reconcile(result) == []
    result.direct.customers_p.add("123456")                       # a customer the rows do not have
    problems = analysis.reconcile(result)
    assert any(f"{SHEET_SECTORS}: grand total customer counts [2, 3, 5, 0, 5, 2]" in p for p in problems)
    assert any(f"{SHEET_REPS}: grand total customer counts" in p for p in problems)
    assert any(f"'{GRAND_TOTAL}' customer counts [2, 3, 5, 0, 5, 2] differ from the independent recount" in p
               for p in problems)


# ------------------------------------------------------------------ sales-only customer eligibility
def _one_group(tmp_path, rows, mode=MODE_YESTERDAY, files=None):
    """Write ``rows`` (or ``files``: name -> rows), run task 5; return (sector counts, rep counts)."""
    for name, file_rows in (files or {"m.csv": rows}).items():
        write_erp_csv(tmp_path / "input" / name, file_rows)
    main_data = load_main_data(write_main(tmp_path / "main.xlsx"))
    result = run_analysis(_config(tmp_path), mode, REF if mode == MODE_YESTERDAY else None, main_data)
    assert analysis.verify_analysis_report(result) == []                     # incl. the independent recount
    sheets = _sheets(result.report_path)
    return _counts(sheets[SHEET_SECTORS]), _counts(sheets[SHEET_REPS])


Y5, P5, F5 = "30-10-2026", "20-10-2026", "31-10-2026"


@pytest.mark.parametrize("rows, want", [
    ([row(Y5, S, "1.00", 100001)], [1, 0, 1]),                                              # 1 sales only
    ([row(P5, S, "1.00", 100001), row(Y5, R, "-1.00", 100001, doc="2")], [0, 1, 1]),       # 2 sales + returns
    ([row(Y5, R, "-1.00", 100001), row(P5, R, "-1.00", 100001, doc="2")], [0, 0, 0]),      # 3 returns only
    ([row(Y5, S, "1.00", 100001, doc=str(i)) for i in range(4)], [1, 0, 1]),               # 4 many sales
    ([row(d, k, "1.00" if k == S else "-1.00", 100001, doc=f"{d}{k}{i}")
      for d in (Y5, P5) for k in (S, R) for i in range(3)], [1, 1, 1]),                    # 5 many sales + returns
    ([row(Y5, S, "1.00", 100001), row(Y5, R, "-1.00", 100001, doc="2")], [1, 0, 1]),       # 6 both yesterday
    ([row(Y5, R, "-1.00", 100001)], [0, 0, 0]),                                            # 7 return yesterday only
    ([row(P5, S, "1.00", 100001), row(P5, R, "-1.00", 100001, doc="2")], [0, 1, 1]),       # 8 both previous
    ([row(P5, R, "-1.00", 100001)], [0, 0, 0]),                                            # 9 return previous only
    ([row(Y5, S, "1.00", 100001), row(P5, S, "1.00", 100001, doc="2")], [1, 1, 1]),        # 10 sales both periods
    ([row(Y5, R, "-1.00", 100001), row(P5, S, "1.00", 100001, doc="2")], [0, 1, 1]),       # 11 return y, sale p
    ([row(Y5, S, "0.00", 100001)], [1, 0, 1]),                                             # zero-value sale
    ([row(F5, S, "1.00", 100001), row(P5, R, "-1.00", 100001, doc="2")], [0, 0, 0]),       # 16 future sale only
])
def test_sales_only_eligibility(tmp_path, rows, want):
    s, r = _one_group(tmp_path, rows)
    assert s["الجملة"] == r["سامح السيد محمود احم"] == s[GRAND_TOTAL] == r[GRAND_TOTAL] == want


def test_sales_only_eligibility_all_transactions_mode(tmp_path):
    rows = [row(Y5, R, "-1.00", 100001), row(F5, S, "1.00", 100002, doc="2"),   # 100002: future sale counts here
            row(P5, R, "-1.00", 100003, doc="3"), row(P5, S, "1.00", 100004, doc="4"), row(Y5, R, "-1.00", 100004, doc="5")]
    s, r = _one_group(tmp_path, rows, MODE_ALL)
    assert s["الجملة"] == [1] and s["صيدليات وسلاسل صيدليات"] == [0] and s["مستشفيات"] == [1]
    assert r["سامح السيد محمود احم"] == [0] and r["داليا صلاح الدين فكر"] == [1] and s[GRAND_TOTAL] == [2]


def test_customer_counts_across_groups_and_files(tmp_path):
    """12: a customer in two files counts once; 13: returns in one file, sales in another qualifies;
    14: returns only in every file never counts. Grand total = distinct over the whole report."""
    files = {
        "a.csv": [row(Y5, S, "1.00", 100001), row(Y5, R, "-1.00", 100002, doc="2"), row(Y5, R, "-1.00", 100003, doc="3")],
        "b.csv": [row(P5, S, "1.00", 100001), row(P5, S, "1.00", 100002, doc="2"), row(P5, R, "-1.00", 100003, doc="3"),
                  row(P5, S, "1.00", 100004, doc="4")],
    }
    s, r = _one_group(tmp_path, None, files=files)
    assert s["الجملة"] == [1, 2, 2]                                  # 100001 (both files) + 100002 (sale in b)
    assert s["صيدليات وسلاسل صيدليات"] == [0, 0, 0]                  # 100003: returns only in both files
    assert r["سامح السيد محمود احم"] == [1, 1, 1]                    # 100001 once; 100003 excluded
    assert r["داليا صلاح الدين فكر"] == [0, 1, 1] and r["هند محمد رضا"] == [0, 1, 1]
    assert s[GRAND_TOTAL] == r[GRAND_TOTAL] == [1, 3, 3]           # 100001, 100002, 100004 - not 1+1+1+...


def test_blank_codes_and_unregistered_returns_only_are_not_counted(tmp_path):
    rows = [row(Y5, S, "5.00", "", "بدون كود"), row(Y5, R, "-2.00", 777777, "مرتجع فقط", doc="2"),
            row(P5, S, "1.00", 888888, "زائر", doc="3")]
    s, _ = _one_group(tmp_path, rows)
    assert s[UNREGISTERED] == [0, 1, 1] and s[GRAND_TOTAL] == [0, 1, 1]         # only 888888
    assert _by_label(_sheets(next((tmp_path / "output").rglob(WORKBOOK_NAME)))[SHEET_SECTORS])[UNREGISTERED] == \
        _d("5.00", "1.00", "6.00", "-2.00", "0", "-2.00", "4.00")               # amounts unchanged


def test_independent_recount_detects_a_returns_only_customer(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    assert analysis.reconcile(result) == []
    result.sectors["صيدليات وسلاسل صيدليات"].customers_y.add("100003")        # 100003 has only a return
    result.direct.customers_y.add("100003")
    result.representatives["سامح السيد محمود احم"].customers_y.add("100003")
    problems = analysis.reconcile(result)
    assert any("'صيدليات وسلاسل صيدليات' customer counts [1, 0, 1, 0, 1, 0] differ from the independent recount" in p
               for p in problems)
    assert any(f"'{GRAND_TOTAL}' customer counts [3, 2, 5, 0, 5, 1] differ" in p for p in problems)


def test_reconcile_detects_customer_in_two_rows(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_ALL, None, setup)
    result.representatives["الاداره"].customers_p.add("100001")    # already counted for سامح
    assert any("more than one row" in p for p in analysis.reconcile(result))


# ------------------------------------------------------------------ mapping
def test_customer_key_matches_text_and_numbers():
    assert {customer_key(v) for v in (100418, 100418.0, "100418", " 100418 ", "١٠٠٤١٨", "0100418")} == {"100418"}


def test_adding_missing_customer_to_main_data_classifies_it_next_run(tmp_path, setup):
    first = run_analysis(_config(tmp_path), MODE_ALL, None, setup)
    assert set(first.missing) == {"555555", "666666"}
    updated = load_main_data(write_main(tmp_path / "data" / "main2.xlsx",
                                        [*MAIN, (555555, "زائر جديد", "عملاء اون لاين", "اون لاين")]))
    second = run_analysis(_config(tmp_path), MODE_ALL, None, updated)
    assert set(second.missing) == {"666666"}
    s = _by_label(_sheets(second.report_path)[SHEET_SECTORS])
    assert s["عملاء اون لاين"] == _d("50.00", "-7.25", "42.75") and s[UNREGISTERED] == _d("0.01", "0", "0.01")


def test_main_data_errors(tmp_path):
    with pytest.raises(MainDataError, match="more than once"):
        load_main_data(write_main(tmp_path / "a.xlsx", [*MAIN, (100001, "x", "مستشفيات", "الاداره")]))
    assert len(load_main_data(write_main(tmp_path / "b.xlsx", [*MAIN, MAIN[0]])).customers) == 4  # same row twice
    wb = Workbook()
    wb.active.append(["code", "name", "sector", "rep"])
    wb.save(tmp_path / "c.xlsx")
    with pytest.raises(MainDataError, match="header"):
        load_main_data(tmp_path / "c.xlsx")
    with pytest.raises(MainDataError, match="exactly one"):
        find_main_data(tmp_path)                                                  # 3 workbooks here
    (tmp_path / "one").mkdir()
    write_main(tmp_path / "one" / "m.xlsx")
    (tmp_path / "one" / "~$m.xlsx").write_bytes(b"lock")
    assert find_main_data(tmp_path / "one").name == "m.xlsx"


def test_unknown_main_data_values_get_their_own_row(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row("30-10-2026", S, "5.00", 100009)])
    main_data = load_main_data(write_main(tmp_path / "main.xlsx", [(100009, "x", "قطاع جديد", "مندوب جديد")]))
    result = run_analysis(_config(tmp_path), MODE_ALL, None, main_data)
    s = _by_label(_sheets(result.report_path)[SHEET_SECTORS])
    assert s["قطاع جديد"] == _d("5.00", "0", "5.00") and result.extra_labels == (["قطاع جديد"], ["مندوب جديد"])
    assert analysis.verify_analysis_report(result) == []


# ------------------------------------------------------------------ report structure
def test_structure_and_formatting(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    wb = load_workbook(result.report_path)
    assert wb.sheetnames == [SHEET_SECTORS, SHEET_REPS, SHEET_MISSING] == [
        "تقرير القطاعات", "تقرير المندوبين", "العملاء غير المسجلين"]
    sectors = [r[1] for r in wb[SHEET_SECTORS].iter_rows(min_row=3, values_only=True)]   # date line, header; م first
    assert sectors == [*SECTORS, UNREGISTERED, GRAND_TOTAL] and len(SECTORS) == 8
    reps = [r[1] for r in wb[SHEET_REPS].iter_rows(min_row=3, values_only=True)]
    assert reps == [*GROUP_ONE, "إجمالي المباشر", *GROUP_TWO, "إجمالي الوكلاء", UNREGISTERED,
                    "إجمالي المجموعة الثالثة", GRAND_TOTAL]                     # Cover order and groups
    assert sorted(GROUP_ONE + GROUP_TWO) == sorted(REPORT_REPRESENTATIVES) and len(REPORT_REPRESENTATIVES) == 16
    assert len(REPRESENTATIVES) == 17 and "دينا كرم ابراهيم" not in reps                  # excluded from Task 5
    for ws in wb.worksheets[:3]:                                       # Master Data: its own test
        assert ws.sheet_view.rightToLeft and ws.freeze_panes == "A3"
        assert ws.cell(row=1, column=1).value == "تاريخ التقرير: 30/10/2026"
        header = ws.cell(row=2, column=1)
        assert header.fill.fgColor.rgb.endswith("1F4E78") and header.font.name == "Segoe UI" and header.font.bold
        last = ws.max_row
        assert ws.cell(row=last, column=2).value == GRAND_TOTAL and ws.cell(row=last, column=2).font.bold
        assert ws.auto_filter.ref.endswith(str(last - 1))                         # total row outside the filter
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=last, column=col)
            is_count = ws.cell(row=2, column=col).value in _dated(*YESTERDAY_COUNT_COLUMNS)
            if isinstance(cell.value, (int, float)):
                assert cell.number_format == ("#,##0" if is_count else "#,##0.00")
                assert not str(cell.value).startswith("=")
            assert ws.column_dimensions[cell.column_letter].width >= 12
        assert all(isinstance(v, str) and v.isascii() is False
                   for v in next(ws.iter_rows(min_row=2, max_row=2, values_only=True)))
    assert wb[SHEET_SECTORS].cell(row=4, column=1).fill.fgColor.rgb.endswith("F2F2F2")   # striped rows


def test_no_missing_customers_sheet_has_only_grand_total(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row("30-10-2026", S, "5.00", 100001)])
    result = run_analysis(_config(tmp_path), MODE_ALL, None, load_main_data(write_main(tmp_path / "main.xlsx")))
    assert _sheets(result.report_path)[SHEET_MISSING][1:] == [[GRAND_TOTAL, None, *_d("0", "0", "0")]]
    assert _by_label(_sheets(result.report_path)[SHEET_SECTORS])[UNREGISTERED] == _d("0", "0", "0")


def test_reconcile_detects_mismatch(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_ALL, None, setup)
    result.representatives["الاداره"].add(analysis.Category.SALES, False, D("1.00"))  # counted twice
    problems = analysis.reconcile(result)
    assert any("sector and representative grand totals differ" in p for p in problems)
    assert any("exactly once" in p for p in problems)


# ------------------------------------------------------------------ menu and CLI
def _cli(tmp_path, answers, *extra):
    asked = []

    def fake_input(prompt):
        asked.append(prompt)
        return answers.pop(0)

    code = main(["--input-dir", str(tmp_path / "input"), "--output-dir", str(tmp_path / "out"),
                 "--main-data", str(tmp_path / "data" / "main.xlsx"), *extra], input_fn=fake_input)
    return code, asked


def test_menu_option_1_asks_date(tmp_path, setup, capsys):
    code, asked = _cli(tmp_path, ["5", "1", "31/02/2026", "abc", "30/10/2026"], "--no-move")
    assert code == EXIT_OK
    assert asked == ["Select task number: ", "Select report type: ", *["Enter yesterday's date (DD/MM/YYYY): "] * 3]
    out = capsys.readouterr().out
    assert "5. Sector & Representative Sales Analysis" in out
    assert "1. Yesterday & Previous Days" in out and "2. All Transactions" in out
    assert out.count("التاريخ غير صحيح") == 2 and "Reconciliation           : PASS" in out


def test_menu_option_2_asks_no_date(tmp_path, setup):
    code, asked = _cli(tmp_path, ["5", "3", "x", "2"], "--no-move")
    assert code == EXIT_OK and asked == ["Select task number: ", *["Select report type: "] * 3]
    (report,) = (tmp_path / "out").glob(f"Task5_*/{WORKBOOK_NAME}")
    assert _sheets(report)[SHEET_SECTORS][0] == ["قطاع العملاء", *ALL_COLUMNS, *ALL_COUNT_COLUMNS]


def test_task_argument(tmp_path, setup):
    base = ["--input-dir", str(tmp_path / "input"), "--output-dir", str(tmp_path / "out"),
            "--main-data", str(tmp_path / "data" / "main.xlsx"), "--non-interactive", "--no-move", "--task", "5"]
    assert main([*base, "--ref-date", "30/10/2026"]) == EXIT_OK                    # type 1 implied
    assert main([*base, "--report-type", "2"]) == EXIT_OK
    assert main(base) == EXIT_USAGE                                                 # type required
    assert main([*base, "--report-type", "2", "--ref-date", "30/10/2026"]) == EXIT_USAGE
    assert main([*base, "--report-type", "1"]) == EXIT_USAGE                        # date required
    assert main([*base, "--ref-date", "2026/10/30"]) == EXIT_USAGE                  # invalid date
    assert main([*base, "--start-date", "01-10-2026"]) == EXIT_USAGE
    assert main([*base[:-2], "--task", "3", "--report-type", "2"]) == EXIT_USAGE    # task 5 option on task 3
    assert len(list((tmp_path / "out").iterdir())) == 2


def test_missing_main_data_stops_before_processing(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", ROWS)
    code = main(["--input-dir", str(tmp_path / "input"), "--output-dir", str(tmp_path / "out"), "--task", "5",
                 "--report-type", "2", "--main-data", str(tmp_path / "nope.xlsx"), "--non-interactive"])
    assert code == EXIT_USAGE and not (tmp_path / "out").exists() and not (tmp_path / "processed").exists()
    assert [p.name for p in (tmp_path / "input").iterdir()] == ["m.csv"]


# ------------------------------------------------------------------ lifecycle
def _run(tmp_path, *extra):
    return main(["--input-dir", str(tmp_path / "input"), "--output-dir", str(tmp_path / "out"), "--task", "5",
                 "--report-type", "2", "--main-data", str(tmp_path / "data" / "main.xlsx"), "--non-interactive",
                 *extra])


def _sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_success_moves_files_after_verification(tmp_path, setup):
    write_erp_csv(tmp_path / "input" / "second.csv", [row("30-10-2026", S, "1.00", 100002)])
    original = _sha(tmp_path / "input" / "month.csv")
    main_sha = _sha(tmp_path / "data" / "main.xlsx")
    assert _run(tmp_path) == EXIT_OK
    (report_dir,) = (tmp_path / "out").iterdir()
    (archive,) = (tmp_path / "processed").iterdir()
    assert report_dir.name.startswith("Task5_") and archive.name == report_dir.name
    assert sorted(p.name for p in archive.iterdir()) == ["month.csv", "second.csv"]
    assert _sha(archive / "month.csv") == original and list((tmp_path / "input").iterdir()) == []
    assert _sha(tmp_path / "data" / "main.xlsx") == main_sha                       # Main Data never written
    journal = tmp_path / "state" / "task5_journal.jsonl"
    events = [json.loads(line)["event"] for line in journal.read_text(encoding="utf-8").splitlines()]
    assert events == ["RUN_STARTED", "COMMIT", "MOVED", "MOVED", "RUN_FINISHED"]
    for other in ("processing_journal", "period_journal", "customer_journal", "representative_journal"):
        assert not (tmp_path / "state" / f"{other}.jsonl").exists()


def test_failed_and_rejected_files_stay_in_input(tmp_path, setup):
    (tmp_path / "input" / "empty.csv").write_bytes(b"")
    write_erp_csv(tmp_path / "input" / "bad.csv", [*ROWS[:2], row("xx-10-2026", S, "5.00", 100001, doc="99")])
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert sorted(p.name for p in (tmp_path / "input").iterdir()) == ["bad.csv", "empty.csv"]
    assert [p.name for p in (tmp_path / "processed").rglob("*.csv")] == ["month.csv"]


@pytest.mark.parametrize("target,replacement", [
    ("write_analysis_report", lambda result, path: (_ for _ in ()).throw(OSError("disk full"))),
    ("verify_analysis_report", lambda result: ["sheet missing"]),
    ("reconcile", lambda result: ["sector and representative grand totals differ"]),
])
def test_save_verification_or_reconciliation_failure_moves_nothing(tmp_path, setup, monkeypatch, target,
                                                                    replacement):
    monkeypatch.setattr(analysis, target, replacement)
    original = _sha(tmp_path / "input" / "month.csv")
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "input").iterdir()] == ["month.csv"]
    assert _sha(tmp_path / "input" / "month.csv") == original
    assert not (tmp_path / "processed").exists() or not list((tmp_path / "processed").rglob("*.csv"))


def test_no_move(tmp_path, setup):
    assert _run(tmp_path, "--no-move") == EXIT_OK
    assert [p.name for p in (tmp_path / "input").iterdir()] == ["month.csv"]
    assert not (tmp_path / "processed").exists() and not (tmp_path / "state").exists()


def test_never_overwrites_and_unrelated_files_untouched(tmp_path, setup):
    content = (tmp_path / "input" / "month.csv").read_bytes()
    (tmp_path / "input" / "notes.txt").write_text("keep", encoding="utf-8")
    (tmp_path / "input" / "sub").mkdir()
    (tmp_path / "input" / "sub" / "nested.csv").write_bytes(content)
    assert _run(tmp_path) == EXIT_OK
    (tmp_path / "input" / "month.csv").write_bytes(content)
    assert _run(tmp_path) == EXIT_OK                        # existing policy: re-supplied file counted again
    assert sorted(p.name for p in (tmp_path / "input").iterdir()) == ["notes.txt", "sub"]
    assert [p.name for p in (tmp_path / "processed").rglob("*.csv")] == ["month.csv", "month.csv"]
    assert len(list((tmp_path / "out").iterdir())) == 2


def test_empty_input(tmp_path):
    (tmp_path / "input").mkdir()
    write_main(tmp_path / "data" / "main.xlsx")
    write_fixed(tmp_path / "data" / FIXED_CODES_NAME, [100001])
    assert _run(tmp_path) == EXIT_NO_INPUT and not (tmp_path / "out").exists()


def test_crash_recovery_of_interrupted_run(tmp_path, setup):
    """A run that died before COMMIT: its folder is marked void and the files are processed normally."""
    dead = tmp_path / "out" / "Task5_2026-10-30"
    dead.mkdir(parents=True)
    (tmp_path / "state").mkdir()
    (tmp_path / "state" / "task5_journal.jsonl").write_text(json.dumps(
        {"event": "RUN_STARTED", "run_id": dead.name, "run_dir": str(dead), "task": "task5", "files": []}) + "\n",
        encoding="utf-8")
    assert _run(tmp_path) == EXIT_OK
    assert "VOID" in (dead / "RUN_INTERRUPTED.txt").read_text(encoding="utf-8")
    events = [json.loads(line)["event"]
              for line in (tmp_path / "state" / "task5_journal.jsonl").read_text(encoding="utf-8").splitlines()]
    assert events[:2] == ["RUN_STARTED", "RUN_ABANDONED"] and events[-1] == "RUN_FINISHED"
    assert [p.name for p in (tmp_path / "processed").rglob("*.csv")] == ["month.csv"]


# ------------------------------------------------------------------ fixed code.xlsx: the 224-code split
def test_fixed_columns_follow_the_total_column(tmp_path, setup):
    assert YESTERDAY_COUNT_COLUMNS == ("عدد عملاء أمس", "عدد عملاء الأيام السابقة", "إجمالي العملاء المميزين",
                                       "عدد العملاء من الـ224 كود", "عدد العملاء خارج الـ224 كود",
                                       "عدد عملاء المرتجعات فقط")
    assert ALL_COUNT_COLUMNS == ("إجمالي العملاء المميزين", "عدد العملاء من الـ224 كود", "عدد العملاء خارج الـ224 كود",
                                 "عدد عملاء المرتجعات فقط")
    fixed = load_fixed_codes(tmp_path / "data" / FIXED_CODES_NAME)
    yesterday = _sheets(run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, fixed_codes=fixed).report_path)
    assert yesterday[SHEET_SECTORS][0][10:13] == ["إجمالي العملاء المميزين", *FIXED_COLUMNS]  # columns 11, 12, 13
    every = _sheets(run_analysis(_config(tmp_path), MODE_ALL, None, setup, fixed_codes=fixed).report_path)
    assert every[SHEET_REPS][0][4:7] == ["إجمالي العملاء المميزين", *FIXED_COLUMNS]         # columns 5, 6, 7
    assert every[SHEET_MISSING][0] == list(MISSING_COLUMNS)                                   # third sheet unchanged


def test_load_fixed_codes_ignores_blanks_and_repeats(tmp_path):
    wb = Workbook()
    for v in (100001, None, "100001", 100002, " 100002 ", 100001.0, "", "A12", "٥٥٥٥٥٥"):
        wb.active.append([v, 999999])                                        # column B is never read
    wb.create_sheet("other").append([123456])                                # only the first sheet
    wb.save(tmp_path / FIXED_CODES_NAME)
    assert load_fixed_codes(tmp_path / FIXED_CODES_NAME) == {"100001", "100002", "A12", "555555"}
    with pytest.raises(MainDataError, match="cannot open"):
        load_fixed_codes(tmp_path / "missing.xlsx")
    write_fixed(tmp_path / "empty" / FIXED_CODES_NAME, [None, ""])
    with pytest.raises(MainDataError, match="no codes"):
        load_fixed_codes(tmp_path / "empty" / FIXED_CODES_NAME)


def test_find_main_data_ignores_the_fixed_code_list(tmp_path):
    write_main(tmp_path / "شيت المناديب و القطاعات.xlsx")
    write_fixed(tmp_path / FIXED_CODES_NAME, [100001])
    assert find_main_data(tmp_path).name == "شيت المناديب و القطاعات.xlsx"


@pytest.mark.parametrize("mode", [MODE_YESTERDAY, MODE_ALL])
def test_split_inside_and_outside_the_fixed_codes(tmp_path, setup, mode):
    """Eligible (with sales): 100001, 100002, 555555 (unregistered), 666666 (unregistered).
    The list also has 100003, which has only a return: it is not a customer, so it is never counted."""
    fixed = load_fixed_codes(tmp_path / "data" / FIXED_CODES_NAME)
    result = run_analysis(_config(tmp_path), mode, REF if mode == MODE_YESTERDAY else None, setup, fixed_codes=fixed)
    assert analysis.verify_analysis_report(result) == []
    sheets = _sheets(result.report_path)
    s, r = _split(sheets[SHEET_SECTORS]), _split(sheets[SHEET_REPS])
    assert s["الجملة"] == [1, 1]                         # 100001 in the list, 100002 not
    assert s[UNREGISTERED] == r[UNREGISTERED] == [1, 1]  # 555555 in the list (unregistered too), 666666 not
    assert s["صيدليات وسلاسل صيدليات"] == [0, 0]          # 100003: in the list, returns only
    assert r["سامح السيد محمود احم"] == [1, 0] and r["داليا صلاح الدين فكر"] == [0, 1]
    assert s[GRAND_TOTAL] == r[GRAND_TOTAL] == [2, 2]
    for rows in (sheets[SHEET_SECTORS], sheets[SHEET_REPS]):
        total = rows[0].index(TOTAL_CUSTOMERS_COLUMN)
        for row in rows[1:]:                             # in + out = إجمالي العملاء المميزين, every row
            assert row[total + 1] + row[total + 2] == row[total]


def test_split_counts_each_customer_once_and_grand_total_is_a_union(tmp_path):
    files = {
        "a.csv": [row(Y5, S, "1.00", 100001), row(Y5, S, "1.00", 100001, doc="2"), row(Y5, S, "1.00", 777777, doc="3")],
        "b.csv": [row(P5, S, "1.00", 100001), row(P5, R, "-1.00", 100001, doc="2"), row(P5, S, "1.00", 777777, doc="3"),
                  row(P5, S, "1.00", 100004, doc="4")],
    }
    for name, rows in files.items():
        write_erp_csv(tmp_path / "input" / name, rows)
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, load_main_data(write_main(tmp_path / "main.xlsx")),
                          fixed_codes=frozenset({"100001", "777777"}))
    assert analysis.verify_analysis_report(result) == []
    sheets = _sheets(result.report_path)
    s = _split(sheets[SHEET_SECTORS])
    assert s["الجملة"] == [1, 0] and s[UNREGISTERED] == [1, 0] and s["مستشفيات"] == [0, 1]
    assert _counts(sheets[SHEET_SECTORS])[GRAND_TOTAL] == [2, 3, 3]
    assert s[GRAND_TOTAL] == [2, 1]                      # 100001 + 777777 once each (several rows, two files)


def test_reconcile_detects_a_split_that_does_not_add_up(tmp_path, setup, monkeypatch):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, fixed_codes=frozenset({"100001"}))
    assert analysis.reconcile(result) == []
    original = analysis.Amounts.counts
    monkeypatch.setattr(analysis.Amounts, "counts",
                        lambda self, mode, fixed: [*original(self, mode, fixed)[:-2], 1 + original(self, mode, fixed)[-2],
                                                   original(self, mode, fixed)[-1]])
    problems = analysis.reconcile(result)
    assert any(f"differ from إجمالي العملاء المميزين" in p for p in problems)
    assert any("differ from the independent recount" in p for p in problems)


def test_readback_detects_a_changed_split_cell(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_ALL, None, setup, fixed_codes=frozenset({"100001"}))
    wb = load_workbook(result.report_path)
    wb[SHEET_SECTORS].cell(row=wb[SHEET_SECTORS].max_row, column=7).value += 1   # grand total, in the list
    wb.save(result.report_path)
    assert any("customer count" in p for p in analysis.verify_analysis_report(result))


def test_cli_reads_the_fixed_codes_beside_main_data(tmp_path, setup):
    assert _run(tmp_path, "--no-move") == EXIT_OK
    (report,) = (tmp_path / "out").glob(f"Task5_*/{WORKBOOK_NAME}")
    assert _split(_sheets(report)[SHEET_SECTORS])[GRAND_TOTAL] == [2, 2]   # list: 100001, 555555, 100003


def test_missing_fixed_code_list_stops_before_processing(tmp_path, setup, capsys):
    (tmp_path / "data" / FIXED_CODES_NAME).unlink()
    assert _run(tmp_path) == EXIT_USAGE
    assert FIXED_CODES_NAME in capsys.readouterr().err
    assert not (tmp_path / "out").exists() and not (tmp_path / "processed").exists()
    assert [p.name for p in (tmp_path / "input").iterdir()] == ["month.csv"]


# ------------------------------------------------------------------ عدد عملاء المرتجعات فقط
RET_MAIN = [  # code, name, sector, representative
    (200001, "a", "الجملة", "سامح السيد محمود احم"),             # sales only
    (200002, "b", "الجملة", "داليا صلاح الدين فكر"),              # several return invoices, no sales
    (200003, "c", "مستشفيات", "هند محمد رضا"),                   # return in a.csv, sale in b.csv
    (200004, "d", "مستشفيات", "سامح السيد محمود احم"),           # returns only, in both files
    (200005, "e", "صيدليات وسلاسل صيدليات", "هند محمد رضا"),     # return yesterday, sale after the date
]
RET_FILES = {
    "a.csv": [row(Y5, S, "1.00", 200001), row(Y5, R, "-1.00", 200002, doc="2"), row(P5, R, "-2.00", 200002, doc="3"),
              row(P5, R, "-3.00", 200003, doc="4"), row(P5, R, "-1.00", 200004, doc="5"),
              row(Y5, R, "-1.00", 200005, doc="6"), row(F5, S, "9.00", 200005, doc="7"),
              row(Y5, R, "-4.00", "", "بدون كود", doc="8"),                        # blank code: never a customer
              row(P5, R, "-5.00", 777777, "مرتجع غير مسجل", doc="9")],            # unregistered, returns only
    "b.csv": [row(P5, S, "0.00", 200003), row(Y5, R, "-1.00", 200004, doc="2"), row(P5, R, "-1.00", 200002, doc="3")],
}


def _returns_run(tmp_path, mode, fixed=frozenset()):
    for name, rows in RET_FILES.items():
        write_erp_csv(tmp_path / "input" / name, rows)
    main_data = load_main_data(write_main(tmp_path / "ret_main.xlsx", RET_MAIN))
    result = run_analysis(_config(tmp_path), mode, REF if mode == MODE_YESTERDAY else None, main_data,
                          fixed_codes=fixed)
    assert analysis.verify_analysis_report(result) == []
    return result, _sheets(result.report_path)


def test_returns_only_column_is_last_in_both_summary_sheets(tmp_path, setup):
    for mode in (MODE_YESTERDAY, MODE_ALL):
        sheets = _sheets(run_analysis(_config(tmp_path), mode, REF if mode == MODE_YESTERDAY else None,
                                      setup).report_path)
        for name in (SHEET_SECTORS, SHEET_REPS):
            assert sheets[name][0][-1] == "عدد عملاء المرتجعات فقط"
            columns = (YESTERDAY_COLUMNS + YESTERDAY_COUNT_COLUMNS[:-1] if mode == MODE_YESTERDAY
                       else ALL_COLUMNS + ALL_COUNT_COLUMNS[:-1])
            assert sheets[name][0][:-1] == _dated(sheets[name][0][0], *columns,
                                                  ref=REF if mode == MODE_YESTERDAY else None)
        assert sheets[SHEET_MISSING][0] == list(MISSING_COLUMNS)                        # no count column here


def test_returns_only_on_the_standard_data(tmp_path, setup):
    """100003 (صيدليات / سامح) and 100004 (مستشفيات / هند) have only returns; 100001 and 555555 have both."""
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    sheets = _sheets(result.report_path)
    s, r = _returns_only(sheets[SHEET_SECTORS]), _returns_only(sheets[SHEET_REPS])
    assert s["الجملة"] == 0 and s["صيدليات وسلاسل صيدليات"] == 1 and s["مستشفيات"] == 1 and s[UNREGISTERED] == 0
    assert r["سامح السيد محمود احم"] == 1 and r["داليا صلاح الدين فكر"] == 0 and r["هند محمد رضا"] == 1
    assert s[GRAND_TOTAL] == r[GRAND_TOTAL] == 2
    assert _counts(sheets[SHEET_SECTORS])[GRAND_TOTAL] == [2, 2, 4]                       # existing counts unchanged


def test_returns_only_yesterday_and_previous_days(tmp_path):
    _, sheets = _returns_run(tmp_path, MODE_YESTERDAY)
    s, r = _returns_only(sheets[SHEET_SECTORS]), _returns_only(sheets[SHEET_REPS])
    assert s["الجملة"] == 1                    # 200002 once (4 return invoices in 2 files); 200001 sales only
    assert s["مستشفيات"] == 1                  # 200004 once (both files); 200003 has a sale in the other file
    assert s["صيدليات وسلاسل صيدليات"] == 1    # 200005: its only sale is after the date, so it is out of scope
    assert s[UNREGISTERED] == 1                # 777777; the blank code is not a customer
    assert r["سامح السيد محمود احم"] == 1 and r["داليا صلاح الدين فكر"] == 1 and r["هند محمد رضا"] == 1
    assert s[GRAND_TOTAL] == r[GRAND_TOTAL] == 4                         # distinct: 200002, 200004, 200005, 777777
    assert _counts(sheets[SHEET_SECTORS])["مستشفيات"] == [0, 1, 1]      # 200003 is a (zero-value) sales customer


def test_returns_only_all_transactions(tmp_path):
    _, sheets = _returns_run(tmp_path, MODE_ALL)
    s, r = _returns_only(sheets[SHEET_SECTORS]), _returns_only(sheets[SHEET_REPS])
    assert s["صيدليات وسلاسل صيدليات"] == 0    # every date counts here: 200005's sale makes it a sales customer
    assert s["الجملة"] == 1 and s["مستشفيات"] == 1 and s[UNREGISTERED] == 1
    assert r["هند محمد رضا"] == 0 and r["سامح السيد محمود احم"] == 1
    assert s[GRAND_TOTAL] == r[GRAND_TOTAL] == 3


def test_returns_only_grand_total_is_a_distinct_count(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    assert analysis.reconcile(result) == []
    result.direct.customers_r.add("999999")                      # a returns-only customer no row has
    problems = analysis.reconcile(result)
    assert any("grand total customer counts [2, 2, 4, 0, 4, 3]" in p for p in problems)
    assert any("differ from the independent recount" in p for p in problems)


def test_independent_recount_detects_a_wrong_returns_only_count(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_ALL, None, setup)
    result.sectors["الجملة"].customers_r.add("100002")           # 100002 has a sale: never returns-only
    result.sectors["الجملة"].customers_p.discard("100002")
    problems = analysis.reconcile(result)
    assert any("'الجملة' customer counts" in p and "independent recount" in p for p in problems)


# ------------------------------------------------------------------ Master Data
def _master(path):
    return _sheets(path)[SHEET_MASTER]


def test_master_data_rows_yesterday_mode(tmp_path, setup):
    fixed = load_fixed_codes(tmp_path / "data" / FIXED_CODES_NAME)         # 100001, 555555, 100003
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, fixed_codes=fixed)
    assert analysis.verify_analysis_report(result) == []
    wb = load_workbook(result.reps_report_path, data_only=True)
    assert wb.sheetnames[:2] == [SHEET_COVER, SHEET_MASTER]
    rows = [list(r) for r in wb[SHEET_MASTER].iter_rows(values_only=True)][1:]  # row 1: the report date
    assert [r[0] for r in rows] == ["م", *range(1, len(rows) - 1), None]         # column م: 1, 2, 3, ...; total blank
    rows = [r[1:] for r in rows]
    assert rows[0] == list(MASTER_COLUMNS) == ["التاريخ", "المندوب", "القطاع", "كود العميل", "اسم العميل", "المبيعات",
                                               "الصافي", "المرتجعات", "حالة كود العميل"]
    IN, OUT = "داخل قائمة الـ224 كود", "خارج قائمة الـ224 كود"
    *body, total = rows[1:]
    assert [[r[0].date(), *r[1:5], _num(r[5]), _num(r[6]), _num(r[7]), r[8]] for r in body] == [   # file order
        [date(2026, 10, 30), "سامح السيد محمود احم", "الجملة", "100001", "اسم 100001", D("100.10"), D("100.10"),
         D("0.00"), IN],
        [date(2026, 10, 29), "داليا صلاح الدين فكر", "الجملة", "100002", "اسم 100002", D("200.20"), D("200.20"),
         D("0.00"), OUT],
        [date(2026, 10, 30), "سامح السيد محمود احم", "صيدليات وسلاسل صيدليات", "100003", "اسم 100003", D("0.00"),
         D("-30.00"), D("-30.00"), IN],
        [date(2026, 10, 1), "سامح السيد محمود احم", "الجملة", "100001", "اسم 100001", D("0.00"), D("-5.05"),
         D("-5.05"), IN],
        [date(2026, 10, 30), "غير مسجل", "غير مسجل", "555555", "زائر جديد", D("50.00"), D("50.00"), D("0.00"), IN],
        [date(2026, 10, 15), "غير مسجل", "غير مسجل", "555555", "زائر جديد", D("0.00"), D("-7.25"), D("-7.25"), IN],
        [date(2026, 10, 10), "هند محمد رضا", "مستشفيات", "100004", "اسم 100004", D("0.00"), D("2.00"), D("2.00"),
         OUT],                                                                   # a positive return: sign kept
        [date(2026, 10, 12), "غير مسجل", "غير مسجل", "666666", "عميل مؤقت", D("0.01"), D("0.01"), D("0.00"), OUT],
    ]
    # الإجمالي العام: 3 representatives (سامح, داليا, هند; غير مسجل is not one), 6 customer codes, the sums
    assert [total[0], total[1], total[2], total[3], total[4], *map(_num, total[5:8]), total[8]] == [
        GRAND_TOTAL, 3, None, 6, None, D("350.31"), D("310.01"), D("-40.30"), None]
    assert _by_label(_sheets(result.report_path)[SHEET_SECTORS])[GRAND_TOTAL][-1] == D("310.01")
    assert _by_label(_sheets(result.report_path)[SHEET_SECTORS])[GRAND_TOTAL][5] == D("-40.30")
    assert _by_label(_sheets(result.report_path)[SHEET_SECTORS])[GRAND_TOTAL][2] == D("350.31")


def test_master_data_all_transactions_mode(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_ALL, None, setup)
    rows = _master(result.reps_report_path)
    assert len(rows) - 1 == 9 == result.direct.rows                       # + the 31-10 sale; تسوية still out
    assert any(r[0].date() == date(2026, 10, 31) and r[6] == D("999.99") for r in rows[1:])
    assert {r[8] for r in rows[1:]} == {"خارج قائمة الـ224 كود"}             # no fixed codes given
    assert sum((r[6] for r in rows[1:]), D(0)) == D("1310.00")


def test_master_data_excludes_rejected_rows_and_other_types(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [
        row(Y5, S, "1.00", 100001), row("xx-10-2026", S, "5.00", 100001, doc="2"),       # rejected
        row(Y5, "تسوية", "7.00", 100002, doc="3"), row(Y5, S, "abc", 100002, doc="4")])   # other type; rejected
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, load_main_data(write_main(tmp_path / "main.xlsx")))
    assert [r[3] for r in _master(result.reps_report_path)[1:]] == ["100001"]
    assert analysis.verify_analysis_report(result) == []


def test_master_data_blank_code_unregistered_and_blank_main_data_cells(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [row(Y5, S, "5.00", "", "بدون كود"), row(P5, R, "-2.00", 300001, doc="2")])
    write_erp_csv(tmp_path / "input" / "b.csv", [row(P5, S, "3.00", "0300002", "اسم ب")])   # code with leading zero
    main_data = load_main_data(write_main(tmp_path / "main.xlsx", [(300002, "ب", None, None)]))  # blank sector/rep
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, main_data, fixed_codes=frozenset({"300001"}))
    assert analysis.verify_analysis_report(result) == []
    rows = _master(result.reps_report_path)[1:]
    assert [r[1:5] + [r[8]] for r in rows] == [
        ["غير مسجل", "غير مسجل", None, "بدون كود", "خارج قائمة الـ224 كود"],          # blank code: amount kept
        ["غير مسجل", "غير مسجل", "300001", "اسم 300001", "داخل قائمة الـ224 كود"],     # unregistered, in the list
        ["غير محدد", "غير محدد", "300002", "اسم ب", "خارج قائمة الـ224 كود"],          # Main Data cells blank
    ]
    assert _by_label(_sheets(result.report_path)[SHEET_SECTORS])[None] == _d("0", "3.00", "3.00", "0", "0", "0", "3.00")


def test_master_data_customer_in_several_files(tmp_path):
    result, _ = _returns_run(tmp_path, MODE_YESTERDAY, fixed=frozenset({"200004"}))
    rows = _master(result.reps_report_path)[1:]
    assert len(rows) == 11                                      # 9 + 3 rows, less the sale after the date
    d4 = [r for r in rows if r[3] == "200004"]
    assert len(d4) == 2 and {r[8] for r in d4} == {"داخل قائمة الـ224 كود"}
    assert all(r[1:3] == ["سامح السيد محمود احم", "مستشفيات"] for r in d4)


def test_empty_master_data(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row(F5, S, "1.00", 100001), row(Y5, "تسوية", "2.00", 100001, doc="2")])
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, load_main_data(write_main(tmp_path / "main.xlsx")))
    assert result.master == [] and analysis.verify_analysis_report(result) == []
    ws = load_workbook(result.reps_report_path)[SHEET_MASTER]
    assert ws.max_row == 3 and ws.auto_filter.ref == "A2:J2"                    # date line, header, الإجمالي العام
    assert [ws.cell(row=3, column=c).value for c in range(2, 10)] == [GRAND_TOTAL, 0, None, 0, None, 0, 0, 0]


def test_master_data_formatting(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    ws = load_workbook(result.reps_report_path)[SHEET_MASTER]
    assert ws.sheet_view.rightToLeft and ws.freeze_panes == "A3" and ws.auto_filter.ref == "A2:J10"  # not the total
    header = ws.cell(row=2, column=1)
    assert header.fill.fgColor.rgb.endswith("1F4E78") and header.font.name == "Segoe UI" and header.font.bold
    assert header.font.color.rgb.endswith("FFFFFF")
    assert ws.cell(row=3, column=2).number_format == "DD/MM/YYYY" and ws.cell(row=3, column=8).number_format == "#,##0.00"
    assert ws.cell(row=3, column=5).data_type == "s" and ws.cell(row=3, column=3).font.name == "Segoe UI"
    total = ws.max_row
    assert ws.cell(row=total, column=2).value == GRAND_TOTAL and ws.cell(row=total, column=1).value is None
    assert ws.cell(row=total, column=2).fill.fgColor.rgb.endswith("DDEBF7") and ws.cell(row=total, column=2).font.bold
    assert all(ws.column_dimensions[c].width >= 12 for c in "ABCDEFGHIJ")


def test_reconcile_detects_a_missing_master_row(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    result.master.pop()
    problems = analysis.reconcile(result)
    assert any("Master Data: 7 rows, 8 transactions in the totals, 8 in the source files" in p for p in problems)
    assert any("differs from its Master Data rows" in p for p in problems)


def test_reconcile_detects_a_wrong_master_amount_mapping_or_status(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, fixed_codes=frozenset({"100001"}))
    good = list(result.master)
    result.master[0] = replace(good[0], net=good[0].net + D("1.00"))
    problems = analysis.reconcile(result)
    assert any("Master Data: sales" in p for p in problems) and any("Master Data: net" in p for p in problems)
    result.master[:] = good
    result.master[1] = replace(good[1], representative="هند محمد رضا")
    assert any("representative or sector that differs from Main Data" in p for p in analysis.reconcile(result))
    result.master[:] = good
    result.master[1] = replace(good[1], status="داخل قائمة الـ224 كود")              # 100002 is not in the list
    assert any("224-code status differs" in p for p in analysis.reconcile(result))


def test_readback_detects_a_changed_or_missing_master_data_row(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    wb = load_workbook(result.reps_report_path)
    wb[SHEET_MASTER].cell(row=4, column=8).value = 1.0                                   # wrong الصافي
    wb.save(result.reps_report_path)
    assert any("Master Data: 1 row(s) differ" in p for p in analysis.verify_analysis_report(result))
    wb[SHEET_MASTER].delete_rows(5)
    wb.save(result.reps_report_path)
    assert any("Master Data: 7 rows written, 8 expected" in p for p in analysis.verify_analysis_report(result))


def test_master_data_failure_moves_nothing(tmp_path, setup, monkeypatch):
    original = analysis.master_rows
    monkeypatch.setattr(analysis, "master_rows", lambda result: original(result)[:-1])     # one transaction lost
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "input").iterdir()] == ["month.csv"]
    assert not (tmp_path / "processed").exists() or not list((tmp_path / "processed").rglob("*.csv"))


def test_managed_run_writes_master_data_before_moving(tmp_path, setup):
    assert _run(tmp_path) == EXIT_OK
    (report,) = (tmp_path / "out").glob(f"Task5_*/{WORKBOOK_NAME}")
    (reps,) = (tmp_path / "out").glob("Task5_*/Representatives_Analysis_*.xlsx")
    assert list(_sheets(report)) == [SHEET_SECTORS, SHEET_REPS, SHEET_MISSING]
    assert len(_sheets(reps)[SHEET_MASTER]) - 1 == 9                                     # report type 2
    assert [p.name for p in (tmp_path / "processed").rglob("*.csv")] == ["month.csv"]


# ------------------------------------------------------------------ Representatives_Analysis workbook
GROUP_ONE = ["اون لاين", "الاداره", "هند محمد رضا", "اسلام حمد كمال محمد", "سامح السيد محمود احم",
             "داليا صلاح الدين فكر", "سميرة محمد اسماعيل", "محمود الشيمي", "ندا عبد الناصر طه"]
GROUP_TWO = ["هشام احمد محمد", "ابراهيم سلام", "رامي رأفت يوقيم", "محمد رأفت شحاتة محمد", "احمد محمد علام محمد",
             "احمد محمود موسي", "احمد الحريري"]
REP_SHEETS = [*GROUP_ONE, *GROUP_TWO, "غير مسجل"]   # the representative sheets, in Cover order (no extra names here)
COVER_GROUP_ONE = [*GROUP_ONE[:8], "د هاني", GROUP_ONE[8]]   # Cover: د هاني in المباشر, after محمود الشيمي
COVER_ROWS = [*COVER_GROUP_ONE, "إجمالي المباشر", *GROUP_TWO, "إجمالي الوكلاء",
              "غير مسجل", "إجمالي المجموعة الثالثة", "الإجمالي العام"]
RUN_AT = datetime(2026, 11, 7, 9, 30)


def _cover(sheets):
    """Label -> [إجمالي البيع, إجمالي المرتجع, صافي] as Decimals + [هدف العملاء, داخل, خارج] (the target is an int, or
    the text لا يوجد هدف). The percentage has its own tests (read unrounded from the saved cells)."""
    return {r[0]: [*r[1:4], int(r[4]) if isinstance(r[4], Decimal) else r[4], int(r[5]), int(r[6])]
            for r in sheets[SHEET_COVER][1:]}


def _reps_run(tmp_path, mode, fixed=frozenset({"200003", "777777"})):
    for name, rows in RET_FILES.items():
        write_erp_csv(tmp_path / "input" / name, rows)
    main_data = load_main_data(write_main(tmp_path / "ret_main.xlsx", RET_MAIN))
    result = run_analysis(_config(tmp_path), mode, REF if mode == MODE_YESTERDAY else None, main_data,
                          fixed_codes=fixed, now=RUN_AT)
    assert analysis.verify_analysis_report(result) == []
    return result, _sheets(result.reps_report_path)


def test_one_run_writes_two_workbooks_in_the_same_folder(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, now=RUN_AT)
    assert result.reps_report_path.parent == result.report_path.parent == result.run_dir
    assert sorted(p.name for p in result.run_dir.iterdir()) == sorted(
        [WORKBOOK_NAME, "Representatives_Analysis_2026-11-07.xlsx"])
    assert load_workbook(result.report_path).sheetnames == ["تقرير القطاعات", "تقرير المندوبين", "العملاء غير المسجلين"]
    assert load_workbook(result.reps_report_path).sheetnames == ["Cover", "Master Data", *REP_SHEETS,
                                                                 "Pivot for Last Day"]
    assert analysis.verify_analysis_report(result) == []


def test_each_representative_sheet_holds_exactly_its_master_data_rows(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, now=RUN_AT)
    sheets = _sheets(result.reps_report_path)
    master = [tuple(r) for r in sheets[SHEET_MASTER][1:]]
    seen = []
    for name in REP_SHEETS:
        assert sheets[name][0] == list(MASTER_COLUMNS)
        rows = [tuple(r) for r in sheets[name][1:]]
        assert all(r[1] == name for r in rows)                             # only this representative
        assert rows == [r for r in master if r[1] == name]                 # same rows, same order
        seen += rows
    assert sorted(seen, key=str) == sorted(master, key=str) and len(seen) == len(master) == 8   # each exactly once
    assert len(sheets["سامح السيد محمود احم"]) - 1 == 3 and len(sheets["اون لاين"]) == 1


def test_cover_yesterday_and_previous_days(tmp_path):
    result, sheets = _reps_run(tmp_path, MODE_YESTERDAY)
    assert sheets[SHEET_COVER][0] == list(COVER_COLUMNS) == [
        "المندوب", "إجمالي البيع", "إجمالي المرتجع", "صافي", "هدف العملاء", "عدد العملاء داخل قائمة الـ224 كود",
        "عدد العملاء خارج قائمة الـ224 كود", "نسبة تحقيق هدف العملاء"]
    c = _cover(sheets)
    assert list(c) == COVER_ROWS                          # the two groups, the other rows, their subtotals, total
    # سامح: 200001 sales only (outside the list); 200004 returns only, in both files
    assert c["سامح السيد محمود احم"] == [*_d("1.00", "-2.00", "-1.00"), 38, 0, 1]
    # داليا: 200002 several return invoices in two files, no sale: no sales customer
    assert c["داليا صلاح الدين فكر"] == [*_d("0", "-4.00", "-4.00"), 44, 0, 0]
    # هند: 200003 return in a.csv + zero-value sale in b.csv (inside the list); 200005 sale only after the date
    assert c["هند محمد رضا"] == [*_d("0", "-4.00", "-4.00"), 25, 1, 0]
    # غير مسجل: 777777 returns only (in the list, but not a sales customer); the blank code is never counted
    assert c["غير مسجل"] == [*_d("0", "-9.00", "-9.00"), "لا يوجد هدف", 0, 0]
    assert c["الاداره"] == [*_d("0", "0", "0"), "لا يوجد هدف", 0, 0]
    assert c[GRAND_TOTAL] == [*_d("1.00", "-19.00", "-18.00"), 219, 1, 1]


def test_cover_all_transactions(tmp_path):
    _, sheets = _reps_run(tmp_path, MODE_ALL)
    c = _cover(sheets)
    assert c["هند محمد رضا"] == [*_d("9.00", "-4.00", "5.00"), 25, 1, 1]  # 200005's later sale counts here
    assert c["سامح السيد محمود احم"] == [*_d("1.00", "-2.00", "-1.00"), 38, 0, 1]
    assert c[GRAND_TOTAL] == [*_d("10.00", "-19.00", "-9.00"), 219, 1, 2]
    assert len(sheets[SHEET_MASTER]) - 1 == 12


def test_cover_net_adds_negative_returns(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row(Y5, S, "500,000.00", 100001), row(P5, R, "-25,000.00", 100001, doc="2")])
    result = run_analysis(_config(tmp_path), MODE_ALL, None, load_main_data(write_main(tmp_path / "main.xlsx")))
    assert analysis.verify_analysis_report(result) == []
    c = _cover(_sheets(result.reps_report_path))
    assert c["سامح السيد محمود احم"][:3] == c[GRAND_TOTAL][:3] == _d("500000.00", "-25000.00", "475000.00")


def test_cover_agrees_with_the_report_representative_sheet(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, fixed_codes=frozenset({"100001"}))
    report = _sheets(result.report_path)[SHEET_REPS]
    cover = _cover(_sheets(result.reps_report_path))
    for row in report[1:]:
        if row[0] in ("إجمالي المباشر", "إجمالي الوكلاء", "إجمالي المجموعة الثالثة"):
            continue                                                            # the report's own subtotals
        label = {UNREGISTERED: "غير مسجل"}.get(row[0], row[0])
        assert cover[label][4:] == [int(v) for v in row[-3:-1]]                # inside / outside the 224 list
        assert cover[label][:3] == [row[3], row[6], row[7]]                    # إجمالي المبيعات، المرتجعات، الصافي


def test_representative_sheet_for_a_blank_main_data_cell(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row(P5, S, "3.00", 300002, "اسم ب"), row(Y5, R, "-1.00", 300002, doc="2")])
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF,
                          load_main_data(write_main(tmp_path / "main.xlsx", [(300002, "ب", "الجملة", None)])))
    assert analysis.verify_analysis_report(result) == []
    sheets = _sheets(result.reps_report_path)
    assert list(sheets)[-3:] == ["غير محدد", "غير مسجل", "Pivot for Last Day"]  # extra row, unregistered, pivot
    assert [r[1:4] for r in sheets["غير محدد"][1:]] == [["غير محدد", "الجملة", "300002"]] * 2
    assert _cover(sheets)["غير محدد"] == [*_d("3.00", "-1.00", "2.00"), "لا يوجد هدف", 0, 1]


def test_sheet_names_are_valid_unique_and_keep_the_full_name_in_the_data(tmp_path):
    long_a = "مندوب باسم طويل جدا جدا يتجاوز الحد المسموح أ"
    long_b = "مندوب باسم طويل جدا جدا يتجاوز الحد المسموح ب"            # same first 31 characters
    odd = "فرع/الاسكندرية: [1]*?"
    main = [(400001, "x", "الجملة", long_a), (400002, "y", "الجملة", long_b), (400003, "z", "الجملة", odd),
            (400004, "w", "الجملة", "cover")]
    write_erp_csv(tmp_path / "input" / "m.csv", [row(Y5, S, "1.00", 400001 + i, doc=str(i)) for i in range(4)])
    result = run_analysis(_config(tmp_path), MODE_ALL, None, load_main_data(write_main(tmp_path / "main.xlsx", main)))
    assert analysis.verify_analysis_report(result) == []
    names = load_workbook(result.reps_report_path).sheetnames
    assert len(names) == len({n.casefold() for n in names}) == 2 + 16 + 4 + 1 + 1          # + Pivot for Last Day
    assert all(len(n) <= 31 and not any(ch in n for ch in "[]:*?/\\") for n in names)
    sheets = _sheets(result.reps_report_path)
    assert {long_a, long_b, odd, "cover"} <= set(_cover(sheets))                # full names in Cover
    named = analysis.representative_sheet_names(result)
    for label in (long_a, long_b, odd, "cover"):
        assert [r[1] for r in sheets[named[label]][1:]] == [label]            # full name in the data


def test_empty_master_data_gives_empty_representative_sheets(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row(F5, S, "1.00", 100001), row(Y5, "تسوية", "2.00", 100001, doc="2")])
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, load_main_data(write_main(tmp_path / "main.xlsx")))
    assert analysis.verify_analysis_report(result) == []
    sheets = _sheets(result.reps_report_path)
    assert all(sheets[name] == [list(MASTER_COLUMNS)] for name in [SHEET_MASTER, *REP_SHEETS])
    assert _cover(sheets)[GRAND_TOTAL] == [*_d("0", "0", "0"), 219, 0, 0]


def test_cover_formatting(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    wb = load_workbook(result.reps_report_path)
    ws = wb[SHEET_COVER]
    last = ws.max_row
    assert ws.sheet_view.rightToLeft and ws.freeze_panes == "A3" and ws.auto_filter.ref == f"A2:I{last - 1}"
    header = ws.cell(row=2, column=1)                                            # black and white
    assert header.fill.fgColor.rgb.endswith("D9D9D9") and header.font.bold and header.font.color.rgb.endswith("000000")
    assert ws.cell(row=last, column=2).value == GRAND_TOTAL and ws.cell(row=last, column=2).font.bold
    assert ws.cell(row=last, column=2).border.top.style == "double"              # emphasised total row
    sameh = next(r for r in range(3, last) if ws.cell(row=r, column=2).value == "سامح السيد محمود احم")
    assert [ws.cell(row=sameh, column=c).number_format for c in range(3, 10)] == \
        ["#,##0.00"] * 3 + ["#,##0"] * 3 + ["0.00%"]
    for name in wb.sheetnames[1:-1]:                                            # not Pivot for Last Day
        detail = wb[name]
        assert detail.sheet_view.rightToLeft and detail.freeze_panes == "A3"
        assert detail.auto_filter.ref == f"A2:J{detail.max_row - 1}"            # the transactions, not the total


def test_cover_grand_total_customers_are_recounted_not_summed(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    assert analysis.reconcile(result) == []
    result.direct.customers_p.add("999999")                                     # no representative has it
    assert any("Cover: grand total customers" in p for p in analysis.reconcile(result))
    result.direct.customers_p.discard("999999")
    result.representatives["داليا صلاح الدين فكر"].customers_p.add("100001")    # counted twice
    assert any("Cover: 'داليا صلاح الدين فكر' customers inside / outside" in p for p in analysis.reconcile(result))


def test_corrupted_cover_amount_is_detected(tmp_path, setup, monkeypatch):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, now=RUN_AT)
    wb = load_workbook(result.reps_report_path)
    wb[SHEET_COVER].cell(row=3, column=3).value = 123.45                         # saved إجمالي البيع, first row
    wb.save(result.reps_report_path)
    problems = analysis.verify_analysis_report(result)
    assert any(p.startswith("Representatives_Analysis_2026-11-07.xlsx / Cover: row 3") and "إجمالي البيع" in p
               for p in problems)
    original = analysis.cover_table
    monkeypatch.setattr(analysis, "cover_table",
                        lambda r: [[row[0], row[1] + D("1.00"), *row[2:]] for row in original(r)])
    problems = analysis.reconcile(result)
    assert any("differ from its 3 Master Data rows" in p for p in problems)
    assert any("صافي" in p and "is not sales + returns" in p for p in problems)


def test_missing_representative_transaction_is_detected(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, now=RUN_AT)
    wb = load_workbook(result.reps_report_path)
    wb["سامح السيد محمود احم"].delete_rows(3)                                 # its first transaction
    wb.save(result.reps_report_path)
    problems = analysis.verify_analysis_report(result)
    assert any("/ سامح السيد محمود احم: 2 rows written, 3 expected" in p for p in problems)
    assert any("miss 1 and duplicate 0 Master Data transaction(s)" in p for p in problems)


def test_duplicated_transaction_is_detected(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, now=RUN_AT)
    wb = load_workbook(result.reps_report_path)
    source, target = wb["سامح السيد محمود احم"], wb["داليا صلاح الدين فكر"]
    target.insert_rows(target.max_row)                                          # above its الإجمالي العام
    for c, cell in enumerate(source[3], start=1):                                # the same transaction twice
        target.cell(row=target.max_row - 1, column=c).value = cell.value
    wb.save(result.reps_report_path)
    problems = analysis.verify_analysis_report(result)
    assert any("/ داليا صلاح الدين فكر: 2 rows written, 1 expected" in p for p in problems)
    assert any("miss 0 and duplicate 1 Master Data transaction(s)" in p for p in problems)


def test_corrupted_representatives_workbook_moves_nothing(tmp_path, setup, monkeypatch):
    def broken(result, path):
        path.write_bytes(b"not an xlsx file")
        return path
    monkeypatch.setattr(analysis, "write_reps_report", broken)
    original = _sha(tmp_path / "input" / "month.csv")
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "input").iterdir()] == ["month.csv"]
    assert _sha(tmp_path / "input" / "month.csv") == original
    assert not (tmp_path / "processed").exists() or not list((tmp_path / "processed").rglob("*.csv"))


def test_both_workbooks_are_verified_before_any_file_moves(tmp_path, setup, monkeypatch):
    monkeypatch.setattr(analysis, "verify_reps_report", lambda result: ["Representatives_Analysis: Cover broken"])
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "input").iterdir()] == ["month.csv"]
    monkeypatch.undo()
    assert _run(tmp_path) == EXIT_OK                                            # the same file, both workbooks verified
    reports = sorted(p.name for p in (tmp_path / "out").rglob("*.xlsx"))
    assert len(reports) == 4 and sum(n.startswith("Representatives_Analysis_") for n in reports) == 2
    assert [p.name for p in (tmp_path / "processed").rglob("*.csv")] == ["month.csv"]


# ------------------------------------------------------------------ dynamic date headers
@pytest.mark.parametrize("ref, day, before", [
    (date(2026, 11, 5), "05/11/2026", "04/11/2026"),
    (date(2026, 11, 1), "01/11/2026", "31/10/2026"),     # first day of a month
    (date(2027, 1, 1), "01/01/2027", "31/12/2026"),      # first day of a year
    (date(2028, 3, 1), "01/03/2028", "29/02/2028"),      # after a leap day
])
def test_date_headers(ref, day, before):
    assert analysis.date_headers(YESTERDAY_COLUMNS, ref) == (
        f"مبيعات يوم {day}", f"مبيعات حتى {before}", "إجمالي المبيعات",
        f"مرتجعات يوم {day}", f"مرتجعات حتى {before}", "إجمالي المرتجعات", "الصافي")
    assert analysis.date_headers(YESTERDAY_COUNT_COLUMNS, ref)[:3] == (
        f"عدد عملاء يوم {day}", f"عدد عملاء حتى {before}", "إجمالي العملاء المميزين")
    assert not any("الأيام السابقة" in h for h in analysis.date_headers(YESTERDAY_COLUMNS + YESTERDAY_COUNT_COLUMNS, ref))
    assert analysis.date_headers(ALL_COLUMNS + ALL_COUNT_COLUMNS, None) == ALL_COLUMNS + ALL_COUNT_COLUMNS


def test_report_headers_use_the_selected_reference_date_not_today(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row("05-11-2026", S, "10.00", 100001), row("04-11-2026", R, "-2.00", 100001, doc="2")])
    ref = date(2026, 11, 5)
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, ref, load_main_data(write_main(tmp_path / "main.xlsx")),
                          now=datetime(2027, 3, 15, 8, 0))                     # run months later
    assert analysis.verify_analysis_report(result) == []
    sheets = _sheets(result.report_path)
    want = ["مبيعات يوم 05/11/2026", "مبيعات حتى 04/11/2026", "إجمالي المبيعات",
            "مرتجعات يوم 05/11/2026", "مرتجعات حتى 04/11/2026", "إجمالي المرتجعات", "الصافي",
            "عدد عملاء يوم 05/11/2026", "عدد عملاء حتى 04/11/2026"]
    for name, first in ((SHEET_SECTORS, "قطاع العملاء"), (SHEET_REPS, "المندوب")):
        header = sheets[name][0]
        assert header[0] == first and header[1:10] == want                       # same labels in both sheets
        assert not any("أمس" in h for h in header) and not any("15/03/2027" in h for h in header)
    assert _by_label(sheets[SHEET_SECTORS])["الجملة"] == _d("10.00", "0", "10.00", "0", "-2.00", "-2.00", "8.00")
    assert result.reps_report_path.name == "Representatives_Analysis_2027-03-15.xlsx"   # the file: run date


def test_all_transactions_headers_carry_no_date(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_ALL, None, setup)
    for name in (SHEET_SECTORS, SHEET_REPS):
        header = _sheets(result.report_path)[name][0]
        assert header[1:] == [*ALL_COLUMNS, *ALL_COUNT_COLUMNS]
        assert not any("يوم" in h or "أمس" in h or "حتى" in h for h in header)


# ------------------------------------------------------------------ formatting and printing
def _all_sheets(result):
    return [(path.name, ws) for path in (result.report_path, result.reps_report_path) for ws in load_workbook(path).worksheets]


def test_every_sheet_prints_a4_landscape_one_page_wide_with_repeated_header(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    sheets = _all_sheets(result)
    assert len(sheets) == 3 + 2 + len(REP_SHEETS) + 1                         # + Pivot for Last Day
    for book, ws in sheets:
        where = (book, ws.title)
        assert ws.page_setup.orientation == "landscape", where
        assert int(ws.page_setup.paperSize) == int(ws.PAPERSIZE_A4) == 9, where     # A4
        assert ws.sheet_properties.pageSetUpPr.fitToPage and ws.page_setup.fitToWidth == 1, where
        assert ws.print_title_rows == "$1:$2", where                             # date line + header
        assert ws.print_options.horizontalCentered, where
        last = get_column_letter(ws.max_column)
        assert ws.print_area.endswith(f"$A$1:${last}${ws.max_row}"), where       # every column, every row
        assert 0.25 <= ws.page_margins.left <= 1 and 0.25 <= ws.page_margins.right <= 1, where
        assert ws.sheet_view.rightToLeft and ws.freeze_panes == "A3", where


def test_page_height_rule(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    fit = {(book, ws.title): ws.page_setup.fitToHeight for book, ws in _all_sheets(result)}
    report, reps = result.report_path.name, result.reps_report_path.name
    assert fit[(report, SHEET_SECTORS)] == fit[(report, SHEET_REPS)] == fit[(reps, SHEET_COVER)] == 1  # one page
    assert fit[(reps, SHEET_MASTER)] == 0 and all(fit[(reps, name)] == 0 for name in REP_SHEETS)  # pages down


def test_long_unregistered_list_is_not_squeezed_onto_one_page(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row(Y5, S, "1.00", 900000 + i, doc=str(i)) for i in range(40)])
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, load_main_data(write_main(tmp_path / "main.xlsx")))
    assert analysis.verify_analysis_report(result) == []
    ws = load_workbook(result.report_path)[SHEET_MISSING]
    assert ws.max_row == 43 and ws.page_setup.fitToHeight == 0 and ws.page_setup.fitToWidth == 1


def test_header_and_body_fonts(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    for book, ws in _all_sheets(result):
        if ws.title in ("Pivot for Last Day", SHEET_COVER):                       # black and white: own tests
            continue
        cover = False
        assert ws.row_dimensions[2].height >= 36, (book, ws.title)
        for col in range(1, ws.max_column + 1):
            h = ws.cell(row=2, column=col)
            assert h.font.name == "Segoe UI" and h.font.bold and h.font.size >= 12 and h.font.color.rgb.endswith("FFFFFF")
            assert h.fill.fgColor.rgb.endswith("1F4E78")
            assert (h.alignment.horizontal, h.alignment.vertical, h.alignment.wrap_text) == ("center", "center", True)
            assert ws.column_dimensions[get_column_letter(col)].width >= 14
        for r in range(3, ws.max_row + 1):
            assert ws.row_dimensions[r].height >= (34 if cover else 22), (book, ws.title, r)
            for col in range(1, ws.max_column + 1):
                c = ws.cell(row=r, column=col)
                assert c.font.name == "Segoe UI" and c.font.bold and c.font.size >= (13 if cover else 11)
                assert (c.alignment.horizontal, c.alignment.vertical, c.alignment.wrap_text) == ("center", "center", True)


def test_cover_print_layout(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup)
    ws = load_workbook(result.reps_report_path)[SHEET_COVER]
    last = ws.max_row
    assert ws.page_setup.fitToHeight == 1 and ws.page_setup.fitToWidth == 1 and ws.page_setup.orientation == "landscape"
    assert int(ws.page_setup.paperSize) == 9 and ws.page_setup.blackAndWhite is True
    assert ws.print_area.endswith(f"$A$1:$I${last}") and ws.print_title_rows == "$1:$2"
    assert ws.print_options.horizontalCentered and 0.25 <= ws.page_margins.left <= 0.5
    assert ws.cell(row=2, column=1).font.size == 13 and ws.cell(row=3, column=1).font.size == 12
    assert ws.cell(row=last, column=2).value == GRAND_TOTAL and ws.cell(row=last, column=2).font.size == 14
    assert ws.row_dimensions[last].height >= 38 and ws.cell(row=last, column=2).fill.fill_type is None
    for r in range(2, last + 1):                                                 # black text, white / grey only
        for c in range(1, 10):
            cell = ws.cell(row=r, column=c)
            assert cell.font.color.rgb.endswith("000000") and cell.font.bold and cell.font.name == "Segoe UI"
            assert cell.fill.fill_type is None or cell.fill.fgColor.rgb[-6:] in ("F2F2F2", "D9D9D9")
            assert cell.border.left.color.rgb.endswith("000000")


def test_long_text_wraps_into_a_taller_row(tmp_path):
    long_name = "اسم عميل طويل جدا " * 6
    write_erp_csv(tmp_path / "input" / "m.csv", [row(Y5, S, "1.00", 900001, long_name.strip()), row(Y5, S, "1.00", 100001, doc="2")])
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, load_main_data(write_main(tmp_path / "main.xlsx")))
    ws = load_workbook(result.reps_report_path)[SHEET_MASTER]
    assert ws.column_dimensions["F"].width <= 50                                 # capped, so the name wraps
    assert ws.row_dimensions[3].height > ws.row_dimensions[4].height            # the long name gets more lines
    assert ws.cell(row=3, column=6).alignment.wrap_text and ws.cell(row=3, column=8).number_format == "#,##0.00"
    assert ws.cell(row=3, column=2).number_format == "DD/MM/YYYY"


def test_formatting_changes_no_value(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, fixed_codes=frozenset({"100001"}))
    assert analysis.verify_analysis_report(result) == []                         # every saved cell = the tables
    s = _by_label(_sheets(result.report_path)[SHEET_SECTORS])
    assert s[GRAND_TOTAL] == _d("150.10", "200.21", "350.31", "-30.00", "-10.30", "-40.30", "310.01")
    assert _counts(_sheets(result.report_path)[SHEET_SECTORS])[GRAND_TOTAL] == [2, 2, 4]


# ------------------------------------------------------------------ column م, Cover groups, excluded representative
DINA = "دينا كرم ابراهيم"
GROUP_MAIN = [
    (500001, "a", "الجملة", "اون لاين"),
    (500002, "b", "الجملة", "هند محمد رضا"),
    (500003, "c", "مستشفيات", "هند محمد رضا"),
    (500004, "d", "الجملة", "هشام احمد محمد"),
    (500005, "e", "الجملة", "احمد الحريري"),
    (500006, "f", "الجملة", "مندوب جديد"),          # not in the two groups: the third group
    (500007, "g", "الجملة", DINA),                  # excluded from every output
]
GROUP_ROWS = [
    row(Y5, S, "100.00", 500001), row(Y5, R, "-10.00", 500001, doc="2"),
    row(P5, S, "50.00", 500002, doc="3"), row(P5, R, "-5.00", 500003, doc="4"),
    row(Y5, S, "200.00", 500004, doc="5"), row(P5, R, "-20.00", 500004, doc="6"),
    row(Y5, R, "-7.00", 500005, doc="7"), row(P5, S, "30.00", 500006, doc="8"),
    row(Y5, S, "1000.00", 500007, doc="9"), row(Y5, R, "-100.00", 500007, doc="10"),
    row(P5, S, "9.00", 888888, "زائر", doc="11"),
    row("xx-10-2026", S, "5.00", 500001, doc="12"), row(Y5, "تسوية", "4.00", 500001, doc="13"),  # rejected, other type
]
GROUPED_COVER = [*COVER_GROUP_ONE, "إجمالي المباشر", *GROUP_TWO, "إجمالي الوكلاء",
                 "مندوب جديد", "غير مسجل", "إجمالي المجموعة الثالثة", "الإجمالي العام"]


def _group_run(tmp_path, mode=MODE_YESTERDAY):
    write_erp_csv(tmp_path / "input" / "g.csv", GROUP_ROWS)
    main_data = load_main_data(write_main(tmp_path / "group_main.xlsx", GROUP_MAIN))
    result = run_analysis(_config(tmp_path), mode, REF if mode == MODE_YESTERDAY else None, main_data,
                          fixed_codes=frozenset({"500001", "500004", "888888"}), now=RUN_AT)
    assert analysis.verify_analysis_report(result) == []
    return result


def _raw(path):
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in load_workbook(path).worksheets}


def test_cover_groups_match_the_representative_list_exactly():
    assert [list(g) for g in analysis.COVER_GROUPS] == [GROUP_ONE, GROUP_TWO]
    named = GROUP_ONE + GROUP_TWO
    assert len(named) == len(set(named)) == 16 and analysis.EXCLUDED_REPRESENTATIVE == DINA
    assert sorted(named + [DINA]) == sorted(REPRESENTATIVES)                      # each name matches once
    assert DINA not in REPORT_REPRESENTATIVES and len(REPORT_REPRESENTATIVES) == 16


def test_cover_groups_subtotals_and_grand_total(tmp_path):
    c = _cover(_sheets(_group_run(tmp_path).reps_report_path))
    assert list(c) == GROUPED_COVER
    assert c["اون لاين"] == [*_d("100.00", "-10.00", "90.00"), "لا يوجد هدف", 1, 0]
    assert c["هند محمد رضا"] == [*_d("50.00", "-5.00", "45.00"), 25, 0, 1]
    assert c["الاداره"] == [*_d("0", "0", "0"), "لا يوجد هدف", 0, 0]                    # no transactions: zeros, still shown
    assert c["د هاني"] == [*_d("0", "0", "0"), 3, 0, 0]                          # in المباشر, also without data
    assert c["إجمالي المباشر"] == [*_d("150.00", "-15.00", "135.00"), 206, 1, 1]
    assert c["هشام احمد محمد"] == [*_d("200.00", "-20.00", "180.00"), 2, 1, 0]
    assert c["احمد الحريري"] == [*_d("0", "-7.00", "-7.00"), 1, 0, 0]
    assert c["إجمالي الوكلاء"] == [*_d("200.00", "-27.00", "173.00"), 13, 1, 0]
    assert c["مندوب جديد"] == [*_d("30.00", "0", "30.00"), "لا يوجد هدف", 0, 1]            # extra name kept
    assert c["غير مسجل"] == [*_d("9.00", "0", "9.00"), "لا يوجد هدف", 1, 0]
    assert c["إجمالي المجموعة الثالثة"] == [*_d("39.00", "0", "39.00"), "لا يوجد هدف", 1, 1]
    assert c[GRAND_TOTAL] == [*_d("389.00", "-42.00", "347.00"), 219, 3, 2]     # دينا's 1000 / -100 not in it


def test_cover_groups_all_transactions(tmp_path):
    c = _cover(_sheets(_group_run(tmp_path, MODE_ALL).reps_report_path))
    assert list(c) == GROUPED_COVER
    assert c[GRAND_TOTAL] == [*_d("389.00", "-42.00", "347.00"), 219, 3, 2]


def test_group_subtotal_counts_distinct_customers_of_the_group(tmp_path):
    result = _group_run(tmp_path)
    result.representatives["اون لاين"].customers_p.add("500002")                  # 500002 under two reps of group 1
    rows = {r[0]: r for r in analysis.cover_table(result)}
    assert rows["اون لاين"][6] + rows["هند محمد رضا"][6] == 2                     # outside the list: 500002 twice
    assert rows["إجمالي المباشر"][6] == 1                                         # a union, not 1 + 1
    assert any("'اون لاين' customers inside / outside" in p for p in analysis.reconcile(result))


def test_excluded_representative_is_absent_from_both_workbooks(tmp_path):
    result = _group_run(tmp_path)
    assert result.excluded_rows == 2 and result.excluded_net == D("900.00")
    for path in (result.report_path, result.reps_report_path):
        for title, rows in _raw(path).items():
            assert DINA not in title
            assert not any(isinstance(v, str) and "دينا" in v for r in rows for v in r), (path.name, title)
    assert all(r.code != "500007" for r in result.master)
    report = _sheets(result.report_path)
    assert DINA not in _by_label(report[SHEET_REPS])
    assert _by_label(report[SHEET_SECTORS])["الجملة"] == _d("300.00", "80.00", "380.00", "-17.00", "-20.00", "-37.00",
                                                            "343.00")            # her 1000 / -100 left out
    assert _by_label(report[SHEET_SECTORS])[GRAND_TOTAL][-1] == D("347.00") == _cover(_sheets(result.reps_report_path))[GRAND_TOTAL][2]
    sheets = load_workbook(result.reps_report_path).sheetnames
    assert sheets == ["Cover", "Master Data", *GROUP_ONE, *GROUP_TWO, "مندوب جديد", "غير مسجل", "Pivot for Last Day"]


def test_reconcile_detects_an_excluded_transaction(tmp_path):
    result = _group_run(tmp_path)
    good = result.master[0]
    result.master.append(replace(good, representative=DINA, rep_row=DINA))
    problems = analysis.reconcile(result)
    assert any(f"rows of {DINA}, who is excluded" in p for p in problems)


def test_serial_column_in_every_sheet(tmp_path):
    result = _group_run(tmp_path)
    for path in (result.report_path, result.reps_report_path):
        for title, rows in _raw(path).items():
            if title == "Pivot for Last Day":                                    # no م: its own tests
                continue
            rows = rows[1:]                                                      # row 1: the report date
            assert rows[0][0] == "م", (path.name, title)
            labels = [r[1] for r in rows[1:]]
            totals = {i for i, label in enumerate(labels)
                      if label in (GRAND_TOTAL, *analysis.SUBTOTAL_LABELS, *analysis.REP_SUBTOTAL_LABELS)}
            assert labels[-1] == GRAND_TOTAL                                     # every table ends with it
            want, n = [], 0
            for i in range(len(labels)):
                n += i not in totals
                want.append(None if i in totals else n)
            assert [r[0] for r in rows[1:]] == want, (path.name, title)      # 1, 2, 3 ...; blank on total rows
    cover = _raw(result.reps_report_path)[SHEET_COVER][1:]
    assert [r[0] for r in cover[1:]][:11] == [*range(1, 11), None]             # المباشر (with د هاني), its subtotal
    assert [r[0] for r in cover[1:]][-4:] == [18, 19, None, None]               # continuous across groups


def test_serial_numbers_after_exclusion_and_rejected_rows(tmp_path):
    result = _group_run(tmp_path)
    master = _raw(result.reps_report_path)[SHEET_MASTER][1:]
    assert [r[0] for r in master[1:]] == [*range(1, 10), None]                 # 13 rows: 2 excluded, 2 not counted
    hind = _raw(result.reps_report_path)["هند محمد رضا"][1:]
    assert [r[0] for r in hind[1:]] == [1, 2, None] and {r[4] for r in hind[1:-1]} == {"500002", "500003"}


def test_serial_column_formatting_and_print_area(tmp_path):
    result = _group_run(tmp_path)
    for path in (result.report_path, result.reps_report_path):
        for ws in load_workbook(path).worksheets:
            if ws.title == "Pivot for Last Day":                                  # no م: its own tests
                continue
            assert ws.cell(row=2, column=1).value == "م" and ws.column_dimensions["A"].width >= 14
            a = ws.cell(row=3, column=1)                                         # a data row or الإجمالي العام
            assert a.font.name == "Segoe UI" and a.font.bold and a.alignment.horizontal == "center"
            assert ws.print_area.endswith(f"$A$1:${get_column_letter(ws.max_column)}${ws.max_row}")
            assert ws.auto_filter.ref.startswith("A2:") and ws.freeze_panes == "A3"
    cover = load_workbook(result.reps_report_path)[SHEET_COVER]
    labels = {cover.cell(row=r, column=2).value: r for r in range(3, cover.max_row + 1)}
    sub, total = labels["إجمالي المباشر"], labels[GRAND_TOTAL]
    assert cover.cell(row=sub, column=2).fill.fgColor.rgb.endswith("D9D9D9")         # subtotal: grey, bold
    assert cover.cell(row=sub, column=2).border.top.style == "medium"
    assert cover.cell(row=total, column=2).border.top.style == "double"              # grand total: most prominent
    assert cover.cell(row=total, column=2).font.size > cover.cell(row=sub, column=2).font.size
    assert cover.row_dimensions[total].height > cover.row_dimensions[sub].height >= 34


def test_saved_cover_subtotal_serial_and_excluded_name_are_verified(tmp_path):
    result = _group_run(tmp_path)
    wb = load_workbook(result.reps_report_path)
    ws = wb[SHEET_COVER]
    sub = next(r for r in range(3, ws.max_row + 1) if ws.cell(row=r, column=2).value == "إجمالي المباشر")
    ws.cell(row=sub, column=3).value = 1.0                                        # its إجمالي البيع
    ws.cell(row=4, column=1).value = 7                                            # serial number
    ws.cell(row=5, column=2).value = DINA                                         # a name that must not appear
    wb.save(result.reps_report_path)
    problems = analysis.verify_analysis_report(result)
    assert any("saved 'إجمالي المباشر' amounts differ from the sum of its group's rows" in p for p in problems)
    assert any("column م is not 1, 2, 3" in p for p in problems)
    assert any(f"{DINA} appears although she is excluded" in p for p in problems)


def test_reconcile_detects_a_wrong_group_subtotal(tmp_path, monkeypatch):
    result = _group_run(tmp_path)
    original = analysis._merged
    monkeypatch.setattr(analysis, "_merged", lambda parts: original(list(parts)[:-1] or list(parts)))
    problems = analysis.reconcile(result)
    assert any("'إجمالي المجموعة الثالثة' amounts" in p for p in problems)      # غير مسجل dropped from its subtotal
    assert any("'إجمالي المجموعة الثالثة' customer counts" in p for p in problems)


# ------------------------------------------------------------------ Master Data column المرتجعات
RETURNS_ROWS = [  # the specification's example (سامح) + a returns-only customer with two return invoices (داليا)
    row(Y5, S, "1,000.00", 100001), row(Y5, R, "-300.00", 100001, doc="2"),
    row(P5, S, "500.00", 100003, doc="3"), row(P5, R, "-100.00", 100003, doc="4"),
    row(Y5, R, "-50.00", 100002, doc="5"), row(P5, R, "-25.00", 100002, doc="6"),
]


def _returns_column_run(tmp_path, rows=RETURNS_ROWS):
    write_erp_csv(tmp_path / "input" / "r.csv", rows)
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, load_main_data(write_main(tmp_path / "main.xlsx")),
                          now=RUN_AT)
    assert analysis.verify_analysis_report(result) == []
    return result, _sheets(result.reps_report_path)


def test_returns_column_follows_net_in_master_data_and_every_representative_sheet(tmp_path):
    _, sheets = _returns_column_run(tmp_path)
    header = ["التاريخ", "المندوب", "القطاع", "كود العميل", "اسم العميل", "المبيعات", "الصافي", "المرتجعات",
              "حالة كود العميل"]
    for name in [SHEET_MASTER, *REP_SHEETS]:                                     # empty sheets included
        assert sheets[name][0] == header, name
    raw = load_workbook(_returns_column_run(tmp_path / "again")[0].reps_report_path)[SHEET_MASTER]
    assert [raw.cell(row=2, column=c).value for c in range(1, 11)] == ["م", *header]


def test_returns_column_values_and_totals(tmp_path):
    result, sheets = _returns_column_run(tmp_path)
    rows = sheets[SHEET_MASTER][1:]
    assert [(r[5], r[6], r[7]) for r in rows] == [                              # المبيعات، الصافي، المرتجعات
        (D("1000.00"), D("1000.00"), D("0.00")), (D("0.00"), D("-300.00"), D("-300.00")),   # sale; return
        (D("500.00"), D("500.00"), D("0.00")), (D("0.00"), D("-100.00"), D("-100.00")),
        (D("0.00"), D("-50.00"), D("-50.00")), (D("0.00"), D("-25.00"), D("-25.00"))]       # returns only
    sales, net, returns = (sum(r[c] for r in rows) for c in (5, 6, 7))
    assert net == D("1025.00") and returns == D("-475.00") and net == sales + returns == D("1500.00") + returns
    assert (result.direct.net, result.direct.returns, result.direct.sales) == (D("1025.00"), D("-475.00"), D("1500.00"))
    sameh = sheets["سامح السيد محمود احم"][1:]                                   # the specification's table
    assert [r[7] for r in sameh] == _d("0", "-300", "0", "-100")
    assert sum(r[6] for r in sameh) == D("1100.00") and sum(r[7] for r in sameh) == D("-400.00")
    assert sum(r[5] for r in sameh) == D("1500.00")                              # 1,100 = 1,500 + (-400)
    dalia = sheets["داليا صلاح الدين فكر"][1:]                                   # one customer, two return invoices
    assert [r[7] for r in dalia] == [r[6] for r in dalia] == _d("-50", "-25")


def test_returns_column_sales_only_and_empty_sheets(tmp_path):
    _, sheets = _returns_column_run(tmp_path, [row(Y5, S, "10.00", 100001), row(P5, S, "0.00", 100001, doc="2")])
    assert [r[7] for r in sheets[SHEET_MASTER][1:]] == _d("0", "0")
    assert all(sheets[name] == [sheets[SHEET_MASTER][0]] for name in REP_SHEETS if name != "سامح السيد محمود احم")


def test_each_representative_sheet_reconciles_with_master_data_including_returns(tmp_path):
    _, sheets = _returns_column_run(tmp_path)
    master = sheets[SHEET_MASTER][1:]
    for name in REP_SHEETS:
        mine = [r for r in master if r[1] == name]
        assert sheets[name][1:] == mine
        assert sum((r[7] for r in sheets[name][1:]), D(0)) == sum((r[7] for r in mine), D(0))


def test_returns_column_formatting(tmp_path):
    result, _ = _returns_column_run(tmp_path)
    for ws in load_workbook(result.reps_report_path).worksheets[1:-1]:          # Master Data + representative sheets
        h = ws.cell(row=2, column=9)
        assert h.value == "المرتجعات" and h.font.name == "Segoe UI" and h.font.bold and h.fill.fgColor.rgb.endswith("1F4E78")
        assert ws.auto_filter.ref.startswith("A2:J") and ws.print_area.endswith(f"$A$1:$J${ws.max_row}")
        assert ws.freeze_panes == "A3" and ws.sheet_view.rightToLeft
        for r in range(3, ws.max_row + 1):                                       # the transactions and the total
            assert ws.cell(row=r, column=9).number_format == "#,##0.00" and ws.cell(row=r, column=9).font.bold


def test_wrong_returns_column_is_detected(tmp_path, monkeypatch):
    result, _ = _returns_column_run(tmp_path)
    wb = load_workbook(result.reps_report_path)
    wb[SHEET_MASTER].cell(row=3, column=9).value = -1000.0                        # a sale's المرتجعات
    wb.save(result.reps_report_path)
    problems = analysis.verify_analysis_report(result)
    assert any("Master Data: 1 row(s) differ" in p for p in problems)
    assert any("saved المبيعات 1500 / الصافي 1025 / المرتجعات -1475 differ from the report sales 1500.00 / net 1025.00"
               " / returns -475.00" in p for p in problems)
    monkeypatch.setattr(analysis.MasterRow, "returns", property(lambda self: -abs(self.net)))   # sales as returns
    assert any("column المرتجعات" in p for p in analysis.reconcile(result))


# ------------------------------------------------------------------ report groups, detail totals, report date line
REPORT_GROUPED = [*GROUP_ONE, "إجمالي المباشر", *GROUP_TWO, "إجمالي الوكلاء", "مندوب جديد", UNREGISTERED,
                  "إجمالي المجموعة الثالثة", GRAND_TOTAL]


def test_report_representatives_follow_the_cover_groups_with_subtotals(tmp_path):
    result = _group_run(tmp_path)
    reps = _sheets(result.report_path)[SHEET_REPS]
    assert [r[0] for r in reps[1:]] == REPORT_GROUPED
    rows, counts = _by_label(reps), _counts(reps)
    split = {r[0]: [int(v) for v in r[-3:]] for r in reps[1:]}                  # in, out, returns only
    assert rows["إجمالي المباشر"] == _d("100", "50", "150", "-10", "-5", "-15", "135")
    assert counts["إجمالي المباشر"] == [1, 1, 2] and split["إجمالي المباشر"] == [1, 1, 1]
    assert rows["إجمالي الوكلاء"] == _d("200", "0", "200", "-7", "-20", "-27", "173")
    assert counts["إجمالي الوكلاء"] == [1, 0, 1] and split["إجمالي الوكلاء"] == [1, 0, 1]
    assert rows["إجمالي المجموعة الثالثة"] == _d("0", "39", "39", "0", "0", "0", "39")
    assert counts["إجمالي المجموعة الثالثة"] == [0, 2, 2] and split["إجمالي المجموعة الثالثة"] == [1, 1, 0]
    assert rows[GRAND_TOTAL] == _d("300", "89", "389", "-17", "-25", "-42", "347")
    assert counts[GRAND_TOTAL] == [2, 3, 5] and DINA not in rows                  # distinct over the report


def test_third_report_subtotal_only_when_its_group_has_transactions(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row(Y5, S, "1.00", 100001)])
    result = run_analysis(_config(tmp_path), MODE_ALL, None, load_main_data(write_main(tmp_path / "main.xlsx")))
    assert analysis.verify_analysis_report(result) == []
    labels = [r[0] for r in _sheets(result.report_path)[SHEET_REPS][1:]]
    assert labels == [*GROUP_ONE, "إجمالي المباشر", *GROUP_TWO, "إجمالي الوكلاء", UNREGISTERED, GRAND_TOTAL]


def test_report_subtotal_counts_distinct_customers_and_are_reconciled(tmp_path, monkeypatch):
    result = _group_run(tmp_path)
    result.representatives["اون لاين"].customers_p.add("500002")                  # 500002 under two reps of a group
    rows = {r[0]: r for r in analysis.representative_display_table(result)}
    width = len(result.columns)
    assert rows["اون لاين"][1 + width + 2] + rows["هند محمد رضا"][1 + width + 2] == 3
    assert rows["إجمالي المباشر"][1 + width + 2] == 2                             # a union, not 1 + 2
    result.representatives["اون لاين"].customers_p.discard("500002")
    original = analysis._merged
    monkeypatch.setattr(analysis, "_merged", lambda parts: original(list(parts)[:-1] or list(parts)))
    assert any("تقرير المندوبين: 'إجمالي الوكلاء' amounts" in p for p in analysis.reconcile(result))


def test_saved_report_subtotal_is_verified(tmp_path):
    result = _group_run(tmp_path)
    wb = load_workbook(result.report_path)
    ws = wb[SHEET_REPS]
    sub = next(r for r in range(3, ws.max_row + 1) if ws.cell(row=r, column=2).value == "إجمالي المباشر")
    assert ws.cell(row=sub, column=1).value is None and ws.cell(row=sub, column=2).fill.fgColor.rgb.endswith("E7E6E6")
    ws.cell(row=sub, column=5).value = 1.0                                        # its إجمالي المبيعات
    wb.save(result.report_path)
    problems = analysis.verify_analysis_report(result)
    assert any("saved 'إجمالي المباشر' amounts differ from the sum of its group's rows" in p for p in problems)


def test_every_representative_sheet_ends_with_a_grand_total(tmp_path):
    result = _group_run(tmp_path)
    raw = _raw(result.reps_report_path)
    names = list(raw)[2:-1]                                                     # not Pivot for Last Day
    assert names == [*GROUP_ONE, *GROUP_TWO, "مندوب جديد", "غير مسجل"] and list(raw)[-1] == "Pivot for Last Day"
    for name in names:
        total = raw[name][-1]
        assert total[0] is None and total[1] == GRAND_TOTAL, name                 # no serial number
    hind = raw["هند محمد رضا"][-1]                                                # 1 representative, 2 customers
    assert [hind[2], hind[4], *map(_num, hind[6:9])] == [1, 2, *_d("50", "45", "-5")]
    empty = raw["الاداره"]                                                        # no transactions: still a total
    assert len(empty) == 3 and [empty[-1][c] for c in (2, 4, 6, 7, 8)] == [0, 0, 0, 0, 0]
    unregistered = raw["غير مسجل"][-1]
    assert unregistered[2] == 0 and unregistered[4] == 1                           # غير مسجل is no representative


def test_master_data_grand_total(tmp_path):
    result = _group_run(tmp_path)
    total = _raw(result.reps_report_path)[SHEET_MASTER][-1]
    assert total[0] is None and total[1] == GRAND_TOTAL
    # 5 representatives (اون لاين, هند, هشام, احمد الحريري, مندوب جديد), 7 customer codes (500001-500006, 888888)
    assert [total[2], total[4]] == [5, 7]
    sales, net, returns = map(_num, total[6:9])
    assert (sales, net, returns) == tuple(_d("389", "347", "-42")) and net == sales + returns
    assert (result.direct.sales, result.direct.net, result.direct.returns) == (sales, net, returns)


def test_master_data_grand_total_ignores_blank_codes_and_representatives(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [row(Y5, S, "5.00", "", "بدون كود"), row(P5, R, "-2.00", 300001, doc="2"),
                                                row(P5, S, "3.00", 300002, "اسم ب", doc="3")])
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF,
                          load_main_data(write_main(tmp_path / "main.xlsx", [(300002, "ب", "الجملة", None)])))
    assert analysis.verify_analysis_report(result) == []
    total = _raw(result.reps_report_path)[SHEET_MASTER][-1]
    assert [total[2], total[4], *map(_num, total[6:9])] == [0, 2, *_d("8", "6", "-2")]   # غير مسجل / غير محدد: no rep


def test_wrong_detail_totals_are_detected(tmp_path, monkeypatch):
    result = _group_run(tmp_path)
    wb = load_workbook(result.reps_report_path)
    wb[SHEET_MASTER].cell(row=wb[SHEET_MASTER].max_row, column=5).value = 99      # distinct customers
    wb["هند محمد رضا"].cell(row=wb["هند محمد رضا"].max_row, column=8).value = 1.0   # its net
    wb.save(result.reps_report_path)
    problems = analysis.verify_analysis_report(result)
    assert any("/ Master Data: الإجمالي العام" in p for p in problems)
    assert any("/ هند محمد رضا: saved الإجمالي العام does not match its saved rows" in p for p in problems)
    monkeypatch.setattr(analysis, "detail_total", lambda rows: [GRAND_TOTAL, 0, "", 0, "", D(0), D(0), D(0), ""])
    assert any("Master Data: الإجمالي العام" in p and "differs from the source rows" in p
               for p in analysis.reconcile(result))


def test_report_date_line_on_every_sheet(tmp_path, setup):
    for mode, ref, want in ((MODE_YESTERDAY, REF, "تاريخ التقرير: 30/10/2026"),
                            (MODE_ALL, None, "تاريخ التقرير: 07/11/2026")):    # all transactions: the run date
        write_erp_csv(tmp_path / mode / "input" / "month.csv", ROWS)
        result = run_analysis(PipelineConfig(input_dir=tmp_path / mode / "input", output_dir=tmp_path / mode / "out",
                                             reference_date=REF), mode, ref, setup, now=RUN_AT)
        assert analysis.verify_analysis_report(result) == []
        for book, ws in _all_sheets(result):
            where = (book, ws.title)
            last = get_column_letter(ws.max_column)
            if ws.title == "Pivot for Last Day":                                  # its title also names the day
                assert ws.cell(row=1, column=1).value.startswith("تاريخ تشغيل التقرير: 07/11/2026  |  "), where
                assert f"A1:{last}1" in {str(r) for r in ws.merged_cells.ranges} and ws.print_title_rows == "$1:$2"
                continue
            assert ws.cell(row=1, column=1).value == want, where
            assert f"A1:{last}1" in {str(r) for r in ws.merged_cells.ranges}, where   # across the whole table
            a1 = ws.cell(row=1, column=1)
            assert a1.alignment.horizontal == "left" and a1.font.bold, where      # left side of the printed page
            assert ws.cell(row=2, column=1).value == "م", where                     # the header below it
            assert ws.print_title_rows == "$1:$2" and ws.print_area.endswith(f"$A$1:${last}${ws.max_row}"), where


def test_date_boundary_total_includes_the_report_date(tmp_path):
    rows = [row("01-11-2026", S, "100.00", 100001), row("31-10-2026", S, "40.00", 100001, doc="2"),
            row("02-11-2026", S, "7.00", 100001, doc="3"),                       # after the date: not counted
            row("01-11-2026", R, "-10.00", 100002, doc="4"), row("30-10-2026", R, "-5.00", 100002, doc="5")]
    write_erp_csv(tmp_path / "input" / "m.csv", rows)
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, date(2026, 11, 1),
                          load_main_data(write_main(tmp_path / "main.xlsx")))
    assert analysis.verify_analysis_report(result) == []
    sheet = _sheets(result.report_path)[SHEET_SECTORS]
    assert sheet[0][1:8] == ["مبيعات يوم 01/11/2026", "مبيعات حتى 31/10/2026", "إجمالي المبيعات",
                             "مرتجعات يوم 01/11/2026", "مرتجعات حتى 31/10/2026", "إجمالي المرتجعات", "الصافي"]
    # إجمالي المبيعات = يوم 01/11 (100) + حتى 31/10 (40): the report date included once, the 02/11 sale left out
    assert _by_label(sheet)[GRAND_TOTAL] == _d("100", "40", "140", "-10", "-5", "-15", "125")
    assert _counts(sheet)[GRAND_TOTAL] == [1, 1, 1]                               # one customer, counted once


# ------------------------------------------------------------------ Pivot for Last Day
PIVOT = "Pivot for Last Day"
PIVOT_MAIN = [
    (600001, "عميل أ", "الجملة", "سامح السيد محمود احم"),
    (600002, "عميل ب", "الجملة", "سامح السيد محمود احم"),
    (600003, "عميل ج", "مستشفيات", "هند محمد رضا"),
    (600004, "عميل د", "الجملة", "هشام احمد محمد"),
]
PIVOT_ROWS = [
    row(Y5, S, "100.00", 600001, "اسم في الملف", doc="1001"),       # one invoice, three product rows
    row(Y5, S, "50.00", 600001, doc="1001"), row(Y5, S, "25.00", 600001, doc="1001"),
    row(Y5, S, "30.00", 600001, doc="1002"),                         # a second invoice of the same customer
    row(Y5, R, "-20.00", 600001, doc="501"),                         # returns: never an invoice
    row(Y5, R, "-5.00", 600001, doc="1002"),                         # a return numbered like a sales invoice
    row(Y5, S, "60.00", 600002, doc="1003"),
    row(P5, S, "999.00", 600002, doc="1004"),                        # an earlier day: not in the pivot
    row(Y5, R, "-15.00", 600003, doc="502"),                         # returns only
    row(Y5, S, "80.00", 600004, doc="2001"),                         # another representative
    row(F5, S, "500.00", 600004, doc="2002"),                        # after the date: not counted at all
    row(Y5, S, "10.00", 777777, "زائر", doc="3001"),                 # unregistered: CSV name
]
PIVOT_EXPECTED = [                                                  # customer names: CSV column G, never Main Data
    ["هند محمد رضا", "اسم 600003", 0, *_d("0", "-15", "-15")],       # Cover order: هند before سامح
    ["إجمالي هند محمد رضا", None, 0, *_d("0", "-15", "-15")],
    ["سامح السيد محمود احم", "اسم 600002", 1, *_d("60", "0", "60")],  # customers by name
    ["سامح السيد محمود احم", "اسم في الملف", 2, *_d("205", "-25", "180")],  # 600001: its first CSV name
    ["إجمالي سامح السيد محمود احم", None, 3, *_d("265", "-25", "240")],
    ["هشام احمد محمد", "اسم 600004", 1, *_d("80", "0", "80")],
    ["إجمالي هشام احمد محمد", None, 1, *_d("80", "0", "80")],
    ["غير مسجل", "زائر", 1, *_d("10", "0", "10")],
    ["إجمالي غير مسجل", None, 1, *_d("10", "0", "10")],
    [GRAND_TOTAL, None, 5, *_d("355", "-40", "315")],
]


def _pivot_run(tmp_path, ref=REF, mode=MODE_YESTERDAY, rows=PIVOT_ROWS):
    write_erp_csv(tmp_path / "input" / "p.csv", rows)
    result = run_analysis(_config(tmp_path), mode, ref if mode == MODE_YESTERDAY else None,
                          load_main_data(write_main(tmp_path / "pivot_main.xlsx", PIVOT_MAIN)), now=RUN_AT)
    assert analysis.verify_analysis_report(result) == []
    return result, _sheets(result.reps_report_path)[PIVOT]


def test_pivot_is_the_last_sheet_with_the_six_columns(tmp_path):
    result, pivot = _pivot_run(tmp_path)
    wb = load_workbook(result.reps_report_path)
    assert wb.sheetnames[-1] == PIVOT and wb.sheetnames[:2] == ["Cover", "Master Data"]
    assert pivot[0] == ["اسم المندوب", "العميل", "عدد الفواتير", "إجمالي المبيعات", "إجمالي المرتجعات", "الصافي"]
    assert wb[PIVOT].cell(row=1, column=1).value == (                          # run 07/11, transactions of 30/10
        "تاريخ تشغيل التقرير: 07/11/2026  |  حركات يوم 30/10/2026 حسب المندوب والعميل")


def test_pivot_rows_and_grand_total(tmp_path):
    """Representatives, customers, invoices once per document, returns signed, net = sales + returns,
    no earlier or later day, grand total from the transactions."""
    result, pivot = _pivot_run(tmp_path)
    assert pivot[1:] == PIVOT_EXPECTED
    for r in pivot[1:]:
        assert r[5] == r[3] + r[4]                                               # الصافي = sales + returns
    day = [m for m in result.master if m.day == REF]
    assert pivot[-1][3:] == [sum((m.sales for m in day), D(0)), sum((m.returns for m in day), D(0)),
                             sum((m.net for m in day), D(0))]
    assert pivot[-1][2] == len({(m.code, m.doc_no) for m in day if m.category is analysis.Category.SALES}) == 5
    assert (pivot[-1][3], pivot[-1][4]) == (result.direct.sales_y, result.direct.returns_y)  # = مبيعات / مرتجعات يوم


def test_pivot_follows_the_selected_date(tmp_path):
    _, pivot = _pivot_run(tmp_path, ref=date(2026, 10, 20))
    assert pivot[1:] == [["سامح السيد محمود احم", "اسم 600002", 1, *_d("999", "0", "999")],
                         ["إجمالي سامح السيد محمود احم", None, 1, *_d("999", "0", "999")],
                         [GRAND_TOTAL, None, 1, *_d("999", "0", "999")]]


def test_pivot_empty_day_gives_zero_totals(tmp_path):
    result, pivot = _pivot_run(tmp_path, ref=date(2026, 10, 25))
    assert pivot[1:] == [[GRAND_TOTAL, None, 0, *_d("0", "0", "0")]]
    ws = load_workbook(result.reps_report_path)[PIVOT]
    assert ws.auto_filter.ref == "A2:F2" and ws.max_row == 3


def test_pivot_all_transactions_has_no_day(tmp_path):
    result, pivot = _pivot_run(tmp_path, mode=MODE_ALL)
    assert pivot[1:] == [[GRAND_TOTAL, None, 0, *_d("0", "0", "0")]]
    title = load_workbook(result.reps_report_path)[PIVOT].cell(row=1, column=1).value
    assert title.startswith("تاريخ تشغيل التقرير: 07/11/2026  |  تقرير كل الحركات: لا يوجد يوم مختار")


def test_pivot_excludes_the_excluded_representative(tmp_path):
    write_erp_csv(tmp_path / "input" / "p.csv", [row(Y5, S, "40.00", 600005, doc="9"), *PIVOT_ROWS])
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF,
                          load_main_data(write_main(tmp_path / "m.xlsx", [*PIVOT_MAIN, (600005, "x", "الجملة",
                                                                                         "دينا كرم ابراهيم")])))
    assert analysis.verify_analysis_report(result) == []
    assert _sheets(result.reps_report_path)[PIVOT][1:] == PIVOT_EXPECTED


def test_pivot_print_setup_black_and_white(tmp_path):
    result, _ = _pivot_run(tmp_path)
    ws = load_workbook(result.reps_report_path)[PIVOT]
    last = ws.max_row
    assert int(ws.page_setup.paperSize) == 9 and ws.page_setup.orientation == "landscape"
    assert ws.page_setup.blackAndWhite is True
    assert ws.sheet_properties.pageSetUpPr.fitToPage and ws.page_setup.fitToWidth == 1 and ws.page_setup.fitToHeight == 1
    assert ws.print_area.endswith(f"$A$1:$F${last}") and ws.print_title_rows == "$1:$2"
    assert ws.print_options.horizontalCentered and 0.25 <= ws.page_margins.left <= 0.5
    assert ws.sheet_view.rightToLeft and ws.freeze_panes == "A3" and ws.auto_filter.ref == f"A2:F{last - 1}"
    for c in range(1, 7):                                                        # black text, grey / white only
        header, body, total = ws.cell(row=2, column=c), ws.cell(row=3, column=c), ws.cell(row=last, column=c)
        assert header.font.bold and header.font.color.rgb.endswith("000000") and header.fill.fgColor.rgb.endswith("D9D9D9")
        assert body.font.color.rgb.endswith("000000") and body.border.left.color.rgb.endswith("000000")
        assert total.font.bold and total.border.top.style == "double" and total.border.bottom.style == "medium"
        assert ws.column_dimensions[get_column_letter(c)].width >= 14
    fills = {ws.cell(row=r, column=1).fill.fgColor.rgb[-6:] for r in range(3, last)
             if ws.cell(row=r, column=1).fill.fill_type}
    assert fills <= {"F2F2F2", "D9D9D9"}                                         # light grey stripes / subtotals
    assert ws.cell(row=1, column=1).font.bold and ws.cell(row=1, column=1).font.color.rgb.endswith("000000")
    assert [ws.cell(row=3, column=c).number_format for c in range(3, 7)] == ["#,##0"] + ["#,##0.00"] * 3


def test_long_pivot_prints_one_page_wide_several_pages_down(tmp_path):
    rows = [row(Y5, S, "1.00", 900000 + i, f"زائر {i}", doc=str(i)) for i in range(40)]
    result, pivot = _pivot_run(tmp_path, rows=rows)
    assert len(pivot) == 43                                       # header + 40 customers + their subtotal + total
    ws = load_workbook(result.reps_report_path)[PIVOT]
    assert ws.page_setup.fitToWidth == 1 and ws.page_setup.fitToHeight == 0 and ws.print_title_rows == "$1:$2"


def test_wrong_pivot_is_detected(tmp_path, monkeypatch):
    result, _ = _pivot_run(tmp_path)
    wb = load_workbook(result.reps_report_path)
    wb[PIVOT].cell(row=wb[PIVOT].max_row, column=3).value = 9                     # الإجمالي العام: عدد الفواتير
    wb.save(result.reps_report_path)
    problems = analysis.verify_analysis_report(result)
    assert any("/ Pivot for Last Day: row 12 عدد الفواتير 9 differs from 5" in p for p in problems)
    assert any("/ Pivot for Last Day: saved الإجمالي العام" in p for p in problems)
    original = analysis.pivot_table
    monkeypatch.setattr(analysis, "pivot_table",
                        lambda r: [[*row[:2], row[2] + 1, *row[3:]] for row in original(r)])  # one invoice per row too many
    assert any("Pivot for Last Day: customer rows differ from the day's transactions" in p
               for p in analysis.reconcile(result))
    assert any("Pivot for Last Day: الإجمالي العام" in p for p in analysis.reconcile(result))


def test_pivot_leaves_every_other_sheet_unchanged(tmp_path, setup):
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF, setup, now=RUN_AT)
    sheets = _sheets(result.reps_report_path)
    assert list(sheets)[-1] == PIVOT and len(sheets) == 2 + len(REP_SHEETS) + 1
    assert _cover(sheets)[GRAND_TOTAL] == [*_d("350.31", "-40.30", "310.01"), 219, 0, 4]   # as before the pivot
    assert len(sheets[SHEET_MASTER]) - 1 == 8


# ------------------------------------------------------------------ Pivot: CSV names, subtotals, run date
def _blank_name(r):
    """An ERP row whose name cell (column G) is empty."""
    r = list(r)
    r[6] = ""
    return r


def test_pivot_customer_name_is_csv_column_g_only(tmp_path):
    """Main Data names are 'عميل أ' ... 'عميل د'; the pivot shows the CSV names (column G) only."""
    rows = [*PIVOT_ROWS, row(Y5, S, "5.00", 600003, "اسم آخر في الملف", doc="1005"),   # a second CSV spelling
            _blank_name(row(Y5, S, "7.00", 600004, doc="2003"))]                        # blank G: the first non-blank
    _, pivot = _pivot_run(tmp_path, rows=rows)
    names = [r[1] for r in pivot[1:] if r[1]]
    assert names == ["اسم 600003", "اسم 600002", "اسم في الملف", "اسم 600004", "زائر"]
    assert not any(n.startswith("عميل ") for n in names)                       # never the Main Data name
    write_erp_csv(tmp_path / "blank" / "input" / "p.csv", [_blank_name(row(Y5, S, "9.00", 600001, doc="1"))])
    result = run_analysis(_config(tmp_path / "blank"), MODE_YESTERDAY, REF,
                          load_main_data(write_main(tmp_path / "blank" / "m.xlsx", PIVOT_MAIN)), now=RUN_AT)
    assert analysis.verify_analysis_report(result) == []
    assert _sheets(result.reps_report_path)[PIVOT][1][1] == "600001"             # no CSV name: the code, not Main Data


def test_pivot_names_match_the_csv_rows_of_the_day(tmp_path):
    result, pivot = _pivot_run(tmp_path)
    csv_names = {}
    for r in result.master:                                                     # Master Data = CSV column G
        if r.day == REF and r.name:
            csv_names.setdefault((r.representative, r.code), r.name)
    assert sorted(r[1] for r in pivot[1:] if r[1]) == sorted(csv_names.values())


def test_pivot_representative_subtotals(tmp_path):
    result, pivot = _pivot_run(tmp_path)
    body = pivot[1:-1]
    for i, r in enumerate(body):
        if r[1] is None:                                                        # a subtotal row
            rep = r[0].removeprefix("إجمالي ")
            customers = []
            for prev in reversed(body[:i]):
                if prev[1] is None:
                    break
                customers.append(prev)
            assert customers and {c[0] for c in customers} == {rep}             # right after its own customers
            assert r[2:] == [sum(c[2] for c in customers), sum(c[3] for c in customers), sum(c[4] for c in customers),
                             sum(c[5] for c in customers)]
    subs = [r for r in body if r[1] is None]
    assert [r[0] for r in subs] == ["إجمالي هند محمد رضا", "إجمالي سامح السيد محمود احم", "إجمالي هشام احمد محمد",
                                    "إجمالي غير مسجل"]
    assert pivot[-1][2:] == [sum(r[2] for r in subs), sum(r[3] for r in subs), sum(r[4] for r in subs),
                             sum(r[5] for r in subs)]
    # invoices: 1001 (3 rows) + 1002 for سامح's first customer, 1003 for the second; returns never count
    assert next(r for r in subs if "سامح" in r[0])[2] == 3
    ws = load_workbook(result.reps_report_path)[PIVOT]
    sub_row = next(r for r in range(3, ws.max_row) if ws.cell(row=r, column=1).value == "إجمالي سامح السيد محمود احم")
    cell = ws.cell(row=sub_row, column=1)
    assert cell.font.bold and cell.border.top.style == "medium" and cell.border.bottom.style == "medium"
    assert ws.cell(row=sub_row - 1, column=1).border.top.style == "thin"         # customers: thin borders


def test_pivot_run_date_differs_from_the_transaction_date(tmp_path):
    rows = [row("11-11-2026", S, "10.00", 600001, doc="1"), row("10-11-2026", S, "20.00", 600002, doc="2")]
    titles, data = {}, {}
    for ref in (date(2026, 11, 11), date(2026, 11, 10)):                        # same run, two selected days
        write_erp_csv(tmp_path / str(ref) / "input" / "p.csv", rows)
        result = run_analysis(PipelineConfig(input_dir=tmp_path / str(ref) / "input",
                                             output_dir=tmp_path / str(ref) / "out", reference_date=ref),
                              MODE_YESTERDAY, ref, load_main_data(write_main(tmp_path / str(ref) / "m.xlsx", PIVOT_MAIN)),
                              now=datetime(2026, 11, 12, 9, 15))
        assert analysis.verify_analysis_report(result) == []
        titles[ref] = load_workbook(result.reps_report_path)[PIVOT].cell(row=1, column=1).value
        data[ref] = _sheets(result.reps_report_path)[PIVOT][1:]
    assert titles[date(2026, 11, 11)] == "تاريخ تشغيل التقرير: 12/11/2026  |  حركات يوم 11/11/2026 حسب المندوب والعميل"
    assert titles[date(2026, 11, 10)] == "تاريخ تشغيل التقرير: 12/11/2026  |  حركات يوم 10/11/2026 حسب المندوب والعميل"
    assert data[date(2026, 11, 11)][-1][3] == D("10.00") and data[date(2026, 11, 10)][-1][3] == D("20.00")
    # the other sheets keep the selected report date
    other = load_workbook(result.reps_report_path)["Cover"].cell(row=1, column=1).value
    assert other == "تاريخ التقرير: 10/11/2026"


def test_pivot_run_date_is_the_actual_run_day(tmp_path, setup):
    """Through the real command: the run date is the computer's date at Run time, the data the selected day."""
    write_erp_csv(tmp_path / "input" / "month.csv", ROWS)
    assert main(["--input-dir", str(tmp_path / "input"), "--output-dir", str(tmp_path / "out"), "--task", "5",
                 "--ref-date", "30/10/2026", "--main-data", str(tmp_path / "data" / "main.xlsx"), "--non-interactive",
                 "--no-move"]) == EXIT_OK
    (reps,) = (tmp_path / "out").glob("Task5_*/Representatives_Analysis_*.xlsx")
    title = load_workbook(reps)[PIVOT].cell(row=1, column=1).value
    today = date.today().strftime("%d/%m/%Y")
    assert title == f"تاريخ تشغيل التقرير: {today}  |  حركات يوم 30/10/2026 حسب المندوب والعميل"
    assert today != "30/10/2026"


def test_wrong_pivot_subtotal_is_detected(tmp_path, monkeypatch):
    result, _ = _pivot_run(tmp_path)
    wb = load_workbook(result.reps_report_path)
    ws = wb[PIVOT]
    sub_row = next(r for r in range(3, ws.max_row) if ws.cell(row=r, column=1).value == "إجمالي سامح السيد محمود احم")
    ws.cell(row=sub_row, column=4).value = 1.0                                   # its إجمالي المبيعات
    wb.save(result.reps_report_path)
    problems = analysis.verify_analysis_report(result)
    assert any("saved 'إجمالي سامح السيد محمود احم' differs from the sum of its customers" in p for p in problems)
    original = analysis.pivot_table

    def wrong(r):
        rows = original(r)
        return [[*row[:3], row[3] + D("1"), row[4], row[5] + D("1")] if row[0] == "إجمالي هند محمد رضا" else row
                for row in rows]
    monkeypatch.setattr(analysis, "pivot_table", wrong)
    assert any("subtotal 'إجمالي هند محمد رضا'" in p and "does not match its customers" in p
               for p in analysis.reconcile(result))



# ------------------------------------------------------------------ Cover: targets, المباشر / الوكلاء, achievement %
def _cover_raw(path):
    """Label -> the saved Cover row (unrounded cells, without م), from row 3."""
    ws = load_workbook(path)[SHEET_COVER]
    return {r[1]: list(r[1:]) for r in ws.iter_rows(min_row=3, values_only=True)}


def test_cover_targets_are_the_fixed_business_targets():
    expected = {"اسلام حمد كمال محمد": 25, "سامح السيد محمود احم": 38, "داليا صلاح الدين فكر": 44,
                "سميرة محمد اسماعيل": 25, "محمود الشيمي": 21, "د هاني": 3, "ندا عبد الناصر طه": 25,
                "هشام احمد محمد": 2, "ابراهيم سلام": 4, "رامي رأفت يوقيم": 1, "محمد رأفت شحاتة محمد": 2,
                "احمد محمد علام محمد": 1, "احمد محمود موسي": 2, "احمد الحريري": 1, "هند محمد رضا": 25}
    assert analysis.COVER_TARGETS == expected
    for name, target in expected.items():
        assert analysis.cover_target(name) == target
    assert analysis.cover_target("اون لاين") is None and analysis.cover_target("الاداره") is None
    assert analysis.cover_target("غير مسجل") is None and analysis.cover_target("مندوب جديد") is None
    assert sum(expected[n] for n in COVER_GROUP_ONE if n in expected) == 206
    assert sum(expected[n] for n in GROUP_TWO) == 13


def test_cover_ratio_is_inside_over_target_with_safe_division():
    assert analysis._ratio(10, 38) == D("0.2632")                                # shown as 26.32%
    assert analysis._ratio(0, 25) == D("0") and analysis._ratio(3, 3) == D("1")
    assert analysis._ratio(5, 2) == D("2.5")                                      # above the target: 250.00%
    assert analysis._ratio(4, None) == analysis._ratio(4, 0) == analysis.NO_RATIO  # never a division by zero


def test_cover_labels_groups_and_direct_addition(tmp_path):
    result = _group_run(tmp_path)
    labels = [r[0] for r in analysis.cover_table(result)]
    assert labels == GROUPED_COVER
    assert "إجمالي المجموعة الأولى" not in labels and "إجمالي المجموعة الثانية" not in labels
    assert labels.index("د هاني") == labels.index("محمود الشيمي") + 1 < labels.index("إجمالي المباشر")
    assert DINA not in labels
    assert "د هاني" not in _sheets(result.reps_report_path)                       # no data: no sheet of his own


def test_cover_direct_addition_with_transactions(tmp_path):
    """د هاني with data: in المباشر (not the third group) and in its subtotal; reconciliation still passes."""
    main_rows = [*GROUP_MAIN, (500008, "h", "الجملة", "د هاني"), (500009, "i", "الجملة", "د هاني")]
    rows = [*GROUP_ROWS, row(Y5, S, "60.00", 500008, doc="20"), row(Y5, S, "10.00", 500009, doc="21")]
    write_erp_csv(tmp_path / "input" / "g.csv", rows)
    result = run_analysis(_config(tmp_path), MODE_YESTERDAY, REF,
                          load_main_data(write_main(tmp_path / "m.xlsx", main_rows)),
                          fixed_codes=frozenset({"500001", "500004", "888888", "500008"}), now=RUN_AT)
    assert analysis.verify_analysis_report(result) == []
    c = _cover(_sheets(result.reps_report_path))
    labels = list(c)
    assert labels.count("د هاني") == 1 and labels.index("د هاني") == labels.index("محمود الشيمي") + 1
    assert labels.index("د هاني") < labels.index("إجمالي المباشر")
    assert c["د هاني"] == [*_d("70.00", "0", "70.00"), 3, 1, 1]
    assert c["إجمالي المباشر"] == [*_d("220.00", "-15.00", "205.00"), 206, 2, 2]
    assert c["إجمالي المجموعة الثالثة"] == [*_d("39.00", "0", "39.00"), "لا يوجد هدف", 1, 1]
    raw = _cover_raw(result.reps_report_path)
    assert raw["د هاني"][-1] == pytest.approx(0.3333)                            # 1 / 3
    assert raw["إجمالي المباشر"][-1] == pytest.approx(0.0049)                    # 1 / 206 (اون لاين has no target)
    assert raw[GRAND_TOTAL][-1] == pytest.approx(0.0091)                         # (1 + 1) / 219


def test_cover_shows_every_representative_even_at_zero(tmp_path):
    result = _group_run(tmp_path)
    c = _cover(_sheets(result.reps_report_path))
    for name in [*COVER_GROUP_ONE, *GROUP_TWO]:
        assert name in c
    assert c["سميرة محمد اسماعيل"] == [*_d("0", "0", "0"), 25, 0, 0]              # nothing: shown, with its target
    assert c["ابراهيم سلام"] == [*_d("0", "0", "0"), 4, 0, 0]
    raw = _cover_raw(result.reps_report_path)
    assert raw["سميرة محمد اسماعيل"][-1] == 0 and raw["الاداره"][4:] == ["لا يوجد هدف", 0, 0, "—"]


def test_cover_subtotal_and_grand_total_ratios_are_not_sums_of_percentages(tmp_path):
    result = _group_run(tmp_path)
    raw = _cover_raw(result.reps_report_path)
    assert raw["هشام احمد محمد"][-1] == pytest.approx(0.5)                       # 1 / 2
    assert raw["إجمالي الوكلاء"][-1] == pytest.approx(0.0769)                    # 1 / 13, not 50% + 0% + ...
    assert raw["إجمالي المباشر"][-1] == 0                                        # 0 / 206 (اون لاين has no target)
    assert raw["إجمالي المجموعة الثالثة"][4:] == ["لا يوجد هدف", 1, 1, "—"]
    assert raw[GRAND_TOTAL][4] == 219 and raw[GRAND_TOTAL][-1] == pytest.approx(0.0046)   # 1 / 219
    for label, saved in raw.items():                                              # every representative row
        if isinstance(saved[4], int) and label not in (*analysis.SUBTOTAL_LABELS, GRAND_TOTAL):
            assert saved[-1] == pytest.approx(round(saved[5] / saved[4], 4)), label


def test_cover_net_is_sales_plus_returns(tmp_path):
    result = _group_run(tmp_path)
    for label, saved in _cover_raw(result.reps_report_path).items():
        assert saved[3] == pytest.approx(saved[1] + saved[2]), label


def test_wrong_cover_target_or_ratio_is_detected(tmp_path, monkeypatch):
    result = _group_run(tmp_path)
    original = analysis.cover_table
    monkeypatch.setattr(analysis, "cover_table", lambda r: [
        [*row[:4], 99, *row[5:]] if row[0] == "هشام احمد محمد" else row for row in original(r)])
    assert any("هشام احمد محمد" in p and "target" in p for p in analysis.reconcile(result))
    monkeypatch.setattr(analysis, "cover_table", lambda r: [
        [*row[:7], D("0.9")] if row[0] == "إجمالي الوكلاء" else row for row in original(r)])
    assert any("إجمالي الوكلاء" in p and "percentage" in p for p in analysis.reconcile(result))
    monkeypatch.setattr(analysis, "cover_table",
                        lambda r: [row for row in original(r) if row[0] != "سميرة محمد اسماعيل"])
    assert any("سميرة محمد اسماعيل" in p for p in analysis.reconcile(result))   # a representative left out


def test_saved_cover_ratio_is_verified(tmp_path):
    result = _group_run(tmp_path)
    wb = load_workbook(result.reps_report_path)
    ws = wb[SHEET_COVER]
    r = next(r for r in range(3, ws.max_row + 1) if ws.cell(row=r, column=2).value == "هشام احمد محمد")
    ws.cell(row=r, column=9).value = 0.9
    wb.save(result.reps_report_path)
    assert any("Cover" in p for p in analysis.verify_analysis_report(result))


def test_cover_print_setup_black_and_white_a4_landscape(tmp_path):
    result = _group_run(tmp_path)
    ws = load_workbook(result.reps_report_path)[SHEET_COVER]
    ps = ws.page_setup
    assert int(ps.paperSize) == 9 and ps.orientation == "landscape" and ps.blackAndWhite is True
    assert ps.fitToWidth == 1 and ws.sheet_properties.pageSetUpPr.fitToPage
    assert ws.print_title_rows == "$1:$2" and ws.sheet_view.rightToLeft
    assert [ws.cell(row=2, column=c).value for c in range(1, 10)] == ["م", *COVER_COLUMNS]
    assert all(ws.cell(row=2, column=c).alignment.wrap_text for c in range(1, 10))
