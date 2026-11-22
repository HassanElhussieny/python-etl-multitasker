"""Task 5 option 3 (Receivables Update): its own tests - the receivables workbook, the customer-level Master Data,
the representative sheets, Cover, the checks, the command line, and that nothing else runs or changes."""

import hashlib
from datetime import date
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook

from sales_etl import analysis
from sales_etl import receivables as rv
from sales_etl.analysis import FIXED_CODES_NAME, GRAND_TOTAL, load_fixed_codes, load_main_data
from sales_etl.cli import EXIT_AUDIT_FAILED, EXIT_NO_INPUT, EXIT_OK, EXIT_USAGE, main
from sales_etl.config import PipelineConfig
from sales_etl.fixed_customers import load_fixed_customers
from tests.conftest import R, S, write_erp_csv
from tests.test_analysis import _cli, row, write_fixed, write_main

D = Decimal
P5, Y5 = "20-10-2026", "30-10-2026"
RECV_MAIN = [
    (700001, "صيدلية النور", "الجملة", "مندوب تجريبي 01"),
    (700002, "صيدلية النور", "الجملة", "مندوب تجريبي 01"),     # the same name, another code
    (700003, "مستشفى", "مستشفيات", "مندوب تجريبي 08"),
    (700004, "وكيل", "وكلاء المنطقة أ", "مندوب تجريبي 03"),
    (700005, "عميل مندوب 13", "الجملة", "مندوب تجريبي 13"),           # the excluded representative
]
RECV_CSV = [
    row(Y5, S, "60.00", 700001, "صيدلية النور", doc="1"), row(Y5, S, "40.00", 700001, doc="1"),   # two rows, 1 invoice
    row(P5, R, "-10.00", 700001, doc="2"),
    row(P5, S, "50.00", 700002, "صيدلية النور", doc="3"),
    row(Y5, R, "-5.00", 700003, doc="4"),                                                        # returns only
    row(Y5, S, "30.00", 700004, doc="5"),                                                        # not in the file
    row(Y5, S, "9.00", 999001, "زائر", doc="6"),                                                 # not in Main Data
    row(Y5, S, "1000.00", 700005, doc="7"),                                                      # excluded
]
HEADER = ["م", "كود العميل", "اسم العميل", "المندوب", "قيمة الفواتير المفتوحة", "المستحق 7-10", "المحصل من مستحق 7-10",
          "الشيكات المرتدة", "المرتجعات", "مصروفات وتسويات", "المتبقي من 7-10", "المستحق الجديد", "المحصل على الجديد"]
RECV_ROWS = [  # code, 7-10 (column F), new (column L)
    (700001, 1000.5, 200), (700002, 0, None), (700003, -50, 75.25), (999001, 5, 6), (888888, 7, 8),   # 888888: no sale
]


def write_receivables(path, rows=RECV_ROWS, header=HEADER, title=True):
    """The real layout: a title row, an empty row, the header on row 3, the customers, then a total and a note
    row without a code."""
    wb = Workbook()
    ws = wb.active
    ws.title = "القديم والجديد"
    if title:
        ws.append([None, "متابعة التحصيل: القديم والجديد"])
        ws.append([])
    ws.append(header)
    for i, (code, old, new) in enumerate(rows, start=1):
        line = [None] * len(header)
        line[0], line[1], line[2], line[3] = i, code, f"عميل {code}", "مندوب"
        line[5], line[11] = old, new                                              # columns F and L
        ws.append(line)
    ws.append([None, None, "الإجمالي", None, None, 962.5])
    ws.append([None, None, "منها مستحق 7-10"])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _setup(tmp_path, csv_rows=RECV_CSV, recv_rows=RECV_ROWS):
    main_path = write_main(tmp_path / "data" / "main.xlsx", RECV_MAIN)
    write_fixed(tmp_path / "data" / FIXED_CODES_NAME, [700001])
    sales = write_erp_csv(tmp_path / "sales" / "sales.csv", csv_rows)
    recv = write_receivables(tmp_path / "recv" / "مستحقات.xlsx", recv_rows)
    return main_path, sales, recv


def _run(tmp_path, **kw):
    main_path, sales, recv = _setup(tmp_path, **kw)
    config = PipelineConfig(input_dir=sales.parent, output_dir=tmp_path / "out", reference_date=date.today())
    main_data = load_main_data(main_path)
    fixed_path = tmp_path / "data" / FIXED_CODES_NAME                         # 700001: الشهر الأعلى قيمة 1000
    result = rv.run_receivables_update(config, sales, recv, main_data, fixed_codes=load_fixed_codes(fixed_path),
                                       fixed_customers=load_fixed_customers(fixed_path), fixed_path=fixed_path)
    return result


def _saved(result):
    wb = load_workbook(result.report_path)
    return wb, {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


# ------------------------------------------------------------------ the receivables workbook
def test_load_receivables_reads_the_real_layout(tmp_path):
    f = rv.load_receivables(write_receivables(tmp_path / "r.xlsx"))
    assert (f.sheet, f.header_row, f.rows_without_code) == ("القديم والجديد", 3, 2)   # total + note rows skipped
    assert sorted(f.records) == ["700001", "700002", "700003", "888888", "999001"]
    assert (f.records["700001"].old, f.records["700001"].new) == (D("1000.5"), D("200"))
    assert (f.records["700002"].old, f.records["700002"].new) == (D("0"), None)        # zero kept, empty = None
    assert (f.records["700003"].old, f.records["700003"].new) == (D("-50"), D("75.25"))   # negative kept
    assert f.records["700001"].row == 4


def test_wrong_receivables_workbooks_are_refused(tmp_path):
    with pytest.raises(rv.ReceivablesError, match="no sheet has the columns"):
        rv.load_receivables(write_receivables(tmp_path / "a.xlsx", header=[h.replace("المستحق الجديد", "المستحق الجديد حتي اليوم")
                                                                            for h in HEADER]))
    moved = ["م", "اسم العميل", "كود العميل", *HEADER[3:]]                       # the code in column C
    with pytest.raises(rv.ReceivablesError, match="expected column B"):
        rv.load_receivables(write_receivables(tmp_path / "b.xlsx", header=moved))
    with pytest.raises(rv.ReceivablesError, match=r"1 customer code\(s\) appear more than once.*700001 \(rows 4, 6\)"):
        rv.load_receivables(write_receivables(tmp_path / "c.xlsx", rows=[*RECV_ROWS[:2], (700001, 3, 4)]))
    with pytest.raises(rv.ReceivablesError, match="'abc' is not an amount"):
        rv.load_receivables(write_receivables(tmp_path / "d.xlsx", rows=[(700001, "abc", 1)]))
    path = write_receivables(tmp_path / "e.xlsx", rows=[(700001, 1, 2)])
    wb = load_workbook(path)
    wb.active.cell(row=4, column=6).value = "=1+1"                              # never calculated: no saved value
    wb.save(path)
    with pytest.raises(rv.ReceivablesError, match="formula without a saved value"):
        rv.load_receivables(path)


# ------------------------------------------------------------------ the workbook
def test_master_data_one_row_per_customer(tmp_path):
    result = _run(tmp_path)
    assert result.succeeded, result.problems
    assert result.report_path.name.startswith("Representatives_Receivables_Update_")
    assert result.report_path.parent.name.startswith("Task5_Receivables_")
    _, sheets = _saved(result)
    master = sheets["Master Data"]
    assert master[1] == list(rv.MASTER_HEADERS)
    body, total = master[2:-1], master[-1]
    assert [r[1] for r in body] == [700001, 700002, 700003, 700004, 999001]      # once each, CSV order; no 700005
    a, b = body[0], body[1]
    assert a[2] == b[2] == "صيدلية النور" and a[13:15] == [1000.5, 200] and b[13:15] == [0, None]   # by code only
    assert a[11:13] == ["نعم", None] and b[11:13] == [None, "نعم"]               # in / out of the 224 list
    assert all((r[11] == "نعم") != (r[12] == "نعم") for r in body)               # always exactly one of the two
    assert a[5] == 3 and a[8:11] == [100, -10, 90]                                # 3 transactions summed
    assert body[2][8:11] == [0, -5, -5] and body[2][13:15] == [-50, 75.25]
    assert body[3][13:16] == [None, None, rv.NOT_FOUND_NOTE]                      # not in the file: empty
    assert body[4][3] == "غير مسجل" and body[4][13:15] == [5, 6]
    assert total[:3] == [None, GRAND_TOTAL, "5 عميل"] and total[5] == 7 and total[11:13] == [1, 4]
    assert total[8:11] == [189, -15, 174] and total[13:15] == [955.5, 281.25]
    detail = sheets["تفاصيل الحركات"]
    assert len(detail) - 3 == 7 and detail[-1][1] == GRAND_TOTAL                  # every original transaction


def test_representative_sheets_carry_each_customer_once(tmp_path):
    result = _run(tmp_path)
    wb, sheets = _saved(result)
    expected = list(analysis.representative_sheet_names(result.analysis).values())
    assert wb.sheetnames == ["Cover", "Master Data", *expected, "تفاصيل الحركات"]  # every representative, also empty
    sameh = sheets["مندوب تجريبي 01"]
    assert sameh[1] == list(rv.MASTER_HEADERS) and [r[1] for r in sameh[2:-1]] == [700001, 700002]
    assert sameh[-1][13:15] == [1000.5, 200] and sameh[-1][8:11] == [150, -10, 140] and sameh[-1][11:13] == [1, 1]
    assert sheets["مندوب تجريبي 03"][-1][13:15] == [0, 0]                          # its customer not in the file
    assert sheets["غير مسجل"][-1][13:15] == [5, 6]
    assert sheets["اون لاين"][2:] == [[None, GRAND_TOTAL, "0 عميل", None, None, 0, None, None, 0, 0, 0, 0, 0, 0, 0, None]]
    codes = [r[1] for name in expected for r in sheets[name][2:-1]]
    assert sorted(codes) == sorted(r[1] for r in sheets["Master Data"][2:-1])    # none lost, none twice
    master = {r[1]: r for r in sheets["Master Data"][2:-1]}
    for name in expected:
        for r in sheets[name][2:-1]:
            assert r[1:] == master[r[1]][1:]                                      # the Master Data values (م: own)


def test_cover_shows_money_who_sold_and_receivables_per_representative(tmp_path):
    result = _run(tmp_path)
    _, sheets = _saved(result)
    cover = sheets["Cover"]
    assert cover[1] == ["م", "اسم المندوب", "الشهر الأعلى قيمة", "إجمالي المبيعات خلال الشهر",
                        "إجمالي المرتجعات خلال الشهر", "إجمالي الصافي", "عملاء داخل الـ224 حققوا مبيعات",
                        "عملاء خارج الـ224 حققوا مبيعات", "إجمالي المستحق 7-10", "إجمالي المستحق الجديد"]
    for gone in ("إجمالي العملاء", "عدد العملاء داخل قائمة الـ224 كود", "عدد العملاء خارج قائمة الـ224 كود",
                 "هدف العملاء", "نسبة تحقيق هدف العملاء"):                      # 1: no listed-customer counts
        assert gone not in cover[1]
    rows = {r[1]: r[2:] for r in cover[2:]}
    # مندوب 01: 700001 (in the list, sold 100, C = 1000) + 700002 (outside, sold 50)
    assert rows["مندوب تجريبي 01"] == [1000, 150, -10, 140, 1, 1, 1000.5, 200]
    assert rows["مندوب تجريبي 08"] == [0, 0, -5, -5, 0, 0, -50, 75.25]             # returns only: never "sold"
    assert rows["مندوب تجريبي 03"] == [0, 30, 0, 30, 0, 1, 0, 0]                  # not in the receivables file
    assert rows["إجمالي المباشر"] == [1000, 150, -15, 135, 1, 1, 950.5, 275.25]
    assert rows["إجمالي الوكلاء"] == [0, 30, 0, 30, 0, 1, 0, 0]
    assert rows["غير مسجل"] == rows["إجمالي المجموعة الثالثة"] == [0, 9, 0, 9, 0, 1, 5, 6]
    assert rows[GRAND_TOTAL] == [1000, 189, -15, 174, 1, 3, 955.5, 281.25]
    assert [r[0] for r in cover[2:]][-1] is None and cover[-1][1] == GRAND_TOTAL
    reps = [sheets[n][-1] for n in analysis.representative_sheet_names(result.analysis).values()]
    assert [sum(r[13] for r in reps), sum(r[14] for r in reps)] == rows[GRAND_TOTAL][6:]       # 6, 7
    assert [sum(r[8] for r in reps), sum(r[9] for r in reps)] == rows[GRAND_TOTAL][1:3]        # 5


def test_reconciliation_statistics(tmp_path):
    result = _run(tmp_path)
    st = result.stats
    assert (st["csv_codes"], st["found"], st["not_found"], st["duplicates"]) == (5, 4, 1, 0)
    assert st["missing_codes"] == ["700004"] and st["rows_without_code"] == 2
    assert (st["source_old"], st["source_new"]) == (D("955.5"), D("281.25"))
    assert st["master_old"] == st["cover_old"] == D("955.5") and st["master_new"] == st["cover_new"] == D("281.25")
    assert (st["sales"], st["returns"], st["net"], st["transactions"]) == (D("189.00"), D("-15.00"), D("174.00"), 7)


def test_wrong_tables_are_detected(tmp_path, monkeypatch):
    result = _run(tmp_path)
    assert rv.reconcile(result) == [] and rv.verify_workbook(result) == []
    result.customers[0].old = D("1")                                              # a wrong lookup
    assert any("customer 700001 receivables" in p for p in rv.reconcile(result))
    result.customers[0].old = D("1000.5")
    lost = result.customers.pop(1)                                                # a customer lost
    assert any("a customer lost or repeated" in p for p in rv.reconcile(result))
    result.customers.insert(1, lost)
    wb = load_workbook(result.report_path)
    wb["مندوب تجريبي 01"].cell(row=3, column=14).value = 1.0                 # a saved receivable
    wb.save(result.report_path)
    problems = rv.verify_workbook(result)
    assert any("مندوب تجريبي 01: row 3 المستحق 7-10" in p for p in problems)


def test_source_files_are_never_changed_or_moved(tmp_path):
    main_path, sales, recv = _setup(tmp_path)
    before = [_sha(p) for p in (main_path, sales, recv)]
    config = PipelineConfig(input_dir=sales.parent, output_dir=tmp_path / "out", reference_date=date.today())
    result = rv.run_receivables_update(config, sales, recv, load_main_data(main_path))
    assert result.succeeded and [_sha(p) for p in (main_path, sales, recv)] == before
    assert sorted(p.name for p in sales.parent.iterdir()) == ["sales.csv"]
    assert [p.name for p in (tmp_path / "out").iterdir()] == [result.run_dir.name]   # no Task5_ reports


# ------------------------------------------------------------------ the command line
def _args(tmp_path, *extra):
    return ["--input-dir", str(tmp_path / "sales"), "--output-dir", str(tmp_path / "out"), "--main-data",
            str(tmp_path / "data" / "main.xlsx"), "--task", "5", "--non-interactive", *extra]


def test_cli_option_3_writes_only_the_receivables_report(tmp_path, capsys):
    _, sales, recv = _setup(tmp_path)
    code = main(_args(tmp_path, "--report-type", "3", "--sales-csv", str(sales), "--receivables", str(recv)))
    out = capsys.readouterr().out
    assert code == EXIT_OK, out
    assert "Report type              : 3. Receivables Update" in out and "Receivables check        : PASS" in out
    assert "Customer codes (CSV)     : 5  | found 4 | not found 1 | duplicates 0" in out
    assert "Reconciliation" not in out and "Customer report check" not in out     # reports 1-3 never ran
    (path,) = (tmp_path / "out").rglob("*.xlsx")
    assert path.name.startswith("Representatives_Receivables_Update_") and path.parent.name.startswith("Task5_Receivables_")
    assert sales.is_file() and not (tmp_path / "processed").exists() and not (tmp_path / "state").exists()


def _input_folder(tmp_path, *, sales=1, books=1, extras=True):
    """The project-style input folder: Sales CSV(s) and receivables workbook(s) under names that do not say what
    they are, plus - with ``extras`` - files that must be ignored (another CSV, another Excel, an Excel lock file)."""
    write_main(tmp_path / "data" / "main.xlsx", RECV_MAIN)
    write_fixed(tmp_path / "data" / FIXED_CODES_NAME, [700001])
    folder = tmp_path / "input"
    for i in range(sales):
        write_erp_csv(folder / f"ملف {i + 1}.csv", RECV_CSV)
    for i in range(books):
        write_receivables(folder / f"book {i + 1}.xlsx")
    if extras:
        (folder / "notes.csv").write_text("a;b\n1;2\n", encoding="utf-8")               # not a sales file
        other = Workbook()
        other.active.append(["كود العميل", "اسم العميل", "المبلغ"])
        other.save(folder / "other.xlsx")                                           # not a receivables file
        (folder / "~$book 1.xlsx").write_bytes(b"lock")
    return folder


def _snapshot(folder):
    return sorted((p.name, _sha(p)) for p in folder.iterdir())


def test_cli_option_3_finds_both_files_in_the_input_folder(tmp_path, capsys):
    folder = _input_folder(tmp_path)
    before = _snapshot(folder)
    code, asked = _cli(tmp_path, ["5", "3"])
    out = capsys.readouterr().out
    assert code == EXIT_OK, out
    assert asked == ["Select task number: ", "Select report type: "]                # no path asked
    assert "Sales CSV (found)        : ملف 1.csv" in out and "Receivables (found)      : book 1.xlsx" in out
    assert "Receivables check        : PASS" in out
    assert _snapshot(folder) == before                                           # nothing renamed, moved, changed
    (report,) = (tmp_path / "out").rglob("*.xlsx")
    assert report.name.startswith("Representatives_Receivables_Update_")


def test_cli_option_3_names_the_missing_file(tmp_path, capsys):
    folder = _input_folder(tmp_path, books=0)
    assert _cli(tmp_path, ["5", "3"])[0] == EXIT_NO_INPUT
    err = capsys.readouterr().err
    assert "no receivables Excel file" in err and "no Sales CSV" not in err
    assert "ignored: other.xlsx - no sheet with the columns" in err
    for p in folder.glob("*.csv"):
        p.unlink()
    assert _cli(tmp_path, ["5", "3"])[0] == EXIT_NO_INPUT
    err = capsys.readouterr().err
    assert "no Sales CSV" in err and "no receivables Excel file" in err
    assert not (tmp_path / "out").exists() or not list((tmp_path / "out").rglob("*.xlsx"))


def test_cli_option_3_asks_which_file_when_there_are_several(tmp_path, capsys):
    _input_folder(tmp_path, sales=2, books=2)
    code, asked = _cli(tmp_path, ["5", "3", "9", "2", "1"])
    out = capsys.readouterr().out
    assert code == EXIT_OK
    assert asked == ["Select task number: ", "Select report type: ", "Which Sales CSV file (1-2)? ",
                     "Which Sales CSV file (1-2)? ", "Which receivables Excel file (1-2)? "]   # 9: not offered
    assert "Sales CSV (found)        : ملف 2.csv" in out and "Receivables (found)      : book 1.xlsx" in out
    code = main(["--input-dir", str(tmp_path / "input"), "--output-dir", str(tmp_path / "out2"), "--main-data",
                 str(tmp_path / "data" / "main.xlsx"), "--task", "5", "--report-type", "3", "--non-interactive"])
    err = capsys.readouterr().err
    assert code == EXIT_USAGE and "2 Sales CSV files in the input folder - choose one with --sales-csv" in err
    assert "1. ملف 1.csv" in err and "2. ملف 2.csv" in err


def test_cli_option_3_given_files_are_used_as_they_are(tmp_path, capsys):
    _input_folder(tmp_path, sales=2, books=2)
    code = main(["--input-dir", str(tmp_path / "input"), "--output-dir", str(tmp_path / "out"), "--main-data",
                 str(tmp_path / "data" / "main.xlsx"), "--task", "5", "--report-type", "3", "--non-interactive",
                 "--sales-csv", str(tmp_path / "input" / "ملف 2.csv"),
                 "--receivables", str(tmp_path / "input" / "book 2.xlsx")])
    assert code == EXIT_OK and "Sales CSV (found)        : ملف 2.csv" in capsys.readouterr().out


def test_cli_option_3_argument_rules(tmp_path, capsys):
    _, sales, recv = _setup(tmp_path)
    assert main(_args(tmp_path, "--report-type", "3")) == EXIT_NO_INPUT          # the folder has no receivables file
    assert main(_args(tmp_path, "--report-type", "3", "--sales-csv", str(sales), "--receivables", str(recv),
                      "--ref-date", "30/10/2026")) == EXIT_USAGE                  # no date
    assert main(_args(tmp_path, "--report-type", "2", "--sales-csv", str(sales))) == EXIT_USAGE
    assert main(_args(tmp_path, "--report-type", "3", "--sales-csv", str(recv), "--receivables", str(recv))) == EXIT_USAGE
    dup = write_receivables(tmp_path / "dup.xlsx", rows=[*RECV_ROWS, (700001, 1, 1)])
    assert main(_args(tmp_path, "--report-type", "3", "--sales-csv", str(sales), "--receivables", str(dup))) \
        == EXIT_AUDIT_FAILED
    err = capsys.readouterr().err
    assert "appear more than once" in err and "No receivables report was written" in err
    assert not list((tmp_path / "out").rglob("*.xlsx")) if (tmp_path / "out").exists() else True


def test_options_1_and_2_are_unchanged(tmp_path):
    """Option 2 on the same CSV: the three Task 5 reports, no receivables report, input moved as before."""
    _setup(tmp_path)
    assert main(_args(tmp_path, "--report-type", "2", "--no-move")) == EXIT_OK
    names = sorted(p.name for p in (tmp_path / "out").rglob("*.xlsx"))
    assert len(names) == 3 and not any(n.startswith("Representatives_Receivables") for n in names)



# ------------------------------------------------------------------ Cover: customers who sold (distinct codes, sales > 0)
SOLD_ROWS = [
    row(Y5, S, "60.00", 700001, doc="1"), row(Y5, S, "40.00", 700001, doc="2"),   # two invoices: once
    row(P5, R, "-500.00", 700001, doc="3"),                                       # net < 0, sales > 0: sold
    row(Y5, S, "0.00", 700002, doc="4"),                                           # a zero sale: not sold
    row(Y5, R, "-5.00", 700003, doc="5"),                                          # returns only: not sold
    row(Y5, S, "30.00", 700004, doc="6"), row(P5, S, "1.00", 700004, doc="7"),    # outside, two invoices: once
]


def test_cover_counts_customers_who_sold_by_distinct_code(tmp_path):
    result = _run(tmp_path, csv_rows=SOLD_ROWS)
    assert result.succeeded, result.problems
    rows = {r[1]: list(r[2:]) for r in load_workbook(result.report_path)["Cover"].iter_rows(min_row=3, values_only=True)}
    assert rows["مندوب تجريبي 01"][4:6] == [1, 0]                           # 700001 yes, 700002 (0.00) no
    assert rows["مندوب تجريبي 08"][4:6] == [0, 0]                                   # returns only
    assert rows["مندوب تجريبي 03"][4:6] == [0, 1]                                 # two invoices: one customer
    assert rows[GRAND_TOTAL][4:6] == [1, 1] and rows[GRAND_TOTAL][1:4] == [131, -505, -374]
    master = [list(r) for r in load_workbook(result.report_path)["Master Data"].iter_rows(min_row=3, values_only=True)]
    assert len(master) - 1 == 4 and master[-1][11:13] == [1, 3]                  # listed: all four, by column


def test_wrong_cover_is_detected(tmp_path):
    result = _run(tmp_path)
    result.customers[1].sales = D("0")                                            # 700002 would not have sold
    problems = rv.reconcile(result)
    assert any("Cover: 'مندوب تجريبي 01'" in p and "differs from the source files" in p for p in problems)
    result.customers[1].sales = D("50.00")
    result.customers[0].top = D("1")                                              # a wrong الشهر الأعلى قيمة
    assert any("Cover: 'مندوب تجريبي 01'" in p for p in rv.reconcile(result))
