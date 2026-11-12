"""Task 7: عدد العملاء الشهري لكل ملف - distinct customer codes per file per month (A date, F code)."""

import json
from datetime import date, timedelta

import pytest

from openpyxl import load_workbook

from sales_etl import monthly_file_customers
from sales_etl.cli import EXIT_AUDIT_FAILED, EXIT_NO_INPUT, EXIT_OK, EXIT_USAGE, main
from sales_etl.config import PipelineConfig
from sales_etl.dates import parse_date
from sales_etl.monthly_file_customers import (
    FILE_COLUMN,
    JOURNAL_NAME,
    MONTHS,
    REPORT_NAME,
    REPORT_SHEET,
    TASK_PREFIX,
    YEAR_COLUMN,
    accounting_month,
    reconcile,
    run_file_customers,
    verify_file_customers_report,
)
from tests.conftest import ERP_HEADER, R, S, alex8_rows, write_erp_csv


def row(day, code, kind=S, net="10.00", name="", doc="1"):
    """ERP row (21 cells): A date, B type, F code, G name, T صافى."""
    return [day, kind, doc, "93", "", code, name, "", "7", "rep", net, *["0"] * 8, net, ""]


def _config(tmp_path):
    return PipelineConfig(input_dir=tmp_path / "input", output_dir=tmp_path / "output",
                          reference_date=date(2026, 1, 1))


def _rows(path):
    wb = load_workbook(path, data_only=True)
    assert wb.sheetnames == [REPORT_SHEET]
    return [list(r) for r in wb[REPORT_SHEET].iter_rows(values_only=True)]


# ------------------------------------------------------------------ calculation
def test_repeated_customer_in_one_month_counts_once(tmp_path):
    write_erp_csv(tmp_path / "input" / "file1.csv", [
        row("05-01-2026", "100001"), row("08-01-2026", "100001", R, "-5.00", doc="2"),
        row("12-01-2026", "100002", doc="3"), row("03-02-2026", "100001", doc="4"),
        *[row("20-01-2026", "100001", doc=str(10 + i)) for i in range(10)],
    ])
    result = run_file_customers(_config(tmp_path))
    rows = _rows(result.report_path)
    assert result.report_path.name == REPORT_NAME == "تقرير عدد العملاء الشهري لكل ملف.xlsx"
    assert rows == [["اسم الملف", "يناير", "فبراير", "إجمالي العملاء المميزين خلال السنة"], ["file1.csv", 2, 1, 2]]
    assert result.complete and verify_file_customers_report(result) == []


def test_same_customer_in_two_months_counts_in_each(tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [row("01-03-2026", "7"), row("31-03-2026", "7", doc="2"),
                                                 row("01-05-2026", "7", doc="3")])
    rows = _rows(run_file_customers(_config(tmp_path)).report_path)
    assert rows[0] == [FILE_COLUMN, *MONTHS[:5], YEAR_COLUMN]          # يناير .. مايو, nothing after
    assert rows[1] == ["f.csv", 0, 0, 1, 1, 1, 1]                      # 31-03 is أبريل (from the 26th)


def test_same_customer_in_two_files_counted_in_each_file(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [row("01-01-2026", "100001"), row("02-01-2026", "100002", doc="2")])
    write_erp_csv(tmp_path / "input" / "b.csv", [row("09-01-2026", "100001"), row("01-02-2026", "100001", doc="2")])
    rows = _rows(run_file_customers(_config(tmp_path)).report_path)
    assert rows[1:] == [["a.csv", 2, 0, 2], ["b.csv", 1, 1, 1]]        # no merging across files


def test_months_are_the_union_of_all_files(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [row("15-02-2026", "1")])
    write_erp_csv(tmp_path / "input" / "b.csv", [row("15-10-2026", "1")])
    rows = _rows(run_file_customers(_config(tmp_path)).report_path)
    assert rows[0] == [FILE_COLUMN, *MONTHS[:10], YEAR_COLUMN]
    assert rows[1:] == [["a.csv", 0, 1, *[0] * 8, 1], ["b.csv", *[0] * 9, 1, 1]]


def test_every_document_type_counts_and_amounts_are_ignored(tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [
        row("01-01-2026", "1", S, "0.00"), row("01-01-2026", "2", R, "-99999.00", doc="2"),
        row("01-01-2026", "3", "تسوية", "5.00", doc="3"),           # another document type: counted too
    ])
    assert _rows(run_file_customers(_config(tmp_path)).report_path)[1] == ["f.csv", 3, 3]


def test_codes_kept_as_written_and_names_not_used(tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [
        row("01-01-2026", "100", name="س"), row("01-01-2026", "0100", name="س", doc="2"),
        row("01-01-2026", "١٠٠", name="س", doc="3"),                 # Arabic digits: a different code
        row("02-01-2026", "100", name="اسم آخر", doc="4"),          # same code, other name: one customer
    ], encoding="utf-8")
    result = run_file_customers(_config(tmp_path))
    assert _rows(result.report_path)[1] == ["f.csv", 3, 3]
    assert result.rows[0].months[1] == {"100", "0100", "١٠٠"}


def test_future_dates_are_counted(tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [row("15-12-2099", "1")])
    rows = _rows(run_file_customers(_config(tmp_path)).report_path)
    assert rows[0][-2] == "ديسمبر" and rows[1][-2:] == [1, 1]


def test_real_alex8_content_matches_an_independent_count(tmp_path):
    txns = alex8_rows()
    write_erp_csv(tmp_path / "input" / "alex8.csv", txns)
    result = run_file_customers(_config(tmp_path))
    expected: dict[int, set[str]] = {}
    for t in txns:
        expected.setdefault(_period(parse_date(t.date)), set()).add(t.customer)
    last = max(expected)
    assert _rows(result.report_path)[1] == ["alex8.csv", *[len(expected.get(m, ())) for m in range(1, last + 1)],
                                            len({t.customer for t in txns})]


def test_files_listed_in_numeric_name_order(tmp_path):
    for name in ("3omla210.csv", "3omla29.csv", "a.csv"):
        write_erp_csv(tmp_path / "input" / name, [row("01-01-2026", "1")])
    rows = _rows(run_file_customers(_config(tmp_path)).report_path)
    assert [r[0] for r in rows[1:]] == ["3omla29.csv", "3omla210.csv", "a.csv"]


# ------------------------------------------------------------------ year column (distinct in the whole file)
def test_year_column_counts_each_customer_once_per_file(tmp_path):
    write_erp_csv(tmp_path / "input" / "several_months.csv", [                  # one customer, many months
        row("05-01-2026", "100"), row("05-02-2026", "100", doc="2"), row("05-03-2026", "100", doc="3")])
    write_erp_csv(tmp_path / "input" / "same_month.csv", [                      # one customer, 5 times in a month
        *[row("05-01-2026", "0100", doc=str(i)) for i in range(5)]])
    write_erp_csv(tmp_path / "input" / "different.csv", [                       # other customers every month
        row("05-01-2026", "1"), row("05-02-2026", "2", doc="2"), row("05-03-2026", "3", doc="3")])
    write_erp_csv(tmp_path / "input" / "mixed.csv", [                           # 2 + 2 + 2 monthly, 4 in the year
        row("05-01-2026", "A"), row("06-01-2026", "B", doc="2"), row("05-02-2026", "B", doc="3"),
        row("06-02-2026", "C", doc="4"), row("05-03-2026", "C", doc="5"), row("06-03-2026", "D", doc="6")])
    write_erp_csv(tmp_path / "input" / "empty.csv", [])
    result = run_file_customers(_config(tmp_path))
    rows = _rows(result.report_path)
    assert rows[0] == [FILE_COLUMN, "يناير", "فبراير", "مارس", YEAR_COLUMN]
    assert rows[1:] == [
        ["different.csv", 1, 1, 1, 3],
        ["empty.csv", 0, 0, 0, 0],
        ["mixed.csv", 2, 2, 2, 4],                                   # not 6: the monthly counts are not added
        ["same_month.csv", 1, 0, 0, 1],
        ["several_months.csv", 1, 1, 1, 1],
    ]
    assert verify_file_customers_report(result) == []


def test_year_column_matches_an_independent_distinct_count_of_the_file(tmp_path):
    txns = alex8_rows()
    write_erp_csv(tmp_path / "input" / "alex8.csv", txns)
    write_erp_csv(tmp_path / "input" / "b.csv", [row("01-01-2026", t.customer) for t in txns[:40]])
    result = run_file_customers(_config(tmp_path))
    rows = _rows(result.report_path)
    assert rows[1][-1] == len({t.customer for t in txns})
    assert rows[2][-1] == len({t.customer for t in txns[:40]})                 # its own file only
    assert rows[1][-1] <= sum(rows[1][1:-1])                                   # never more than the months


def test_reconcile_and_verify_detect_a_wrong_year_value(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [row("01-01-2026", "1"), row("01-02-2026", "1", doc="2")])
    result = run_file_customers(_config(tmp_path))
    wb = load_workbook(result.report_path)
    wb[REPORT_SHEET]["D2"] = 2                                                 # the sum of the months
    wb.save(result.report_path)
    assert any("cell 2 differs from 1" in p for p in verify_file_customers_report(result))
    result.rows[0].codes.add("9")
    assert any(f"a.csv: {YEAR_COLUMN} 2 differs from 1" in p for p in reconcile(result))


# ------------------------------------------------------------------ accounting months (26th .. 25th)
def _period(day):
    """Independent: the period 26-<previous month> .. 25-<month> that holds ``day``."""
    periods = [(12, date(2025, 11, 26), date(2025, 12, 25))]
    periods += [(m, date(2025 if m == 1 else 2026, 12 if m == 1 else m - 1, 26), date(2026, m, 25))
                for m in range(1, 13)]
    (month,) = [m for m, start, end in periods if start <= day <= end]
    return month


@pytest.mark.parametrize("day, month", [
    ("25-09-2026", "سبتمبر"), ("26-09-2026", "أكتوبر"), ("30-09-2026", "أكتوبر"), ("01-10-2026", "أكتوبر"),
    ("25-10-2026", "أكتوبر"), ("26-10-2026", "نوفمبر"), ("25-12-2025", "ديسمبر"), ("26-12-2025", "يناير"),
])
def test_accounting_month_boundaries(tmp_path, day, month):
    write_erp_csv(tmp_path / "input" / "f.csv", [row(day, "1")])
    result = run_file_customers(_config(tmp_path))
    header, values = _rows(result.report_path)
    assert header[-2] == month and values == ["f.csv", *[0] * (len(header) - 3), 1, 1]
    assert verify_file_customers_report(result) == []


def test_december_to_january_transition(tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [
        row("20-12-2025", "A"), row("25-12-2025", "B", doc="2"),                         # ديسمبر 2025
        row("26-12-2025", "A", doc="3"), row("31-12-2025", "C", doc="4"),                # يناير 2026
        row("01-01-2026", "C", doc="5"), row("25-01-2026", "D", doc="6"),                # يناير 2026
        row("26-01-2026", "D", doc="7"),                                                 # فبراير 2026
    ])
    result = run_file_customers(_config(tmp_path))
    rows = _rows(result.report_path)
    assert rows[0] == [FILE_COLUMN, *MONTHS, YEAR_COLUMN]
    assert rows[1] == ["f.csv", 3, 1, *[0] * 9, 2, 4]                  # يناير: A, C, D once each; year: A-D
    assert verify_file_customers_report(result) == []


def test_accounting_month_matches_the_periods_for_every_day():
    day = date(2025, 11, 26)
    while day <= date(2026, 12, 25):
        assert accounting_month(day) == _period(day), day
        day += timedelta(days=1)


def test_reconcile_uses_the_accounting_periods(tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [row("26-09-2026", "1")])
    result = run_file_customers(_config(tmp_path))
    assert reconcile(result) == []
    result.rows[0].months = {9: {"1"}}                                 # as if calendar months were used
    assert any("month columns" in p for p in reconcile(result))


# ------------------------------------------------------------------ empty files
def test_file_without_transactions_gets_a_row_of_zeros_and_moves(tmp_path, capsys):
    write_erp_csv(tmp_path / "in" / "a.csv", [row("10-03-2026", "1")])
    write_erp_csv(tmp_path / "in" / "header_only.csv", [])
    (tmp_path / "in" / "zero_bytes.csv").write_bytes(b"")
    assert _run(tmp_path) == EXIT_OK
    (report,) = (tmp_path / "out").glob(f"*/{REPORT_NAME}")
    assert _rows(report)[1:] == [["a.csv", 0, 0, 1, 1], ["header_only.csv", 0, 0, 0, 0],
                                   ["zero_bytes.csv", 0, 0, 0, 0]]
    assert list((tmp_path / "in").iterdir()) == []
    assert sorted(p.name for p in (tmp_path / "processed").rglob("*.csv")) == [
        "a.csv", "header_only.csv", "zero_bytes.csv"]
    assert "zero_bytes.csv: no transactions, all months 0" in capsys.readouterr().out


def test_only_empty_files_give_file_rows_without_month_columns(tmp_path):
    write_erp_csv(tmp_path / "input" / "e.csv", [])
    result = run_file_customers(_config(tmp_path))
    assert _rows(result.report_path) == [["اسم الملف", YEAR_COLUMN], ["e.csv", 0]]
    assert verify_file_customers_report(result) == []


# ------------------------------------------------------------------ errors
def test_row_without_customer_code_is_rejected_with_line_and_file_stays(tmp_path, capsys):
    write_erp_csv(tmp_path / "in" / "bad.csv", [row("01-01-2026", "1"), row("02-01-2026", "", doc="2")])
    write_erp_csv(tmp_path / "in" / "good.csv", [row("01-01-2026", "2")])
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert "REJECTED bad.csv line 10: العميل: empty" in capsys.readouterr().out
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["bad.csv"]
    assert [p.name for p in (tmp_path / "processed").rglob("*.csv")] == ["good.csv"]
    (report,) = (tmp_path / "out").glob(f"*/{REPORT_NAME}")
    assert _rows(report)[1:] == [["bad.csv", 1, 1], ["good.csv", 1, 1]]  # valid rows still counted


def test_invalid_date_rejected_with_line_and_file_stays(tmp_path, capsys):
    write_erp_csv(tmp_path / "in" / "bad.csv", [row("xx-01-2026", "1"), row("01-01-2026", "2", doc="2")])
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert "REJECTED bad.csv line 9: التاريخ" in capsys.readouterr().out
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["bad.csv"]


def test_unreadable_file_fails_stays_in_input_and_has_no_row(tmp_path, capsys):
    (tmp_path / "in").mkdir()
    (tmp_path / "in" / "broken.csv").write_bytes("لا يوجد صف عناوين;x\r\n1;2\r\n".encode("cp1256"))
    write_erp_csv(tmp_path / "in" / "good.csv", [row("01-01-2026", "1")])
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert "FAILED  broken.csv  ->  MISSING_REQUIRED_COLUMNS" in capsys.readouterr().out
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["broken.csv"]
    (report,) = (tmp_path / "out").glob(f"*/{REPORT_NAME}")
    assert _rows(report)[1:] == [["good.csv", 1, 1]]


def test_missing_customer_column_fails_the_file(tmp_path):
    header = ["xx" if h == "العميل" else h for h in ERP_HEADER]
    write_erp_csv(tmp_path / "input" / "m.csv", [row("01-01-2026", "1")], header=header)
    result = run_file_customers(_config(tmp_path))
    assert result.failed[0].error == "MISSING_COLUMN: required column 'العميل' not found"
    assert result.rows == []


# ------------------------------------------------------------------ validation
def test_reconcile_detects_wrong_counts_rows_and_names(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [row("01-01-2026", "1"), row("01-02-2026", "2", doc="2")])
    write_erp_csv(tmp_path / "input" / "b.csv", [row("01-01-2026", "1")])
    result = run_file_customers(_config(tmp_path))
    assert reconcile(result) == []
    result.rows[1].months[1].add("1 ")                                  # an extra customer in b / يناير
    assert any("b.csv: يناير 2 differs from 1" in p for p in reconcile(result))
    result.rows[1].months[1].discard("1 ")
    result.rows[0].item.file_name = "other.csv"
    assert any("file names" in p for p in reconcile(result))
    result.rows[0].item.file_name = "a.csv"
    del result.rows[1]
    assert any("file rows" in p for p in reconcile(result))


def test_verify_detects_a_changed_cell_or_name(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [row("01-01-2026", "1")])
    result = run_file_customers(_config(tmp_path))
    wb = load_workbook(result.report_path)
    wb[REPORT_SHEET]["B2"] = 2
    wb[REPORT_SHEET]["A2"] = "x.csv"
    wb.save(result.report_path)
    problems = verify_file_customers_report(result)
    assert any("cell 2 differs from 1" in p for p in problems)
    assert any("file name cell 'x.csv'" in p for p in problems)


def test_report_format(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [row("01-01-2026", "1"), row("01-02-2026", "1")])
    ws = load_workbook(run_file_customers(_config(tmp_path)).report_path)[REPORT_SHEET]
    assert ws.sheet_view.rightToLeft and ws.freeze_panes == "A2"
    assert ws["A1"].fill.fgColor.rgb.endswith("1F4E78") and ws["A1"].font.bold
    assert ws["B2"].number_format == "#,##0" and ws.max_row == 2 and ws.max_column == 4   # no total row
    assert ws["D1"].value == YEAR_COLUMN and ws["D2"].value == 1 and ws["D2"].number_format == "#,##0"


# ------------------------------------------------------------------ menu and lifecycle
def _run(tmp_path, *extra):
    return main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"), "--task", "7",
                 "--non-interactive", *extra])


def test_menu_choice_7_asks_nothing_else_and_moves_files(tmp_path, capsys):
    path = write_erp_csv(tmp_path / "in" / "a.csv", [row("01-01-2026", "1")])
    before = path.read_bytes()
    answers, asked = ["7"], []

    def fake_input(prompt):
        asked.append(prompt)
        return answers.pop(0)

    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")], input_fn=fake_input)
    assert code == EXIT_OK and asked == ["Select task number: "]
    out = capsys.readouterr().out
    assert "6. صافي مبيعات العملاء الشهري" in out and "7. عدد العملاء الشهري لكل ملف" in out
    (folder,) = (tmp_path / "out").iterdir()
    (archive,) = (tmp_path / "processed").iterdir()
    assert folder.name.startswith(f"{TASK_PREFIX}_") and archive.name == folder.name
    assert TASK_PREFIX == "عملاء_الملفات_الشهري"
    assert [p.name for p in folder.iterdir()] == [REPORT_NAME]
    assert (archive / "a.csv").read_bytes() == before and list((tmp_path / "in").iterdir()) == []
    journal = tmp_path / "state" / JOURNAL_NAME
    assert JOURNAL_NAME == "سجل_عملاء_الملفات_الشهري.jsonl"
    events = [json.loads(line)["event"] for line in journal.read_text(encoding="utf-8").splitlines()]
    assert events == ["RUN_STARTED", "COMMIT", "MOVED", "RUN_FINISHED"]
    assert [p.name for p in (tmp_path / "state").glob("*.jsonl")] == [JOURNAL_NAME]


def test_date_arguments_are_rejected(tmp_path):
    (tmp_path / "in").mkdir()
    assert _run(tmp_path, "--ref-date", "01/01/2026") == EXIT_USAGE
    assert _run(tmp_path, "--start-date", "01-01-2026") == EXIT_USAGE
    assert _run(tmp_path, "--report-type", "2") == EXIT_USAGE


def test_report_failure_moves_nothing(tmp_path, monkeypatch):
    write_erp_csv(tmp_path / "in" / "a.csv", [row("01-01-2026", "1")])
    monkeypatch.setattr(monthly_file_customers, "verify_file_customers_report", lambda result: ["header mismatch"])
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["a.csv"]
    assert not (tmp_path / "processed").exists() or not list((tmp_path / "processed").rglob("*.csv"))


def test_same_file_again_is_counted_again_and_nothing_overwritten(tmp_path):
    content = write_erp_csv(tmp_path / "in" / "a.csv", [row("01-01-2026", "1")]).read_bytes()
    assert _run(tmp_path) == EXIT_OK
    (tmp_path / "in" / "a.csv").write_bytes(content)
    (tmp_path / "in" / "copy.csv").write_bytes(content)
    assert _run(tmp_path) == EXIT_OK                                   # no duplicate detection
    first, second = sorted((tmp_path / "out").iterdir())
    assert _rows(first / REPORT_NAME)[1:] == [["a.csv", 1, 1]]
    assert _rows(second / REPORT_NAME)[1:] == [["a.csv", 1, 1], ["copy.csv", 1, 1]]
    assert sorted(p.name for p in (tmp_path / "processed").rglob("*.csv")) == ["a.csv", "a.csv", "copy.csv"]


def test_no_move_and_empty_input(tmp_path):
    path = write_erp_csv(tmp_path / "in" / "a.csv", [row("01-01-2026", "1")])
    assert _run(tmp_path, "--no-move") == EXIT_OK
    assert path.exists() and not (tmp_path / "processed").exists() and not (tmp_path / "state").exists()
    path.unlink()
    assert _run(tmp_path) == EXIT_NO_INPUT
