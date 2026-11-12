"""Task 6: صافي مبيعات العملاء الشهري - net per customer code per month (A date, F/G customer, J rep, T net)."""

import json
from datetime import date
from decimal import Decimal

import pytest
from openpyxl import load_workbook

from sales_etl import monthly_customer
from sales_etl.cli import EXIT_AUDIT_FAILED, EXIT_NO_INPUT, EXIT_OK, EXIT_USAGE, main
from sales_etl.config import PipelineConfig
from sales_etl.monthly_customer import (
    FIXED_COLUMNS,
    JOURNAL_NAME,
    MONTHS,
    REPORT_NAME,
    REPORT_SHEET,
    TASK_PREFIX,
    reconcile,
    run_monthly,
    verify_monthly_report,
)
from tests.conftest import ERP_HEADER, R, S, alex8_rows, write_erp_csv

D = Decimal


def row(day, kind, net, code, name="", rep="", doc="1"):
    """ERP row (21 cells): A date, B type, F code, G name, I rep code, J rep name, T صافى."""
    return [day, kind, doc, "93", "", code, name, "", "7", rep, net, *["0"] * 8, net, ""]


ROWS = [
    row("05-01-2026", S, "100.00", "400100", "عميل أ", "مندوب 1", "1"),
    row("20-01-2026", R, "-30.00", "400100", "عميل أ", "مندوب 2", "2"),     # second representative
    row("10-03-2026", S, "50.50", "400100", "عميل أ", "مندوب 1", "3"),      # first one again: shown once
    row("01-03-2026", S, "70.00", "400200", "عميل ب", "مندوب 2", "4"),
    row("15-03-2026", R, "-70.00", "400200", "عميل ب", "مندوب 2", "5"),     # net 0: still a row
    row("28-04-2026", R, "5.00", "400300", "عميل أ", "مندوب 3", "6"),       # same name, other code; + return
    row("12-02-2026", "تسوية", "999.00", "400100", "عميل أ", "مندوب 9", "7"),  # other type: not counted
]


def _config(tmp_path):
    return PipelineConfig(input_dir=tmp_path / "input", output_dir=tmp_path / "output",
                          reference_date=date(2026, 1, 1))


def _rows(path):
    wb = load_workbook(path, data_only=True)
    assert wb.sheetnames == [REPORT_SHEET]
    return [list(r) for r in wb[REPORT_SHEET].iter_rows(values_only=True)]


# ------------------------------------------------------------------ calculation
def test_monthly_net_per_customer(tmp_path):
    write_erp_csv(tmp_path / "input" / "year.csv", ROWS)
    result = run_monthly(_config(tmp_path))
    rows = _rows(result.report_path)
    assert result.report_path.name == REPORT_NAME == "تقرير صافي مبيعات العملاء الشهري.xlsx"
    assert rows[0] == ["كود العميل", "اسم العميل", "المندوب", "يناير", "فبراير", "مارس", "أبريل"]
    assert rows[1:] == [
        ["400100", "عميل أ", "مندوب 1، مندوب 2", 70, 0, 50.5, 0],   # يناير = 100 - 30
        ["400200", "عميل ب", "مندوب 2", 0, 0, 0, 0],
        ["400300", "عميل أ", "مندوب 3", 0, 0, 0, 5],               # positive return: sign kept
    ]
    assert result.other_type_rows == 1 and result.complete
    assert verify_monthly_report(result) == []


def test_columns_start_in_january_and_stop_at_the_last_month(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row("30-09-2026", S, "10.00", "1", "x", "r"),
                                                 row("02-10-2026", R, "-4.00", "1", "x", "r", "2")])
    rows = _rows(run_monthly(_config(tmp_path)).report_path)
    assert rows[0] == [*FIXED_COLUMNS, *MONTHS[:10]]                     # يناير .. أكتوبر, no نوفمبر
    assert rows[1] == ["1", "x", "r", 0, 0, 0, 0, 0, 0, 0, 0, 10, -4]


def test_future_dates_are_counted(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row("31-12-2026", S, "1.00", "1", "x", "r")])
    rows = _rows(run_monthly(_config(tmp_path)).report_path)
    assert rows[0][-1] == "ديسمبر" and rows[1][-1] == 1


def test_multiple_files_are_combined_per_code(tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", ROWS[:2])
    write_erp_csv(tmp_path / "input" / "b.csv", [row("07-01-2026", S, "0.10", "400100", "عميل أ", "مندوب 4")])
    rows = _rows(run_monthly(_config(tmp_path)).report_path)
    assert rows[1:] == [["400100", "عميل أ", "مندوب 1، مندوب 4، مندوب 2", 70.1]]   # 05-01, 07-01, 20-01


def test_representatives_ordered_by_date_not_by_file_name(tmp_path):
    """3omla210.csv (October) is read before 3omla29.csv (September): September's rep still comes first."""
    write_erp_csv(tmp_path / "input" / "3omla210.csv", [row("02-10-2026", S, "1.00", "1", "x", "مندوب أكتوبر")])
    write_erp_csv(tmp_path / "input" / "3omla29.csv", [row("15-09-2026", S, "1.00", "1", "x", "مندوب سبتمبر")])
    result = run_monthly(_config(tmp_path))
    assert [f.file_name for f in result.files] == ["3omla210.csv", "3omla29.csv"]
    assert _rows(result.report_path)[1][2] == "مندوب سبتمبر، مندوب أكتوبر"
    assert verify_monthly_report(result) == []


def test_real_alex8_content_net_equals_erp_footer(tmp_path):
    """All 235 real transactions: the sum of every monthly net = the file's ERP صافى footer."""
    write_erp_csv(tmp_path / "input" / "alex8.csv", alex8_rows(), footer_net=D("1105646.62"), footer_count=235)
    result = run_monthly(_config(tmp_path))
    total = sum((v for r in monthly_customer.monthly_table(result) for v in r[3:]), D(0))
    assert total == D("1105646.62") and verify_monthly_report(result) == []


# ------------------------------------------------------------------ customers and representatives
def test_codes_are_kept_exactly_as_written(tmp_path):
    rows = [row("01-01-2026", S, "1.00", "400100", "a", "r"), row("01-01-2026", S, "2.00", "0400100", "a", "r", "2"),
            row("01-01-2026", S, "4.00", "٤٠٠١٠٠", "a", "r", "3")]
    write_erp_csv(tmp_path / "input" / "m.csv", rows, encoding="utf-8")
    table = _rows(run_monthly(_config(tmp_path)).report_path)[1:]
    assert sorted(r[0] for r in table) == sorted(["400100", "0400100", "٤٠٠١٠٠"])       # three customers
    assert {r[0]: r[3] for r in table} == {"400100": 1, "0400100": 2, "٤٠٠١٠٠": 4}


def test_same_name_different_codes_are_not_merged(tmp_path):
    write_erp_csv(tmp_path / "input" / "m.csv", [row("01-01-2026", S, "1.00", "1", "نفس الاسم", "r"),
                                                 row("01-01-2026", S, "2.00", "2", "نفس الاسم", "r", "2")])
    assert [r[:4] for r in _rows(run_monthly(_config(tmp_path)).report_path)[1:]] == [
        ["1", "نفس الاسم", "r", 1], ["2", "نفس الاسم", "r", 2]]


def test_representatives_first_appearance_no_repeats_no_blanks(tmp_path):
    rows = [row("01-02-2026", S, "1.00", "1", "x", "احمد علي"),
            row("01-01-2026", S, "1.00", "1", "x", "", "2"),                 # blank J: not listed
            row("01-03-2026", S, "1.00", "1", "x", "سامي", "3"),
            row("01-04-2026", S, "1.00", "1", "x", "أحمد  علي", "4"),       # same name, other spelling
            row("01-05-2026", S, "1.00", "1", "x", "سامي", "5")]
    write_erp_csv(tmp_path / "input" / "m.csv", rows)
    result = run_monthly(_config(tmp_path))
    assert _rows(result.report_path)[1][2] == "احمد علي، سامي"
    assert verify_monthly_report(result) == []


def test_name_differences_use_first_name_and_warn(tmp_path, capsys):
    write_erp_csv(tmp_path / "in" / "m.csv", [row("01-01-2026", S, "1.00", "1", "اسم أول", "r"),
                                              row("01-02-2026", S, "1.00", "1", "اسم ثان", "r", "2")])
    assert main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"), "--task", "6",
                 "--non-interactive", "--no-move"]) == EXIT_OK
    assert "WARNING code 1 has 2 names; the report uses 'اسم أول'" in capsys.readouterr().out


# ------------------------------------------------------------------ rejected rows and failures
def test_row_without_customer_code_is_rejected_and_file_stays(tmp_path, capsys):
    write_erp_csv(tmp_path / "in" / "bad.csv", [row("01-01-2026", S, "5.00", "1", "x", "r"),
                                                row("02-01-2026", S, "9.00", "", "بدون كود", "r", "2")])
    write_erp_csv(tmp_path / "in" / "good.csv", [row("01-01-2026", S, "1.00", "2", "y", "r")])
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    out = capsys.readouterr().out
    assert "bad.csv line" in out and "without a customer code" in out
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["bad.csv"]
    assert [p.name for p in (tmp_path / "processed").rglob("*.csv")] == ["good.csv"]
    (report,) = (tmp_path / "out").glob(f"*/{REPORT_NAME}")
    assert _rows(report)[1:] == [["1", "x", "r", 5], ["2", "y", "r", 1]]          # the 9.00 is not counted


def test_invalid_date_or_amount_rejected_and_file_stays(tmp_path):
    write_erp_csv(tmp_path / "in" / "bad.csv", [row("xx-01-2026", S, "5.00", "1", "x", "r"),
                                                row("01-01-2026", S, "abc", "1", "x", "r", "2")])
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["bad.csv"]


def test_missing_representative_column_fails_the_file(tmp_path):
    header = ["xx" if h == "المندوب" else h for h in ERP_HEADER]
    write_erp_csv(tmp_path / "input" / "m.csv", [row("01-01-2026", S, "1.00", "1", "x", "r")], header=header)
    result = run_monthly(_config(tmp_path))
    assert result.failed[0].error == "MISSING_COLUMN: required column 'المندوب' not found"


# ------------------------------------------------------------------ validation
def test_reconcile_detects_wrong_month_rep_and_total(tmp_path):
    write_erp_csv(tmp_path / "input" / "year.csv", ROWS)
    result = run_monthly(_config(tmp_path))
    assert reconcile(result) == []
    result.customers["400200"].months[2] = D("1.00")
    result.customers["400100"].reps_seen["x"] = (date(2026, 1, 1), -1, "مندوب 1")      # listed twice
    problems = reconcile(result)
    assert any("code 400200: فبراير" in p for p in problems)
    assert any("code 400100: representatives" in p for p in problems)
    assert any("sum of all monthly nets" in p for p in problems)
    del result.customers["400300"]
    assert any("customer rows" in p for p in reconcile(result))


def test_verify_detects_a_changed_cell(tmp_path):
    write_erp_csv(tmp_path / "input" / "year.csv", ROWS)
    result = run_monthly(_config(tmp_path))
    wb = load_workbook(result.report_path)
    wb[REPORT_SHEET]["D2"] = 71
    wb.save(result.report_path)
    assert any("code 400100" in p for p in verify_monthly_report(result))


def test_report_format(tmp_path):
    write_erp_csv(tmp_path / "input" / "year.csv", ROWS)
    ws = load_workbook(run_monthly(_config(tmp_path)).report_path)[REPORT_SHEET]
    assert ws.sheet_view.rightToLeft and ws.freeze_panes == "A2"
    assert ws["A1"].fill.fgColor.rgb.endswith("1F4E78") and ws["A1"].font.bold
    assert ws["D2"].number_format.startswith("#,##0.00") and ws.max_row == 4               # no total row
    assert ws.max_column == 7                                                               # no total column


# ------------------------------------------------------------------ menu and lifecycle
def _run(tmp_path, *extra):
    return main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out"), "--task", "6",
                 "--non-interactive", *extra])


def test_menu_choice_6_asks_nothing_else_and_moves_file(tmp_path, capsys):
    path = write_erp_csv(tmp_path / "in" / "year.csv", ROWS)
    before = path.read_bytes()
    answers, asked = ["6"], []

    def fake_input(prompt):
        asked.append(prompt)
        return answers.pop(0)

    code = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")], input_fn=fake_input)
    assert code == EXIT_OK and asked == ["Select task number: "]
    assert "6. صافي مبيعات العملاء الشهري" in capsys.readouterr().out
    (folder,) = (tmp_path / "out").iterdir()
    (archive,) = (tmp_path / "processed").iterdir()
    assert folder.name.startswith(f"{TASK_PREFIX}_") and archive.name == folder.name
    assert [p.name for p in folder.iterdir()] == [REPORT_NAME]
    assert (archive / "year.csv").read_bytes() == before and list((tmp_path / "in").iterdir()) == []
    journal = tmp_path / "state" / JOURNAL_NAME
    assert JOURNAL_NAME == "سجل_صافي_العملاء_الشهري.jsonl"
    events = [json.loads(line)["event"] for line in journal.read_text(encoding="utf-8").splitlines()]
    assert events == ["RUN_STARTED", "COMMIT", "MOVED", "RUN_FINISHED"]
    for other in ("processing", "period", "customer", "representative", "task5"):
        assert not (tmp_path / "state" / f"{other}_journal.jsonl").exists()


def test_date_arguments_are_rejected(tmp_path):
    (tmp_path / "in").mkdir()
    assert _run(tmp_path, "--ref-date", "01/01/2026") == EXIT_USAGE
    assert _run(tmp_path, "--start-date", "01-01-2026") == EXIT_USAGE
    assert _run(tmp_path, "--report-type", "2") == EXIT_USAGE


def test_report_failure_moves_nothing(tmp_path, monkeypatch):
    write_erp_csv(tmp_path / "in" / "year.csv", ROWS)
    monkeypatch.setattr(monthly_customer, "verify_monthly_report", lambda result: ["header mismatch"])
    assert _run(tmp_path) == EXIT_AUDIT_FAILED
    assert [p.name for p in (tmp_path / "in").iterdir()] == ["year.csv"]
    assert not (tmp_path / "processed").exists() or not list((tmp_path / "processed").rglob("*.csv"))


def test_same_file_again_is_counted_again_and_nothing_overwritten(tmp_path):
    content = write_erp_csv(tmp_path / "in" / "year.csv", ROWS).read_bytes()
    assert _run(tmp_path) == EXIT_OK
    (tmp_path / "in" / "year.csv").write_bytes(content)
    (tmp_path / "in" / "copy.csv").write_bytes(content)
    assert _run(tmp_path) == EXIT_OK                               # no duplicate detection
    second = sorted((tmp_path / "out").iterdir())[-1] / REPORT_NAME
    assert _rows(second)[1][3:] == [140, 0, 101, 0]                  # both copies counted
    assert sorted(p.name for p in (tmp_path / "processed").rglob("*.csv")) == ["copy.csv", "year.csv", "year.csv"]


def test_no_move_and_empty_input(tmp_path):
    path = write_erp_csv(tmp_path / "in" / "year.csv", ROWS)
    assert _run(tmp_path, "--no-move") == EXIT_OK
    assert path.exists() and not (tmp_path / "processed").exists() and not (tmp_path / "state").exists()
    path.unlink()
    assert _run(tmp_path) == EXIT_NO_INPUT
