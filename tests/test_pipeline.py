"""End-to-end scenarios with small synthetic datasets and known expected totals."""

from decimal import Decimal

from openpyxl import load_workbook

from sales_etl.config import GRAND_TOTAL_LABEL, REPORT_COLUMNS
from sales_etl.models import CheckStatus
from tests.conftest import R, S, Txn, check, outcome, write_erp_csv

D = Decimal

# Known dataset relative to reference date 08/10/2026
MIXED = [
    Txn("08-10-2026", S, "100.10", "11"),      # sales yesterday
    Txn("08-10-2026", S, "200.20", "12"),      # sales yesterday
    Txn("07-10-2026", S, "50.05", "13"),       # sales previous
    Txn("01-10-2026", S, "1000.00", "14"),     # sales previous
    Txn("26-09-2026", S, "0.00", "15"),        # sales previous (zero net, series 96 style)
    Txn("08-10-2026", R, "-30.00", "21"),      # returns yesterday
    Txn("05-10-2026", R, "-12.34", "22"),      # returns previous
    Txn("26-09-2026", R, "-0.66", "23"),       # returns previous
    Txn("09-10-2026", S, "999.99", "16"),      # future -> excluded
    Txn("19-10-2026", R, "-88.00", "24"),      # future -> excluded
]


def _report_values(result):
    ws = load_workbook(result.report_path, data_only=True)["التقرير"]
    return [list(r) for r in ws.iter_rows(values_only=True)]


NEW = 3  # leading columns: إجمالي المبيعات، إجمالي المرتجعات، الصافي


def _original_columns(rows):
    """The five pre-existing columns (اسم الملف + 4 metrics), exactly as before."""
    return [r[NEW:NEW + 5] for r in rows]


# Scenarios 1-5, 12: the four metrics, future exclusion, sign of returns
def test_four_metrics_and_future_exclusion(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "rep1.csv", MIXED)
    result = run("08/10/2026")
    m = outcome(result, "rep1.csv").metrics
    assert m.sales_yesterday == D("300.30")
    assert m.sales_previous == D("1050.05")
    assert m.returns_yesterday == D("-30.00")       # negative preserved, not negated
    assert m.returns_previous == D("-13.00")
    assert m.future_count == 2 and m.future_net == D("911.99")
    assert check(result, "FUTURE_TRANSACTIONS_EXCLUDED")[0].status is CheckStatus.PASS
    assert check(result, "ERP_CONTROL_TOTAL_NET")[0].status is CheckStatus.PASS
    assert result.audit_status is CheckStatus.PASS


def test_reference_date_itself_not_in_previous_days(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [Txn("08-10-2026", S, "5"), Txn("07-10-2026", S, "7")])
    m = outcome(run("08/10/2026"), "f.csv").metrics
    assert (m.sales_yesterday, m.sales_previous) == (D("5"), D("7"))


# Scenario 6
def test_no_transactions_on_reference_date_gives_zero(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [Txn("01-10-2026", S, "10"), Txn("02-10-2026", R, "-4")])
    result = run("08/10/2026")
    m = outcome(result, "f.csv").metrics
    assert m.sales_yesterday == 0 and m.returns_yesterday == 0
    assert isinstance(m.sales_yesterday, Decimal)
    row = _original_columns(_report_values(result))[1]
    assert row == ["f.csv", 0, 10, 0, -4]


# Scenario 7
def test_no_transactions_before_reference_date_gives_zero(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [Txn("08-10-2026", S, "10"), Txn("08-10-2026", R, "-4")])
    m = outcome(run("08/10/2026"), "f.csv").metrics
    assert m.sales_previous == 0 and m.returns_previous == 0
    assert (m.sales_yesterday, m.returns_yesterday) == (D("10"), D("-4"))


# Scenario 8
def test_empty_file_fails_but_batch_continues(run, tmp_path):
    (tmp_path / "input" / "empty.csv").write_bytes(b"")
    write_erp_csv(tmp_path / "input" / "good.csv", MIXED)
    result = run()
    assert outcome(result, "empty.csv").error_code == "EMPTY_FILE"
    assert outcome(result, "good.csv").succeeded
    assert [r[NEW] for r in _report_values(result)[1:]] == ["good.csv", GRAND_TOTAL_LABEL]
    assert check(result, "ALL_FILES_PROCESSED")[0].status is CheckStatus.FAIL
    assert result.audit_status is CheckStatus.FAIL  # an incomplete batch is never reported as PASS


def test_header_only_file_is_processed_with_zeros(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "header_only.csv", [], footer=False)
    result = run()
    o = outcome(result, "header_only.csv")
    assert o.succeeded and o.metrics.sales_yesterday == 0
    assert any(i.code == "NO_TRANSACTIONS" for i in o.parsed.issues)
    assert _original_columns(_report_values(result))[1] == ["header_only.csv", 0, 0, 0, 0]


# Scenario 9
def test_invalid_dates_and_amounts_are_rejected_and_logged(run, tmp_path):
    rows = [
        Txn("08-10-2026", S, "100"),
        Txn("31-02-2026", S, "999"),          # invalid calendar date
        Txn("08-10-2026", S, "abc"),          # invalid monetary value
        Txn("08-10-2026", S, ""),             # empty monetary value
        Txn("07-10-2026", R, "-5"),
    ]
    write_erp_csv(tmp_path / "input" / "f.csv", rows, footer_net=D("-0.01"))
    result = run()
    o = outcome(result, "f.csv")
    assert o.metrics.sales_yesterday == D("100")    # bad rows NOT counted as zero or included
    assert o.metrics.returns_previous == D("-5")
    assert sorted(r.line_no for r in o.parsed.rejected) == [10, 11, 12]
    reasons = " | ".join("; ".join(r.reasons) for r in o.parsed.rejected)
    assert "invalid calendar date" in reasons and "invalid amount" in reasons and "empty amount" in reasons
    assert check(result, "NO_REJECTED_ROWS")[0].status is CheckStatus.FAIL
    assert result.audit_status is CheckStatus.FAIL


def test_unexpected_document_type_is_reported_and_excluded(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [Txn("08-10-2026", S, "10"), Txn("08-10-2026", "تحويل", "77")])
    result = run()
    o = outcome(result, "f.csv")
    assert o.metrics.sales_yesterday == D("10") and o.metrics.other_type_net == D("77")
    assert any(i.code == "UNEXPECTED_DOC_TYPE" for i in o.parsed.issues)
    assert check(result, "PARTITION_COMPLETE")[0].status is CheckStatus.PASS


def test_footer_control_total_mismatch_fails(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [Txn("08-10-2026", S, "10")], footer_net=D("11"))
    result = run()
    c = check(result, "ERP_CONTROL_TOTAL_NET")[0]
    assert c.status is CheckStatus.FAIL and c.difference == "-1.00"


def test_column_shift_row_is_rejected(run, tmp_path):
    shifted = Txn("08-10-2026", S, "10").cells()
    shifted.insert(6, "اسم فيه ; فاصلة")  # an extra cell shifts صافى right
    write_erp_csv(tmp_path / "input" / "f.csv", [Txn("08-10-2026", S, "1"), shifted])
    o = outcome(run(), "f.csv")
    assert len(o.parsed.rejected) == 1 and o.metrics.sales_yesterday == D("1")


# Scenario 10 + 11
def test_multiple_files_mixed_encodings_one_reference_date(run, tmp_path):
    inp = tmp_path / "input"
    write_erp_csv(inp / "a.csv", [Txn("08-10-2026", S, "1,000.50", "1"), Txn("01-10-2026", R, "-2", "2")])
    write_erp_csv(inp / "b.csv", [Txn("08/10/2026", S, "3", "3")], encoding="utf-8-sig")
    write_erp_csv(inp / "c.csv", [Txn("09-10-2026", S, "4", "4")], encoding="utf-8", delimiter=",")
    (inp / "notes.txt").write_text("not a csv")
    result = run("08/10/2026")
    assert [p.name for p in result.discovered] == ["a.csv", "b.csv", "c.csv"]
    assert {o.file_name: o.parsed.encoding for o in result.outcomes} == {
        "a.csv": "cp1256", "b.csv": "utf-8-sig", "c.csv": "utf-8"}
    rows = _report_values(result)
    assert rows[0] == list(REPORT_COLUMNS)
    assert _original_columns(rows)[1:] == [
        ["a.csv", 1000.5, 0, 0, -2],
        ["b.csv", 3, 0, 0, 0],
        ["c.csv", 0, 0, 0, 0],             # only a future transaction -> zeros
        [GRAND_TOTAL_LABEL, 1003.5, 0, 0, -2],
    ]
    assert result.audit_status is CheckStatus.PASS


# Scenario 13
def test_grand_total_reconciliation(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", MIXED)
    write_erp_csv(tmp_path / "input" / "b.csv", [Txn("08-10-2026", S, "0.10", "91"),
                                                 Txn("08-10-2026", R, "-0.20", "92")])
    result = run()
    gt = result.grand_total
    assert gt.sales_yesterday == D("300.40")  # 300.30 + 0.10 exactly, no float drift
    assert gt.returns_yesterday == D("-30.20")
    assert gt.sales_previous == D("1050.05") and gt.returns_previous == D("-13.00")
    grand_checks = [c for c in result.checks if c.check.startswith("GRAND_TOTAL ")]
    assert len(grand_checks) == 4 and all(c.status is CheckStatus.PASS for c in grand_checks)
    assert check(result, "REPORT_READBACK")[0].status is CheckStatus.PASS
    assert _original_columns(_report_values(result))[-1] == [GRAND_TOTAL_LABEL, 300.4, 1050.05, -30.2, -13]


# Scenario 14
def test_repeated_rows_across_files_are_all_counted(run, tmp_path):
    shared = Txn("08-10-2026", S, "500", "777", "400555")
    write_erp_csv(tmp_path / "input" / "a.csv", [shared, Txn("08-10-2026", S, "1", "778")])
    write_erp_csv(tmp_path / "input" / "b.csv", [shared])
    result = run()
    assert result.grand_total.sales_yesterday == D("1001")
    assert not any(c.check == "DUPLICATE_SCAN" for c in result.checks)
    assert result.audit_status is CheckStatus.PASS


def test_repeated_rows_and_document_numbers_in_one_file_are_all_counted(run, tmp_path):
    same = Txn("08-10-2026", S, "100", "5", "400001")
    write_erp_csv(tmp_path / "input" / "a.csv", [same, same, Txn("07-10-2026", S, "7", "5", "400002")])
    result = run()
    o = outcome(result, "a.csv")
    assert len(o.parsed.transactions) == 3 and not o.parsed.rejected        # nothing rejected
    assert (o.metrics.sales_yesterday, o.metrics.sales_previous) == (D("200"), D("7"))
    assert result.audit_status is CheckStatus.PASS


def test_identical_files_are_both_counted(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", MIXED)
    write_erp_csv(tmp_path / "input" / "a_copy.csv", MIXED)
    result = run()
    assert [o.file_name for o in result.successful] == ["a.csv", "a_copy.csv"]
    assert result.grand_total.sales_yesterday == 2 * outcome(result, "a.csv").metrics.sales_yesterday


# Safe file handling
def test_sources_untouched_and_runs_never_overwrite(run, tmp_path):
    path = write_erp_csv(tmp_path / "input" / "a.csv", MIXED)
    before = path.read_bytes()
    first, second = run(), run()
    assert path.read_bytes() == before
    assert check(first, "SOURCE_FILES_UNCHANGED")[0].status is CheckStatus.PASS
    assert first.run_dir != second.run_dir and first.report_path.exists() and second.report_path.exists()
    assert sorted(p.name for p in (tmp_path / "input").iterdir()) == ["a.csv"]


def test_outputs_written(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", MIXED)
    result = run()
    import json
    payload = json.loads(result.audit_json_path.read_text(encoding="utf-8"))
    assert payload["audit_status"] == "PASS" and payload["reference_date"] == "2026-10-08"
    assert payload["report"]["rows"][0]["returns_yesterday"] == "-30.00"
    wb = load_workbook(result.audit_xlsx_path)
    assert wb.sheetnames == ["Summary", "Checks", "Files", "Issues", "Transactions", "Rejected rows"]
    assert "duplicates" not in payload
    meta = load_workbook(result.report_path)["بيانات التقرير"]
    assert meta["B2"].value.date().isoformat() == "2026-10-08"
    assert result.log_path.read_text(encoding="utf-8").count("\n") >= 3


def test_no_csv_files_is_a_fail(run, tmp_path):
    result = run()
    assert check(result, "FILES_DISCOVERED")[0].status is CheckStatus.FAIL
    assert result.audit_status is CheckStatus.FAIL
