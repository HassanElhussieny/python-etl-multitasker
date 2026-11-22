"""Regression test on a synthetic 235-document export (tests/data/branch_a_compact.txt, fictional values) rebuilt in
the ERP layout (cp1256).

Expected values were computed independently (awk over the transcription) and the
transcription itself reconciles with the original file's footer:
235 documents, اجمالى 2,660,538.01, صافى 2,434,046.23.
"""

from decimal import Decimal

from sales_etl.models import CheckStatus
from tests.conftest import branch_a_rows, check, outcome, write_erp_csv

FOOTER_NET = Decimal("2434046.23")


def _write(tmp_path, name="branch_a.csv"):
    return write_erp_csv(tmp_path / "input" / name, branch_a_rows(), footer_net=FOOTER_NET, footer_count=235)


def test_branch_a_reference_08_10_2026(run, tmp_path):
    _write(tmp_path)
    result = run("08/10/2026")
    o = outcome(result, "branch_a.csv")
    m = o.metrics
    assert o.parsed.encoding == "cp1256" and o.parsed.delimiter == ";"
    assert len(o.parsed.transactions) == 235 and not o.parsed.rejected
    assert m.sales_yesterday == Decimal("47760.09")
    assert m.sales_previous == Decimal("627717.40")
    assert m.returns_yesterday == Decimal("0")
    assert m.returns_previous == Decimal("-113577.28")
    assert m.future_count == 163
    assert check(result, "ERP_CONTROL_TOTAL_NET")[0].status is CheckStatus.PASS
    assert check(result, "ERP_CONTROL_DOC_COUNT")[0].status is CheckStatus.PASS
    assert result.audit_status is CheckStatus.PASS


def test_real_erp_footer_labels_are_control_totals(run, tmp_path):
    """The ERP export prints 'عدد الفواتير' / 'إجمالى عام' (seen in branch_a.csv and branch_b.csv)."""
    write_erp_csv(tmp_path / "input" / "branch_a.csv", branch_a_rows(), footer_net=FOOTER_NET,
                  footer_count=235, footer_labels=("عدد الفواتير", "إجمالى عام"))
    result = run("18/10/2026")
    o = outcome(result, "branch_a.csv")
    assert not o.parsed.rejected
    assert (o.parsed.control.doc_count, o.parsed.control.net_total) == (235, FOOTER_NET)
    assert check(result, "ERP_CONTROL_TOTAL_NET")[0].status is CheckStatus.PASS
    assert o.metrics.sales_yesterday == Decimal("165628.40")      # 18/10/2026 = yesterday
    assert o.metrics.sales_previous == Decimal("2425858.59")
    assert o.metrics.returns_previous == Decimal("-157440.76")
    assert result.audit_status is CheckStatus.PASS


def test_branch_a_total_columns_match_erp_footer(run, tmp_path):
    """18/10/2026: no row is dated later, so الصافي must equal the ERP 'إجمالى عام' of صافى."""
    from openpyxl import load_workbook

    write_erp_csv(tmp_path / "input" / "branch_a.csv", branch_a_rows(), footer_net=FOOTER_NET,
                  footer_count=235, footer_labels=("عدد الفواتير", "إجمالى عام"))
    result = run("18/10/2026")
    row = list(list(load_workbook(result.report_path, data_only=True)["التقرير"].iter_rows(values_only=True))[1])
    assert row[:4] == [2591486.99, -157440.76, 2434046.23, "branch_a.csv"]  # 165,628.40 + 2,425,858.59
    assert row[4:8] == [165628.40, 2425858.59, 0, -157440.76]               # existing columns unchanged
    assert Decimal(str(row[2])) == FOOTER_NET


def test_branch_a_all_dates_included_when_reference_is_period_end(run, tmp_path):
    _write(tmp_path)
    m = outcome(run("19/10/2026"), "branch_a.csv").metrics
    total = m.sales_yesterday + m.sales_previous + m.returns_yesterday + m.returns_previous
    assert total == FOOTER_NET and m.future_count == 0
