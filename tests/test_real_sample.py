"""Regression test on the real alex8.csv content (rebuilt in its original ERP layout, cp1256).

Expected values were computed independently (awk over the transcription) and the
transcription itself reconciles with the original file's footer:
235 documents, اجمالى 1,780,805.00, صافى 1,105,646.62.
"""

from decimal import Decimal

from sales_etl.models import CheckStatus
from tests.conftest import alex8_rows, check, outcome, write_erp_csv

FOOTER_NET = Decimal("1105646.62")


def _write(tmp_path, name="alex8.csv"):
    return write_erp_csv(tmp_path / "input" / name, alex8_rows(), footer_net=FOOTER_NET, footer_count=235)


def test_alex8_reference_08_10_2026(run, tmp_path):
    _write(tmp_path)
    result = run("08/10/2026")
    o = outcome(result, "alex8.csv")
    m = o.metrics
    assert o.parsed.encoding == "cp1256" and o.parsed.delimiter == ";"
    assert len(o.parsed.transactions) == 235 and not o.parsed.rejected
    assert m.sales_yesterday == Decimal("108163.37")
    assert m.sales_previous == Decimal("432555.72")
    assert m.returns_yesterday == Decimal("0")
    assert m.returns_previous == Decimal("-178019.57")
    assert m.future_count == 163
    assert check(result, "ERP_CONTROL_TOTAL_NET")[0].status is CheckStatus.PASS
    assert check(result, "ERP_CONTROL_DOC_COUNT")[0].status is CheckStatus.PASS
    assert result.audit_status is CheckStatus.PASS


def test_real_erp_footer_labels_are_control_totals(run, tmp_path):
    """The real export prints 'عدد الفواتير' / 'إجمالى عام' (seen in alex8.csv and kenna9.csv)."""
    write_erp_csv(tmp_path / "input" / "alex8.csv", alex8_rows(), footer_net=FOOTER_NET,
                  footer_count=235, footer_labels=("عدد الفواتير", "إجمالى عام"))
    result = run("18/10/2026")
    o = outcome(result, "alex8.csv")
    assert not o.parsed.rejected
    assert (o.parsed.control.doc_count, o.parsed.control.net_total) == (235, FOOTER_NET)
    assert check(result, "ERP_CONTROL_TOTAL_NET")[0].status is CheckStatus.PASS
    assert o.metrics.sales_yesterday == Decimal("8772.84")      # 18/10/2026 = yesterday
    assert o.metrics.sales_previous == Decimal("1279161.34")
    assert o.metrics.returns_previous == Decimal("-182287.56")
    assert result.audit_status is CheckStatus.PASS


def test_alex8_total_columns_match_erp_footer(run, tmp_path):
    """18/10/2026: no row is dated later, so الصافي must equal the ERP 'إجمالى عام' of صافى."""
    from openpyxl import load_workbook

    write_erp_csv(tmp_path / "input" / "alex8.csv", alex8_rows(), footer_net=FOOTER_NET,
                  footer_count=235, footer_labels=("عدد الفواتير", "إجمالى عام"))
    result = run("18/10/2026")
    row = list(list(load_workbook(result.report_path, data_only=True)["التقرير"].iter_rows(values_only=True))[1])
    assert row[:4] == [1287934.18, -182287.56, 1105646.62, "alex8.csv"]  # 8,772.84 + 1,279,161.34
    assert row[4:8] == [8772.84, 1279161.34, 0, -182287.56]               # existing columns unchanged
    assert Decimal(str(row[2])) == FOOTER_NET


def test_alex8_all_dates_included_when_reference_is_period_end(run, tmp_path):
    _write(tmp_path)
    m = outcome(run("19/10/2026"), "alex8.csv").metrics
    total = m.sales_yesterday + m.sales_previous + m.returns_yesterday + m.returns_previous
    assert total == FOOTER_NET and m.future_count == 0
