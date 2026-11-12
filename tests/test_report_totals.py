"""The three leading report columns: إجمالي المبيعات، إجمالي المرتجعات، الصافي."""

import json
import shutil
from decimal import Decimal

from openpyxl import load_workbook

from sales_etl.config import GRAND_TOTAL_LABEL, METRIC_COLUMNS, REPORT_COLUMNS, REPORT_FILE_COLUMN
from sales_etl.models import CheckStatus
from sales_etl.report import report_totals, verify_report
from tests.conftest import R, S, Txn, check, write_erp_csv

D = Decimal

# reference date 08/10/2026
FILE_A = [
    Txn("08-10-2026", S, "100.10", "1"), Txn("08-10-2026", S, "200.20", "2"),   # sales yesterday   300.30
    Txn("07-10-2026", S, "50.05", "3"), Txn("01-10-2026", S, "1000.00", "4"),   # sales previous   1050.05
    Txn("08-10-2026", R, "-30.00", "5"),                                        # returns yesterday -30.00
    Txn("05-10-2026", R, "-12.34", "6"), Txn("26-09-2026", R, "-0.66", "7"),    # returns previous  -13.00
    Txn("09-10-2026", S, "999.99", "8"), Txn("19-10-2026", R, "-88.00", "9"),   # future: excluded everywhere
]
FILE_B = [
    Txn("08-10-2026", S, "0.10", "1"), Txn("06-10-2026", S, "5.00", "2"),       # sales  5.10
    Txn("08-10-2026", R, "-0.20", "3"), Txn("02-10-2026", R, "-1.50", "4"),     # returns -1.70
]


def _rows(result):
    ws = load_workbook(result.report_path, data_only=True)["التقرير"]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def test_column_order_new_columns_first_then_existing_unchanged():
    assert REPORT_COLUMNS == (
        "إجمالي المبيعات", "إجمالي المرتجعات", "الصافي",
        "اسم الملف", "إجمالي مبيعات أمس", "إجمالي مبيعات الأيام السابقة",
        "إجمالي مرتجعات أمس", "إجمالي مرتجعات الأيام السابقة",
        "عدد عملاء أمس", "عدد عملاء الأيام السابقة", "إجمالي العملاء المميزين",
    )
    assert REPORT_COLUMNS[3:8] == (REPORT_FILE_COLUMN, *METRIC_COLUMNS.values())


def test_formulas():
    sales_y, sales_p, ret_y, ret_p = D("300.30"), D("1050.05"), D("-30.00"), D("-13.00")
    sales_total, returns_total, net = report_totals([sales_y, sales_p, ret_y, ret_p])
    assert sales_total == sales_y + sales_p == D("1350.35")
    assert returns_total == ret_y + ret_p == D("-43.00")          # sign kept: negative stays negative
    assert net == sales_total + returns_total == D("1307.35")      # literally sales + returns


def test_values_per_file_and_grand_total(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", FILE_A)
    write_erp_csv(tmp_path / "input" / "b.csv", FILE_B)
    result = run("08/10/2026")
    rows = _rows(result)
    assert rows[0] == list(REPORT_COLUMNS)
    assert [r[:8] for r in rows[1:]] == [                           # the 8 money/name columns
        [1350.35, -43, 1307.35, "a.csv", 300.3, 1050.05, -30, -13],
        [5.1, -1.7, 3.4, "b.csv", 0.1, 5, -0.2, -1.5],
        [1355.45, -44.7, 1310.75, GRAND_TOTAL_LABEL, 300.4, 1055.05, -30.2, -14.5],
    ]
    # Grand Total of the new columns = sum of the per-file values in each column
    for col in range(3):
        assert round(sum(r[col] for r in rows[1:-1]), 2) == rows[-1][col]
    assert check(result, "REPORT_READBACK")[0].status is CheckStatus.PASS
    assert result.audit_status is CheckStatus.PASS


def test_return_signs_are_preserved(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "r.csv", [
        Txn("08-10-2026", R, "-10.00", "1"),
        Txn("07-10-2026", R, "3.00", "2"),        # unusual positive return: kept as +3, not negated
    ])
    rows = _rows(run("08/10/2026"))
    assert rows[1][:8] == [0, -7, -7, "r.csv", 0, 0, -10, 3]


def test_cells_are_numbers_with_money_format(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", FILE_A)
    result = run("08/10/2026")
    ws = load_workbook(result.report_path)["التقرير"]
    for row in (2, 3):  # file row and Grand Total row
        for col in (1, 2, 3):
            cell = ws.cell(row=row, column=col)
            assert isinstance(cell.value, (int, float)) and not str(cell.value).startswith("=")
            assert cell.number_format == "#,##0.00;[Red]-#,##0.00;0.00"
    assert ws.cell(row=3, column=4).value == GRAND_TOTAL_LABEL


def test_audit_json_carries_exact_totals(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", FILE_A)
    write_erp_csv(tmp_path / "input" / "b.csv", FILE_B)
    result = run("08/10/2026")
    report = json.loads(result.audit_json_path.read_text(encoding="utf-8"))["report"]
    assert report["columns"] == list(REPORT_COLUMNS)
    assert {k: report["rows"][0][k] for k in ("sales_total", "returns_total", "net_total")} == {
        "sales_total": "1350.35", "returns_total": "-43.00", "net_total": "1307.35"}
    assert report["rows"][0]["sales_yesterday"] == "300.30"   # existing fields unchanged
    assert report["grand_total_columns"] == {
        "sales_total": "1355.45", "returns_total": "-44.70", "net_total": "1310.75"}


def test_readback_check_covers_the_new_columns(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", FILE_A)
    result = run("08/10/2026")
    tampered = tmp_path / "tampered.xlsx"
    shutil.copy(result.report_path, tampered)
    wb = load_workbook(tampered)
    wb["التقرير"].cell(row=2, column=3).value = 1.0   # wrong الصافي
    wb.save(tampered)
    outcome = verify_report(tampered, result)
    assert outcome.status is CheckStatus.FAIL and "1307.35" in outcome.detail
