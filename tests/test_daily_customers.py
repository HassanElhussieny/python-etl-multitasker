"""Task 1: distinct customer counts (العميل, column F) - the three columns at the end of the report.

A customer is counted only with at least one valid مبيعات row; a customer with only مرتجعات is never counted.
"""

import json
import shutil

from openpyxl import load_workbook

from sales_etl.config import CUSTOMER_COUNT_COLUMNS, GRAND_TOTAL_LABEL, REPORT_COLUMNS
from sales_etl.models import CheckStatus, CustomerSets
from sales_etl.reconcile import reconcile_customers
from sales_etl.report import verify_report
from tests.conftest import R, S, Txn, check, outcome, write_erp_csv

# reference date 08/10/2026: "y" = 08-10-2026, "p" = an earlier day, "f" = a later day
Y, P, F = "08-10-2026", "03-10-2026", "09-10-2026"


def t(day, kind, code, net=None, doc="1"):
    return Txn(day, kind, net or ("10.00" if kind == S else "-5.00"), doc, code)


def _rows(result):
    ws = load_workbook(result.report_path, data_only=True)["التقرير"]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _counts(result):
    """{file or Grand Total: [أمس, الأيام السابقة, الإجمالي]} as read from the saved report."""
    return {r[3]: r[8:] for r in _rows(result)[1:]}


def _one(run, tmp_path, rows):
    write_erp_csv(tmp_path / "input" / "f.csv", rows)
    result = run("08/10/2026")
    assert check(result, "CUSTOMER_COUNTS")[0].status is CheckStatus.PASS
    assert result.audit_status is CheckStatus.PASS
    return _counts(result)["f.csv"]


def test_columns_are_appended_last_in_this_order():
    assert REPORT_COLUMNS[8:] == ("عدد عملاء أمس", "عدد عملاء الأيام السابقة",
                                  "إجمالي العملاء المميزين") == tuple(CUSTOMER_COUNT_COLUMNS.values())
    assert "عدد عملاء المرتجعات فقط" not in REPORT_COLUMNS and len(REPORT_COLUMNS) == 11


def test_case_1_sales_only_is_counted(run, tmp_path):
    assert _one(run, tmp_path, [t(Y, S, "X")]) == [1, 0, 1]


def test_case_2_sales_and_returns_counted_once(run, tmp_path):
    assert _one(run, tmp_path, [t(P, S, "X"), t(Y, R, "X", doc="2")]) == [0, 1, 1]


def test_case_3_returns_only_is_excluded_everywhere(run, tmp_path):
    assert _one(run, tmp_path, [t(Y, R, "X"), t(P, R, "X", doc="2")]) == [0, 0, 0]


def test_case_4_many_sales_counted_once(run, tmp_path):
    assert _one(run, tmp_path, [t(Y, S, "X", doc=str(i)) for i in range(5)]) == [1, 0, 1]


def test_case_5_many_sales_and_returns_counted_once(run, tmp_path):
    rows = [t(d, k, "X", doc=f"{d}{k}{i}") for d in (Y, P) for k in (S, R) for i in range(3)]
    assert _one(run, tmp_path, rows) == [1, 1, 1]


def test_case_6_sales_and_returns_yesterday_counted_yesterday(run, tmp_path):
    assert _one(run, tmp_path, [t(Y, S, "X"), t(Y, R, "X", doc="2")]) == [1, 0, 1]


def test_case_7_returns_yesterday_only_not_counted_yesterday(run, tmp_path):
    assert _one(run, tmp_path, [t(Y, R, "X"), t(Y, S, "Z", doc="2")]) == [1, 0, 1]   # only Z


def test_case_8_sales_and_returns_previous_counted_once(run, tmp_path):
    assert _one(run, tmp_path, [t(P, S, "X"), t(P, R, "X", doc="2"), t(P, S, "X", doc="3")]) == [0, 1, 1]


def test_case_9_returns_previous_only_excluded(run, tmp_path):
    assert _one(run, tmp_path, [t(P, R, "X"), t(P, S, "Z", doc="2")]) == [0, 1, 1]   # only Z


def test_case_10_sales_both_periods_once_in_total(run, tmp_path):
    assert _one(run, tmp_path, [t(Y, S, "X"), t(P, S, "X", doc="2")]) == [1, 1, 1]   # total 1, not 2


def test_case_11_returns_yesterday_sales_previous(run, tmp_path):
    assert _one(run, tmp_path, [t(Y, R, "X"), t(P, S, "X", doc="2")]) == [0, 1, 1]


def test_case_12_customer_in_several_files_counts_once_in_grand_total(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [t(Y, S, "X"), t(P, S, "Y", doc="2")])
    write_erp_csv(tmp_path / "input" / "b.csv", [t(Y, S, "X"), t(P, S, "Y", doc="2"), t(P, S, "Z", doc="3")])
    result = run("08/10/2026")
    assert _counts(result) == {"a.csv": [1, 1, 2], "b.csv": [1, 2, 3],
                               GRAND_TOTAL_LABEL: [1, 2, 3]}               # not the sums 2, 3, 5
    assert check(result, "GRAND_TOTAL_CUSTOMERS")[0].status is CheckStatus.PASS


def test_case_13_returns_in_one_file_sales_in_another_qualifies(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [t(Y, R, "W")])
    write_erp_csv(tmp_path / "input" / "b.csv", [t(P, S, "W")])
    counts = _counts(run("08/10/2026"))
    assert counts == {"a.csv": [0, 0, 0], "b.csv": [0, 1, 1], GRAND_TOTAL_LABEL: [0, 1, 1]}


def test_case_14_returns_only_across_all_files_excluded(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [t(Y, R, "X"), t(P, S, "Q", doc="2")])
    write_erp_csv(tmp_path / "input" / "b.csv", [t(P, R, "X")])
    result = run("08/10/2026")
    assert _counts(result)[GRAND_TOTAL_LABEL] == [0, 1, 1]                  # only Q, never X
    assert check(result, "GRAND_TOTAL_CUSTOMERS")[0].status is CheckStatus.PASS


def test_case_15_blank_customer_code_is_not_counted(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [t(Y, S, ""), t(P, R, "", doc="2"), t(P, S, "X", doc="3")])
    row = _rows(run("08/10/2026"))[1]
    assert row[8:] == [0, 1, 1]
    assert row[4:8] == [10, 10, 0, -5]                                     # its amounts are still counted


def test_case_16_rows_after_the_reference_date_are_not_counted(run, tmp_path):
    assert _one(run, tmp_path, [
        t(F, S, "X"), t(F, R, "Y", doc="2"),            # only future rows: not counted at all
        t(P, R, "Z", doc="3"), t(F, S, "Z", doc="4"),   # a future sale does not qualify Z's past return
    ]) == [0, 0, 0]


def test_case_17_existing_numbers_unchanged(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [
        Txn(Y, S, "100.10", "1", "A"), Txn(P, S, "50.05", "2", "B"), Txn(Y, R, "-30.00", "3", "A"),
        Txn(P, R, "-12.34", "4", "C"), Txn(F, S, "999.99", "5", "D")])
    rows = _rows(run("08/10/2026"))
    assert rows[1][:8] == [150.15, -42.34, 107.81, "f.csv", 100.1, 50.05, -30, -12.34]
    assert rows[2][:8] == [150.15, -42.34, 107.81, GRAND_TOTAL_LABEL, 100.1, 50.05, -30, -12.34]
    assert rows[1][8:] == rows[2][8:] == [1, 1, 2]                          # A, B; C has only returns


def test_zero_value_sale_still_qualifies(run, tmp_path):
    assert _one(run, tmp_path, [t(Y, S, "X", net="0.00")]) == [1, 0, 1]


def test_other_document_types_and_invalid_rows_are_not_counted(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [
        t(Y, S, "X"), Txn(Y, "تسوية", "5.00", "2", "Y"),               # other type: not a customer here
        Txn("xx-10-2026", S, "5.00", "3", "Z"), Txn(Y, S, "abc", "4", "W"),  # rejected rows
    ])
    result = run("08/10/2026")
    assert _rows(result)[1][8:] == [1, 0, 1]
    assert check(result, "CUSTOMER_COUNTS")[0].status is CheckStatus.PASS


def test_codes_are_kept_as_written(run, tmp_path):
    assert _one(run, tmp_path, [t(Y, S, "100"), t(Y, S, "0100", doc="2")]) == [2, 0, 2]


def test_empty_and_header_only_files_keep_their_behaviour(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "h.csv", [])
    (tmp_path / "input" / "e.csv").write_bytes(b"")
    write_erp_csv(tmp_path / "input" / "x.csv", [t(Y, S, "X")])
    result = run("08/10/2026")
    assert outcome(result, "e.csv").error_code == "EMPTY_FILE"                # still a file failure
    assert _counts(result) == {"h.csv": [0, 0, 0], "x.csv": [1, 0, 1], GRAND_TOTAL_LABEL: [1, 0, 1]}


# ------------------------------------------------------------------ reconciliation and read-back
def test_reconciliation_detects_wrong_sets(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [t(Y, S, "X"), t(P, R, "Y", doc="2")])
    result = run("08/10/2026")
    o = outcome(result, "f.csv")
    txns = o.parsed.transactions
    assert reconcile_customers("f.csv", "CUSTOMER_COUNTS", o.customers, txns, result.reference_date).status \
        is CheckStatus.PASS
    for wrong in (CustomerSets(yesterday=frozenset({"X", "Q"}), previous=o.customers.previous),
                  CustomerSets(yesterday=o.customers.yesterday, previous=frozenset({"Y"}))):  # returns-only Y
        assert reconcile_customers("f.csv", "CUSTOMER_COUNTS", wrong, txns, result.reference_date).status \
            is CheckStatus.FAIL


def test_grand_total_counts_are_a_union_not_a_sum(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "a.csv", [t(Y, S, "X")])
    write_erp_csv(tmp_path / "input" / "b.csv", [t(Y, S, "X")])
    result = run("08/10/2026")
    file_rows = [r for r in _rows(result)[1:-1]]
    assert sum(r[10] for r in file_rows) == 2 and _rows(result)[-1][10] == 1
    assert result.grand_customers.counts() == [1, 0, 1]


def test_readback_detects_a_changed_count_cell(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [t(Y, S, "X"), t(P, S, "Y", doc="2")])
    result = run("08/10/2026")
    assert check(result, "REPORT_READBACK")[0].status is CheckStatus.PASS
    for col in (9, 10, 11):
        tampered = tmp_path / f"tampered{col}.xlsx"
        shutil.copy(result.report_path, tampered)
        wb = load_workbook(tampered)
        wb["التقرير"].cell(row=3, column=col).value += 1              # a Grand Total count cell
        wb.save(tampered)
        readback = verify_report(tampered, result)
        assert readback.status is CheckStatus.FAIL and "customers" in readback.detail


def test_count_cells_are_integers_with_count_format(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [t(Y, S, "X")])
    ws = load_workbook(run("08/10/2026").report_path)["التقرير"]
    for row in (2, 3):
        for col in range(9, 12):
            cell = ws.cell(row=row, column=col)
            assert type(cell.value) is int and cell.number_format == "#,##0"
        assert ws.cell(row=row, column=12).value is None                    # no fourth count column
    assert ws.cell(row=2, column=8).number_format == "#,##0.00;[Red]-#,##0.00;0.00"   # money unchanged


def test_audit_json_carries_the_counts(run, tmp_path):
    write_erp_csv(tmp_path / "input" / "f.csv", [t(Y, S, "X"), t(P, R, "Y", doc="2")])
    result = run("08/10/2026")
    report = json.loads(result.audit_json_path.read_text(encoding="utf-8"))["report"]
    assert report["columns"] == list(REPORT_COLUMNS)
    assert {k: report["rows"][0][k] for k in CUSTOMER_COUNT_COLUMNS} == {
        "customers_yesterday": 1, "customers_previous": 0, "customers_total": 1}
    assert report["grand_total_customer_counts"] == {
        "customers_yesterday": 1, "customers_previous": 0, "customers_total": 1}
