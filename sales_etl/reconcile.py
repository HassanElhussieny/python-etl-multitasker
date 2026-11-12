"""Independent reconciliation.

The checks here deliberately recompute everything with a second, plain-Python
implementation that works from each transaction's original date and document
type text - not from the pandas masks or the pre-computed ``period`` /
``category`` fields - so a bug in one path shows up as a FAIL.
"""

from __future__ import annotations

from dataclasses import fields
from datetime import date
from decimal import Decimal

import pandas as pd

from sales_etl.amounts import fmt_amount
from sales_etl.config import (
    COL_NET,
    CUSTOMER_COUNT_COLUMNS,
    DOC_TYPE_RETURNS,
    DOC_TYPE_SALES,
    METRIC_COLUMNS,
    METRIC_KEYS,
)
from sales_etl.metrics import metric_masks
from sales_etl.models import (
    ZERO,
    CheckResult,
    CheckStatus,
    CustomerSets,
    FileMetrics,
    FileOutcome,
    ParsedFile,
    Transaction,
)
from sales_etl.text import clean_cell

#: metric key -> (document type text, date relation)
_METRIC_RULES: dict[str, tuple[str, str]] = {
    "sales_yesterday": (DOC_TYPE_SALES, "eq"),
    "sales_previous": (DOC_TYPE_SALES, "lt"),
    "returns_yesterday": (DOC_TYPE_RETURNS, "eq"),
    "returns_previous": (DOC_TYPE_RETURNS, "lt"),
}


def independent_totals(
    transactions: list[Transaction], reference_date: date
) -> dict[str, tuple[int, Decimal]]:
    """Recompute ``(count, sum)`` per metric with a simple loop."""
    result = {key: (0, ZERO) for key in _METRIC_RULES}
    for txn in transactions:
        doc_type = clean_cell(txn.doc_type)
        for key, (wanted_type, relation) in _METRIC_RULES.items():
            if doc_type != wanted_type:
                continue
            same_day = txn.txn_date == reference_date
            earlier = txn.txn_date < reference_date
            if (relation == "eq" and same_day) or (relation == "lt" and earlier):
                count, amount = result[key]
                result[key] = (count + 1, amount + txn.net)
    return result


def independent_customer_counts(transactions: list[Transaction], reference_date: date) -> list[int]:
    """Recompute the three customer counts with a plain loop (``CUSTOMER_COUNT_KEYS`` order).

    Only مبيعات rows qualify a customer; مرتجعات rows are skipped entirely.
    """
    seen: dict[str, set[str]] = {}  # code -> {"yesterday", "previous"}
    for txn in transactions:
        doc_type, code = clean_cell(txn.doc_type), txn.customer
        if not code or doc_type != DOC_TYPE_SALES or txn.txn_date > reference_date:
            continue
        seen.setdefault(code, set()).add("yesterday" if txn.txn_date == reference_date else "previous")
    return [
        sum("yesterday" in f for f in seen.values()),
        sum("previous" in f for f in seen.values()),
        len(seen),
    ]


def reconcile_customers(scope: str, check: str, customers: CustomerSets,
                        transactions: list[Transaction], reference_date: date) -> CheckResult:
    """The three customer counts vs an independent recount over ``transactions``."""
    expected = independent_customer_counts(transactions, reference_date)
    actual = customers.counts()
    labels = list(CUSTOMER_COUNT_COLUMNS.values())
    return CheckResult(
        scope, check, _status(actual == expected),
        expected=", ".join(f"{label} {n}" for label, n in zip(labels, expected)),
        actual=", ".join(f"{label} {n}" for label, n in zip(labels, actual)),
        detail="distinct customer codes (sets) vs independent transaction-level recount",
    )


def _status(ok: bool, failure: CheckStatus = CheckStatus.FAIL) -> CheckStatus:
    return CheckStatus.PASS if ok else failure


def _diff(expected: Decimal, actual: Decimal) -> str:
    return fmt_amount(actual - expected)


def reconcile_file(
    parsed: ParsedFile, ledger: pd.DataFrame, metrics: FileMetrics, reference_date: date
) -> list[CheckResult]:
    """All per-file checks; see README for the meaning of each."""
    scope = parsed.file_name
    checks: list[CheckResult] = []
    independent = independent_totals(parsed.transactions, reference_date)

    for key in METRIC_KEYS:
        exp_count, exp_amount = independent[key]
        act_amount, act_count = metrics.value(key), metrics.count(key)
        checks.append(CheckResult(
            scope, f"METRIC {METRIC_COLUMNS[key]}",
            _status(exp_amount == act_amount and exp_count == act_count),
            expected=f"{fmt_amount(exp_amount)} ({exp_count} rows)",
            actual=f"{fmt_amount(act_amount)} ({act_count} rows)",
            difference=_diff(exp_amount, act_amount),
            detail="pandas result vs independent transaction-level recomputation",
        ))

    masks = metric_masks(ledger, reference_date)
    included = masks["sales_yesterday"] | masks["sales_previous"] | masks["returns_yesterday"] | masks["returns_previous"]
    leaked = int((ledger.loc[included, "date_ordinal"] > reference_date.toordinal()).sum())
    future = [t for t in parsed.transactions if t.txn_date > reference_date]
    checks.append(CheckResult(
        scope, "FUTURE_TRANSACTIONS_EXCLUDED", _status(leaked == 0),
        expected="0 future-dated rows in metrics", actual=f"{leaked} future-dated rows in metrics",
        detail=f"{len(future)} transaction(s) dated after the reference date excluded "
               f"(صافى {fmt_amount(sum((t.net for t in future), ZERO))})",
    ))

    all_net = sum((t.net for t in parsed.transactions), ZERO)
    parts = (sum((metrics.value(k) for k in METRIC_KEYS), ZERO)
             + metrics.future_net + metrics.other_type_net)
    part_count = (sum(metrics.count(k) for k in METRIC_KEYS)
                  + metrics.future_count + metrics.other_type_count)
    checks.append(CheckResult(
        scope, "PARTITION_COMPLETE",
        _status(parts == all_net and part_count == len(parsed.transactions)),
        expected=f"{fmt_amount(all_net)} ({len(parsed.transactions)} rows)",
        actual=f"{fmt_amount(parts)} ({part_count} rows)",
        difference=_diff(all_net, parts),
        detail="4 metrics + future-dated + other document types must equal all valid rows "
               "(nothing lost, nothing counted twice)",
    ))

    control = parsed.control
    if control.net_total is None:
        checks.append(CheckResult(scope, "ERP_CONTROL_TOTAL_NET", CheckStatus.NOT_APPLICABLE,
                                  detail="file has no ERP grand-total footer"))
    else:
        checks.append(CheckResult(
            scope, "ERP_CONTROL_TOTAL_NET", _status(all_net == control.net_total),
            expected=fmt_amount(control.net_total), actual=fmt_amount(all_net),
            difference=_diff(control.net_total, all_net),
            detail=f"sum of {COL_NET} over all valid rows vs ERP footer (line {control.net_total_line})",
        ))
    if control.doc_count is None:
        checks.append(CheckResult(scope, "ERP_CONTROL_DOC_COUNT", CheckStatus.NOT_APPLICABLE,
                                  detail="file has no ERP document-count footer"))
    else:
        checks.append(CheckResult(
            scope, "ERP_CONTROL_DOC_COUNT",
            _status(parsed.candidate_rows == control.doc_count, CheckStatus.WARN),
            expected=str(control.doc_count), actual=str(parsed.candidate_rows),
            difference=str(parsed.candidate_rows - control.doc_count),
            detail="transaction rows read vs ERP document count",
        ))

    rejected_lines = ", ".join(str(r.line_no) for r in parsed.rejected[:20])
    checks.append(CheckResult(
        scope, "NO_REJECTED_ROWS", _status(not parsed.rejected),
        expected="0", actual=str(len(parsed.rejected)),
        detail=(f"rejected lines: {rejected_lines}{' ...' if len(parsed.rejected) > 20 else ''}; "
                "their amounts are NOT in the totals" if parsed.rejected else "all rows parsed"),
    ))
    return checks


def sum_metrics(items: list[FileMetrics]) -> FileMetrics:
    """Field-wise sum of per-file metrics (the Grand Total)."""
    empty = FileMetrics()
    totals = {f.name: getattr(empty, f.name) for f in fields(FileMetrics)}
    for item in items:
        for name in totals:
            totals[name] = totals[name] + getattr(item, name)
    return FileMetrics(**totals)


def reconcile_grand_total(
    outcomes: list[FileOutcome], reference_date: date
) -> tuple[FileMetrics, list[CheckResult]]:
    """Grand Total from per-file totals, reconciled against all source transactions."""
    succeeded = [o for o in outcomes if o.succeeded and o.metrics and o.parsed]
    grand = sum_metrics([o.metrics for o in succeeded])  # type: ignore[misc]
    all_transactions = [t for o in succeeded for t in o.parsed.transactions]  # type: ignore[union-attr]
    independent = independent_totals(all_transactions, reference_date)
    checks = []
    for key in METRIC_KEYS:
        count, amount = independent[key]
        checks.append(CheckResult(
            "GRAND TOTAL", f"GRAND_TOTAL {METRIC_COLUMNS[key]}",
            _status(amount == grand.value(key) and count == grand.count(key)),
            expected=f"{fmt_amount(amount)} ({count} rows)",
            actual=f"{fmt_amount(grand.value(key))} ({grand.count(key)} rows)",
            difference=_diff(amount, grand.value(key)),
            detail=f"sum of {len(succeeded)} per-file totals vs recomputation over "
                   f"{len(all_transactions)} source transactions",
        ))
    return grand, checks


def reconcile_grand_customers(
    outcomes: list[FileOutcome], reference_date: date
) -> tuple[CustomerSets, list[CheckResult]]:
    """Grand Total customers = union of the per-file sets (never a sum of counts), rechecked."""
    succeeded = [o for o in outcomes if o.succeeded and o.customers and o.parsed]
    grand = CustomerSets()
    for o in succeeded:
        grand = grand.union(o.customers)  # type: ignore[arg-type]
    all_transactions = [t for o in succeeded for t in o.parsed.transactions]  # type: ignore[union-attr]
    return grand, [reconcile_customers("GRAND TOTAL", "GRAND_TOTAL_CUSTOMERS", grand, all_transactions,
                                       reference_date)]
