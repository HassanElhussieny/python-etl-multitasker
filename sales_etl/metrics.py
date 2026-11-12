"""Per-file metric calculation (primary path, pandas + Decimal).

Only the ``صافى`` column is summed. Amounts stay :class:`~decimal.Decimal`
(object dtype) so no floating-point rounding is introduced.
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal

import pandas as pd

from sales_etl.config import COL_CUSTOMER, COL_DATE, COL_DOC_NO, COL_DOC_TYPE, COL_NET
from sales_etl.models import Category, CustomerSets, FileMetrics, Transaction

LEDGER_COLUMNS = (
    "اسم الملف", "line_no", COL_DATE, "date_ordinal", COL_DOC_TYPE, "category",
    COL_DOC_NO, COL_CUSTOMER, COL_NET, "period",
)


def build_ledger(transactions: list[Transaction]) -> pd.DataFrame:
    """Transaction-level DataFrame for one or more files."""
    return pd.DataFrame({
        "اسم الملف": pd.Series([t.source_file for t in transactions], dtype=object),
        "line_no": pd.Series([t.line_no for t in transactions], dtype="int64"),
        COL_DATE: pd.Series([t.txn_date for t in transactions], dtype=object),
        "date_ordinal": pd.Series([t.txn_date.toordinal() for t in transactions], dtype="int64"),
        COL_DOC_TYPE: pd.Series([t.doc_type for t in transactions], dtype=object),
        "category": pd.Series([t.category.value for t in transactions], dtype=object),
        COL_DOC_NO: pd.Series([t.doc_no for t in transactions], dtype=object),
        COL_CUSTOMER: pd.Series([t.customer for t in transactions], dtype=object),
        COL_NET: pd.Series([t.net for t in transactions], dtype=object),
        "period": pd.Series([t.period.value for t in transactions], dtype=object),
    }, columns=list(LEDGER_COLUMNS))


def metric_masks(ledger: pd.DataFrame, reference_date: date) -> dict[str, pd.Series]:
    """Boolean row selections for the four metrics (built from the raw date, not ``period``)."""
    ref = reference_date.toordinal()
    sales = ledger["category"] == Category.SALES.value
    returns = ledger["category"] == Category.RETURNS.value
    on_ref = ledger["date_ordinal"] == ref
    before = ledger["date_ordinal"] < ref
    return {
        "sales_yesterday": sales & on_ref,
        "sales_previous": sales & before,
        "returns_yesterday": returns & on_ref,
        "returns_previous": returns & before,
    }


def _sum(ledger: pd.DataFrame, mask: pd.Series) -> Decimal:
    return sum(ledger.loc[mask, COL_NET].tolist(), Decimal("0"))


def compute_customer_sets(ledger: pd.DataFrame, reference_date: date) -> CustomerSets:
    """Distinct customer codes with a مبيعات row; returns, blank codes and future dates never count."""
    masks = metric_masks(ledger, reference_date)
    has_code = ledger[COL_CUSTOMER] != ""

    def codes(mask: pd.Series) -> frozenset[str]:
        return frozenset(ledger.loc[mask & has_code, COL_CUSTOMER].tolist())

    return CustomerSets(yesterday=codes(masks["sales_yesterday"]), previous=codes(masks["sales_previous"]))


def compute_file_metrics(ledger: pd.DataFrame, reference_date: date) -> FileMetrics:
    """Compute the four metrics; zero (never NaN) when nothing matches."""
    masks = metric_masks(ledger, reference_date)
    after = ledger["date_ordinal"] > reference_date.toordinal()
    in_scope = ledger["category"].isin([Category.SALES.value, Category.RETURNS.value])
    other = ~in_scope
    return FileMetrics(
        **{key: _sum(ledger, mask) for key, mask in masks.items()},
        **{f"{key}_count": int(mask.sum()) for key, mask in masks.items()},
        future_count=int((after & in_scope).sum()),
        future_net=_sum(ledger, after & in_scope),
        other_type_count=int(other.sum()),
        other_type_net=_sum(ledger, other),
    )
