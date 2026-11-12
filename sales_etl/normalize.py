"""Row classification, type conversion and row-level validation.

Every body row ends up in exactly one bucket - nothing is dropped silently:

* blank row                      -> counted
* repeated header row            -> counted, INFO issue
* ERP footer (control totals)    -> captured for reconciliation
* valid transaction              -> :class:`Transaction`
* anything else                  -> :class:`RejectedRow` + ERROR issue
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal

from sales_etl.amounts import AmountParseError, parse_amount
from sales_etl.config import (
    COL_CUSTOMER,
    COL_DATE,
    COL_DOC_NO,
    COL_DOC_TYPE,
    COL_NET,
    DOC_TYPE_RETURNS,
    DOC_TYPE_SALES,
    FOOTER_DOC_COUNT_LABELS,
    FOOTER_GRAND_TOTAL_LABELS,
    PREAMBLE_FROM_LABEL,
    PREAMBLE_TO_LABEL,
)
from sales_etl.dates import DateParseError, format_dmy, parse_date
from sales_etl.ingest import RawTable, Row
from sales_etl.models import (
    Category,
    ControlTotals,
    Issue,
    ParsedFile,
    Period,
    RejectedRow,
    Severity,
    Transaction,
)
from sales_etl.text import ascii_digits, canonical, clean_cell

_SALES_KEY = canonical(DOC_TYPE_SALES)
_RETURNS_KEY = canonical(DOC_TYPE_RETURNS)
_COUNT_LABELS = {canonical(label) for label in FOOTER_DOC_COUNT_LABELS}
_TOTAL_LABELS = {canonical(label) for label in FOOTER_GRAND_TOTAL_LABELS}


def classify_period(txn_date: date, reference_date: date) -> Period:
    """Single definition of yesterday / previous / future relative to the reference date."""
    if txn_date == reference_date:
        return Period.YESTERDAY
    if txn_date < reference_date:
        return Period.PREVIOUS
    return Period.FUTURE


def classify_doc_type(value: str) -> Category:
    key = canonical(value)
    if key == _SALES_KEY:
        return Category.SALES
    if key == _RETURNS_KEY:
        return Category.RETURNS
    return Category.OTHER


def _row_key(cells: list[str]) -> tuple[str, ...]:
    keys = [canonical(c) for c in cells]
    while keys and not keys[-1]:
        keys.pop()
    return tuple(keys)


def _preamble_period(preamble: list[Row]) -> tuple[date | None, date | None]:
    """Extract the ``من <date> الى <date>`` report period, if the ERP printed one."""
    start = end = None
    from_key, to_key = canonical(PREAMBLE_FROM_LABEL), canonical(PREAMBLE_TO_LABEL)
    for _, cells in preamble:
        values = [clean_cell(c) for c in cells]
        for i, value in enumerate(values[:-1]):
            key = canonical(value)
            if key not in (from_key, to_key):
                continue
            following = next((v for v in values[i + 1:] if v), "")
            try:
                parsed = parse_date(following)
            except DateParseError:
                continue
            if key == from_key and start is None:
                start = parsed
            elif key == to_key and end is None:
                end = parsed
    return start, end


def parse_transactions(raw: RawTable, reference_date: date) -> ParsedFile:
    """Convert a :class:`RawTable` into validated transactions for one file."""
    name = raw.file_name
    idx = raw.column_index
    i_date, i_type, i_net = idx[COL_DATE], idx[COL_DOC_TYPE], idx[COL_NET]
    i_doc, i_cust = idx.get(COL_DOC_NO), idx.get(COL_CUSTOMER)
    width = len(raw.header_cells)
    header_key = _row_key(raw.header_cells)

    transactions: list[Transaction] = []
    rejected: list[RejectedRow] = []
    issues: list[Issue] = list(raw.issues)
    control = ControlTotals()
    blank_rows = repeated_headers = 0

    for line_no, cells in raw.body:
        cleaned = [clean_cell(c) for c in cells]
        raw_text = raw.delimiter.join(cells)
        if not any(cleaned):
            blank_rows += 1
            continue
        if _row_key(cells) == header_key:
            repeated_headers += 1
            issues.append(Issue(name, line_no, Severity.INFO, "REPEATED_HEADER",
                                "repeated header row skipped"))
            continue

        label = canonical(next(c for c in cleaned if c))
        if label in _COUNT_LABELS:
            values = [c for c in cleaned if c][1:]
            try:
                control.doc_count = int(ascii_digits(values[0]).replace(",", ""))
                control.doc_count_line = line_no
            except (IndexError, ValueError):
                issues.append(Issue(name, line_no, Severity.WARNING, "FOOTER_UNPARSEABLE",
                                    "document-count footer has no readable number", raw_text))
            continue
        if label in _TOTAL_LABELS:
            try:
                control.net_total = parse_amount(cleaned[i_net] if i_net < len(cleaned) else "")
                control.net_total_line = line_no
            except AmountParseError as exc:
                issues.append(Issue(name, line_no, Severity.WARNING, "FOOTER_UNPARSEABLE",
                                    f"grand-total footer has no readable صافى ({exc})", raw_text))
            continue

        # --- transaction candidate -----------------------------------------------
        reasons: list[str] = []
        # Any extra cell (even an empty one) means a delimiter inside a value may have
        # shifted columns: the header's trailing blank column would otherwise absorb the
        # shift and a neighbouring column would be read as صافى.
        if len(cells) > width:
            reasons.append(
                f"row has {len(cells)} cells but the header has {width}; "
                "columns may be shifted, so no value is trusted"
            )
        padded = cleaned + [""] * max(0, width - len(cleaned))
        txn_date = net = None
        doc_type = padded[i_type]
        if not reasons:
            try:
                txn_date = parse_date(padded[i_date])
            except DateParseError as exc:
                reasons.append(f"{COL_DATE}: {exc}")
            if not doc_type:
                reasons.append(f"{COL_DOC_TYPE}: empty")
            try:
                net = parse_amount(padded[i_net])
            except AmountParseError as exc:
                reasons.append(f"{COL_NET}: {exc}")

        if reasons or txn_date is None or net is None:
            rejected.append(RejectedRow(name, line_no, tuple(reasons), raw_text))
            issues.append(Issue(name, line_no, Severity.ERROR, "ROW_REJECTED",
                                "; ".join(reasons), raw_text))
            continue

        category = classify_doc_type(doc_type)
        if category is Category.OTHER:
            issues.append(Issue(name, line_no, Severity.WARNING, "UNEXPECTED_DOC_TYPE",
                                f"{COL_DOC_TYPE} '{doc_type}' is neither '{DOC_TYPE_SALES}' nor "
                                f"'{DOC_TYPE_RETURNS}'; kept in the ledger, excluded from the "
                                "four metrics", raw_text))
        elif (category is Category.SALES and net < 0) or (category is Category.RETURNS and net > 0):
            issues.append(Issue(name, line_no, Severity.WARNING, "UNUSUAL_SIGN",
                                f"{doc_type} with {COL_NET} {net}; sign preserved as-is", raw_text))

        transactions.append(Transaction(
            source_file=name, line_no=line_no, txn_date=txn_date, doc_type=doc_type,
            category=category, period=classify_period(txn_date, reference_date), net=net,
            doc_no=padded[i_doc] if i_doc is not None else "",
            customer=padded[i_cust] if i_cust is not None else "",
            cells=tuple(padded[:width]),
        ))

    period_from, period_to = _preamble_period(raw.preamble)
    if not transactions and not rejected:
        issues.append(Issue(name, None, Severity.WARNING, "NO_TRANSACTIONS",
                            "file has a header but no transaction rows; metrics are zero"))
    elif not any(t.period is Period.YESTERDAY for t in transactions):
        issues.append(Issue(name, None, Severity.INFO, "NO_TRANSACTIONS_ON_REFERENCE_DATE",
                            f"no valid transactions dated {format_dmy(reference_date)}"))
    if period_from and period_to and not period_from <= reference_date <= period_to:
        issues.append(Issue(name, None, Severity.WARNING, "REFERENCE_DATE_OUTSIDE_FILE_PERIOD",
                            f"file period {format_dmy(period_from)} - {format_dmy(period_to)} "
                            f"does not contain the reference date {format_dmy(reference_date)}"))
    if control.net_total is None:
        issues.append(Issue(name, None, Severity.INFO, "NO_CONTROL_TOTAL",
                            "no ERP grand-total footer found; footer reconciliation not possible"))

    return ParsedFile(
        file_name=name, path=raw.path, sha256=raw.sha256, size_bytes=raw.size_bytes,
        encoding=raw.encoding, encoding_reason=raw.encoding_reason, delimiter=raw.delimiter,
        header_line_no=raw.header_line_no, columns=raw.columns, column_index=raw.column_index,
        period_from=period_from, period_to=period_to, transactions=transactions,
        rejected=rejected, issues=issues, control=control, blank_rows=blank_rows,
        repeated_headers=repeated_headers,
    )


def total(values: list[Decimal]) -> Decimal:
    """Exact Decimal sum that is ``Decimal('0')`` for an empty list."""
    return sum(values, Decimal("0"))
