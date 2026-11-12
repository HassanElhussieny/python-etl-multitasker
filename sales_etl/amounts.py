"""Strict monetary parsing into :class:`decimal.Decimal`.

Accepted: optional sign, digits with optional *correct* thousands grouping
(``1,234,567``), optional decimal part with ``.``. Anything else - empty cells,
text, European ``1.234,56`` notation, malformed grouping - is rejected rather
than coerced to zero.
"""

from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation

from sales_etl.text import ascii_digits, clean_cell

_AMOUNT = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$")


class AmountParseError(ValueError):
    """A value could not be interpreted as a monetary amount."""


def parse_amount(text: str | None) -> Decimal:
    """Parse a monetary string, raising :class:`AmountParseError` when invalid."""
    raw = ascii_digits(clean_cell(text))
    if not raw:
        raise AmountParseError("empty amount")
    if not _AMOUNT.match(raw):
        raise AmountParseError(f"invalid amount {text.strip()!r}")
    try:
        return Decimal(raw.replace(",", ""))
    except InvalidOperation as exc:  # pragma: no cover - regex already guarantees validity
        raise AmountParseError(f"invalid amount {text!r}") from exc


def fmt_amount(value: Decimal | None) -> str:
    """Display form with thousands separators (no rounding beyond the value's own scale)."""
    if value is None:
        return ""
    exponent = value.as_tuple().exponent
    places = max(2, -exponent) if isinstance(exponent, int) else 2
    return f"{value:,.{places}f}"
