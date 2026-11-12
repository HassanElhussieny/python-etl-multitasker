"""Strict, day-first date parsing.

Accepted forms (separator must be consistent within a value):

* ``DD/MM/YYYY``, ``DD-MM-YYYY``, ``DD.MM.YYYY`` - always day first, never month first
* ``YYYY-MM-DD`` (ISO, unambiguous)

Two-digit years are rejected. Transaction dates may carry a time-of-day suffix,
which is ignored; the reference date may not.
"""

from __future__ import annotations

import re
from datetime import date

from sales_etl.text import ascii_digits, clean_cell

_DMY = re.compile(
    r"^(?P<d>\d{1,2})(?P<sep>[/\-.])(?P<m>\d{1,2})(?P=sep)(?P<y>\d{4})"
    r"(?P<time>[ T]+\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:\s*[AaPp][Mm])?)?$"
)
_ISO = re.compile(
    r"^(?P<y>\d{4})-(?P<m>\d{2})-(?P<d>\d{2})"
    r"(?P<time>[ T]+\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?$"
)
_WEEKDAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
_MONTHS = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)


class DateParseError(ValueError):
    """A value could not be interpreted as a calendar date."""


def parse_date(text: str | None, *, allow_time: bool = True) -> date:
    """Parse a day-first or ISO date string, raising :class:`DateParseError`."""
    raw = ascii_digits(clean_cell(text))
    if not raw:
        raise DateParseError("empty date")
    match = _ISO.match(raw) or _DMY.match(raw)
    if match is None:
        raise DateParseError(f"unrecognised date {text!r}; expected DD/MM/YYYY")
    if match.group("time") and not allow_time:
        raise DateParseError(f"time of day is not allowed here: {text!r}")
    try:
        return date(int(match.group("y")), int(match.group("m")), int(match.group("d")))
    except ValueError as exc:
        raise DateParseError(f"invalid calendar date {text!r} ({exc})") from exc


def parse_reference_date(text: str | None) -> date:
    """Parse the user-entered Reference Yesterday Date (DD/MM/YYYY, day first)."""
    value = parse_date(text, allow_time=False)
    if not 2000 <= value.year <= 2100:
        raise DateParseError(f"year {value.year} is outside the supported range 2000-2100")
    return value


def describe_date(value: date) -> str:
    """Unambiguous human form, e.g. ``Thursday 08 October 2026 (ISO 2026-10-08)``."""
    return (
        f"{_WEEKDAYS[value.weekday()]} {value.day:02d} {_MONTHS[value.month - 1]} "
        f"{value.year} (ISO {value.isoformat()})"
    )


def format_dmy(value: date) -> str:
    """``DD/MM/YYYY``."""
    return f"{value.day:02d}/{value.month:02d}/{value.year:04d}"
