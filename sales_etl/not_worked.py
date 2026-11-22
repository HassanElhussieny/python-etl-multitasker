"""Task 5 option 1, report 4: Customers_Not_Worked_<Month>_<year>.xlsx - the customers without sales in the month.

Built after report 3 (Fixed_Customers_Report) from the same ``AnalysisResult`` and the same customer rows
(``fixed_customers.customers_table``: the fixed list + every other customer with activity in the period, the
المسؤول, the collection values), which it only reads - reports 1-3 never change. A customer is shown when its sales
of the month (from the 1st up to the reference date) are exactly 0, also when it has returns; الشهر الأعلى قيمة is
shown, never used to choose. One sheet per المسؤول - the same people, order and sheet names as report 3, also a
person without such a customer - with one الإجمالي العام row right below its customers (``person_total``).
"""

from __future__ import annotations

import logging
from collections import Counter
from collections.abc import Sequence
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook

from sales_etl import fixed_customers as fc
from sales_etl.analysis import (
    GRAND_TOTAL,
    MODE_YESTERDAY,
    ZERO,
    _TOLERANCE,
    AnalysisResult,
    _check_title,
    _ensure_new,
    _money,
    canonical,
    customer_key,
    format_dmy,
)

logger = logging.getLogger(__name__)

WORKBOOK_PREFIX = "Customers_Not_Worked"
ENGLISH_MONTHS: tuple[str, ...] = ("January", "February", "March", "April", "May", "June", "July", "August",
                                   "September", "October", "November", "December")


def enabled(result: AnalysisResult) -> bool:
    """Report 4 is part of option 1 (Yesterday & Previous Days) only, with the fixed customer list."""
    return result.mode == MODE_YESTERDAY and bool(result.fixed_customers)


def workbook_name(result: AnalysisResult) -> str:
    start, _ = fc.customers_period(result)
    return f"{WORKBOOK_PREFIX}_{ENGLISH_MONTHS[start.month - 1]}_{start.year}.xlsx"


def not_worked(row: Sequence[Any]) -> bool:
    """The only condition: the month's sales are 0 (returns, net and الشهر الأعلى قيمة do not matter)."""
    return row[4] == ZERO


def sheets_table(result: AnalysisResult) -> list[tuple[str, list[list[Any]]]]:
    """(المسؤول, its customers without sales) - report 3's groups (same people, same order, same customers per
    person), each keeping only its rows with sales = 0, in report 3's order."""
    detail = fc.customers_table(result)[:-1]
    return [(label, [r for r in mine if not_worked(r)]) for label, mine in fc.customers_by_responsible(detail)]


def sheet_title(result: AnalysisResult, label: str) -> str:
    start, end = fc.customers_period(result)
    return (f"العملاء الذين لم يحققوا مبيعات خلال شهر {fc.customers_month_label(result)}  |  الفترة من "
            f"{format_dmy(start)} إلى {format_dmy(end)}  |  المسؤول: {label}")


def build_not_worked_report(result: AnalysisResult, run_dir: Path) -> Path | None:
    """Task 5 option 1, step 4: write the workbook beside reports 1-3 (only when report 3 was written: it uses
    the same rows). A failure is kept in ``result.not_worked_error`` (reports 1-3 stay as written)."""
    if not enabled(result) or result.customers_report_path is None:
        return None
    try:
        result.not_worked_path = write_not_worked_report(result, run_dir / workbook_name(result))
    except Exception as exc:  # noqa: BLE001 - report 4 fails on its own
        result.not_worked_error = f"{WORKBOOK_PREFIX} could not be written: {type(exc).__name__}: {exc}"
        logger.exception("report 4 failed", extra={"event": "task5_not_worked_failed"})
    return result.not_worked_path


def write_not_worked_report(result: AnalysisResult, path: Path) -> Path:
    """One sheet per المسؤول (report 3's names): title, the 11 columns of report 3, the customers without sales,
    then its الإجمالي العام row - the black-and-white design of report 3's person sheets, without their block."""
    _ensure_new(path)
    headers = fc.customers_columns(result)
    collection = fc.collection_enabled(result)
    groups = sheets_table(result)
    names = fc.customers_sheet_names([label for label, _ in groups])
    wb = Workbook()
    for i, (label, mine) in enumerate(groups):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = names[label]
        fc._write_customer_table(ws, headers, mine, fc.person_total(mine, collection), None,
                                 sheet_title(result, label))
    wb.save(path)
    return path


# ------------------------------------------------------------------ checks
def _same(got: Any, want: Any) -> bool:
    if isinstance(want, (Decimal, float)) and not isinstance(want, bool):
        return isinstance(got, (int, float)) and abs(Decimal(str(got)) - Decimal(str(want))) <= _TOLERANCE
    if isinstance(want, int) and not isinstance(want, bool):
        return got == want
    return (got if got is not None else "") == (want if want is not None else "")


def _recount(rows: Sequence[list[Any]], collection: bool) -> list[Any]:
    """The total row recounted with plain loops from saved rows."""
    def amounts(i: int) -> Decimal:
        return sum((Decimal(str(r[i])) for r in rows if isinstance(r[i], (int, float))), ZERO)

    out = [None, GRAND_TOTAL, amounts(2), None, amounts(4), amounts(5), amounts(6),
           len({customer_key(r[0]) for r in rows if r[7] == fc.CUSTOMER_MARK}),
           len({customer_key(r[0]) for r in rows if r[8] == fc.CUSTOMER_MARK})]
    return out + ([amounts(9), amounts(10)] if collection else [])


def verify_not_worked_report(result: AnalysisResult) -> list[str]:
    """Report 4's own check: every sheet cell by cell; on the saved values every customer's sales = 0, each code
    once in the workbook, on the sheet of its own المسؤول, the total row = its customers recounted; then
    :func:`reconcile_not_worked` from the raw source cells."""
    if not enabled(result):
        return []
    if result.not_worked_error:
        return [result.not_worked_error]
    if result.customers_report_path is None:
        return []                                       # report 3 not written: its problem is reported already
    path = result.not_worked_path
    if path is None or not path.is_file() or path.stat().st_size == 0:
        return [f"missing or empty report 4: {path}"]
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        saved = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
        wb.close()
    except Exception as exc:  # noqa: BLE001
        return [f"cannot re-open {path.name}: {type(exc).__name__}: {exc}"]
    headers = fc.customers_columns(result)
    collection = fc.collection_enabled(result)
    groups = sheets_table(result)
    names = fc.customers_sheet_names([label for label, _ in groups])
    if list(saved) != list(names.values()):
        return [f"{path.name}: sheets {list(saved)} differ from report 3's المسؤول sheets {list(names.values())}"]
    problems: list[str] = []
    codes: list[str] = []
    for label, mine in groups:
        where = f"{path.name} / {names[label]}"
        sheet = saved[names[label]]
        problems += _check_title(where, sheet, sheet_title(result, label))
        if len(sheet) < 2 or tuple(sheet[1]) != tuple(headers):
            problems.append(f"{where}: header mismatch")
            continue
        expected = [*mine, fc.person_total(mine, collection)]
        body = sheet[2:]
        if len(body) != len(expected):
            problems.append(f"{where}: {len(body)} rows written, {len(expected)} expected (customers + one total)")
            continue
        for line, (got, want) in enumerate(zip(body, expected), start=3):
            for g, w, header in zip(got, want, headers):
                if not _same(g, w):
                    problems.append(f"{where}: row {line} {header} {g!r} differs from {w!r}")
        rows, total = body[:-1], body[-1]
        for line, r in enumerate(rows, start=3):
            key = customer_key(r[0])
            codes.append(key)
            if _money(r[4]) != ZERO:
                problems.append(f"{where}: row {line} customer {key} has sales {r[4]} (only 0 belongs here)")
            if canonical(fc._responsible_label(r[3])) != canonical(label):
                problems.append(f"{where}: customer {key} (المسؤول {r[3]!r}) is not on its own sheet")
            if abs(_money(r[4]) + _money(r[5]) - _money(r[6])) > 2 * _TOLERANCE:
                problems.append(f"{where}: row {line} net is not sales + returns")
        recount = _recount(rows, collection)
        if any(not _same(g, w) for g, w in zip(total, recount)) or len(total) != len(recount):
            problems.append(f"{where}: saved الإجمالي العام {total} differs from its customers recounted {recount}")
    repeated = sorted(k for k, n in Counter(codes).items() if n > 1)
    if repeated:
        problems.append(f"{path.name}: customer codes on more than one row: {repeated[:5]}")
    return problems + reconcile_not_worked(result, set(codes))


def reconcile_not_worked(result: AnalysisResult, saved_codes: set[str]) -> list[str]:
    """The saved customers vs a recount from the raw source cells (``independent_customers``): exactly the
    report's customers (fixed list + others with activity) whose recounted sales of the period are 0, with the
    recounted sales / returns, and - option 1 with the collection workbook - columns F / L read again."""
    where = f"{workbook_name(result)} (recount)"
    _, _, sums = fc.independent_customers(result)
    listed = {c.key for c in result.fixed_customers}
    candidates = listed | set(sums)
    want = {k for k in candidates if sums.get(k, [ZERO, ZERO])[0] == ZERO}
    problems: list[str] = []
    if saved_codes != want:
        problems.append(f"{where}: customers {sorted(saved_codes - want)[:5]} should not be there, "
                        f"{sorted(want - saved_codes)[:5]} are missing")
    rows = [r for _, mine in sheets_table(result) for r in mine]
    owed = fc._independent_collection(result.collection.path) if fc.collection_enabled(result) else None
    for r in rows:
        key = customer_key(r[0])
        if [r[4], r[5]] != sums.get(key, [ZERO, ZERO]):
            problems.append(f"{where}: customer {key} sales / returns {[r[4], r[5]]} differ from the source rows "
                            f"{sums.get(key, [ZERO, ZERO])}")
        if owed is not None and r[9:11] != list(owed.get(key, (None, None))):
            problems.append(f"{where}: customer {key} المستحق {r[9:11]} differs from the collection workbook "
                            f"{list(owed.get(key, (None, None)))}")
    return problems


def counts(result: AnalysisResult) -> dict[str, Any]:
    """For the summary: customers per sheet, in total, with returns and without."""
    groups = sheets_table(result)
    rows = [r for _, mine in groups for r in mine]
    return {"per_sheet": [(label, len(mine)) for label, mine in groups], "customers": len(rows),
            "with_returns": sum(1 for r in rows if r[5] != ZERO), "no_returns": sum(1 for r in rows if r[5] == ZERO)}
