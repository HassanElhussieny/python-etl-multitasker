"""Task 5, report 3: Fixed_Customers_Report_<run date>.xlsx - the 224-code customer report.

One of Task 5's three reports, kept in its own module: it is built after reports 1 (تقرير تحليل المبيعات
والمرتجعات) and 2 (Representatives_Analysis) from the same ``AnalysisResult``, which it only reads - building or
checking it never changes the other two. Its own steps: :func:`load_fixed_customers` (fixed code.xlsx),
:func:`build_customers_report` (the workbook: the main sheet, then one sheet per المسؤول), and
:func:`verify_customers_report` (the saved workbook cell by cell + :func:`reconcile_customers` from the raw source
cells). ``analysis.verify_task5`` runs it after reports 1-2 and keeps its result apart.
"""

from __future__ import annotations

import hashlib
import logging
from collections import Counter
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from sales_etl.business import BUSINESS
from sales_etl.analysis import (
    COUNT_FORMAT,
    DOC_TYPE_RETURNS,
    DOC_TYPE_SALES,
    FIXED_CODES_NAME,
    GRAND_TOTAL,
    MASTER_UNREGISTERED,
    MASTER_UNSPECIFIED,
    MODE_YESTERDAY,
    MONEY_FORMAT,
    SHEET_MASTER,
    ZERO,
    _BW_BORDER,
    _BW_HEADER_BORDER,
    _BW_HEADER_FILL,
    _BW_MARGINS,
    _BW_STRIPE,
    _BW_TOTAL_BORDER,
    _BW_TOTAL_FILL,
    _CENTER,
    _FONT,
    _ONE_PAGE_ROWS,
    _SHEET_NAME_FORBIDDEN,
    _TOLERANCE,
    AnalysisResult,
    Category,
    MainDataError,
    _cell_lines,
    _check_title,
    _code_key,
    _ensure_new,
    _excluded_cells,
    _is_excluded,
    _layout,
    _money,
    ascii_digits,
    canonical,
    clean_cell,
    customer_key,
    format_dmy,
)
from sales_etl.receivables import COL_NEW, COL_OLD, Receivable, ReceivablesError, ReceivablesFile, load_receivables

logger = logging.getLogger(__name__)


# ------------------------------------------------------------------ third workbook: the 224-code customer report
#: Fixed_Customers_Report_<run date>.xlsx: every customer of fixed code.xlsx (its order), then the other customers
#: with activity in the period; the month's sales, returns and net per customer code
CUSTOMERS_WORKBOOK_PREFIX = "Fixed_Customers_Report"
SHEET_CUSTOMERS = "تقرير عملاء الـ224"
#: the mark in عملاء داخل قائمة الـ224 or عملاء خارج قائمة الـ224 (never both; the other cell stays empty)
CUSTOMER_MARK = "نعم"
#: the header of fixed code.xlsx, columns A-D (compared with canonical(): spaces and spelling variants fold)
FIXED_LIST_HEADERS: tuple[str, ...] = ("كود العميل", "اسم العميل", "الشهر الاعلي قيمه", "المسئول")
ARABIC_MONTHS: tuple[str, ...] = ("يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر",
                                  "أكتوبر", "نوفمبر", "ديسمبر")


@dataclass(frozen=True)
class FixedCustomer:
    """One row of fixed code.xlsx, values exactly as stored (A code, B name, C الشهر الأعلى قيمة, D المسؤول)."""

    key: str            # customer_key of the code: the matching key with the transactions
    code: Any
    name: Any
    top_month: Any
    responsible: Any


def load_fixed_customers(path: Path) -> tuple[FixedCustomer, ...]:
    """The customers of ``fixed code.xlsx`` (first sheet, read-only, never written), in the file's order. The header
    must be FIXED_LIST_HEADERS; blank rows are skipped; a missing or repeated code stops the run."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001 - any failure: stop before processing
        raise MainDataError(f"cannot open the fixed customer list {path}: {type(exc).__name__}: {exc}") from exc
    try:
        rows = [list(r)[:4] + [None] * (4 - len(r)) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()
    header = [canonical(str(v or "")) for v in (rows[0] if rows else [])]
    if header != [canonical(h) for h in FIXED_LIST_HEADERS]:
        raise MainDataError(f"{path.name}: columns A-D must be {list(FIXED_LIST_HEADERS)}, found {rows[0] if rows else []}")
    customers: list[FixedCustomer] = []
    seen: set[str] = set()
    for line, (code, name, top, responsible) in enumerate(rows[1:], start=2):
        if all(v is None or str(v).strip() == "" for v in (code, name, top, responsible)):
            continue
        key = customer_key(code)
        if not key:
            raise MainDataError(f"{path.name}: row {line} has no customer code")
        if key in seen:
            raise MainDataError(f"{path.name}: customer code {key} appears twice (row {line})")
        seen.add(key)
        customers.append(FixedCustomer(key, code, name, top, responsible))
    if not customers:
        raise MainDataError(f"{path.name} has no customers")
    return tuple(customers)


# ------------------------------------------------------------------ the collection workbook (option 1 only)
#: the folder of the current collection workbook, in the project root: read again on every option-1 run
COLLECTION_DIR_NAME = "التحصيل"
#: the two amounts added after the customer columns, and the column each must have in the collection workbook
COLLECTION_HEADERS: tuple[str, ...] = (COL_OLD, COL_NEW)
COLLECTION_LETTERS: dict[str, str] = {COL_OLD: "F", COL_NEW: "L"}
COLLECTION_SUMMARY_HEADERS: tuple[str, ...] = (f"إجمالي {COL_OLD}", f"إجمالي {COL_NEW}")
COLLECTION_SUFFIXES: tuple[str, ...] = (".xlsx", ".xlsm", ".xls")


class CollectionError(ValueError):
    """The collection folder or workbook cannot be used: report 3 is not written (reports 1-2 are not affected)."""


@dataclass(frozen=True)
class CollectionData:
    """The collection workbook as read for this run (never written): its records by customer_key and the file's
    SHA-256 at reading time (checked again after the run: the source must not change)."""

    file: ReceivablesFile
    sha256: str

    @property
    def path(self) -> Path:
        return self.file.path

    @property
    def records(self) -> dict[str, Receivable]:
        return self.file.records


def find_collection_workbook(folder: Path) -> Path:
    """The single Excel workbook of ``folder`` (any name; Excel's ~$ lock files and hidden files are ignored).
    A missing folder, no workbook or more than one workbook stops report 3 with the reason."""
    if not folder.is_dir():
        raise CollectionError(f"the collection folder does not exist: {folder}")
    candidates = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in COLLECTION_SUFFIXES
                        and not p.name.startswith(("~$", ".")))
    if not candidates:
        raise CollectionError(f"the collection folder {folder} has no Excel workbook (.xlsx / .xlsm)")
    if len(candidates) > 1:
        raise CollectionError(f"the collection folder {folder} has {len(candidates)} workbooks - keep only the "
                              f"current one: {[p.name for p in candidates]}")
    return candidates[0]


def load_collection(folder: Path) -> CollectionData:
    """Read the current collection workbook of ``folder`` (no cache: called on every run). The reading rules of
    ``receivables.load_receivables``: the sheet whose header row names كود العميل (column B), المستحق 7-10 and
    المستحق الجديد; rows without a code skipped; an unreadable amount, a formula without a saved value or a repeated
    code stops (every repeated code listed - nothing is summed or chosen). Here also: the two amounts must be
    columns F and L."""
    path = find_collection_workbook(folder)
    try:
        sha = hashlib.sha256(path.read_bytes()).hexdigest()
        data = load_receivables(path)
    except (OSError, ReceivablesError) as exc:
        raise CollectionError(f"collection workbook {path.name}: {exc}") from exc
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            header = next(wb[data.sheet].iter_rows(min_row=data.header_row, max_row=data.header_row,
                                                   values_only=True))
        finally:
            wb.close()
    except Exception as exc:  # noqa: BLE001
        raise CollectionError(f"collection workbook {path.name}: {type(exc).__name__}: {exc}") from exc
    letters = {canonical(str(v)): get_column_letter(i) for i, v in enumerate(header, start=1) if v is not None}
    for name, letter in COLLECTION_LETTERS.items():
        if letters.get(canonical(name)) != letter:
            raise CollectionError(f"collection workbook {path.name} / {data.sheet}: '{name}' is column "
                                  f"{letters.get(canonical(name))}, expected column {letter}")
    logger.info("collection read", extra={"event": "task5_collection_read", "file": path.name, "sheet": data.sheet,
                                          "customers": len(data.records)})
    return CollectionData(data, sha)


def collection_enabled(result: AnalysisResult) -> bool:
    """True when report 3 carries المستحق 7-10 / المستحق الجديد (option 1 with its collection workbook read)."""
    return getattr(result, "collection", None) is not None


def collection_stats(result: AnalysisResult) -> dict[str, int]:
    """Matching of the report's customers with the collection workbook (by customer_key)."""
    keys = {customer_key(r[0]) for r in customers_table(result)[:-1]}
    records = result.collection.records
    return {"customers": len(keys), "matched": len(keys & set(records)), "unmatched": len(keys - set(records)),
            "file_customers": len(records), "file_only": len(set(records) - keys), "duplicates": 0,
            "rows_without_code": result.collection.file.rows_without_code}


def _amount(value: Any) -> Decimal:
    return value if isinstance(value, Decimal) else ZERO


def customers_period(result: AnalysisResult) -> tuple[date, date | None]:
    """(first day, last day or None) of the third report: the 1st of the reference date's month up to the reference
    date (yesterday report); all of the latest month with counted transactions (all transactions; None: no end)."""
    if result.reference_date is not None:
        return result.reference_date.replace(day=1), result.reference_date
    latest = max((r.day for r in result.master), default=result.started_at.date())
    return latest.replace(day=1), None


def customers_month_label(result: AnalysisResult) -> str:
    start, _ = customers_period(result)
    return f"{ARABIC_MONTHS[start.month - 1]} {start.year}"


def customers_columns(result: AnalysisResult) -> tuple[str, ...]:
    month = customers_month_label(result)
    return ("كود العميل", "اسم العميل", "الشهر الأعلى قيمة", "اسم المسؤول عن العميل", f"مبيعات شهر {month}",
            f"مرتجعات شهر {month}", f"صافي شهر {month}", "عملاء داخل قائمة الـ224", "عملاء خارج قائمة الـ224",
            *(COLLECTION_HEADERS if collection_enabled(result) else ()))


def customers_title(result: AnalysisResult) -> str:
    start, end = customers_period(result)
    if end is not None:
        return f"تقرير عملاء قائمة الـ224  |  الفترة من {format_dmy(start)} إلى {format_dmy(end)}"
    return (f"تقرير عملاء قائمة الـ224  |  كل الحركات: أحدث شهر في ملفات الإدخال {customers_month_label(result)} "
            f"(من {format_dmy(start)})")


#: one sheet per المسؤول of the customer report, these first and in this order (spelling variants fold:
#: مندوب تجريبى 10 / مندوب تجريبي 10 are one sheet), then any other المسؤول of the report in order of appearance
CUSTOMER_RESPONSIBLES: tuple[str, ...] = BUSINESS.customer_responsibles   # sales_etl/business.py
#: the الإجمالي العام block under a customer table: its header row, then the values row (the counts of customers
#: who sold stand under the two list columns)
CUSTOMERS_SUMMARY_HEADERS: tuple[str, ...] = ("", "إجمالي العملاء", "عملاء داخل الـ224", "عملاء خارج الـ224",
                                              "إجمالي المبيعات", "إجمالي المرتجعات", "إجمالي الصافي",
                                              "عملاء داخل الـ224 حققوا مبيعات", "عملاء خارج الـ224 حققوا مبيعات")


def customer_inside(row: Sequence[Any]) -> bool:
    return row[7] == CUSTOMER_MARK


def customer_outside(row: Sequence[Any]) -> bool:
    return row[8] == CUSTOMER_MARK


def summary_headers(collection: bool) -> tuple[str, ...]:
    return (*CUSTOMERS_SUMMARY_HEADERS, *(COLLECTION_SUMMARY_HEADERS if collection else ()))


def customers_summary(rows: Sequence[list[Any]], collection: bool = False) -> list[Any]:
    """الإجمالي العام of customer rows: distinct customer codes (in + out, never rows), inside / outside the list,
    sales, returns (signed), net = sales + returns, then the distinct codes inside / outside the list whose sales of
    the period are above zero (a customer who sold: never by net, returns-only and zero sales do not count); with
    ``collection`` also the sums of المستحق 7-10 / المستحق الجديد (empty cells count as nothing)."""
    extra = [sum((_amount(r[i]) for r in rows), ZERO) for i in (9, 10)] if collection else []
    inside = {customer_key(r[0]) for r in rows if customer_inside(r)}
    outside = {customer_key(r[0]) for r in rows if customer_outside(r)}
    sold: dict[str, Decimal] = {}
    for r in rows:
        sold[customer_key(r[0])] = sold.get(customer_key(r[0]), ZERO) + r[4]
    sales = sum((r[4] for r in rows), ZERO)
    returns = sum((r[5] for r in rows), ZERO)
    return [GRAND_TOTAL, len(inside) + len(outside), len(inside), len(outside), sales, returns, sales + returns,
            sum(1 for k in inside if sold[k] > 0), sum(1 for k in outside if sold[k] > 0), *extra]


def _number_or_zero(value: Any) -> Decimal:
    """A numeric cell as Decimal (exact); empty or text -> 0 (never part of a sum)."""
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        return ZERO
    return value if isinstance(value, Decimal) else Decimal(str(value))


def person_total(rows: Sequence[list[Any]], collection: bool = False) -> list[Any]:
    """The الإجمالي العام row right below a المسؤول sheet's customer rows (those rows only): the label in اسم العميل
    (كود العميل and اسم المسؤول empty), the sums of الشهر الأعلى قيمة, sales, returns (signed) and net, the distinct
    codes marked نعم in داخل / خارج قائمة الـ224, and - with ``collection`` - the sums of المستحق 7-10 / الجديد."""
    inside = {customer_key(r[0]) for r in rows if customer_inside(r)}
    outside = {customer_key(r[0]) for r in rows if customer_outside(r)}
    sales = sum((r[4] for r in rows), ZERO)
    returns = sum((r[5] for r in rows), ZERO)
    extra = [sum((_amount(r[i]) for r in rows), ZERO) for i in (9, 10)] if collection else []
    return ["", GRAND_TOTAL, sum((_number_or_zero(r[2]) for r in rows), ZERO), "", sales, returns, sales + returns,
            len(inside), len(outside), *extra]


def _responsible_label(value: Any) -> str:
    return str(value).strip() if value not in (None, "") and str(value).strip() else MASTER_UNSPECIFIED


def customers_by_responsible(detail: Sequence[list[Any]]) -> list[tuple[str, list[list[Any]]]]:
    """The customer rows (main sheet order) by المسؤول: CUSTOMER_RESPONSIBLES first (also without customers), then
    every other المسؤول in order of appearance. The key is the name as written (spelling variants fold), so no
    customer moves to another person; each row is in exactly one group."""
    groups: dict[str, tuple[str, list[list[Any]]]] = {canonical(name): (name, []) for name in CUSTOMER_RESPONSIBLES}
    for r in detail:
        label = _responsible_label(r[3])
        groups.setdefault(canonical(label), (label, []))[1].append(r)
    return list(groups.values())


def customers_sheet_names(labels: Sequence[str]) -> dict[str, str]:
    """المسؤول -> a valid, unique worksheet name (31 characters at most, no [ ] : * ? / \\, not the main sheet);
    the full name stays in the sheet's title line and in its rows."""
    used = {SHEET_CUSTOMERS.casefold()}
    names: dict[str, str] = {}
    for label in labels:
        base = _SHEET_NAME_FORBIDDEN.sub("_", label).strip("' ")[:31] or "Sheet"
        name, n = base, 1
        while name.casefold() in used:
            n += 1
            suffix = f" ({n})"
            name = base[:31 - len(suffix)].rstrip() + suffix
        used.add(name.casefold())
        names[label] = name
    return names


def customers_sheet_title(result: AnalysisResult, label: str) -> str:
    return f"{customers_title(result)}  |  المسؤول: {label}"


def _responsible(result: AnalysisResult, key: str) -> str:
    """المسؤول of a customer outside the list: its Main Data representative (غير محدد if blank, غير مسجل if the code
    is not in Main Data)."""
    known = result.main_data.customers.get(key)
    return (known.representative or MASTER_UNSPECIFIED) if known else MASTER_UNREGISTERED


def customers_table(result: AnalysisResult) -> list[list[Any]]:
    """The third report, in customers_columns order: every fixed-list customer (file order, its own name, value and
    المسؤول, داخل الـ224 - at zero without activity), then every other customer code with a transaction in the period
    (by code; CSV name, Main Data المسؤول, خارج الـ224), then الإجمالي العام. Amounts per customer code from the counted
    transactions (Master Data) of the period: sales, returns with their sign, net = sales + returns."""
    start, end = customers_period(result)
    sums: dict[str, list[Decimal]] = {}
    names: dict[str, str] = {}
    for r in result.master:
        if not r.code or r.day < start or (end is not None and r.day > end):
            continue
        s = sums.setdefault(r.code, [ZERO, ZERO])
        s[0 if r.category is Category.SALES else 1] += r.net
        if r.name and r.code not in names:
            names[r.code] = r.name
    rows: list[list[Any]] = []
    listed = {c.key for c in result.fixed_customers}
    for c in result.fixed_customers:
        sales, returns = sums.get(c.key, [ZERO, ZERO])
        rows.append([c.code, c.name, c.top_month, c.responsible, sales, returns, sales + returns, CUSTOMER_MARK, ""])
    for key in sorted(set(sums) - listed, key=_code_key):
        sales, returns = sums[key]
        known = result.main_data.customers.get(key)
        name = names.get(key) or (known.name if known else "") or key
        rows.append([int(key) if key.isdigit() else key, name, "", _responsible(result, key), sales, returns,
                     sales + returns, "", CUSTOMER_MARK])
    sales = sum((r[4] for r in rows), ZERO)
    returns = sum((r[5] for r in rows), ZERO)
    total = [GRAND_TOTAL, "", "", "", sales, returns, sales + returns, "", ""]
    if collection_enabled(result):     # by customer code; not in the collection workbook: both cells empty
        records = result.collection.records
        for r in rows:
            found = records.get(customer_key(r[0]))
            r += [found.old, found.new] if found is not None else [None, None]
        total += [sum((_amount(r[i]) for r in rows), ZERO) for i in (9, 10)]
    return [*rows, total]


def customers_workbook_name(result: AnalysisResult) -> str:
    return f"{CUSTOMERS_WORKBOOK_PREFIX}_{result.started_at:%Y-%m-%d}.xlsx"


def _write_customer_table(ws: Worksheet, headers: Sequence[str], detail: Sequence[list[Any]],
                          total: list[Any] | None, summary: list[Any] | None, title: str,
                          block_label: str = "") -> None:
    """A customer table in the black-and-white design: title on row 1, header on row 2, the customer rows, the
    ``total`` row when given (the main sheet), then the الإجمالي العام block (its header row and ``summary``);
    ``block_label`` (a person's sheet: the المسؤول) fills the block header's first cell, right above الإجمالي العام."""
    ws.sheet_view.rightToLeft = True
    size = 11
    header_font = Font(name=_FONT, bold=True, color="000000", size=size + 1)
    body_font = Font(name=_FONT, bold=True, color="000000", size=size)
    total_font = Font(name=_FONT, bold=True, color="000000", size=size + 1)

    def money(cell: Any, column: int, value: Any) -> None:
        if column in (3, 5, 6, 7, 10, 11) and isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            cell.number_format = MONEY_FORMAT
        elif column == 1 and isinstance(value, int):
            cell.number_format = "0"

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font, cell.fill, cell.border, cell.alignment = header_font, _BW_HEADER_FILL, _BW_HEADER_BORDER, _CENTER
    for r, values in enumerate(detail, start=3):
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=value if value != "" else None)
            cell.font, cell.border, cell.alignment = body_font, _BW_BORDER, _CENTER
            if r % 2 == 0:
                cell.fill = _BW_STRIPE
            money(cell, c, value)
    table = [*detail, *([total] if total is not None else [])]
    last = len(table) + 2
    if total is not None:
        for c, value in enumerate(total, start=1):
            cell = ws.cell(row=last, column=c, value=value if value != "" else None)
            cell.font, cell.fill, cell.border, cell.alignment = total_font, _BW_TOTAL_FILL, _BW_TOTAL_BORDER, _CENTER
            if isinstance(value, Decimal):
                cell.number_format = MONEY_FORMAT
            elif isinstance(value, int) and not isinstance(value, bool):
                cell.number_format = COUNT_FORMAT
    # widths, title, print setup from the table itself (the main sheet keeps its layout), the block below it
    _layout(ws, headers, table or [[""] * len(headers)], size=size, cap=48, body_height=22,
            one_page=len(table) <= _ONE_PAGE_ROWS, title=title, min_width=8, pad=2)
    ws.cell(row=1, column=1).font = Font(name=_FONT, bold=True, color="000000", size=size + 2)
    if total is not None:
        ws.row_dimensions[last].height = max(ws.row_dimensions[last].height or 0, 28)
    label_width = round(len(GRAND_TOTAL) * size * 1.15 / 11) + 2  # the block's label fits on one line
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, label_width)
    block = last + 2                                         # one empty row between the table and the block
    end = last                                               # the last printed row
    if summary is not None:                                  # (report 4 has no block: its total row ends the sheet)
        end = block + 1
        block_headers = (block_label, *summary_headers(len(headers) > len(CUSTOMERS_SUMMARY_HEADERS))[1:])
        for c, header in enumerate(block_headers, start=1):
            cell = ws.cell(row=block, column=c, value=header or None)
            cell.font, cell.fill, cell.border, cell.alignment = header_font, _BW_HEADER_FILL, _BW_HEADER_BORDER, _CENTER
        for c, value in enumerate(summary, start=1):
            cell = ws.cell(row=block + 1, column=c, value=value if value != "" else None)
            cell.font, cell.fill, cell.border, cell.alignment = total_font, _BW_TOTAL_FILL, _BW_TOTAL_BORDER, _CENTER
            if isinstance(value, Decimal):
                cell.number_format = MONEY_FORMAT
            elif isinstance(value, int):
                cell.number_format = COUNT_FORMAT
    for col in range(8, len(headers) + 1):                   # the who-sold / collection total headers
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, 16)
    if summary is not None:
        lines = max(_cell_lines(h, ws.column_dimensions[get_column_letter(c)].width or 8.43, size + 1)
                    for c, h in enumerate(block_headers, start=1) if h)
        ws.row_dimensions[block].height = max(40, round((lines + 1) * (size + 1) * 1.5 + 6))   # nothing cut
        ws.row_dimensions[block + 1].height = 28
    ws.page_margins = _BW_MARGINS
    ws.page_setup.blackAndWhite = False
    ws.page_setup.fitToHeight = 1 if end <= _ONE_PAGE_ROWS + 2 else 0
    ws.print_options.horizontalCentered = True
    ws.print_area = f"A1:{get_column_letter(len(headers))}{end}"
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{len(detail) + 2}"   # the customer rows only


def build_customers_report(result: AnalysisResult, run_dir: Path) -> Path | None:
    """Task 5 step 3: write Fixed_Customers_Report_<run date>.xlsx into the run folder (beside reports 1-2) from the
    result, which it only reads. A failure is recorded in ``result.customers_error`` (reported as report 3's
    problem; reports 1-2 stay as written) instead of stopping the run."""
    if not result.fixed_customers:
        return None
    result.collection = None
    if result.mode == MODE_YESTERDAY and getattr(result, "collection_dir", None) is not None:   # option 1 only
        try:
            result.collection = load_collection(result.collection_dir)       # the current file, every run
        except CollectionError as exc:
            result.customers_error = (f"Fixed_Customers_Report not written - the collection step stopped "
                                      f"(reports 1-2 are not affected): {exc}")
            logger.error(result.customers_error, extra={"event": "task5_collection_failed"})
            return None
    try:
        result.customers_report_path = write_customers_report(result, run_dir / customers_workbook_name(result))
    except Exception as exc:  # noqa: BLE001 - report 3 fails on its own
        result.customers_error = f"Fixed_Customers_Report could not be written: {type(exc).__name__}: {exc}"
        logger.exception("report 3 failed", extra={"event": "task5_customers_report_failed"})
    return result.customers_report_path


def write_customers_report(result: AnalysisResult, path: Path) -> Path:
    """The third workbook: the main sheet (every customer, its الإجمالي العام row, then the الإجمالي العام block), then
    one sheet per المسؤول (:func:`customers_by_responsible`, made from the main sheet's rows) with its own block.
    Black-and-white design of تقرير تحليل المبيعات والمرتجعات, right to left, filter on the customer rows only,
    frozen header, A4 landscape one page wide (further pages down for long lists), rows 1-2 repeated."""
    _ensure_new(path)
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_CUSTOMERS
    headers, rows = customers_columns(result), customers_table(result)
    *detail, total = rows
    collection = collection_enabled(result)
    _write_customer_table(ws, headers, detail, total, customers_summary(detail, collection), customers_title(result))
    groups = customers_by_responsible(detail)
    names = customers_sheet_names([label for label, _ in groups])
    for label, mine in groups:
        _write_customer_table(wb.create_sheet(names[label]), headers, mine, person_total(mine, collection),
                              customers_summary(mine, collection), customers_sheet_title(result, label),
                              block_label=label)
    wb.save(path)
    return path


def independent_customers(result: AnalysisResult) -> tuple[date, date | None, dict[str, list[Decimal]]]:
    """The third report recomputed with a plain loop over the raw source cells (``source_docs``), not from Master
    Data: the period from the reference date (or the latest month of the counted rows), the document type text, the
    صافى text parsed again, the excluded representative's customers left out; customer code -> [sales, returns]."""
    customers = result.main_data.customers
    counted = []
    for _, _, day, kind_text, _, code_text, _, net_text in result.source_docs:
        kind = clean_cell(kind_text)
        if kind not in (DOC_TYPE_SALES, DOC_TYPE_RETURNS):
            continue
        key = customer_key(code_text)
        known = customers.get(key)
        if known is not None and _is_excluded(known.representative):
            continue
        if result.reference_date is not None and day > result.reference_date:
            continue
        counted.append((day, kind, key, Decimal(ascii_digits(clean_cell(net_text)).replace(",", ""))))
    if result.reference_date is not None:
        start, end = result.reference_date.replace(day=1), result.reference_date
    else:
        start, end = max((c[0] for c in counted), default=result.started_at.date()).replace(day=1), None
    sums: dict[str, list[Decimal]] = {}
    for day, kind, key, net in counted:
        if key and day >= start and (end is None or day <= end):
            s = sums.setdefault(key, [ZERO, ZERO])
            s[0 if kind == DOC_TYPE_SALES else 1] += net
    return start, end, sums


def reconcile_customers(result: AnalysisResult) -> list[str]:
    """The third report vs fixed code.xlsx, Main Data and :func:`independent_customers`: the fixed customers all
    there, in order, once, داخل الـ224 with their own values; the others once each (by code), خارج الـ224, never a
    fixed code, with their Main Data المسؤول; every row's sales / returns = the recount, net = sales + returns;
    الإجمالي العام = the recount's totals and the period's Master Data totals."""
    where = SHEET_CUSTOMERS
    problems: list[str] = []
    rows = customers_table(result)
    *detail, total = rows
    start, end, sums = independent_customers(result)
    if (start, end) != customers_period(result):
        problems.append(f"{where}: period {customers_period(result)} differs from the recount's {(start, end)}")
    fixed = list(result.fixed_customers)
    keys = [customer_key(r[0]) for r in detail]
    if len(keys) != len(set(keys)):
        problems.append(f"{where}: a customer code appears more than once: "
                        f"{sorted(k for k, n in Counter(keys).items() if n > 1)[:5]}")
    head = detail[:len(fixed)]
    if [customer_key(r[0]) for r in head] != [c.key for c in fixed]:
        problems.append(f"{where}: the {len(fixed)} fixed customers are not all there in the order of {FIXED_CODES_NAME}")
    for r, c in zip(head, fixed):
        if r[1:4] != [c.name, c.top_month, c.responsible] or r[7:9] != [CUSTOMER_MARK, ""]:
            problems.append(f"{where}: fixed customer {c.key}: name / value / المسؤول / list columns "
                            f"{r[1:4] + r[7:9]} differ from {FIXED_CODES_NAME}")
    listed = {c.key for c in fixed}
    rest = detail[len(fixed):]
    want_rest = sorted(set(sums) - listed, key=_code_key)
    if [customer_key(r[0]) for r in rest] != want_rest:
        problems.append(f"{where}: customers outside the list {[customer_key(r[0]) for r in rest][:5]} differ from "
                        f"the customers with activity {want_rest[:5]}")
    for r in rest:
        key = customer_key(r[0])
        if key in listed or r[7:9] != ["", CUSTOMER_MARK]:
            problems.append(f"{where}: customer {key} is classified {r[7:9]!r} (in the list: {key in listed})")
        if r[3] != _responsible(result, key):
            problems.append(f"{where}: customer {key} المسؤول {r[3]!r} differs from Main Data {_responsible(result, key)!r}")
    for r in detail:
        want = sums.get(customer_key(r[0]), [ZERO, ZERO])
        if [r[4], r[5]] != want:
            problems.append(f"{where}: customer {customer_key(r[0])} sales / returns {[r[4], r[5]]} differ from the "
                            f"source rows {want}")
        if r[6] != r[4] + r[5]:
            problems.append(f"{where}: customer {customer_key(r[0])} net is not sales + returns")
    sales = sum((s[0] for s in sums.values()), ZERO)
    returns = sum((s[1] for s in sums.values()), ZERO)
    if total[4:7] != [sales, returns, sales + returns] or total[0] != GRAND_TOTAL:
        problems.append(f"{where}: الإجمالي العام {total[4:7]} differs from the source rows {[sales, returns, sales + returns]}")
    period = [m for m in result.master if m.code and m.day >= start and (end is None or m.day <= end)]
    if [sum((m.sales for m in period), ZERO), sum((m.returns for m in period), ZERO)] != [sales, returns]:
        problems.append(f"{where}: the period's {SHEET_MASTER} totals differ from the source rows")
    # the الإجمالي العام block: counted again from the fixed list and the recount (distinct codes, never rows)
    sold_in = sum(1 for c in fixed if sums.get(c.key, [ZERO])[0] > 0)              # sales above zero, never net
    sold_out = sum(1 for k in want_rest if sums[k][0] > 0)
    want = [GRAND_TOTAL, len(fixed) + len(want_rest), len(fixed), len(want_rest), sales, returns, sales + returns,
            sold_in, sold_out]
    collection = collection_enabled(result)
    if collection:      # المستحق 7-10 / المستحق الجديد: a second, plain reading of the collection workbook
        problems += _reconcile_collection(result, detail, total)
        source = _independent_collection(result.collection.path)
        want += [sum((_amount(source.get(k, (None, None))[i]) for k in set(keys)), ZERO) for i in (0, 1)]
    if customers_summary(detail, collection) != want:
        problems.append(f"{where}: الإجمالي العام block {customers_summary(detail, collection)[1:]} differs from "
                        f"{want[1:]}")
    both = [customer_key(r[0]) for r in detail if customer_inside(r) == customer_outside(r)]
    if both:
        problems.append(f"{where}: customers marked in both or neither list column: {both[:5]}")
    # one sheet per المسؤول: every customer in exactly one, the people's blocks add up to the main block
    groups = customers_by_responsible(detail)
    labels = [canonical(label) for label, _ in groups]
    if len(labels) != len(set(labels)):
        problems.append(f"{where}: an المسؤول has more than one sheet")
    wanted = {canonical(n) for n in CUSTOMER_RESPONSIBLES} | {canonical(_responsible_label(r[3])) for r in detail}
    if set(labels) != wanted:
        problems.append(f"{where}: المسؤول sheets {sorted(labels)} differ from {sorted(wanted)}")
    seen: dict[str, str] = {}
    for label, mine in groups:
        for r in mine:
            key = customer_key(r[0])
            if canonical(_responsible_label(r[3])) != canonical(label):
                problems.append(f"{where}: customer {key} is on the sheet of '{label}' but its المسؤول is {r[3]!r}")
            if key in seen:
                problems.append(f"{where}: customer {key} is on the sheets of '{seen[key]}' and '{label}'")
            seen[key] = label
    if sorted(seen, key=_code_key) != sorted((customer_key(r[0]) for r in detail), key=_code_key):
        problems.append(f"{where}: the المسؤول sheets miss customers "
                        f"{sorted(set(customer_key(r[0]) for r in detail) - set(seen))[:5]}")
    parts = [customers_summary(mine, collection) for _, mine in groups]
    added = [sum(p[i] for p in parts) for i in range(1, 4)] + [sum((p[i] for p in parts), ZERO) for i in range(4, 7)] \
        + [sum(p[i] for p in parts) for i in (7, 8)] \
        + [sum((p[i] for p in parts), ZERO) for i in range(9, len(want))]
    if added != want[1:]:
        problems.append(f"{where}: the المسؤول totals {added} do not add up to الإجمالي العام {want[1:]}")
    # each المسؤول sheet's total row, counted again from fixed code.xlsx, the recount and the collection workbook
    tops = {c.key: _number_or_zero(c.top_month) for c in fixed}
    owed = _independent_collection(result.collection.path) if collection else {}
    for label, mine in groups:
        codes = {customer_key(r[0]) for r in mine}
        s = sum((sums.get(k, [ZERO, ZERO])[0] for k in codes), ZERO)
        t = sum((sums.get(k, [ZERO, ZERO])[1] for k in codes), ZERO)
        expect = ["", GRAND_TOTAL, sum((tops.get(k, ZERO) for k in codes), ZERO), "", s, t, s + t,
                  len(codes & listed), len(codes - listed)]
        if collection:
            expect += [sum((_amount(owed.get(k, (None, None))[i]) for k in codes), ZERO) for i in (0, 1)]
        if person_total(mine, collection) != expect:
            problems.append(f"{where}: the total row of '{label}' {person_total(mine, collection)[2:]} differs from "
                            f"the recount {expect[2:]}")
    return problems


def _independent_collection(path: Path) -> dict[str, tuple[Decimal | None, Decimal | None]]:
    """A second, plain reading of the collection workbook: code (column B) -> (column F, column L) of the sheet
    whose header row (first 15 rows) has كود العميل in B, المستحق 7-10 in F and المستحق الجديد in L."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            rows = [list(r) + [None] * 12 for r in ws.iter_rows(values_only=True)]
            for i, row in enumerate(rows[:15]):
                names = [clean_cell(str(v)) if v is not None else "" for v in row[:12]]
                if names[1] == "كود العميل" and names[5] == COL_OLD and names[11] == COL_NEW:
                    return {customer_key(r[1]): tuple(None if r[c] in (None, "") else Decimal(str(r[c]))
                                                      for c in (5, 11))
                            for r in rows[i + 1:] if customer_key(r[1])}
    finally:
        wb.close()
    return {}


def _reconcile_collection(result: AnalysisResult, detail: Sequence[list[Any]], total: list[Any]) -> list[str]:
    """المستحق 7-10 / المستحق الجديد of every customer row = columns F / L of its code in the collection workbook
    (read again here, independently; not in the file: both empty), the total row = their sums, and the workbook
    unchanged since it was read (same SHA-256)."""
    where = f"{SHEET_CUSTOMERS} / {result.collection.path.name}"
    problems: list[str] = []
    try:
        if hashlib.sha256(result.collection.path.read_bytes()).hexdigest() != result.collection.sha256:
            problems.append(f"{where}: the collection workbook changed during the run")
        source = _independent_collection(result.collection.path)
    except Exception as exc:  # noqa: BLE001 - not verified
        return [f"{where}: cannot re-read the collection workbook: {type(exc).__name__}: {exc}"]
    if not source:
        return [f"{where}: no sheet with كود العميل in B, {COL_OLD} in F and {COL_NEW} in L on re-reading"]
    for r in detail:
        want = list(source.get(customer_key(r[0]), (None, None)))
        if r[9:11] != want:
            problems.append(f"{where}: customer {customer_key(r[0])} {COL_OLD} / {COL_NEW} {r[9:11]} differ from "
                            f"the collection workbook {want}")
    sums = [sum((_amount(r[i]) for r in detail), ZERO) for i in (9, 10)]
    if total[9:11] != sums:
        problems.append(f"{where}: الإجمالي العام {total[9:11]} differs from the sum of the rows {sums}")
    return problems


def _verify_customers(where: str, saved: list[list[Any]], rows: list[list[Any]], headers: Sequence[str],
                      title: str, has_total: bool = True, block_label: str = "") -> list[str]:
    """A saved customer sheet: title, header, every cell of the customer rows (and the main sheet's total row), an
    empty row, the الإجمالي العام block; on the saved values: net = sales + returns in every row, the total row =
    the sum of the customer rows, the block = distinct codes inside / outside (in + out) and the sums of the rows."""
    problems = _check_title(where, saved, title)
    table = saved[1:]
    if not table or tuple(table[0]) != tuple(headers):
        return problems + [f"{where}: header mismatch {table[0] if table else None}"]
    body = table[1:]
    detail = rows[:-1] if has_total else rows
    collection = len(headers) > len(CUSTOMERS_SUMMARY_HEADERS)
    summary = customers_summary(detail, collection)
    expected = [*rows, [None] * len(headers), [h or None for h in (block_label, *summary_headers(collection)[1:])],
                [v if v != "" else None for v in summary]]
    if len(body) != len(expected):
        return problems + [f"{where}: {len(body)} rows written, {len(expected)} expected"]
    for line, (got, want) in enumerate(zip(body, expected), start=3):
        for g, w, header in zip(got, want, headers):
            if isinstance(w, Decimal) or (isinstance(w, float) and not isinstance(w, bool)):
                bad = not isinstance(g, (int, float)) or abs(Decimal(str(g)) - Decimal(str(w))) > _TOLERANCE
            elif isinstance(w, int) and not isinstance(w, bool):
                bad = g != w
            else:
                bad = (g if g is not None else "") != (w if w is not None else "")
            if bad:
                problems.append(f"{where}: row {line} {header} {g!r} differs from {w!r}")
    saved_detail = body[:len(detail)]
    sums = [ZERO] * 3
    for line, row in enumerate(saved_detail, start=3):
        amounts = [_money(row[c]) for c in (4, 5, 6)]
        if abs(amounts[0] + amounts[1] - amounts[2]) > 2 * _TOLERANCE:
            problems.append(f"{where}: row {line} net is not sales + returns")
        sums = [a + b for a, b in zip(sums, amounts)]
    tolerance = _TOLERANCE * max(len(saved_detail), 1)
    if has_total and any(abs(a - _money(b)) > tolerance for a, b in zip(sums, body[len(detail)][4:7])):
        problems.append(f"{where}: saved الإجمالي العام {body[len(detail)][4:7]} differs from the sum of the rows {sums}")
    block = body[-1]
    inside = {customer_key(r[0]) for r in saved_detail if r[7] == CUSTOMER_MARK}
    outside = {customer_key(r[0]) for r in saved_detail if r[8] == CUSTOMER_MARK}
    marks = [customer_key(r[0]) for r in saved_detail if (r[7] == CUSTOMER_MARK) == (r[8] == CUSTOMER_MARK)]
    if marks:
        problems.append(f"{where}: customers marked in both or neither list column: {marks[:5]}")
    sold_in = {customer_key(r[0]) for r in saved_detail if r[7] == CUSTOMER_MARK and _money(r[4]) > 0}
    sold_out = {customer_key(r[0]) for r in saved_detail if r[8] == CUSTOMER_MARK and _money(r[4]) > 0}
    if block[0] != GRAND_TOTAL or block[1:4] != [len(inside) + len(outside), len(inside), len(outside)] \
            or block[7:9] != [len(sold_in), len(sold_out)] \
            or any(abs(a - _money(b)) > tolerance for a, b in zip(sums, block[4:7])):
        problems.append(f"{where}: saved الإجمالي العام block {block} differs from the saved customer rows "
                        f"({len(inside)} inside, {len(outside)} outside, {sums}, sold {len(sold_in)} / {len(sold_out)})")
    if has_total and block_label:   # a المسؤول sheet's total row, recounted from its saved customer rows only
        saved_total = body[len(detail)]
        top = sum((_money(r[2]) for r in saved_detail if isinstance(r[2], (int, float))), ZERO)
        if saved_total[0] is not None or saved_total[1] != GRAND_TOTAL or saved_total[3] is not None \
                or saved_total[7:9] != [len(inside), len(outside)] \
                or abs(_money(saved_total[2]) - top) > tolerance:
            problems.append(f"{where}: saved total row {saved_total[:9]} differs from its customer rows (الشهر الأعلى "
                            f"قيمة {top}, {len(inside)} inside, {len(outside)} outside)")
    if collection:      # المستحق 7-10 / المستحق الجديد: the saved totals = the sums of the saved rows
        owed = [sum((_money(r[c]) for r in saved_detail if r[c] is not None), ZERO) for c in (9, 10)]
        totals = [*([body[len(detail)][9:11]] if has_total else []), block[9:11]]
        if any(abs(a - _money(b)) > tolerance for t in totals for a, b in zip(owed, t)):
            problems.append(f"{where}: saved {COL_OLD} / {COL_NEW} totals {totals} differ from the sum of the "
                            f"saved rows {owed}")
    return problems


def verify_customers_report(result: AnalysisResult) -> list[str]:
    """Report 3's own check: re-open the workbook - the main sheet and every المسؤول sheet cell by cell (the latter
    made from the main sheet's rows, each with its المسؤول above الإجمالي العام), every saved customer row on exactly
    one المسؤول sheet - then :func:`reconcile_customers`."""
    if result.customers_error:
        return [result.customers_error]
    path = result.customers_report_path
    if path is None or not path.is_file() or path.stat().st_size == 0:
        return [f"missing or empty customer report: {path}"]
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        saved = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
        wb.close()
    except Exception as exc:  # noqa: BLE001 - any failure means "not verified"
        return [f"cannot re-open {path.name}: {type(exc).__name__}: {exc}"]
    headers, rows = customers_columns(result), customers_table(result)
    groups = customers_by_responsible(rows[:-1])
    names = customers_sheet_names([label for label, _ in groups])
    expected = [SHEET_CUSTOMERS, *names.values()]
    if list(saved) != expected:
        return [f"{path.name}: sheets {list(saved)} differ from {expected}"]
    problems = _verify_customers(f"{path.name} / {SHEET_CUSTOMERS}", saved[SHEET_CUSTOMERS], rows, headers,
                                 customers_title(result))
    collection = collection_enabled(result)
    for label, mine in groups:
        problems += _verify_customers(f"{path.name} / {names[label]}", saved[names[label]],
                                      [*mine, person_total(mine, collection)], headers,
                                      customers_sheet_title(result, label), block_label=label)
    main_rows = saved[SHEET_CUSTOMERS][2:2 + len(rows) - 1]
    split = [tuple(r) for label in names.values() for r in saved[label][2:-4]]    # rows above its total row
    if Counter(split) != Counter(map(tuple, main_rows)):
        problems.append(f"{path.name}: the المسؤول sheets do not hold every saved customer row exactly once")
    return problems + _excluded_cells(path.name, saved) + reconcile_customers(result)
