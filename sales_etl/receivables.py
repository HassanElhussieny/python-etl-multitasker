"""Task 5, option 3: Receivables Update - an independent workflow.

From one Sales CSV (the structure Task 5 reads, the same reading rules) and one receivables workbook
(``المستحق 7-10`` / ``المستحق الجديد`` per customer code in column B) it writes
``Representatives_Receivables_Update_<run date>.xlsx``:

* ``Cover`` - the Task 5 Cover table (same groups, rows and numbers) plus the receivables of each representative;
* ``Master Data`` - one row per customer code (all its transactions summed: sales, returns, net with their signs)
  with the customer's two receivables (looked up once, by code);
* one sheet per representative (the customers of Master Data, once each) with its الإجمالي العام;
* ``تفاصيل الحركات`` - the original transactions (Task 5's Master Data layout) for review.

It never runs options 1-2, never writes their reports, never moves an input file and keeps no journal. The saved
workbook is re-opened and checked cell by cell, then reconciled independently (the raw CSV cells and a second
reading of the receivables workbook); only then is the run a success.
"""

from __future__ import annotations

import logging
from collections import Counter
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from sales_etl import analysis
from sales_etl.amounts import AmountParseError, parse_amount
from sales_etl.analysis import (
    COUNT_FORMAT,
    GRAND_TOTAL,
    MODE_ALL,
    MONEY_FORMAT,
    PERCENT_FORMAT,
    ZERO,
    _BW_BORDER,
    _BW_HEADER_BORDER,
    _BW_HEADER_FILL,
    _BW_MARGINS,
    _BW_STRIPE,
    _BW_SUBTOTAL_BORDER,
    _BW_SUBTOTAL_FILL,
    _BW_TOTAL_BORDER,
    _BW_TOTAL_FILL,
    _CENTER,
    _FONT,
    _ONE_PAGE_ROWS,
    AnalysisResult,
    Category,
    MainData,
    _code_key,
    _ensure_new,
    _master_label,
    canonical,
    clean_cell,
    customer_key,
)
from sales_etl.config import DOC_TYPE_RETURNS, DOC_TYPE_SALES, PipelineConfig
from sales_etl.grouped import _free_stamp

logger = logging.getLogger(__name__)

RUN_PREFIX = "Task5_Receivables"                     # output/Task5_Receivables_<stamp>/ (never a Task5_ folder)
WORKBOOK_PREFIX = "Representatives_Receivables_Update"
COL_CODE, COL_OLD, COL_NEW = "كود العميل", "المستحق 7-10", "المستحق الجديد"
SHEET_COVER, SHEET_MASTER, SHEET_DETAIL = "Cover", "Master Data", "تفاصيل الحركات"
NOT_FOUND_NOTE = "غير موجود في ملف المستحقات"
MASTER_HEADERS: tuple[str, ...] = (
    "م", "كود العميل", "اسم العميل", "المندوب", "القطاع", "عدد الحركات", "أول تاريخ", "آخر تاريخ", "المبيعات",
    "المرتجعات", "الصافي", "داخل قائمة الـ224", "خارج قائمة الـ224", COL_OLD, COL_NEW, "ملاحظة المستحقات",
)
MARK = "نعم"                       # داخل / خارج قائمة الـ224: one of the two columns, the other empty
#: Cover: one row per representative (Task 5's groups and their subtotals), then الإجمالي العام - the money and
#: the customers who sold (distinct codes, sales > 0), never the customers listed
COVER_HEADERS: tuple[str, ...] = (
    "م", "اسم المندوب", "الشهر الأعلى قيمة", "إجمالي المبيعات خلال الشهر", "إجمالي المرتجعات خلال الشهر",
    "إجمالي الصافي", "عملاء داخل الـ224 حققوا مبيعات", "عملاء خارج الـ224 حققوا مبيعات", "إجمالي المستحق 7-10",
    "إجمالي المستحق الجديد",
)
# Master Data / representative sheets (0-based): transactions, sales, returns, net, in, out, 7-10, new
I_TX, I_SALES, I_RET, I_NET, I_IN, I_OUT, I_OLD, I_NEW = 5, 8, 9, 10, 11, 12, 13, 14
# Cover (0-based): الشهر الأعلى قيمة, sales, returns, net, sold in, sold out, 7-10, new
C_TOP, C_SALES, C_RET, C_NET, C_SOLD_IN, C_SOLD_OUT, C_OLD, C_NEW = 2, 3, 4, 5, 6, 7, 8, 9
DATE_FORMAT = "dd/mm/yyyy"
_TOL = Decimal("0.005")


class ReceivablesError(ValueError):
    """The receivables workbook cannot be used as it is (nothing is written)."""


# ------------------------------------------------------------------ the receivables workbook
@dataclass(frozen=True)
class Receivable:
    key: str
    code: Any
    old: Decimal | None      # المستحق 7-10 (None: the cell is empty)
    new: Decimal | None      # المستحق الجديد
    row: int


@dataclass
class ReceivablesFile:
    path: Path
    sheet: str
    header_row: int
    records: dict[str, Receivable]
    rows_without_code: int = 0       # rows below the header without a customer code (totals, notes)


def _find_header(rows: list[list[Any]]) -> tuple[int, dict[str, int]] | None:
    """(index, column of each needed header) of the first of the first 15 rows naming the three columns."""
    wanted = {canonical(h): h for h in (COL_CODE, COL_OLD, COL_NEW)}
    for i, row in enumerate(rows[:15]):
        found: dict[str, list[int]] = {}
        for c, value in enumerate(row):
            key = canonical(str(value)) if value is not None else ""
            if key in wanted:
                found.setdefault(wanted[key], []).append(c)
        if len(found) == 3:
            repeated = [h for h, cols in found.items() if len(cols) > 1]
            if repeated:
                raise ReceivablesError(f"header row {i + 1}: column(s) {repeated} appear more than once")
            return i, {h: cols[0] for h, cols in found.items()}
    return None


def _receivable_value(value: Any, formula: Any, where: str) -> Decimal | None:
    """A receivable cell: a number (kept exactly), empty -> None; never a silent zero."""
    if value is None or (isinstance(value, str) and not value.strip()):
        if isinstance(formula, str) and formula.startswith("="):
            raise ReceivablesError(f"{where}: a formula without a saved value - open the file in Excel and save it")
        return None
    if isinstance(value, bool):
        raise ReceivablesError(f"{where}: {value!r} is not an amount")
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    if isinstance(value, str):
        try:
            return parse_amount(value)
        except AmountParseError as exc:
            raise ReceivablesError(f"{where}: {value!r} is not an amount") from exc
    raise ReceivablesError(f"{where}: {value!r} is not an amount")


def load_receivables(path: Path) -> ReceivablesFile:
    """The receivables workbook (read only): the first sheet whose header row (within the first 15 rows) names
    كود العميل, المستحق 7-10 and المستحق الجديد exactly; كود العميل must be column B. Rows without a code (totals,
    notes) are skipped and counted; an unreadable amount, a formula without a saved value or a repeated code stops
    the run (listing every repeated code and its rows) - nothing is summed or chosen at random."""
    try:
        values_wb = load_workbook(path, data_only=True)
        formulas_wb = load_workbook(path)
    except Exception as exc:  # noqa: BLE001 - any failure: stop before writing
        raise ReceivablesError(f"cannot open the receivables workbook {path}: {type(exc).__name__}: {exc}") from exc
    for ws in values_wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        found = _find_header(rows)
        if found is None:
            continue
        index, cols = found
        if cols[COL_CODE] != 1:
            raise ReceivablesError(f"{path.name} / {ws.title}: '{COL_CODE}' is column "
                                   f"{get_column_letter(cols[COL_CODE] + 1)}, expected column B")
        formulas = [list(r) for r in formulas_wb[ws.title].iter_rows(values_only=True)]
        records: dict[str, list[Receivable]] = {}
        without_code = 0
        for i in range(index + 1, len(rows)):
            row = rows[i] + [None] * (max(cols.values()) + 1 - len(rows[i]))
            frow = formulas[i] + [None] * (max(cols.values()) + 1 - len(formulas[i]))
            key = customer_key(row[1])
            if not key:
                without_code += 1
                continue
            where = f"{path.name} / {ws.title} row {i + 1}"
            records.setdefault(key, []).append(Receivable(
                key, row[1], _receivable_value(row[cols[COL_OLD]], frow[cols[COL_OLD]], f"{where} {COL_OLD}"),
                _receivable_value(row[cols[COL_NEW]], frow[cols[COL_NEW]], f"{where} {COL_NEW}"), i + 1))
        repeated = {k: [r.row for r in v] for k, v in records.items() if len(v) > 1}
        if repeated:
            listed = "; ".join(f"{k} (rows {', '.join(map(str, v))})" for k, v in sorted(repeated.items()))
            raise ReceivablesError(f"{path.name} / {ws.title}: {len(repeated)} customer code(s) appear more than "
                                   f"once - correct the file (nothing is summed or chosen): {listed}")
        logger.info("receivables read", extra={"event": "receivables_read", "file": path.name, "sheet": ws.title,
                                                "customers": len(records)})
        return ReceivablesFile(path, ws.title, index + 1, {k: v[0] for k, v in records.items()}, without_code)
    raise ReceivablesError(f"{path.name}: no sheet has the columns {COL_CODE}, {COL_OLD} and {COL_NEW} "
                           "in its header row")


# ------------------------------------------------------------------ the two inputs, found in the input folder
@dataclass
class InputScan:
    """What :func:`find_inputs` found in the input folder: the usable Sales CSV files and receivables workbooks
    (recognised by their structure, never by their names) and every other CSV / Excel file with the reason."""

    folder: Path
    sales: list[Path] = field(default_factory=list)
    receivables: list[Path] = field(default_factory=list)
    ignored: list[tuple[Path, str]] = field(default_factory=list)


def _is_sales_csv(path: Path, config: PipelineConfig) -> str | None:
    """None when ``path`` reads as a Task 5 sales file (the same reading rules: encoding, delimiter, header with
    العميل, نوع المستند and صافى); else the reason."""
    from sales_etl.config import COL_CUSTOMER, COL_DOC_TYPE, COL_NET
    from sales_etl.ingest import read_raw_table

    try:
        raw = read_raw_table(path, encodings=config.encodings, delimiters=config.delimiters,
                             forced_encoding=config.forced_encoding, forced_delimiter=config.forced_delimiter)
    except Exception as exc:  # noqa: BLE001 - not a usable sales file
        return f"not a readable sales CSV ({type(exc).__name__}: {exc})"
    missing = [c for c in (COL_CUSTOMER, COL_DOC_TYPE, COL_NET) if c not in raw.column_index]
    return f"no column(s) {missing}" if missing else None


def _is_receivables_workbook(path: Path) -> str | None:
    """None when a sheet of ``path`` has the receivables header (كود العميل, المستحق 7-10, المستحق الجديد within its
    first 15 rows); else the reason. Only the structure is checked here; load_receivables checks the rest."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001
        return f"cannot be opened ({type(exc).__name__}: {exc})"
    try:
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(max_row=15, values_only=True)]
            try:
                if _find_header(rows) is not None:
                    return None
            except ReceivablesError:
                return None                                   # the columns are there (repeated): load_receivables says so
        return f"no sheet with the columns {COL_CODE}, {COL_OLD}, {COL_NEW}"
    finally:
        wb.close()


def find_inputs(input_dir: Path, config: PipelineConfig) -> InputScan:
    """The project input folder (the one every task uses), read only: every .csv that reads as a Task 5 sales file
    and every .xlsx / .xlsm with the receivables columns; temporary (~$) and hidden files and the output folder are
    left out. Nothing is renamed, moved or written."""
    from sales_etl.ingest import discover_csv_files

    scan = InputScan(input_dir)
    csvs, _ = discover_csv_files(input_dir, exclude_dirs=[config.output_dir])
    for path in csvs:
        reason = _is_sales_csv(path, config)
        (scan.ignored.append((path, reason)) if reason else scan.sales.append(path))
    books = sorted((p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm")
                    and not p.name.startswith(("~$", "."))), key=lambda p: p.name.casefold())
    for path in books:
        reason = _is_receivables_workbook(path)
        (scan.ignored.append((path, reason)) if reason else scan.receivables.append(path))
    return scan


# ------------------------------------------------------------------ customer-level Master Data
@dataclass
class CustomerRow:
    key: str
    code: Any
    name: str
    rep_row: str           # the representative row label (Main Data value, UNREGISTERED, "")
    representative: str    # as shown (غير مسجل / غير محدد)
    sector: str
    transactions: int = 0
    first: date | None = None
    last: date | None = None
    sales: Decimal = ZERO
    returns: Decimal = ZERO
    status: str = ""
    old: Decimal | None = None
    new: Decimal | None = None
    found: bool = False
    top: Decimal | None = None   # الشهر الأعلى قيمة of fixed code.xlsx (column C), only for the list's customers

    @property
    def net(self) -> Decimal:
        return self.sales + self.returns

    @property
    def inside(self) -> bool:
        return self.status == analysis.FIXED_IN_LABEL

    @property
    def sold(self) -> bool:
        """A customer who sold: its sales of the period above zero (never by net)."""
        return bool(self.key) and self.sales > 0

    def values(self, serial: int | None) -> list[Any]:
        note = "" if self.found or not self.key else NOT_FOUND_NOTE
        return [serial, self.code, self.name, self.representative, self.sector, self.transactions, self.first,
                self.last, self.sales, self.returns, self.net, MARK if self.inside else "",
                "" if self.inside else MARK, self.old if self.old is not None else "",
                self.new if self.new is not None else "", note]


def _number(value: Any) -> Decimal | None:
    return Decimal(str(value)) if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def customer_rows(result: AnalysisResult, receivables: ReceivablesFile,
                  top_months: dict[str, Decimal | None] | None = None) -> list[CustomerRow]:
    """Master Data, one row per customer code (in order of first appearance in the CSV): every counted transaction
    of the code summed (sales, returns with their sign, net), its CSV name (first non-blank), Main Data
    representative / sector and 224-code status (one per code), its first / last date and transaction count, and
    its two receivables looked up by code (empty when the code is not in the receivables workbook)."""
    rows: dict[str, CustomerRow] = {}
    for m in result.master:
        c = rows.get(m.code)
        if c is None:
            known = result.main_data.customers.get(m.code) if m.code else None
            c = rows[m.code] = CustomerRow(m.code, int(m.code) if m.code.isdigit() else m.code, "", m.rep_row,
                                           m.representative, m.sector, status=m.status)
            c.top = (top_months or {}).get(m.code) if m.code else None
            rec = receivables.records.get(m.code) if m.code else None
            if rec is not None:
                c.old, c.new, c.found = rec.old, rec.new, True
            c.name = m.name or ""
            if not c.name and known is not None:
                c.name = known.name
        elif not c.name and m.name:
            c.name = m.name
        c.transactions += 1
        c.first = m.day if c.first is None else min(c.first, m.day)
        c.last = m.day if c.last is None else max(c.last, m.day)
        if m.category is Category.SALES:
            c.sales += m.net
        else:
            c.returns += m.net
    for c in rows.values():
        c.name = c.name or (str(c.code) if c.code != "" else "")
    return list(rows.values())


def _amount(value: Decimal | None) -> Decimal:
    return value if value is not None else ZERO


def _total_row(customers: Sequence[CustomerRow]) -> list[Any]:
    return [None, GRAND_TOTAL, f"{len(customers)} عميل", "", "", sum(c.transactions for c in customers), "", "",
            sum((c.sales for c in customers), ZERO), sum((c.returns for c in customers), ZERO),
            sum((c.net for c in customers), ZERO), sum(1 for c in customers if c.inside),
            sum(1 for c in customers if not c.inside), sum((_amount(c.old) for c in customers), ZERO),
            sum((_amount(c.new) for c in customers), ZERO), ""]


def master_table(customers: Sequence[CustomerRow]) -> list[list[Any]]:
    return [*(c.values(i) for i, c in enumerate(customers, start=1)), _total_row(customers)]


def representative_tables(result: AnalysisResult, customers: Sequence[CustomerRow]
                          ) -> list[tuple[str, str, list[list[Any]]]]:
    """(representative row label, sheet name, table) for every Task 5 representative sheet (Cover order, also
    without customers): its customers of Master Data, once each, then الإجمالي العام."""
    names = analysis.representative_sheet_names(result)
    out = []
    for label in analysis.representative_rows(result):
        mine = [c for c in customers if c.rep_row == label]
        out.append((label, names[label], master_table(mine)))
    return out


def cover_values(label: str, members: Sequence[CustomerRow]) -> list[Any]:
    """A Cover row of ``members``: الشهر الأعلى قيمة (column C of fixed code.xlsx, summed), sales, returns, net,
    the distinct customer codes inside / outside the 224 list whose sales are above zero, the two receivables."""
    sold_in = {c.key for c in members if c.sold and c.inside}
    sold_out = {c.key for c in members if c.sold and not c.inside}
    sales = sum((c.sales for c in members), ZERO)
    returns = sum((c.returns for c in members), ZERO)
    return [label, sum((_amount(c.top) for c in members), ZERO), sales, returns, sales + returns, len(sold_in),
            len(sold_out), sum((_amount(c.old) for c in members), ZERO), sum((_amount(c.new) for c in members), ZERO)]


def cover_table(result: AnalysisResult, customers: Sequence[CustomerRow]) -> list[list[Any]]:
    """Cover: one row per representative in Task 5's Cover order and groups (each group followed by its subtotal),
    then الإجمالي العام of every customer (distinct codes over all representatives)."""
    by_rep: dict[str, list[CustomerRow]] = {}
    for c in customers:
        by_rep.setdefault(c.rep_row, []).append(c)
    rows: list[list[Any]] = []
    blanks: set[int] = set()
    for group, subtotal in zip(analysis.cover_groups(result), analysis.SUBTOTAL_LABELS):
        if not group:
            continue
        members: list[CustomerRow] = []
        for name in group:
            mine = by_rep.get(name, [])
            rows.append(cover_values(_master_label(name), mine))
            members += mine
        blanks.add(len(rows))
        rows.append(cover_values(subtotal, members))
    blanks.add(len(rows))
    rows.append(cover_values(GRAND_TOTAL, customers))
    return [[serial, *row] for serial, row in zip(analysis._serials(len(rows), blanks), rows)]


def detail_table(result: AnalysisResult) -> list[list[Any]]:
    """The original transactions (Task 5's Master Data layout) with their الإجمالي العام."""
    rows = [[i, *m.values()] for i, m in enumerate(result.master, start=1)]
    return [*rows, [None, *analysis.detail_total(result.master)]]


# ------------------------------------------------------------------ the workbook
def _write_table(ws: Worksheet, headers: Sequence[str], rows: Sequence[list[Any]], *, title: str,
                 subtotals: set[int] = frozenset(), money: Sequence[int] = (), counts: Sequence[int] = (),
                 dates: Sequence[int] = (), percent: Sequence[int] = ()) -> None:
    """A table in Task 5's black-and-white design (the last row is الإجمالي العام): title on row 1, header on row 2,
    right to left, frozen header, filter on the detail rows, A4 landscape one page wide, rows 1-2 repeated."""
    ws.sheet_view.rightToLeft = True
    size = 11
    header_font = Font(name=_FONT, bold=True, color="000000", size=size + 1)
    body_font = Font(name=_FONT, bold=True, color="000000", size=size)
    total_font = Font(name=_FONT, bold=True, color="000000", size=size + 1)
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font, cell.fill, cell.border, cell.alignment = header_font, _BW_HEADER_FILL, _BW_HEADER_BORDER, _CENTER
    last = len(rows) - 1
    stripe = False
    for index, values in enumerate(rows):
        r = index + 3
        kind = "total" if index == last else ("sub" if index in subtotals else "data")
        if kind == "data":
            stripe = not stripe
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=value if value != "" else None)
            cell.alignment = _CENTER
            if kind == "total":
                cell.font, cell.fill, cell.border = total_font, _BW_TOTAL_FILL, _BW_TOTAL_BORDER
            elif kind == "sub":
                cell.font, cell.fill, cell.border = total_font, _BW_SUBTOTAL_FILL, _BW_SUBTOTAL_BORDER
            else:
                cell.font, cell.border = body_font, _BW_BORDER
                if not stripe:
                    cell.fill = _BW_STRIPE
            if c in percent and isinstance(value, Decimal):
                cell.number_format = PERCENT_FORMAT
            elif c in money and isinstance(value, Decimal):
                cell.number_format = MONEY_FORMAT
            elif c in dates and isinstance(value, date):
                cell.number_format = DATE_FORMAT
            elif c in counts and isinstance(value, int):
                cell.number_format = COUNT_FORMAT
    shown = [[v.strftime("%d/%m/%Y") if isinstance(v, date) else v for v in row] for row in rows]
    analysis._layout(ws, headers, shown, size=size, cap=48, body_height=22, one_page=len(rows) <= _ONE_PAGE_ROWS,
                     title=title, min_width=8, pad=2)
    ws.cell(row=1, column=1).font = Font(name=_FONT, bold=True, color="000000", size=size + 2)
    ws.row_dimensions[last + 3].height = max(ws.row_dimensions[last + 3].height or 0, 28)
    ws.page_margins = _BW_MARGINS
    ws.page_setup.blackAndWhite = False
    ws.print_options.horizontalCentered = True
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{max(len(rows) + 1, 2)}"


MASTER_MONEY, MASTER_COUNTS, MASTER_DATES = (9, 10, 11, 14, 15), (6, 12, 13), (7, 8)
DETAIL_MONEY, DETAIL_DATES = (7, 8, 9), (2,)


@dataclass
class ReceivablesResult:
    """Option 3 of one run."""

    sales_csv: Path
    receivables_path: Path
    started_at: datetime
    analysis: AnalysisResult | None = None
    receivables: ReceivablesFile | None = None
    customers: list[CustomerRow] = field(default_factory=list)
    fixed_path: Path | None = None           # fixed code.xlsx (الشهر الأعلى قيمة), checked again on its own
    run_dir: Path | None = None
    report_path: Path | None = None
    problems: list[str] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)

    @property
    def succeeded(self) -> bool:
        return self.report_path is not None and not self.problems


def title(result: ReceivablesResult) -> str:
    return (f"تحديث المستحقات  |  ملف المبيعات: {result.sales_csv.name}  |  ملف المستحقات: {result.receivables_path.name}"
            f"  |  تاريخ التشغيل: {result.started_at:%d/%m/%Y}")


def write_workbook(result: ReceivablesResult, path: Path) -> Path:
    _ensure_new(path)
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_COVER
    a, customers = result.analysis, result.customers
    cover = cover_table(a, customers)
    subtotals = {i for i, row in enumerate(cover) if row[1] in analysis.SUBTOTAL_LABELS}
    _write_table(ws, COVER_HEADERS, cover, title=title(result), subtotals=subtotals, money=(3, 4, 5, 6, 9, 10),
                 counts=(7, 8))
    _write_table(wb.create_sheet(SHEET_MASTER), MASTER_HEADERS, master_table(customers), title=title(result),
                 money=MASTER_MONEY, counts=MASTER_COUNTS, dates=MASTER_DATES)
    for label, sheet, table in representative_tables(a, customers):
        _write_table(wb.create_sheet(sheet), MASTER_HEADERS, table, title=f"{title(result)}  |  المندوب: "
                     f"{_master_label(label)}", money=MASTER_MONEY, counts=MASTER_COUNTS, dates=MASTER_DATES)
    _write_table(wb.create_sheet(SHEET_DETAIL), ("م", *analysis.MASTER_COLUMNS), detail_table(a), title=title(result),
                 money=DETAIL_MONEY, dates=DETAIL_DATES, counts=(3, 5))
    wb.save(path)
    return path


# ------------------------------------------------------------------ checks
def _same(got: Any, want: Any) -> bool:
    if isinstance(want, Decimal):
        return isinstance(got, (int, float)) and not isinstance(got, bool) and abs(Decimal(str(got)) - want) <= _TOL
    if isinstance(want, date):
        return isinstance(got, (date, datetime)) and (got.date() if isinstance(got, datetime) else got) == want
    if isinstance(want, int) and not isinstance(want, bool):
        return got == want
    return (got if got is not None else "") == (want if want is not None else "")


def _money_of(value: Any) -> Decimal:
    return Decimal(str(value)) if isinstance(value, (int, float)) and not isinstance(value, bool) else ZERO


def _check_sheet(where: str, saved: list[list[Any]], headers: Sequence[str], rows: Sequence[list[Any]],
                 title_text: str) -> list[str]:
    problems = []
    if not saved or saved[0][0] != title_text:
        problems.append(f"{where}: title {saved[0][0] if saved else None!r} differs")
    if len(saved) < 2 or tuple(saved[1][:len(headers)]) != tuple(headers):
        return problems + [f"{where}: header {saved[1] if len(saved) > 1 else None} differs from {list(headers)}"]
    body = saved[2:]
    if len(body) != len(rows):
        return problems + [f"{where}: {len(body)} rows written, {len(rows)} expected"]
    for line, (got, want) in enumerate(zip(body, rows), start=3):
        for g, w, h in zip(got, want, headers):
            if not _same(g, w):
                problems.append(f"{where}: row {line} {h} {g!r} differs from {w!r}")
    return problems


def _saved_totals(rows: Sequence[list[Any]], cols: Sequence[int]) -> list[Decimal]:
    return [sum((_money_of(r[c]) for r in rows), ZERO) for c in cols]


def verify_workbook(result: ReceivablesResult) -> list[str]:
    """The saved workbook re-opened: every sheet cell by cell, then on the saved values: every الإجمالي العام = the
    sum of its rows, every customer once in Master Data and once on a representative sheet, the representative
    sheets' receivables = Cover's rows, Cover's الإجمالي العام = the representative sheets' sum = Master Data."""
    path = result.report_path
    if path is None or not path.is_file() or path.stat().st_size == 0:
        return [f"missing or empty receivables report: {path}"]
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        saved = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
        wb.close()
    except Exception as exc:  # noqa: BLE001
        return [f"cannot re-open {path.name}: {type(exc).__name__}: {exc}"]
    a, customers, t = result.analysis, result.customers, title(result)
    reps = representative_tables(a, customers)
    expected_sheets = [SHEET_COVER, SHEET_MASTER, *(sheet for _, sheet, _ in reps), SHEET_DETAIL]
    if list(saved) != expected_sheets:
        return [f"{path.name}: sheets {list(saved)} differ from {expected_sheets}"]
    cover = cover_table(a, customers)
    problems = _check_sheet(f"{path.name} / {SHEET_COVER}", saved[SHEET_COVER], COVER_HEADERS, cover, t)
    problems += _check_sheet(f"{path.name} / {SHEET_MASTER}", saved[SHEET_MASTER], MASTER_HEADERS,
                             master_table(customers), t)
    for label, sheet, table in reps:
        problems += _check_sheet(f"{path.name} / {sheet}", saved[sheet], MASTER_HEADERS, table,
                                 f"{t}  |  المندوب: {_master_label(label)}")
    problems += _check_sheet(f"{path.name} / {SHEET_DETAIL}", saved[SHEET_DETAIL], ("م", *analysis.MASTER_COLUMNS),
                             detail_table(a), t)
    if problems:
        return problems
    # on the saved values
    sums = (I_TX, I_SALES, I_RET, I_NET, I_OLD, I_NEW)
    master = saved[SHEET_MASTER][2:]
    if _saved_totals(master[:-1], sums) != _saved_totals([master[-1]], sums):
        problems.append(f"{path.name} / {SHEET_MASTER}: الإجمالي العام differs from the sum of its rows")
    keys = [customer_key(r[1]) for r in master[:-1]]
    if len(keys) != len(set(keys)):
        problems.append(f"{path.name} / {SHEET_MASTER}: a customer code appears more than once")
    on_sheets: list[str] = []
    rep_totals = {}
    for label, sheet, _ in reps:
        body = saved[sheet][2:]
        if _saved_totals(body[:-1], sums) != _saved_totals([body[-1]], sums):
            problems.append(f"{path.name} / {sheet}: الإجمالي العام differs from the sum of its rows")
        on_sheets += [customer_key(r[1]) for r in body[:-1]]
        rep_totals[_master_label(label)] = _saved_totals([body[-1]], (I_OLD, I_NEW))
        marks = [r for r in body[:-1] if (r[I_IN] == MARK) == (r[I_OUT] == MARK)]
        if marks:
            problems.append(f"{path.name} / {sheet}: {len(marks)} customer(s) marked in both or neither list column")
    if Counter(on_sheets) != Counter(keys):
        problems.append(f"{path.name}: the representative sheets do not hold every Master Data customer exactly once")
    all_reps = [sum((v[i] for v in rep_totals.values()), ZERO) for i in (0, 1)]
    if all_reps != _saved_totals([master[-1]], (I_OLD, I_NEW)):
        problems.append(f"{path.name}: the representative sheets' receivables {all_reps} differ from Master Data")
    cover_rows = saved[SHEET_COVER][2:]
    for row in cover_rows[:-1]:
        if row[1] in rep_totals and _saved_totals([row], (C_OLD, C_NEW)) != rep_totals[row[1]]:
            problems.append(f"{path.name} / {SHEET_COVER}: '{row[1]}' receivables differ from its sheet "
                            f"{rep_totals[row[1]]}")
    if _saved_totals([cover_rows[-1]], (C_OLD, C_NEW)) != all_reps:
        problems.append(f"{path.name} / {SHEET_COVER}: الإجمالي العام receivables differ from the representative "
                        f"sheets {all_reps}")
    return problems


def _independent_receivables(path: Path) -> dict[str, tuple[Decimal | None, Decimal | None]]:
    """A second, plain reading of the receivables workbook: code -> (7-10, new) of the sheet whose row names the
    three columns (values as numbers, empty -> None)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            for i, row in enumerate(rows[:15]):
                names = [clean_cell(str(v)) if v is not None else "" for v in row]
                if COL_CODE in names and COL_OLD in names and COL_NEW in names:
                    io, inew = names.index(COL_OLD), names.index(COL_NEW)
                    out = {}
                    for r in rows[i + 1:]:
                        r = r + [None] * (max(io, inew) + 1 - len(r))
                        key = customer_key(r[1])
                        if key:
                            out[key] = tuple(None if v is None or v == "" else Decimal(str(v)) for v in (r[io], r[inew]))
                    return out
    finally:
        wb.close()
    return {}


def _independent_tops(path: Path | None) -> dict[str, Decimal | None]:
    """A plain second reading of fixed code.xlsx: code (column A) -> الشهر الأعلى قيمة (column C)."""
    if path is None:
        return {}
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = [list(r) for r in wb.worksheets[0].iter_rows(min_row=2, values_only=True)]
    finally:
        wb.close()
    return {customer_key(r[0]): _number(r[2] if len(r) > 2 else None) for r in rows if r and customer_key(r[0])}


def _reconcile_cover(result: ReceivablesResult, cover: list[list[Any]], raw: dict[str, list[Any]],
                     source: dict[str, tuple[Decimal | None, Decimal | None]]) -> list[str]:
    """Cover recomputed from the raw CSV cells (per code), Main Data (the representative), the 224 list, a second
    reading of fixed code.xlsx (الشهر الأعلى قيمة) and of the receivables workbook: every representative row, every
    group subtotal and الإجمالي العام (customers who sold = distinct codes with sales > 0, over all representatives)."""
    a = result.analysis
    tops = _independent_tops(result.fixed_path)
    per_code: dict[str, list[Any]] = {}
    for key, (sales, returns, _) in raw.items():
        known = a.main_data.customers.get(key) if key else None
        rep = known.representative if known else analysis.UNREGISTERED
        old, new = source.get(key, (None, None)) if key else (None, None)
        per_code[key] = [rep, sales, returns, key in a.fixed_codes, _amount(tops.get(key)) if key else ZERO,
                         _amount(old), _amount(new)]

    def expected(label: str, codes: Sequence[str]) -> list[Any]:
        rows = [per_code[k] for k in codes]
        sales = sum((r[1] for r in rows), ZERO)
        returns = sum((r[2] for r in rows), ZERO)
        return [label, sum((r[4] for r in rows), ZERO), sales, returns, sales + returns,
                len({k for k in codes if k and per_code[k][1] > 0 and per_code[k][3]}),
                len({k for k in codes if k and per_code[k][1] > 0 and not per_code[k][3]}),
                sum((r[5] for r in rows), ZERO), sum((r[6] for r in rows), ZERO)]

    want: list[list[Any]] = []
    for group, subtotal in zip(analysis.cover_groups(a), analysis.SUBTOTAL_LABELS):
        if not group:
            continue
        members: list[str] = []
        for name in group:
            mine = [k for k, v in per_code.items() if v[0] == name]
            want.append(expected(_master_label(name), mine))
            members += mine
        want.append(expected(subtotal, members))
    want.append(expected(GRAND_TOTAL, list(per_code)))
    problems = []
    got = [row[1:] for row in cover]
    if [r[0] for r in got] != [r[0] for r in want]:
        return [f"{SHEET_COVER}: rows {[r[0] for r in got]} differ from {[r[0] for r in want]}"]
    for g, w in zip(got, want):
        if g != w:
            problems.append(f"{SHEET_COVER}: '{g[0]}' {g[1:]} differs from the source files {w[1:]}")
    return problems


def reconcile(result: ReceivablesResult) -> list[str]:
    """Independent of the tables: the raw CSV cells recounted per customer code (excluded representative left out,
    the same document types) and a second reading of the receivables workbook; fills ``result.stats``."""
    a, customers = result.analysis, result.customers
    problems: list[str] = []
    raw: dict[str, list[Any]] = {}
    for _, _, _, kind_text, _, code_text, _, net_text in a.source_docs:
        kind = clean_cell(kind_text)
        if kind not in (DOC_TYPE_SALES, DOC_TYPE_RETURNS):
            continue
        key = customer_key(code_text)
        known = a.main_data.customers.get(key)
        if known is not None and analysis._is_excluded(known.representative):
            continue
        net = Decimal(analysis.ascii_digits(clean_cell(net_text)).replace(",", ""))
        r = raw.setdefault(key, [ZERO, ZERO, 0])
        r[0 if kind == DOC_TYPE_SALES else 1] += net
        r[2] += 1
    source = _independent_receivables(result.receivables_path)
    codes = sorted((k for k in raw if k), key=_code_key)
    found = [k for k in codes if k in source]
    missing = [k for k in codes if k not in source]
    by_key = {c.key: c for c in customers}
    if sorted(by_key, key=_code_key) != sorted(raw, key=_code_key) or len(by_key) != len(customers):
        problems.append(f"{SHEET_MASTER}: customers {len(customers)} differ from the CSV's codes {len(raw)} "
                        "(a customer lost or repeated)")
    for key, (sales, returns, count) in raw.items():
        c = by_key.get(key)
        if c is None:
            continue
        if [c.sales, c.returns, c.transactions] != [sales, returns, count]:
            problems.append(f"{SHEET_MASTER}: customer {key} {[c.sales, c.returns, c.transactions]} differs from the "
                            f"CSV {[sales, returns, count]}")
        want = source.get(key, (None, None)) if key else (None, None)
        if (c.old, c.new) != want or c.found != (bool(key) and key in source):
            problems.append(f"{SHEET_MASTER}: customer {key} receivables {(c.old, c.new)} differ from the file {want}")
    old_total = sum((_amount(source[k][0]) for k in found), ZERO)
    new_total = sum((_amount(source[k][1]) for k in found), ZERO)
    master_total = _total_row(customers)
    if [master_total[I_OLD], master_total[I_NEW]] != [old_total, new_total]:
        problems.append(f"{SHEET_MASTER}: receivables {master_total[I_OLD:I_NEW + 1]} differ from the file's matched customers "
                        f"{[old_total, new_total]}")
    csv_sales = sum((v[0] for v in raw.values()), ZERO)
    csv_returns = sum((v[1] for v in raw.values()), ZERO)
    if master_total[I_SALES:I_NET + 1] != [csv_sales, csv_returns, csv_sales + csv_returns]:
        problems.append(f"{SHEET_MASTER}: sales / returns / net {master_total[I_SALES:I_NET + 1]} differ from the CSV "
                        f"{[csv_sales, csv_returns, csv_sales + csv_returns]}")
    if master_total[I_TX] != sum(v[2] for v in raw.values()) or master_total[I_TX] != len(a.master):
        problems.append(f"{SHEET_MASTER}: {master_total[I_TX]} transactions summed, the CSV has "
                        f"{sum(v[2] for v in raw.values())}")
    reps = representative_tables(a, customers)
    rep_keys = [customer_key(r[1]) for _, _, t in reps for r in t[:-1]]
    if Counter(rep_keys) != Counter(c.key for c in customers):
        problems.append("the representative sheets do not hold every customer exactly once")
    rep_sum = [sum((t[-1][i] for _, _, t in reps), ZERO) for i in (I_OLD, I_NEW)]
    if rep_sum != [old_total, new_total]:
        problems.append(f"the representative sheets' receivables {rep_sum} differ from {[old_total, new_total]}")
    for label, sheet, t in reps:
        mine = [c for c in customers if c.rep_row == label]
        if t[-1][I_OLD:I_NEW + 1] != [sum((_amount(c.old) for c in mine), ZERO),
                                      sum((_amount(c.new) for c in mine), ZERO)]:
            problems.append(f"{sheet}: receivables total differs from its customers")
    cover = cover_table(a, customers)
    if cover[-1][C_OLD:C_NEW + 1] != rep_sum:
        problems.append(f"{SHEET_COVER}: الإجمالي العام receivables {cover[-1][C_OLD:C_NEW + 1]} differ from the "
                        f"representative sheets {rep_sum}")
    problems += _reconcile_cover(result, cover, raw, source)
    result.stats = {
        "csv_codes": len(codes), "found": len(found), "not_found": len(missing), "missing_codes": missing,
        "duplicates": 0, "customers": len(customers), "transactions": master_total[5],
        "source_old": old_total, "source_new": new_total, "master_old": master_total[I_OLD],
        "master_new": master_total[I_NEW], "cover_old": cover[-1][C_OLD], "cover_new": cover[-1][C_NEW],
        "sales": master_total[I_SALES], "returns": master_total[I_RET], "net": master_total[I_NET],
        "rows_without_code": result.receivables.rows_without_code,
        "per_representative": [(_master_label(label), len(t) - 1, t[-1][I_OLD], t[-1][I_NEW]) for label, _, t in reps],
        "cover": cover,
    }
    return problems


# ------------------------------------------------------------------ run
def run_receivables_update(config: PipelineConfig, sales_csv: Path, receivables_path: Path, main_data: MainData, *,
                           fixed_codes: frozenset[str] = frozenset(), now: datetime | None = None,
                           fixed_customers: Sequence[Any] = (), fixed_path: Path | None = None
                           ) -> ReceivablesResult:
    """Option 3: read the Sales CSV (Task 5's reading rules, every transaction) and the receivables workbook, write
    the workbook into output/Task5_Receivables_<stamp>/, re-open and check it, then reconcile it. Input files are
    never moved or written; options 1-2 never run."""
    result = ReceivablesResult(sales_csv, receivables_path, now or datetime.now())
    result.receivables = load_receivables(receivables_path)          # ReceivablesError: stop, nothing written
    a = AnalysisResult(mode=MODE_ALL, reference_date=None, main_data=main_data, input_dir=sales_csv.parent,
                       started_at=result.started_at, fixed_codes=fixed_codes)
    item = analysis._add_file(sales_csv, config, a)
    a.files = [item]
    result.analysis = a
    if item.error:
        raise ReceivablesError(f"the Sales CSV cannot be read: {item.error}")
    if item.rejected_rows:
        result.problems.append(f"{sales_csv.name}: {item.rejected_rows} row(s) could not be read and are not counted")
    result.fixed_path = fixed_path
    tops = {c.key: _number(c.top_month) for c in fixed_customers}       # fixed code.xlsx, column C, as stored
    result.customers = customer_rows(a, result.receivables, tops)
    output = config.output_dir
    result.run_dir = output / f"{RUN_PREFIX}_{_free_stamp(output, None, RUN_PREFIX, result.started_at)}"
    result.run_dir.mkdir(parents=True)
    result.report_path = write_workbook(result, result.run_dir / f"{WORKBOOK_PREFIX}_{result.started_at:%Y-%m-%d}.xlsx")
    result.problems += verify_workbook(result)
    result.problems += reconcile(result)
    logger.info("receivables update", extra={"event": "receivables_update", "report": str(result.report_path),
                                             "problems": len(result.problems)})
    return result
