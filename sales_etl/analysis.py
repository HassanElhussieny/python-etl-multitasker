"""Task 5: sector & representative sales analysis, classified through the Main Data workbook.

Every transaction is read with the existing parser (``ingest`` -> ``normalize``). Its customer
code (CSV column F, header ``العميل``) is looked up in the Main Data workbook (column A); the
sector comes from column C and the representative from column D. The CSV representative
columns (I, J) are never used. ``نوع المستند`` (B) decides مبيعات / مرتجعات, ``صافى`` (T)
is summed with its original sign, in Decimal.

Report types:

* ``MODE_YESTERDAY``: التاريخ == the entered date -> أمس; < the date -> الأيام السابقة;
  later dates are excluded.
* ``MODE_ALL``: every transaction, no date split.

Two workbooks per run, in the same folder:

* the report, exactly three sheets: تقرير القطاعات، تقرير المندوبين، العملاء غير المسجلين;
* ``Representatives_Analysis_<execution date>.xlsx``: Cover (one row per representative row of
  تقرير المندوبين + الإجمالي العام), Master Data (one row per transaction counted in the amounts, no
  totals), then one sheet per representative with its Master Data rows.
A customer code missing from Main Data is not an error: its rows go to the row
``عملاء غير مسجلين`` in both main sheets and are listed per code in the third sheet.

The file lifecycle (lock, journal, verification-before-move) is shared with tasks 3 and 4:
see ``grouped.run_batch_managed``.
"""

from __future__ import annotations

import logging
import re
from collections import Counter
from collections.abc import Collection, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.worksheet import Worksheet

from sales_etl.config import (
    COL_CUSTOMER,
    DOC_TYPE_RETURNS,
    DOC_TYPE_SALES,
    PROJECT_ROOT,
    FileManagementConfig,
    PipelineConfig,
)
from sales_etl.grouped import GroupFile, GroupRunResult, _free_stamp, run_batch_managed
from sales_etl.dates import format_dmy
from sales_etl.ingest import discover_csv_files, read_raw_table
from sales_etl.models import ZERO, Category, FileProcessingError
from sales_etl.normalize import parse_transactions
from sales_etl.report import _ensure_new
from sales_etl.text import ascii_digits, canonical, clean_cell

logger = logging.getLogger(__name__)

TASK_PREFIX = "Task5"
JOURNAL_NAME = "task5_journal.jsonl"
DEFAULT_MAIN_DATA_DIR = PROJECT_ROOT / "data"
WORKBOOK_NAME = "تقرير تحليل المبيعات والمرتجعات.xlsx"

MODE_YESTERDAY, MODE_ALL = "yesterday", "all"

SHEET_SECTORS = "تقرير القطاعات"
SHEET_REPS = "تقرير المندوبين"
SHEET_MISSING = "العملاء غير المسجلين"
SHEETS: tuple[str, ...] = (SHEET_SECTORS, SHEET_REPS, SHEET_MISSING)
#: the second workbook: Cover, Master Data, then one sheet per representative
REPS_WORKBOOK_PREFIX = "Representatives_Analysis"
SHEET_COVER = "Cover"
SHEET_MASTER = "Master Data"
SHEET_PIVOT = "Pivot for Last Day"  # the selected day by representative and customer (last sheet)
UNREGISTERED = "عملاء غير مسجلين"
GRAND_TOTAL = "الإجمالي العام"

SECTORS: tuple[str, ...] = (
    "شركات شحن", "الجملة", "صيدليات وسلاسل صيدليات", "وكلاء الدلتا", "وكلاء الصعيد",
    "مستشفيات", "محلات اطفال", "عملاء اون لاين",
)
REPRESENTATIVES: tuple[str, ...] = (
    "اون لاين", "سامح السيد محمود احم", "داليا صلاح الدين فكر", "هشام احمد محمد", "ابراهيم سلام",
    "رامي رأفت يوقيم", "محمد رأفت شحاتة محمد", "احمد محمد علام محمد", "الاداره", "هند محمد رضا",
    "اسلام حمد كمال محمد", "محمود الشيمي", "سميرة محمد اسماعيل", "ندا عبد الناصر طه",
    "دينا كرم ابراهيم", "احمد محمود موسي", "احمد الحريري",
)
#: Task 5 leaves this representative out of every output: her customers' transactions are not counted
#: anywhere (the source files and Main Data are not changed)
EXCLUDED_REPRESENTATIVE = "دينا كرم ابراهيم"
REPORT_REPRESENTATIVES: tuple[str, ...] = tuple(r for r in REPRESENTATIVES if r != EXCLUDED_REPRESENTATIVE)
#: Cover order: group one, its subtotal, group two, its subtotal, then every other row of تقرير
#: المندوبين (extra Main Data names, غير محدد, غير مسجل) with the third subtotal, then الإجمالي العام
COVER_GROUPS: tuple[tuple[str, ...], ...] = (
    ("اون لاين", "الاداره", "هند محمد رضا", "اسلام حمد كمال محمد", "سامح السيد محمود احم",
     "داليا صلاح الدين فكر", "سميرة محمد اسماعيل", "محمود الشيمي", "ندا عبد الناصر طه"),
    ("هشام احمد محمد", "ابراهيم سلام", "رامي رأفت يوقيم", "محمد رأفت شحاتة محمد", "احمد محمد علام محمد",
     "احمد محمود موسي", "احمد الحريري"),
)
SUBTOTAL_LABELS: tuple[str, ...] = ("إجمالي المباشر", "إجمالي الوكلاء", "إجمالي المجموعة الثالثة")  # Cover
#: Cover only: this representative is shown in المباشر, right after the second name, whatever group
#: تقرير المندوبين gives him (there he stays in the third group)
COVER_DIRECT_ADDITION: tuple[str, str] = ("د هاني", "محمود الشيمي")
#: هدف العملاء of each representative in Cover (fixed business targets; never derived from the data).
#: A representative not listed here has no target: Cover shows NO_TARGET and leaves it out of the percentage.
COVER_TARGETS: dict[str, int] = {
    "هند محمد رضا": 25,
    "اسلام حمد كمال محمد": 25,
    "سامح السيد محمود احم": 38,
    "داليا صلاح الدين فكر": 44,
    "سميرة محمد اسماعيل": 25,
    "محمود الشيمي": 21,
    "د هاني": 3,
    "ندا عبد الناصر طه": 25,
    "هشام احمد محمد": 2,
    "ابراهيم سلام": 4,
    "رامي رأفت يوقيم": 1,
    "محمد رأفت شحاتة محمد": 2,
    "احمد محمد علام محمد": 1,
    "احمد محمود موسي": 2,
    "احمد الحريري": 1,
}
NO_TARGET = "لا يوجد هدف"
NO_RATIO = "—"
PERCENT_FORMAT = "0.00%"
#: the same groups in تقرير المندوبين (the third subtotal only when that group has transactions)
REP_SUBTOTAL_LABELS: tuple[str, ...] = ("إجمالي المباشر", "إجمالي الوكلاء", "إجمالي المجموعة الثالثة")
#: the line above every table: the reference date (yesterday report) or the run date (all transactions)
REPORT_DATE_LABEL = "تاريخ التقرير"
#: the first column of every saved table: 1, 2, 3, ... on the data rows, blank on total rows
SERIAL_COLUMN = "م"
MAIN_DATA_HEADERS: tuple[str, ...] = ("كود العميل", "اسم العميل", "قطاعات العملاء", "المندوب")

YESTERDAY_COLUMNS: tuple[str, ...] = (
    "مبيعات أمس", "مبيعات الأيام السابقة", "إجمالي المبيعات",
    "مرتجعات أمس", "مرتجعات الأيام السابقة", "إجمالي المرتجعات", "الصافي",
)
ALL_COLUMNS: tuple[str, ...] = ("إجمالي المبيعات", "إجمالي المرتجعات", "الصافي")
#: distinct customer codes with at least one مبيعات row (a customer with only مرتجعات is never
#: counted), appended after the money columns of the sector and representative sheets.
#: The next two split إجمالي العملاء المميزين by the codes of ``fixed code.xlsx`` (in + out = total).
#: The last one counts the customers with مرتجعات and no مبيعات in the report's scope.
TOTAL_CUSTOMERS_COLUMN = "إجمالي العملاء المميزين"
FIXED_COLUMNS: tuple[str, ...] = ("عدد العملاء من الـ224 كود", "عدد العملاء خارج الـ224 كود")
RETURNS_ONLY_COLUMN = "عدد عملاء المرتجعات فقط"
YESTERDAY_COUNT_COLUMNS: tuple[str, ...] = ("عدد عملاء أمس", "عدد عملاء الأيام السابقة", TOTAL_CUSTOMERS_COLUMN,
                                            *FIXED_COLUMNS, RETURNS_ONLY_COLUMN)
ALL_COUNT_COLUMNS: tuple[str, ...] = (TOTAL_CUSTOMERS_COLUMN, *FIXED_COLUMNS, RETURNS_ONLY_COLUMN)
#: the fixed customer-code list, read from the folder of the Main Data workbook (never a Main Data candidate)
FIXED_CODES_NAME = "fixed code.xlsx"
MISSING_COLUMNS: tuple[str, ...] = ("كود العميل", "اسم العميل", "إجمالي المبيعات", "إجمالي المرتجعات", "الصافي")
MASTER_COLUMNS: tuple[str, ...] = (
    "التاريخ", "المندوب", "القطاع", "كود العميل", "اسم العميل", "المبيعات", "الصافي", "المرتجعات", "حالة كود العميل",
)
PIVOT_SUBTOTAL_PREFIX = "إجمالي "                 # + the representative: its subtotal row in the pivot
PIVOT_RUN_DATE_LABEL = "تاريخ تشغيل التقرير"      # the pivot shows the RUN date (started_at), not the selected day
PIVOT_COLUMNS: tuple[str, ...] = ("اسم المندوب", "العميل", "عدد الفواتير", "إجمالي المبيعات", "إجمالي المرتجعات",
                                   "الصافي")
COVER_COLUMNS: tuple[str, ...] = (
    "المندوب", "إجمالي البيع", "إجمالي المرتجع", "صافي", "هدف العملاء", "عدد العملاء داخل قائمة الـ224 كود",
    "عدد العملاء خارج قائمة الـ224 كود", "نسبة تحقيق هدف العملاء",
)
_SHEET_NAME_FORBIDDEN = re.compile(r"[\[\]:*?/\\]")  # characters Excel refuses in a sheet name
MASTER_UNREGISTERED = "غير مسجل"   # customer code not in Main Data (or blank)
MASTER_UNSPECIFIED = "غير محدد"    # in Main Data, but its sector / representative cell is blank
FIXED_IN_LABEL, FIXED_OUT_LABEL = "داخل قائمة الـ224 كود", "خارج قائمة الـ224 كود"

#: parse_transactions needs a reference date; MODE_ALL never filters by date.
_NO_REFERENCE_DATE = date(2100, 12, 31)
_TOLERANCE = Decimal("0.005")  # Excel stores IEEE doubles
MONEY_FORMAT = "#,##0.00"
COUNT_FORMAT = "#,##0"
DATE_FORMAT = "DD/MM/YYYY"
_FONT = "Segoe UI"
_NAVY = PatternFill("solid", fgColor="1F4E78")
_STRIPE = PatternFill("solid", fgColor="F2F2F2")
_TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
_THIN = Side(style="thin", color="A6A6A6")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_TOTAL_BORDER = Border(left=_THIN, right=_THIN, top=Side(style="double", color="1F4E78"), bottom=_THIN)
_SUBTOTAL_FILL = PatternFill("solid", fgColor="E7E6E6")
_SUBTOTAL_BORDER = Border(left=_THIN, right=_THIN, top=Side(style="thin", color="1F4E78"),
                          bottom=Side(style="thin", color="1F4E78"))
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ONE_PAGE_ROWS = 25  # a summary table up to this many rows is printed on one page


class MainDataError(ValueError):
    """The Main Data workbook is missing, unreadable or ambiguous; nothing is processed."""


def _is_excluded(representative: str) -> bool:
    return canonical(representative) == canonical(EXCLUDED_REPRESENTATIVE)


# ------------------------------------------------------------------ main data
def customer_key(code: Any) -> str:
    """Matching key for a customer code: CSV text ``"100418"`` == Excel number ``100418``."""
    if code is None:
        return ""
    if isinstance(code, float) and code.is_integer():
        code = int(code)
    text = ascii_digits(clean_cell(str(code)))
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return str(int(text)) if text.isdigit() else text


@dataclass(frozen=True)
class MainCustomer:
    code: str
    name: str
    sector: str          # label of the report row (one of SECTORS, or the value as written)
    representative: str  # label of the report row (one of REPRESENTATIVES, or the value as written)


@dataclass
class MainData:
    path: Path
    customers: dict[str, MainCustomer]
    unknown_sectors: set[str] = field(default_factory=set)
    unknown_representatives: set[str] = field(default_factory=set)


def find_main_data(folder: Path = DEFAULT_MAIN_DATA_DIR) -> Path:
    """The single ``.xlsx`` workbook in ``folder`` (Excel's ``~$`` lock files are ignored)."""
    if not folder.is_dir():
        raise MainDataError(f"Main Data folder not found: {folder}")
    found = sorted(p for p in folder.glob("*.xlsx")
                   if not p.name.startswith(("~$", ".")) and p.name.casefold() != FIXED_CODES_NAME.casefold())
    if len(found) != 1:
        raise MainDataError(f"expected exactly one .xlsx workbook in {folder}, found {len(found)}: "
                            f"{[p.name for p in found]}; use --main-data to choose one")
    return found[0]


def _label(value: str, known: Sequence[str]) -> tuple[str, bool]:
    """``(report label, recognised)``: spelling variants (spaces, أ/ا, ة/ه) map to the fixed list."""
    key = canonical(value)
    for label in known:
        if canonical(label) == key:
            return label, True
    return clean_cell(value), False


def load_main_data(path: Path) -> MainData:
    """Read columns A-D of the first sheet (read-only; the workbook is never written)."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001 - any failure: stop before processing
        raise MainDataError(f"cannot open Main Data workbook {path}: {type(exc).__name__}: {exc}") from exc
    try:
        rows = list(wb.worksheets[0].iter_rows(min_col=1, max_col=4, values_only=True))
    finally:
        wb.close()
    header = [canonical(str(v or "")) for v in (rows[0] if rows else ())]
    if header[:4] != [canonical(h) for h in MAIN_DATA_HEADERS]:
        raise MainDataError(f"Main Data header must be {MAIN_DATA_HEADERS}, found {tuple(rows[0]) if rows else None}")
    data = MainData(path=path, customers={})
    for line, row in enumerate(rows[1:], start=2):
        code, name, sector, rep = (list(row) + [None] * 4)[:4]
        key = customer_key(code)
        if not key:
            if any(v not in (None, "") for v in (name, sector, rep)):
                raise MainDataError(f"Main Data row {line} has no customer code")
            continue  # blank row
        sector_label, sector_ok = _label(str(sector or ""), SECTORS)
        rep_label, rep_ok = _label(str(rep or ""), REPRESENTATIVES)
        if not sector_ok:
            data.unknown_sectors.add(sector_label)
        if not rep_ok:
            data.unknown_representatives.add(rep_label)
        customer = MainCustomer(key, clean_cell(str(name or "")), sector_label, rep_label)
        previous = data.customers.get(key)
        if previous and (previous.sector, previous.representative) != (sector_label, rep_label):
            raise MainDataError(f"customer code {key} appears more than once in Main Data with a different "
                                f"sector or representative (row {line}); fix the workbook")
        data.customers.setdefault(key, customer)
    return data


def load_fixed_codes(path: Path) -> frozenset[str]:
    """Customer codes of ``fixed code.xlsx``: column A of the first sheet (read-only, never written).

    Blank cells and repeated codes are ignored. A code is read as text (``customer_key``, the key
    the distinct-customer columns already use), so the split always adds up to their total.
    """
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001 - any failure: stop before processing
        raise MainDataError(f"cannot open the fixed customer-code list {path}: {type(exc).__name__}: {exc}") from exc
    try:
        codes = frozenset(customer_key(row[0]) for row in wb.worksheets[0].iter_rows(max_col=1, values_only=True)
                          if row and customer_key(row[0]))
    finally:
        wb.close()
    if not codes:
        raise MainDataError(f"the fixed customer-code list {path} has no codes in column A")
    return codes


# ------------------------------------------------------------------ calculation
@dataclass
class Amounts:
    """``*_y`` = أمس (MODE_YESTERDAY only); ``*_p`` = الأيام السابقة, or everything in MODE_ALL."""

    sales_y: Decimal = ZERO
    sales_p: Decimal = ZERO
    returns_y: Decimal = ZERO
    returns_p: Decimal = ZERO
    rows: int = 0
    customers_y: set[str] = field(default_factory=set)  # distinct customer keys with مبيعات (blank codes never added)
    customers_p: set[str] = field(default_factory=set)
    customers_r: set[str] = field(default_factory=set)  # distinct customer keys with مرتجعات (either period)

    def add(self, category: Category, yesterday: bool, amount: Decimal, customer: str = "") -> None:
        self.rows += 1
        if customer and category is Category.SALES:  # returns never make a customer count
            (self.customers_y if yesterday else self.customers_p).add(customer)
        elif customer:
            self.customers_r.add(customer)
        if category is Category.SALES:
            if yesterday:
                self.sales_y += amount
            else:
                self.sales_p += amount
        elif yesterday:
            self.returns_y += amount
        else:
            self.returns_p += amount

    @property
    def sales(self) -> Decimal:
        return self.sales_y + self.sales_p

    @property
    def returns(self) -> Decimal:
        return self.returns_y + self.returns_p

    @property
    def net(self) -> Decimal:
        return self.sales + self.returns

    def values(self, mode: str) -> list[Decimal]:
        if mode == MODE_YESTERDAY:
            return [self.sales_y, self.sales_p, self.sales, self.returns_y, self.returns_p, self.returns, self.net]
        return [self.sales, self.returns, self.net]

    def counts(self, mode: str, fixed: frozenset[str]) -> list[int]:
        """Distinct customers; the period count is a set union, never yesterday + previous days.

        The period's customers are then split: codes in ``fixed`` / the others. Last: customers with
        مرتجعات and no مبيعات in the period (returns only).
        """
        period = self.customers_y | self.customers_p
        split = [len(period & fixed), len(period - fixed), len(self.customers_r - period)]
        if mode == MODE_YESTERDAY:
            return [len(self.customers_y), len(self.customers_p), len(period), *split]
        return [len(period), *split]


@dataclass(frozen=True)
class MasterRow:
    """One Master Data row. ``category``, ``yesterday`` and the two ``*_row`` labels (the summary
    rows it is counted in) are kept for the reconciliation and are not written."""

    day: date
    representative: str
    sector: str
    code: str
    name: str
    net: Decimal
    status: str
    category: Category
    yesterday: bool
    sector_row: str
    rep_row: str
    doc_no: str = ""  # رقم المستند (Pivot for Last Day counts distinct sales documents); not written

    @property
    def sales(self) -> Decimal:
        """المبيعات: الصافي of a sale, with its sign; 0 for a return."""
        return self.net if self.category is Category.SALES else ZERO

    @property
    def returns(self) -> Decimal:
        """المرتجعات: الصافي of a return, with its sign; 0 for a sale."""
        return self.net if self.category is Category.RETURNS else ZERO

    def values(self) -> list[Any]:
        return [self.day, self.representative, self.sector, self.code, self.name, self.sales, self.net, self.returns,
                self.status]


def _master_label(row_label: str) -> str:
    """Master Data text for a summary row label: unregistered / blank Main Data cell get a label."""
    if row_label == UNREGISTERED:
        return MASTER_UNREGISTERED
    return row_label or MASTER_UNSPECIFIED


@dataclass
class MissingCustomer:
    code: str
    name: str = ""
    amounts: Amounts = field(default_factory=Amounts)


@dataclass
class AnalysisResult:
    mode: str
    reference_date: date | None
    main_data: MainData
    input_dir: Path
    started_at: datetime
    fixed_codes: frozenset[str] = frozenset()  # codes of fixed code.xlsx
    files: list[GroupFile] = field(default_factory=list)
    sectors: dict[str, Amounts] = field(default_factory=dict)
    representatives: dict[str, Amounts] = field(default_factory=dict)
    missing: dict[str, MissingCustomer] = field(default_factory=dict)
    direct: Amounts = field(default_factory=Amounts)  # independent: no mapping involved
    #: (document type text, date, customer code as read) of every parsed row: the independent customer recount
    source_rows: list[tuple[str, date, str]] = field(default_factory=list)
    master: list[MasterRow] = field(default_factory=list)  # Master Data, in file and line order
    after_date_rows: int = 0
    other_type_rows: int = 0
    excluded_rows: int = 0      # transactions of EXCLUDED_REPRESENTATIVE's customers, left out of every output
    excluded_net: Decimal = ZERO
    run_dir: Path | None = None
    report_path: Path | None = None
    reps_report_path: Path | None = None  # Representatives_Analysis_<date>.xlsx

    def __post_init__(self) -> None:
        self.sectors = {name: Amounts() for name in SECTORS}
        self.representatives = {name: Amounts() for name in REPORT_REPRESENTATIVES}

    @property
    def report_date(self) -> date:
        """The date of the report: the reference date, or the run date when there is none."""
        return self.reference_date or self.started_at.date()

    @property
    def columns(self) -> tuple[str, ...]:
        return YESTERDAY_COLUMNS if self.mode == MODE_YESTERDAY else ALL_COLUMNS

    @property
    def count_columns(self) -> tuple[str, ...]:
        return YESTERDAY_COUNT_COLUMNS if self.mode == MODE_YESTERDAY else ALL_COUNT_COLUMNS

    @property
    def failed(self) -> list[GroupFile]:
        return [f for f in self.files if f.error is not None]

    @property
    def rejected_rows(self) -> int:
        return sum(f.rejected_rows for f in self.files)

    @property
    def complete(self) -> bool:
        """Every file was read and every transaction row was parsed."""
        return bool(self.files) and not self.failed and not self.rejected_rows

    @property
    def extra_labels(self) -> tuple[list[str], list[str]]:
        """Sector / representative values in Main Data that are not in the fixed lists."""
        return ([k for k in self.sectors if k not in SECTORS and k != UNREGISTERED],
                [k for k in self.representatives if k not in REPORT_REPRESENTATIVES and k != UNREGISTERED])


def _add_file(path: Path, config: PipelineConfig, result: AnalysisResult) -> GroupFile:
    item = GroupFile(file_name=path.name)
    reference = result.reference_date or _NO_REFERENCE_DATE
    try:
        raw = read_raw_table(
            path, encodings=config.encodings, delimiters=config.delimiters,
            forced_encoding=config.forced_encoding, forced_delimiter=config.forced_delimiter,
        )
        if COL_CUSTOMER not in raw.column_index:
            raise FileProcessingError("MISSING_CODE_COLUMN", f"required column '{COL_CUSTOMER}' not found")
        i_name = raw.column_index[COL_CUSTOMER] + 1  # column G: the unlabelled name column after the code
        parsed = parse_transactions(raw, reference)
    except FileProcessingError as exc:
        item.error = f"{exc.code}: {exc.message}"
        return item
    except Exception as exc:  # noqa: BLE001 - isolate unexpected failures per file
        item.error = f"UNEXPECTED_ERROR: {type(exc).__name__}: {exc}"
        logger.exception("unexpected failure", extra={"event": "task5_file_failed", "file": path.name})
        return item
    item.sha256 = raw.sha256
    item.transactions, item.rejected_rows = len(parsed.transactions), len(parsed.rejected)
    customers = result.main_data.customers
    result.source_rows += [(t.doc_type, t.txn_date, t.customer) for t in parsed.transactions]
    for t in parsed.transactions:
        if t.category not in (Category.SALES, Category.RETURNS):
            result.other_type_rows += 1
            continue
        yesterday = False
        if result.mode == MODE_YESTERDAY:
            if t.txn_date > reference:
                result.after_date_rows += 1
                continue
            yesterday = t.txn_date == reference
        key = customer_key(t.customer)
        known = customers.get(key)
        if known is not None and _is_excluded(known.representative):
            result.excluded_rows += 1
            result.excluded_net += t.net
            continue
        result.direct.add(t.category, yesterday, t.net, key)
        if known is None:
            sector = representative = UNREGISTERED
            missing = result.missing.setdefault(key, MissingCustomer(code=clean_cell(t.customer)))
            if not missing.name and i_name < len(t.cells):
                missing.name = t.cells[i_name]
            missing.amounts.add(t.category, yesterday, t.net, key)
        else:
            sector, representative = known.sector, known.representative
        result.sectors.setdefault(sector, Amounts()).add(t.category, yesterday, t.net, key)
        result.representatives.setdefault(representative, Amounts()).add(t.category, yesterday, t.net, key)
        result.master.append(MasterRow(
            t.txn_date, _master_label(representative), _master_label(sector), key,
            t.cells[i_name] if i_name < len(t.cells) else "", t.net,
            FIXED_IN_LABEL if key in result.fixed_codes else FIXED_OUT_LABEL,
            t.category, yesterday, sector, representative, doc_no=t.doc_no))
    return item


def _code_key(code: str) -> tuple[int, int | str]:
    return (0, int(code)) if code.isdigit() else (1, code)


def _ordered(groups: dict[str, Amounts], fixed: Sequence[str]) -> list[str]:
    """The fixed list first, then any extra Main Data values, then عملاء غير مسجلين (always last)."""
    extras = sorted(k for k in groups if k not in fixed and k != UNREGISTERED)
    return [*fixed, *extras, UNREGISTERED]


def _with_total(detail: list[list[Any]], first_money_col: int, counts: Sequence[int] = ()) -> list[list[Any]]:
    """Append الإجمالي العام = the column sums of the money columns, then ``counts`` as given.

    ``counts`` are the distinct customers of the whole report, computed independently: a customer
    is counted once in the total, whatever the number of rows or date groups it appears in.
    """
    width = (len(detail[0]) if detail else first_money_col) - len(counts)
    total: list[Any] = [GRAND_TOTAL, *[""] * (first_money_col - 1)]
    for col in range(first_money_col, width):
        total.append(sum((row[col] for row in detail), ZERO))
    return [*detail, [*total, *counts]]


def sector_table(result: AnalysisResult) -> list[list[Any]]:
    groups = {**result.sectors, UNREGISTERED: result.sectors.get(UNREGISTERED, Amounts())}
    fixed = result.fixed_codes
    return _with_total([[name, *groups[name].values(result.mode), *groups[name].counts(result.mode, fixed)]
                        for name in _ordered(groups, SECTORS)], 1, result.direct.counts(result.mode, fixed))


def representative_table(result: AnalysisResult) -> list[list[Any]]:
    groups = {**result.representatives, UNREGISTERED: result.representatives.get(UNREGISTERED, Amounts())}
    fixed = result.fixed_codes
    return _with_total([[name, *groups[name].values(result.mode), *groups[name].counts(result.mode, fixed)]
                        for name in representative_rows(result)], 1, result.direct.counts(result.mode, fixed))


def representative_display_table(result: AnalysisResult) -> list[list[Any]]:
    """تقرير المندوبين as written: the rows in Cover order, each group followed by its subtotal (amounts summed,
    distinct customers of the group; the third subtotal only when that group has transactions), then
    الإجمالي العام."""
    *detail, total = representative_table(result)
    by_label = {row[0]: row for row in detail}
    rows: list[list[Any]] = []
    for group, label in zip(cover_groups(result), REP_SUBTOTAL_LABELS):
        rows += [by_label[name] for name in group]
        merged = _merged([result.representatives.get(name, Amounts()) for name in group])
        if label != REP_SUBTOTAL_LABELS[2] or merged.rows:
            rows.append([label, *merged.values(result.mode), *merged.counts(result.mode, result.fixed_codes)])
    return [*rows, total]


def rep_subtotal_rows(rows: list[list[Any]]) -> set[int]:
    """Indexes of the group subtotal rows of تقرير المندوبين."""
    return {i for i, row in enumerate(rows) if row[0] in REP_SUBTOTAL_LABELS}


def missing_table(result: AnalysisResult) -> list[list[Any]]:
    rows = [[m.code, m.name, m.amounts.sales, m.amounts.returns, m.amounts.net]
            for m in sorted(result.missing.values(), key=lambda m: _code_key(m.code))]
    return _with_total(rows, 2) if rows else [[GRAND_TOTAL, "", ZERO, ZERO, ZERO]]


def master_rows(result: AnalysisResult) -> list[MasterRow]:
    """Master Data: one row per transaction counted in the amounts (then its الإجمالي العام)."""
    return list(result.master)


def cover_groups(result: AnalysisResult) -> list[list[str]]:
    """Row labels of تقرير المندوبين by Cover group: the two fixed groups in their order, then every
    other row (extra Main Data names, غير محدد, غير مسجل) in the order of تقرير المندوبين."""
    named = {label for group in COVER_GROUPS for label in group}
    others = [label for label in _ordered({**result.representatives, UNREGISTERED: Amounts()}, REPORT_REPRESENTATIVES)
              if label not in named]
    return [*(list(group) for group in COVER_GROUPS), others]


def representative_rows(result: AnalysisResult) -> list[str]:
    """Representative row labels in Cover order: the representative sheets follow it."""
    return [label for group in cover_groups(result) for label in group]


def _merged(parts: Sequence[Amounts]) -> Amounts:
    """One Amounts for several rows: sums, and the union of their customer sets (never a sum of counts)."""
    total = Amounts()
    for a in parts:
        total.sales_y += a.sales_y
        total.sales_p += a.sales_p
        total.returns_y += a.returns_y
        total.returns_p += a.returns_p
        total.rows += a.rows
        total.customers_y |= a.customers_y
        total.customers_p |= a.customers_p
        total.customers_r |= a.customers_r
    return total


def cover_display_groups(result: AnalysisResult) -> list[list[str]]:
    """The Cover groups: those of :func:`cover_groups`, with ``COVER_DIRECT_ADDITION`` moved into المباشر
    (after its anchor) - also when he has no transactions. تقرير المندوبين keeps :func:`cover_groups`."""
    direct, agents, others = cover_groups(result)
    name, after = COVER_DIRECT_ADDITION
    label = next((x for x in others if canonical(x) == canonical(name)), name)
    direct = list(direct)
    direct.insert(direct.index(after) + 1, label)
    return [direct, list(agents), [x for x in others if x != label]]


def cover_target(label: str) -> int | None:
    """هدف العملاء of a representative row (spelling variants match), None when none is configured."""
    return next((target for name, target in COVER_TARGETS.items() if canonical(name) == canonical(label)), None)


def _ratio(inside: int, target: int | None) -> Decimal | str:
    """نسبة تحقيق هدف العملاء as a fraction (shown as 0.00%): customers inside the 224 list / target."""
    if not target:
        return NO_RATIO
    return (Decimal(inside) / Decimal(target)).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)


def _cover_row(label: str, a: Amounts, result: AnalysisResult, target: int | None) -> list[Any]:
    inside, outside = a.counts(result.mode, result.fixed_codes)[-3:-1]
    return [label, a.sales, a.returns, a.net, target if target else NO_TARGET, inside, outside, _ratio(inside, target)]


def cover_table(result: AnalysisResult) -> list[list[Any]]:
    """Cover, in COVER_COLUMNS order: every representative row (never hidden, also at zero), its group's subtotal
    after each group, then الإجمالي العام. Amounts as in تقرير المندوبين; the customers inside / outside the 224
    list are the distinct sales customers (as إجمالي العملاء المميزين); the target comes from COVER_TARGETS.
    A subtotal / the total: summed amounts and targets, distinct customers, and the percentage of the summed
    inside customers of the representatives that have a target over their summed targets (never a sum of %)."""
    rows: list[list[Any]] = []
    grand_target = grand_inside = 0
    for group, subtotal in zip(cover_display_groups(result), SUBTOTAL_LABELS):
        parts, sub_target, sub_inside = [], 0, 0
        for name in group:
            a = result.representatives.get(name, Amounts())
            target = cover_target(name)
            row = _cover_row(_master_label(name), a, result, target)
            rows.append(row)
            parts.append(a)
            if target:
                sub_target, sub_inside = sub_target + target, sub_inside + row[5]
        merged = _merged(parts)
        inside, outside = merged.counts(result.mode, result.fixed_codes)[-3:-1]
        rows.append([subtotal, merged.sales, merged.returns, merged.net, sub_target or NO_TARGET, inside, outside,
                     _ratio(sub_inside, sub_target)])
        grand_target, grand_inside = grand_target + sub_target, grand_inside + sub_inside
    inside, outside = result.direct.counts(result.mode, result.fixed_codes)[-3:-1]
    return [*rows, [GRAND_TOTAL, result.direct.sales, result.direct.returns, result.direct.net,
                    grand_target or NO_TARGET, inside, outside, _ratio(grand_inside, grand_target)]]


def cover_subtotal_rows(rows: list[list[Any]]) -> set[int]:
    """Indexes of the group subtotal rows of a Cover table."""
    return {i for i, row in enumerate(rows) if row[0] in SUBTOTAL_LABELS}


def representative_sheet_names(result: AnalysisResult) -> dict[str, str]:
    """Row label -> a valid, unique worksheet name (31 characters at most; the full name stays in the cells)."""
    used = {SHEET_COVER.casefold(), SHEET_MASTER.casefold()}
    names: dict[str, str] = {}
    for label in representative_rows(result):
        base = _SHEET_NAME_FORBIDDEN.sub("_", _master_label(label)).strip("' ")[:31] or "Sheet"
        name, n = base, 1
        while name.casefold() in used:
            n += 1
            suffix = f" ({n})"
            name = base[:31 - len(suffix)].rstrip() + suffix
        used.add(name.casefold())
        names[label] = name
    return names


def _pivot_groups(result: AnalysisResult) -> dict[tuple[str, str], dict[str, Any]]:
    """(representative row, customer code) -> the selected day's figures. The name is the CSV customer name
    (column G, the cell after العميل) of the customer's first row with one; never Main Data."""
    day = result.reference_date
    groups: dict[tuple[str, str], dict[str, Any]] = {}
    for r in result.master:
        if day is None or r.day != day:
            continue
        g = groups.setdefault((r.rep_row, r.code), dict(rep=r.representative, name="", invoices=set(),
                                                        sales=ZERO, returns=ZERO))
        if not g["name"] and r.name:
            g["name"] = r.name
        if r.category is Category.SALES:
            g["invoices"].add(r.doc_no)
            g["sales"] += r.net
        else:
            g["returns"] += r.net
    return groups


def pivot_subtotal_label(representative: str) -> str:
    return f"{PIVOT_SUBTOTAL_PREFIX}{representative}"


def pivot_table(result: AnalysisResult) -> list[list[Any]]:
    """Pivot for Last Day: for each representative (Cover order) its customers (by name) on the selected
    reference date only, then its subtotal; at the end الإجمالي العام. عدد الفواتير = distinct sales document
    numbers (رقم المستند; a return is never an invoice); sales and returns keep their signs; الصافي = sales +
    returns. Without a reference date (all transactions) there is no selected day: only a zero total."""
    groups = _pivot_groups(result)
    order = {label: i for i, label in enumerate(representative_rows(result))}
    rows: list[list[Any]] = []
    totals = [0, ZERO, ZERO]
    for rep_row in sorted({key[0] for key in groups}, key=lambda label: (order.get(label, len(order)), label)):
        mine = sorted(((code, g) for (label, code), g in groups.items() if label == rep_row),
                      key=lambda item: (item[1]["name"] or item[0], item[0]))
        sub = [0, ZERO, ZERO]
        for code, g in mine:
            rows.append([g["rep"], g["name"] or code, len(g["invoices"]), g["sales"], g["returns"],
                         g["sales"] + g["returns"]])
            sub = [sub[0] + len(g["invoices"]), sub[1] + g["sales"], sub[2] + g["returns"]]
        rows.append([pivot_subtotal_label(mine[0][1]["rep"]), "", sub[0], sub[1], sub[2], sub[1] + sub[2]])
        totals = [totals[0] + sub[0], totals[1] + sub[1], totals[2] + sub[2]]
    return [*rows, [GRAND_TOTAL, "", totals[0], totals[1], totals[2], totals[1] + totals[2]]]


def pivot_subtotal_rows(rows: list[list[Any]]) -> set[int]:
    """Indexes of the representative subtotal rows of a pivot table (label إجمالي <representative>, no customer)."""
    return {i for i, row in enumerate(rows[:-1])
            if str(row[0]).startswith(PIVOT_SUBTOTAL_PREFIX) and row[1] in ("", None)}


def pivot_title(result: AnalysisResult) -> str:
    """Row 1 of the pivot: the RUN date (the day the task was run, ``started_at``), then the transactions' day
    (the selected reference date), which is a different date."""
    run = f"{PIVOT_RUN_DATE_LABEL}: {format_dmy(result.started_at.date())}"
    if result.reference_date is None:
        return (f"{run}  |  تقرير كل الحركات: لا يوجد يوم مختار، لذلك لا توجد بيانات في هذه الورقة "
                "(تُملأ عند اختيار تقرير أمس والأيام السابقة)")
    return f"{run}  |  حركات يوم {format_dmy(result.reference_date)} حسب المندوب والعميل"


def date_headers(columns: Sequence[str], reference_date: date | None) -> tuple[str, ...]:
    """Column headers as printed: with a reference date D, أمس becomes ``يوم D`` and الأيام السابقة
    becomes ``حتى D-1`` (DD/MM/YYYY): مبيعات حتى 06/11/2026 for D = 07/11/2026. Without one (all
    transactions): unchanged. The totals (إجمالي ...) are أمس + the earlier days, so they include D."""
    if reference_date is None:
        return tuple(columns)
    day, before = format_dmy(reference_date), format_dmy(reference_date - timedelta(days=1))
    return tuple(column.replace("الأيام السابقة", f"حتى {before}").replace("أمس", f"يوم {day}") for column in columns)


def sheet_tables(result: AnalysisResult) -> dict[str, tuple[tuple[str, ...], list[list[Any]]]]:
    """Sheet name -> (headers, rows incl. الإجمالي العام) of the report's three sheets, as written
    (تقرير المندوبين with its group subtotals)."""
    return {
        SHEET_SECTORS: (date_headers(("قطاع العملاء", *result.columns, *result.count_columns), result.reference_date),
                        sector_table(result)),
        SHEET_REPS: (date_headers(("المندوب", *result.columns, *result.count_columns), result.reference_date),
                     representative_display_table(result)),
        SHEET_MISSING: (MISSING_COLUMNS, missing_table(result)),
    }


# ------------------------------------------------------------------ reconciliation
def independent_customer_counts(result: AnalysisResult) -> dict[str, dict[str, list[int]]]:
    """Sheet -> row label -> customer counts, recounted with a plain loop over the source rows.

    Works from the document type text and the raw date, not from ``Amounts``: only مبيعات rows
    qualify a customer, مرتجعات never do; a code with مرتجعات and no مبيعات is a returns-only
    customer. ``الإجمالي العام`` is recounted over all rows at once.
    """
    reference = result.reference_date
    seen: dict[str, dict[str, dict[str, set[str]]]] = {SHEET_SECTORS: {}, SHEET_REPS: {}}
    every: dict[str, set[str]] = {}
    for doc_type, day, code in result.source_rows:
        key, kind = customer_key(code), clean_cell(doc_type)
        if not key or kind not in (DOC_TYPE_SALES, DOC_TYPE_RETURNS) or (reference is not None and day > reference):
            continue
        known = result.main_data.customers.get(key)
        if known is not None and _is_excluded(known.representative):
            continue
        if kind == DOC_TYPE_RETURNS:
            when = "returns"
        else:
            when = "yesterday" if reference is not None and day == reference else "previous"
        for sheet, label in ((SHEET_SECTORS, known.sector if known else UNREGISTERED),
                             (SHEET_REPS, known.representative if known else UNREGISTERED)):
            seen[sheet].setdefault(label, {}).setdefault(key, set()).add(when)
        every.setdefault(key, set()).add(when)

    def counts(codes: dict[str, set[str]]) -> list[int]:
        buyers = [code for code, flags in codes.items() if flags - {"returns"}]
        inside = sum(1 for code in buyers if code in result.fixed_codes)
        split = [inside, len(buyers) - inside, sum(1 for flags in codes.values() if flags == {"returns"})]
        if result.mode == MODE_YESTERDAY:
            return [sum("yesterday" in f for f in codes.values()), sum("previous" in f for f in codes.values()),
                    len(buyers), *split]
        return [len(buyers), *split]

    return {sheet: {**{label: counts(codes) for label, codes in groups.items()}, GRAND_TOTAL: counts(every)}
            for sheet, groups in seen.items()}


def reconcile(result: AnalysisResult) -> list[str]:
    """Independent cross-checks of the four sheets; any message means the run must not move files."""
    problems: list[str] = []
    sector_rows, rep_rows, missing_rows = sector_table(result), representative_table(result), missing_table(result)
    width = len(result.columns)

    for name, rows, first, end in ((SHEET_SECTORS, sector_rows, 1, 1 + width), (SHEET_REPS, rep_rows, 1, 1 + width),
                                   (SHEET_MISSING, missing_rows, 2, len(missing_rows[-1]))):
        *detail, total = rows
        labels = [r[0] for r in detail]
        if len(labels) != len(set(labels)):
            problems.append(f"{name}: duplicate rows")
        for col in range(first, end):
            if sum((r[col] for r in detail), ZERO) != total[col]:
                problems.append(f"{name}: grand total column {col + 1} differs from the sum of its rows")

    totals = {name: rows[-1] for name, rows in ((SHEET_SECTORS, sector_rows), (SHEET_REPS, rep_rows))}
    if totals[SHEET_SECTORS][1:] != totals[SHEET_REPS][1:]:
        problems.append("sector and representative grand totals differ")
    expected = result.direct.values(result.mode)
    if totals[SHEET_SECTORS][1:1 + width] != expected:
        problems.append(f"grand totals {totals[SHEET_SECTORS][1:]} differ from the direct sum of all "
                        f"transactions {expected}")
    missing_total = missing_rows[-1][2:]
    for name, rows in ((SHEET_SECTORS, sector_rows), (SHEET_REPS, rep_rows)):
        row = next(r for r in rows if r[0] == UNREGISTERED)
        amounts = row[1:1 + width]
        sales, returns, net = (amounts[2], amounts[5], amounts[6]) if result.mode == MODE_YESTERDAY else amounts
        if [sales, returns, net] != missing_total:
            problems.append(f"{name}: row '{UNREGISTERED}' differs from the {SHEET_MISSING} total")
    for name, groups in ((SHEET_SECTORS, result.sectors), (SHEET_REPS, result.representatives)):
        if sum(a.rows for a in groups.values()) != result.direct.rows:
            problems.append(f"{name}: {sum(a.rows for a in groups.values())} transactions counted, "
                            f"{result.direct.rows} expected (each must be counted exactly once)")
    if sum(m.amounts.rows for m in result.missing.values()) != result.sectors.get(UNREGISTERED, Amounts()).rows:
        problems.append("unregistered customers: transaction count mismatch")

    # distinct customers: the grand total (from all transactions, result.direct) must equal the
    # distinct customers recomputed from the union of the rows' customers
    for name, rows, groups in ((SHEET_SECTORS, sector_rows, result.sectors),
                               (SHEET_REPS, rep_rows, result.representatives)):
        union = Amounts(customers_y=set().union(*(a.customers_y for a in groups.values())),
                        customers_p=set().union(*(a.customers_p for a in groups.values())),
                        customers_r=set().union(*(a.customers_r for a in groups.values())))
        if rows[-1][1 + width:] != union.counts(result.mode, result.fixed_codes):
            problems.append(f"{name}: grand total customer counts {rows[-1][1 + width:]} differ from the distinct "
                            f"customers of its rows {union.counts(result.mode, result.fixed_codes)}")
        # Main Data gives each code one sector and one representative, so no customer may be in two rows
        for col in range(1 + width, len(rows[-1])):
            if sum(r[col] for r in rows[:-1]) != rows[-1][col]:
                problems.append(f"{name}: a customer is counted in more than one row (column {col + 1})")
        total_col = 1 + width + result.count_columns.index(TOTAL_CUSTOMERS_COLUMN)
        for row in rows:  # in the 224 codes + outside them = إجمالي العملاء المميزين, in every row
            if row[total_col + 1] + row[total_col + 2] != row[total_col]:
                problems.append(f"{name}: '{row[0]}' {FIXED_COLUMNS[0]} {row[total_col + 1]} + {FIXED_COLUMNS[1]} "
                                f"{row[total_col + 2]} differ from {TOTAL_CUSTOMERS_COLUMN} {row[total_col]}")
        unregistered = next(r for r in rows if r[0] == UNREGISTERED)[total_col]
        if unregistered != sum(1 for k, m in result.missing.items()
                               if k and (m.amounts.customers_y or m.amounts.customers_p)):
            problems.append(f"{name}: row '{UNREGISTERED}' customer count differs from the {SHEET_MISSING} sheet")

    # every row's customer counts vs the independent recount from the source rows (sales only)
    recount = independent_customer_counts(result)
    zero = [0] * len(result.count_columns)
    for name, rows in ((SHEET_SECTORS, sector_rows), (SHEET_REPS, rep_rows)):
        for row in rows:
            want = recount[name].get(row[0], zero)
            if row[1 + width:] != want:
                problems.append(f"{name}: '{row[0]}' customer counts {row[1 + width:]} differ from the "
                                f"independent recount of customers with sales {want}")
    return (problems + reconcile_master(result, sector_rows, rep_rows) + reconcile_cover(result, rep_rows)
            + reconcile_rep_subtotals(result) + reconcile_pivot(result))


def reconcile_master(result: AnalysisResult, sector_rows: list[list[Any]], rep_rows: list[list[Any]]) -> list[str]:
    """Master Data vs the summary sheets, the source rows, Main Data and the fixed-code list."""
    problems: list[str] = []
    master, direct, reference = result.master, result.direct, result.reference_date

    # every counted transaction exactly once: the source rows filtered independently (type text, raw date)
    customers = result.main_data.customers

    def excluded(code: str) -> bool:
        known = customers.get(customer_key(code))
        return known is not None and _is_excluded(known.representative)

    included = sum(1 for doc_type, day, code in result.source_rows
                   if clean_cell(doc_type) in (DOC_TYPE_SALES, DOC_TYPE_RETURNS)
                   and (reference is None or day <= reference) and not excluded(code))
    if any(_is_excluded(r.rep_row) for r in master):
        problems.append(f"{SHEET_MASTER}: rows of {EXCLUDED_REPRESENTATIVE}, who is excluded")
    if not len(master) == direct.rows == included:
        problems.append(f"{SHEET_MASTER}: {len(master)} rows, {direct.rows} transactions in the totals, "
                        f"{included} in the source files")
    sales = sum((r.net for r in master if r.category is Category.SALES), ZERO)
    returns = sum((r.net for r in master if r.category is Category.RETURNS), ZERO)
    net = sum((r.net for r in master), ZERO)
    for label, got, want in (("sales", sales, direct.sales), ("returns", returns, direct.returns),
                             ("net", net, direct.net)):
        if got != want:
            problems.append(f"{SHEET_MASTER}: {label} {got} differs from the report total {want}")
    for name, category, want_total in (("المبيعات", Category.SALES, direct.sales),
                                       ("المرتجعات", Category.RETURNS, direct.returns)):
        column = [r.values()[MASTER_COLUMNS.index(name)] for r in master]
        wrong = sum(1 for r, v in zip(master, column) if v != (r.net if r.category is category else ZERO))
        if wrong or sum(column, ZERO) != want_total:
            problems.append(f"{SHEET_MASTER}: column {name} {sum(column, ZERO)} ({wrong} wrong row(s)) is not the "
                            f"signed total {want_total}")
    # الإجمالي العام vs the source rows: distinct customer codes and representatives of the counted transactions
    keys, names = set(), set()
    for doc_type, day, code in result.source_rows:
        if (clean_cell(doc_type) not in (DOC_TYPE_SALES, DOC_TYPE_RETURNS) or (reference is not None and day > reference)
                or excluded(code)):
            continue
        key = customer_key(code)
        if key:
            keys.add(key)
        known = customers.get(key) if key else None
        if known is not None and known.representative:
            names.add(known.representative)
    total = detail_total(master)
    if total[1] != len(names) or total[3] != len(keys) or total[5:8] != [direct.sales, direct.net, direct.returns] \
            or total[5] + total[7] != total[6]:
        problems.append(f"{SHEET_MASTER}: الإجمالي العام {total[1:8]} differs from the source rows (representatives "
                        f"{len(names)}, customers {len(keys)}, sales / net / returns {direct.sales} / {direct.net} / "
                        f"{direct.returns})")

    # every sector / representative row = its Master Data rows (amounts and customer counts)
    for name, rows, attr in ((SHEET_SECTORS, sector_rows, "sector_row"), (SHEET_REPS, rep_rows, "rep_row")):
        groups: dict[str, Amounts] = {}
        for r in master:
            groups.setdefault(getattr(r, attr), Amounts()).add(r.category, r.yesterday, r.net, r.code)
        for row in rows[:-1]:
            want = groups.pop(row[0], Amounts())
            if row[1:] != [*want.values(result.mode), *want.counts(result.mode, result.fixed_codes)]:
                problems.append(f"{name}: row '{row[0]}' differs from its {SHEET_MASTER} rows")
        if groups:
            problems.append(f"{name}: {SHEET_MASTER} has rows for {sorted(groups)} that the sheet does not show")

    # every row's representative, sector and 224-code status, re-derived from its customer code
    wrong_mapping = wrong_status = 0
    for r in master:
        known = result.main_data.customers.get(r.code) if r.code else None
        rep, sector = (known.representative, known.sector) if known else (UNREGISTERED, UNREGISTERED)
        if (r.rep_row, r.sector_row, r.representative, r.sector) != (rep, sector, _master_label(rep),
                                                                     _master_label(sector)):
            wrong_mapping += 1
        if r.status != (FIXED_IN_LABEL if r.code in result.fixed_codes else FIXED_OUT_LABEL):
            wrong_status += 1
    if wrong_mapping:
        problems.append(f"{SHEET_MASTER}: {wrong_mapping} row(s) with a representative or sector that differs "
                        "from Main Data")
    if wrong_status:
        problems.append(f"{SHEET_MASTER}: {wrong_status} row(s) whose 224-code status differs from {FIXED_CODES_NAME}")
    return problems


def independent_group_counts(result: AnalysisResult, labels: set[str]) -> list[int]:
    """The customer counts (``count_columns`` order) of the customers whose representative row is in
    ``labels``, recounted with a plain loop over the source rows (distinct codes, never a sum)."""
    reference = result.reference_date
    codes: dict[str, set[str]] = {}
    for doc_type, day, code in result.source_rows:
        key, kind = customer_key(code), clean_cell(doc_type)
        if not key or kind not in (DOC_TYPE_SALES, DOC_TYPE_RETURNS) or (reference is not None and day > reference):
            continue
        known = result.main_data.customers.get(key)
        rep = known.representative if known else UNREGISTERED
        if _is_excluded(rep) or rep not in labels:
            continue
        if kind == DOC_TYPE_RETURNS:
            flag = "return"
        else:
            flag = "yesterday" if reference is not None and day == reference else "previous"
        codes.setdefault(key, set()).add(flag)
    buyers = [code for code, flags in codes.items() if flags - {"return"}]
    inside = sum(1 for code in buyers if code in result.fixed_codes)
    split = [len(buyers), inside, len(buyers) - inside, sum(1 for flags in codes.values() if flags == {"return"})]
    if result.mode == MODE_YESTERDAY:
        return [sum("yesterday" in f for f in codes.values()), sum("previous" in f for f in codes.values()), *split]
    return split


def reconcile_rep_subtotals(result: AnalysisResult) -> list[str]:
    """تقرير المندوبين as written: Cover order, the group subtotals (the third only with transactions), each
    subtotal = the sum of its rows' amounts and an independent recount of its group's distinct customers."""
    problems: list[str] = []
    rows, groups = representative_display_table(result), cover_groups(result)
    detail = {row[0]: row for row in representative_table(result)[:-1]}
    expected: list[str] = []
    for group, label in zip(groups, REP_SUBTOTAL_LABELS):
        expected += group
        if label != REP_SUBTOTAL_LABELS[2] or any(r.rep_row in set(group) for r in result.master):
            expected.append(label)
    expected.append(GRAND_TOTAL)
    if [row[0] for row in rows] != expected:
        return [f"{SHEET_REPS}: rows {[row[0] for row in rows]} differ from {expected}"]
    if groups[:2] != [list(group) for group in COVER_GROUPS]:
        problems.append(f"{SHEET_REPS}: the first two groups are not {COVER_GROUPS}")
    width = len(result.columns)
    for group, label in zip(groups, REP_SUBTOTAL_LABELS):
        row = next((r for r in rows if r[0] == label), None)
        if row is None:
            continue
        money = [sum((detail[name][col] for name in group), ZERO) for col in range(1, 1 + width)]
        if row[1:1 + width] != money:
            problems.append(f"{SHEET_REPS}: '{label}' amounts {row[1:1 + width]} differ from the sum of its rows {money}")
        want = independent_group_counts(result, set(group))
        if row[1 + width:] != want:
            problems.append(f"{SHEET_REPS}: '{label}' customer counts {row[1 + width:]} differ from the distinct "
                            f"customers of its group {want}")
    return problems


def _master_amounts(rows: Sequence[MasterRow]) -> list[Decimal]:
    sales = sum((r.net for r in rows if r.category is Category.SALES), ZERO)
    returns = sum((r.net for r in rows if r.category is Category.RETURNS), ZERO)
    return [sales, returns, sales + returns]


def reconcile_cover(result: AnalysisResult, rep_rows: list[list[Any]]) -> list[str]:
    """Cover: its groups (المباشر with COVER_DIRECT_ADDITION, then الوكلاء, then the rest; nobody hidden), each
    representative row vs تقرير المندوبين / its Master Data rows / the independent customer recount / COVER_TARGETS,
    each subtotal and الإجمالي العام vs Master Data, an independent recount of the distinct customers of the group /
    report, the summed targets, and the percentage recomputed from those (never a sum of percentages)."""
    problems: list[str] = []
    rows, groups = cover_table(result), cover_display_groups(result)
    expected: list[str] = []
    for group, subtotal in zip(groups, SUBTOTAL_LABELS):
        expected += [*(_master_label(label) for label in group), subtotal]
    expected.append(GRAND_TOTAL)
    if [r[0] for r in rows] != expected:
        return [f"{SHEET_COVER}: rows {[r[0] for r in rows]} differ from {expected}"]
    name, after = COVER_DIRECT_ADDITION
    direct = list(COVER_GROUPS[0])
    direct.insert(direct.index(after) + 1, name)
    if [canonical(x) for x in groups[0]] != [canonical(x) for x in direct] or groups[1] != list(COVER_GROUPS[1]):
        problems.append(f"{SHEET_COVER}: the groups are not المباشر {direct} and الوكلاء {list(COVER_GROUPS[1])}")
    if any(_is_excluded(label) for group in groups for label in group):
        problems.append(f"{SHEET_COVER}: {EXCLUDED_REPRESENTATIVE} is shown although she is excluded")
    report = {r[0]: r for r in rep_rows[:-1]}
    shown = {label for group in groups for label in group}
    if set(report) - shown:
        problems.append(f"{SHEET_COVER}: representatives of {SHEET_REPS} missing from Cover: {sorted(set(report) - shown)}")
    recount = independent_customer_counts(result)[SHEET_REPS]
    zero = [0] * len(result.count_columns)
    width = len(result.columns)

    def check(row: list[Any], mine: list[MasterRow], inside_outside: list[int], target: int | None,
              inside_for_ratio: int) -> None:
        if row[1:4] != _master_amounts(mine):
            problems.append(f"{SHEET_COVER}: '{row[0]}' amounts {row[1:4]} differ from its {len(mine)} "
                            f"{SHEET_MASTER} rows {_master_amounts(mine)}")
        if row[5:7] != inside_outside:
            problems.append(f"{SHEET_COVER}: '{row[0]}' customers inside / outside the 224 list {row[5:7]} differ from "
                            f"the independent recount {inside_outside}")
        if row[4] != (target or NO_TARGET) or row[7] != _ratio(inside_for_ratio, target):
            problems.append(f"{SHEET_COVER}: '{row[0]}' target / percentage {row[4]} / {row[7]} differ from "
                            f"{target or NO_TARGET} / {_ratio(inside_for_ratio, target)}")

    position, grand_target, grand_inside = 0, 0, 0
    for group, subtotal in zip(groups, SUBTOTAL_LABELS):
        sub_target = sub_inside = 0
        for label in group:
            row = rows[position]
            position += 1
            rep = report.get(label)
            if rep is not None:
                amounts = rep[1:1 + width]
                money = [amounts[2], amounts[5], amounts[6]] if result.mode == MODE_YESTERDAY else amounts
                if row[1:4] != money:
                    problems.append(f"{SHEET_COVER}: '{row[0]}' differs from its {SHEET_REPS} row")
            counts = recount.get(label, zero)
            target = cover_target(label)
            check(row, [r for r in result.master if r.rep_row == label], [counts[-3], counts[-2]], target, counts[-3])
            if target:
                sub_target, sub_inside = sub_target + target, sub_inside + counts[-3]
        row = rows[position]
        position += 1
        counts = independent_group_counts(result, set(group))
        check(row, [r for r in result.master if r.rep_row in set(group)], [counts[-3], counts[-2]],
              sub_target or None, sub_inside)
        grand_target, grand_inside = grand_target + sub_target, grand_inside + sub_inside
    total = rows[-1]
    everyone = independent_group_counts(result, shown)
    check(total, result.master, [everyone[-3], everyone[-2]], grand_target or None, grand_inside)
    if total[1:4] != [result.direct.sales, result.direct.returns, result.direct.net]:
        problems.append(f"{SHEET_COVER}: grand total amounts {total[1:4]} differ from the report totals")
    if total[5:7] != recount[GRAND_TOTAL][-3:-1]:
        problems.append(f"{SHEET_COVER}: grand total customers {total[5:7]} differ from the distinct customers of the "
                        f"whole report {recount[GRAND_TOTAL][-3:-1]}")
    subtotals = [sum((r[col] for r in rows if r[0] in SUBTOTAL_LABELS), ZERO) for col in (1, 2, 3)]
    if total[1:4] != subtotals:
        problems.append(f"{SHEET_COVER}: grand total amounts {total[1:4]} differ from the group subtotals {subtotals}")
    for row in rows:
        if row[3] != row[1] + row[2]:
            problems.append(f"{SHEET_COVER}: '{row[0]}' صافي {row[3]} is not sales + returns {row[1] + row[2]}")
    return problems


def reconcile_pivot(result: AnalysisResult) -> list[str]:
    """Pivot for Last Day vs a separate plain loop over the transactions of the selected day: every customer row
    (CSV name, distinct sales documents, sales, returns, الصافي), every representative subtotal (placed after its
    customers) and الإجمالي العام; the day's totals also equal مبيعات / مرتجعات يوم D of the report (``direct``)."""
    problems: list[str] = []
    rows = pivot_table(result)
    *body, total = rows
    subtotals = pivot_subtotal_rows(rows)
    day = result.reference_date
    expected: dict[tuple[str, str], list[Any]] = {}  # (representative row, code) -> [rep, CSV name, docs, sales, returns]
    for r in result.master:
        if day is None or r.day != day:
            continue
        want = expected.setdefault((r.rep_row, r.code), [r.representative, "", set(), ZERO, ZERO])
        want[1] = want[1] or r.name
        if r.category is Category.SALES:
            want[2].add(r.doc_no)
            want[3] += r.net
        else:
            want[4] += r.net
    want_rows = Counter((rep, name or code, len(docs), sales, returns, sales + returns)
                        for (_, code), (rep, name, docs, sales, returns) in expected.items())
    customers = [row for i, row in enumerate(body) if i not in subtotals]
    if Counter(tuple(row) for row in customers) != want_rows:
        problems.append(f"{SHEET_PIVOT}: customer rows differ from the day's transactions: written "
                        f"{sorted(map(tuple, customers), key=str)[:5]}, expected {sorted(want_rows, key=str)[:5]}")
    per_rep: dict[str, list[Any]] = {}
    for rep, _, docs, sales, returns in expected.values():
        sub = per_rep.setdefault(rep, [0, ZERO, ZERO])
        sub[0] += len(docs)
        sub[1] += sales
        sub[2] += returns
    seen: list[str] = []
    current: set[str] = set()
    for i, row in enumerate(body):
        if i not in subtotals:
            current.add(row[0])
            continue
        rep = row[0][len(PIVOT_SUBTOTAL_PREFIX):]
        want = per_rep.get(rep, [0, ZERO, ZERO])
        if current != {rep} or row[2:] != [want[0], want[1], want[2], want[1] + want[2]]:
            problems.append(f"{SHEET_PIVOT}: subtotal '{row[0]}' {row[2:]} does not match its customers "
                            f"{sorted(current)} {[want[0], want[1], want[2], want[1] + want[2]]}")
        seen.append(rep)
        current = set()
    if current or sorted(seen) != sorted(per_rep):
        problems.append(f"{SHEET_PIVOT}: representative subtotals {seen} differ from {sorted(per_rep)}")
    invoices = len({(code, doc) for (_, code), (_, _, docs, _, _) in expected.items() for doc in docs})
    sales = sum((v[3] for v in expected.values()), ZERO)
    returns = sum((v[4] for v in expected.values()), ZERO)
    if total != [GRAND_TOTAL, "", invoices, sales, returns, sales + returns]:
        problems.append(f"{SHEET_PIVOT}: الإجمالي العام {total[2:]} differs from the day's transactions "
                        f"{[invoices, sales, returns, sales + returns]}")
    day_sales, day_returns = (result.direct.sales_y, result.direct.returns_y) if day is not None else (ZERO, ZERO)
    if (total[3], total[4]) != (day_sales, day_returns):
        problems.append(f"{SHEET_PIVOT}: الإجمالي العام sales / returns {total[3]} / {total[4]} differ from the report's "
                        f"day totals {day_sales} / {day_returns}")
    return problems


# ------------------------------------------------------------------ workbook
def _display_width(value: Any) -> int:
    if isinstance(value, Decimal):
        return len(f"{value:,.2f}")
    return len(str(value)) if value is not None else 0


def _cell_lines(value: Any, width: float, size: int) -> int:
    """Printed lines a wrapped cell needs at this column width (bold Segoe UI, ``size`` points)."""
    if value is None or value == "":
        return 1
    chars_per_line = max((width - 2) * 11 / (size * 1.15), 1)
    return max(1, -(-_display_width(value) // int(chars_per_line)))


def report_date_title(result: AnalysisResult) -> str:
    """The line printed above every table: the report's date (the reference date, or the run date)."""
    return f"{REPORT_DATE_LABEL}: {format_dmy(result.report_date)}"


def _layout(ws: Worksheet, headers: Sequence[str], rows: list[list[Any]], *, size: int, cap: int,
            body_height: int, one_page: bool, title: str) -> None:
    """Row 1: ``title`` across the table, aligned left (top left of the printed page); row 2: the header;
    then the rows. Column widths that fit the content, row heights that show wrapped text completely,
    and the print setup: A4 landscape, centred, one page wide, rows 1-2 repeated on every page."""
    scale = size * 1.15 / 11  # bold text is wider than the column-width unit
    widths = []
    for c, header in enumerate(headers, start=1):
        longest = max((_display_width(row[c - 1]) for row in rows), default=0)
        longest_word = max((len(word) for word in str(header).split()), default=0)
        width = min(max(round(longest * scale) + 4, round(longest_word * scale) + 4, 14), cap)
        ws.column_dimensions[get_column_letter(c)].width = width
        widths.append(width)
    last_col, last_row = get_column_letter(len(headers)), len(rows) + 2
    ws.merge_cells(f"A1:{last_col}1")
    cell = ws.cell(row=1, column=1, value=title)
    cell.font = Font(name=_FONT, bold=True, size=size + 1, color="1F4E78")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28
    line = size * 1.5
    header_lines = max(_cell_lines(t, w, size + 1) for t, w in zip(headers, widths))
    ws.row_dimensions[2].height = max(36, round(header_lines * (line + 2) + 10))
    for r, values in enumerate(rows, start=3):
        lines = max((_cell_lines(v, w, size) for v, w in zip(values, widths)), default=1)
        ws.row_dimensions[r].height = max(body_height, round(lines * line + 8))
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if one_page else 0  # 0: as many pages down as needed, never tiny text
    ws.print_options.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6, header=0.3, footer=0.3)
    ws.print_title_rows = "1:2"
    ws.print_area = f"A1:{last_col}{last_row}"
    ws.oddFooter.center.text = "صفحة &P من &N"


def _serials(count: int, blanks: Collection[int] = ()) -> list[int | None]:
    """Column م for ``count`` rows: 1, 2, 3, ... on the data rows; ``None`` (blank) on the rows in ``blanks``."""
    serials: list[int | None] = []
    number = 0
    for index in range(count):
        if index in blanks:
            serials.append(None)
        else:
            number += 1
            serials.append(number)
    return serials


def _write_sheet(ws: Worksheet, headers: Sequence[str], rows: list[list[Any]], *, title: str, cover: bool = False,
                 subtotals: Collection[int] = ()) -> None:
    """A table with الإجمالي العام last, after the column م, under the ``title`` line; ``subtotals``: indexes
    of group subtotal rows (no serial number, own style); ``cover``: larger text and taller rows."""
    size = 13 if cover else 11
    ws.sheet_view.rightToLeft = True
    headers = (SERIAL_COLUMN, *headers)
    rows = [[serial, *row] for serial, row in zip(_serials(len(rows), {*subtotals, len(rows) - 1}), rows)]
    *detail, total = rows
    header_font = Font(name=_FONT, bold=True, color="FFFFFF", size=size + 1)
    body_font = Font(name=_FONT, bold=True, size=size)
    subtotal_font = Font(name=_FONT, bold=True, size=size, color="1F4E78")
    total_font = Font(name=_FONT, bold=True, size=size + 1, color="1F4E78")
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font, cell.fill, cell.border, cell.alignment = header_font, _NAVY, _BORDER, _CENTER
    for r, values in enumerate(detail, start=3):
        subtotal = r - 3 in subtotals
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.alignment = _CENTER
            if subtotal:
                cell.font, cell.fill, cell.border = subtotal_font, _SUBTOTAL_FILL, _SUBTOTAL_BORDER
            else:
                cell.font, cell.border = body_font, _BORDER
                if r % 2 == 0:
                    cell.fill = _STRIPE
            if isinstance(value, Decimal):
                cell.number_format = MONEY_FORMAT
            elif isinstance(value, int):
                cell.number_format = COUNT_FORMAT
    last = len(detail) + 3
    for c, value in enumerate(total, start=1):
        cell = ws.cell(row=last, column=c, value=value)
        cell.font, cell.fill, cell.border, cell.alignment = total_font, _TOTAL_FILL, _TOTAL_BORDER, _CENTER
        if isinstance(value, Decimal):
            cell.number_format = MONEY_FORMAT
        elif isinstance(value, int):
            cell.number_format = COUNT_FORMAT
    _layout(ws, headers, rows, size=size, cap=60, body_height=34 if cover else 24,
            one_page=cover or len(rows) <= _ONE_PAGE_ROWS, title=title)
    for index in subtotals:
        ws.row_dimensions[index + 3].height = max(ws.row_dimensions[index + 3].height or 0, 38 if cover else 28)
    if cover:
        ws.row_dimensions[last].height = max(ws.row_dimensions[last].height or 0, 44)
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{len(detail) + 2}"


def detail_total(rows: Sequence[MasterRow]) -> list[Any]:
    """الإجمالي العام of Master Data / a representative sheet, in MASTER_COLUMNS order: the distinct
    representatives (real names: not غير مسجل / غير محدد) under المندوب, the distinct customer codes
    (blank ones excluded) under كود العميل, and the sums of المبيعات, الصافي and المرتجعات."""
    representatives = {r.rep_row for r in rows if r.rep_row not in ("", UNREGISTERED)}
    codes = {r.code for r in rows if r.code}
    sales = sum((r.sales for r in rows), ZERO)
    net = sum((r.net for r in rows), ZERO)
    returns = sum((r.returns for r in rows), ZERO)
    return [GRAND_TOTAL, len(representatives), "", len(codes), "", sales, net, returns, ""]


def _write_master(ws: Worksheet, rows: Sequence[MasterRow], *, title: str) -> None:
    """Master Data / a representative sheet: the column م, one row per transaction, then الإجمالي العام
    (no serial number, outside the filter)."""
    ws.sheet_view.rightToLeft = True
    headers = (SERIAL_COLUMN, *MASTER_COLUMNS)
    values = [[serial, *r.values()] for serial, r in zip(_serials(len(rows)), rows)]
    total = [None, *detail_total(rows)]
    header_font = Font(name=_FONT, bold=True, color="FFFFFF", size=12)
    body_font = Font(name=_FONT, bold=True, size=11)
    total_font = Font(name=_FONT, bold=True, size=12, color="1F4E78")
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font, cell.fill, cell.border, cell.alignment = header_font, _NAVY, _BORDER, _CENTER
    for r, row in enumerate(values, start=3):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.font, cell.border, cell.alignment = body_font, _BORDER, _CENTER
            if isinstance(value, Decimal):
                cell.number_format = MONEY_FORMAT
            elif isinstance(value, date):
                cell.number_format = DATE_FORMAT
    last = len(values) + 3
    for c, value in enumerate(total, start=1):
        cell = ws.cell(row=last, column=c, value=value if value != "" else None)
        cell.font, cell.fill, cell.border, cell.alignment = total_font, _TOTAL_FILL, _TOTAL_BORDER, _CENTER
        if isinstance(value, Decimal):
            cell.number_format = MONEY_FORMAT
        elif isinstance(value, int):
            cell.number_format = COUNT_FORMAT
    _layout(ws, headers, [*values, total], size=11, cap=50, body_height=22, one_page=False, title=title)
    ws.row_dimensions[last].height = max(ws.row_dimensions[last].height or 0, 28)
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{len(values) + 2}"


_BLACK = Side(style="thin", color="000000")
_PIVOT_BORDER = Border(left=_BLACK, right=_BLACK, top=_BLACK, bottom=_BLACK)
_PIVOT_TOTAL_BORDER = Border(left=Side(style="medium", color="000000"), right=Side(style="medium", color="000000"),
                             top=Side(style="double", color="000000"), bottom=Side(style="medium", color="000000"))
_PIVOT_HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
_MEDIUM = Side(style="medium", color="000000")
_PIVOT_SUBTOTAL_BORDER = Border(left=_MEDIUM, right=_MEDIUM, top=_MEDIUM, bottom=_MEDIUM)


def _write_pivot(ws: Worksheet, rows: list[list[Any]], *, title: str) -> None:
    """Pivot for Last Day: a black-and-white table (light grey header and stripes, black text and borders),
    the title on row 1, the header on row 2, a bold subtotal with strong black borders after each
    representative, الإجمالي العام last with the strongest border; A4 landscape, printed in black and white,
    one page wide."""
    ws.sheet_view.rightToLeft = True
    *detail, total = rows
    subtotals = pivot_subtotal_rows(rows)
    header_font = Font(name=_FONT, bold=True, color="000000", size=12)
    body_font = Font(name=_FONT, bold=True, color="000000", size=11)
    total_font = Font(name=_FONT, bold=True, color="000000", size=12)
    for col, header in enumerate(PIVOT_COLUMNS, start=1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font, cell.fill, cell.border, cell.alignment = header_font, _PIVOT_HEADER_FILL, _PIVOT_BORDER, _CENTER
    stripe = False
    for index, values in enumerate(rows):
        r = index + 3
        is_total, is_subtotal = index == len(rows) - 1, index in subtotals
        if not (is_total or is_subtotal):
            stripe = not stripe
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=value if value != "" else None)
            cell.alignment = _CENTER
            if is_total:
                cell.font, cell.border = total_font, _PIVOT_TOTAL_BORDER
            elif is_subtotal:
                cell.font, cell.border, cell.fill = total_font, _PIVOT_SUBTOTAL_BORDER, _PIVOT_HEADER_FILL
            else:
                cell.font, cell.border = body_font, _PIVOT_BORDER
                if not stripe:
                    cell.fill = _STRIPE
            if isinstance(value, Decimal):
                cell.number_format = MONEY_FORMAT
            elif isinstance(value, int):
                cell.number_format = COUNT_FORMAT
    _layout(ws, PIVOT_COLUMNS, rows, size=11, cap=50, body_height=24, one_page=len(rows) <= _ONE_PAGE_ROWS,
            title=title)
    ws.cell(row=1, column=1).font = Font(name=_FONT, bold=True, color="000000", size=13)
    for index in (*subtotals, len(rows) - 1):
        ws.row_dimensions[index + 3].height = max(ws.row_dimensions[index + 3].height or 0, 28)
    ws.page_setup.blackAndWhite = True
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(PIVOT_COLUMNS))}{len(detail) + 2}"


def _write_cover(ws: Worksheet, rows: list[list[Any]], *, title: str, subtotals: Collection[int]) -> None:
    """Cover: a black-and-white table after the column م (light grey header and stripes, black text and borders),
    the title on row 1, the header on row 2, bold group subtotals with strong black borders, الإجمالي العام with
    the strongest border; percentages as 0.00%; A4 landscape printed in black and white, one page wide."""
    ws.sheet_view.rightToLeft = True
    headers = (SERIAL_COLUMN, *COVER_COLUMNS)
    rows = [[serial, *row] for serial, row in zip(_serials(len(rows), {*subtotals, len(rows) - 1}), rows)]
    header_font = Font(name=_FONT, bold=True, color="000000", size=13)
    body_font = Font(name=_FONT, bold=True, color="000000", size=12)
    subtotal_font = Font(name=_FONT, bold=True, color="000000", size=13)
    total_font = Font(name=_FONT, bold=True, color="000000", size=14)
    ratio_col = headers.index("نسبة تحقيق هدف العملاء") + 1
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font, cell.fill, cell.border, cell.alignment = header_font, _PIVOT_HEADER_FILL, _PIVOT_BORDER, _CENTER
    stripe = False
    for index, values in enumerate(rows):
        r = index + 3
        is_total, is_subtotal = index == len(rows) - 1, index in subtotals
        if not (is_total or is_subtotal):
            stripe = not stripe
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=value if value != "" else None)
            cell.alignment = _CENTER
            if is_total:
                cell.font, cell.border = total_font, _PIVOT_TOTAL_BORDER
            elif is_subtotal:
                cell.font, cell.border, cell.fill = subtotal_font, _PIVOT_SUBTOTAL_BORDER, _PIVOT_HEADER_FILL
            else:
                cell.font, cell.border = body_font, _PIVOT_BORDER
                if not stripe:
                    cell.fill = _STRIPE
            if c == ratio_col and isinstance(value, Decimal):
                cell.number_format = PERCENT_FORMAT
            elif isinstance(value, Decimal):
                cell.number_format = MONEY_FORMAT
            elif isinstance(value, int):
                cell.number_format = COUNT_FORMAT
    _layout(ws, headers, rows, size=12, cap=45, body_height=30, one_page=len(rows) <= _ONE_PAGE_ROWS, title=title)
    ws.cell(row=1, column=1).font = Font(name=_FONT, bold=True, color="000000", size=14)
    for index in subtotals:
        ws.row_dimensions[index + 3].height = max(ws.row_dimensions[index + 3].height or 0, 34)
    ws.row_dimensions[len(rows) + 2].height = max(ws.row_dimensions[len(rows) + 2].height or 0, 38)
    ws.page_setup.blackAndWhite = True
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{len(rows) + 1}"


def write_analysis_report(result: AnalysisResult, path: Path) -> Path:
    _ensure_new(path)
    wb = Workbook()
    wb.remove(wb.active)
    title = report_date_title(result)
    for name, (headers, rows) in sheet_tables(result).items():
        subtotals = rep_subtotal_rows(rows) if name == SHEET_REPS else ()
        _write_sheet(wb.create_sheet(name), headers, rows, title=title, subtotals=subtotals)
    wb.save(path)
    return path


def reps_workbook_name(result: AnalysisResult) -> str:
    return f"{REPS_WORKBOOK_PREFIX}_{result.started_at:%Y-%m-%d}.xlsx"


def write_reps_report(result: AnalysisResult, path: Path) -> Path:
    """The second workbook: Cover, Master Data, then one sheet per representative (its own rows, in order)."""
    _ensure_new(path)
    wb = Workbook()
    wb.remove(wb.active)
    title = report_date_title(result)
    cover = cover_table(result)
    _write_cover(wb.create_sheet(SHEET_COVER), cover, title=title, subtotals=cover_subtotal_rows(cover))
    _write_master(wb.create_sheet(SHEET_MASTER), master_rows(result), title=title)
    for label, sheet in representative_sheet_names(result).items():
        _write_master(wb.create_sheet(sheet), [r for r in result.master if r.rep_row == label], title=title)
    _write_pivot(wb.create_sheet(SHEET_PIVOT), pivot_table(result), title=pivot_title(result))
    wb.save(path)
    return path


def _check_serial(where: str, body: list[list[Any]], blanks: Collection[int] = ()) -> list[str]:
    """Saved column م: 1, 2, 3, ... on the data rows, blank on the total rows (body indexes in ``blanks``)."""
    wrong = [index + 3 for index, (row, want) in enumerate(zip(body, _serials(len(body), blanks)))
             if (row[0] if row else None) not in ((None, "") if want is None else (want,))]
    if wrong:
        return [f"{where}: column {SERIAL_COLUMN} is not 1, 2, 3, ... at row(s) {wrong[:5]}"]
    return []


def _check_title(where: str, saved: list[list[Any]], title: str) -> list[str]:
    """Row 1 of a saved sheet must be the report date line."""
    got = saved[0][0] if saved and saved[0] else None
    return [] if got == title else [f"{where}: row 1 is {got!r}, not the report date line {title!r}"]


def _excluded_cells(where: str, saved: dict[str, list[list[Any]]]) -> list[str]:
    """Any sheet name or cell that names the excluded representative."""
    hits = [name for name in saved if _is_excluded(name)]
    hits += [f"{name}!{get_column_letter(c + 1)}{r + 1}" for name, values in saved.items()
             for r, row in enumerate(values) for c, value in enumerate(row)
             if isinstance(value, str) and _is_excluded(value)]
    if hits:
        return [f"{where}: {EXCLUDED_REPRESENTATIVE} appears although she is excluded: {hits[:5]}"]
    return []


def _money(value: Any) -> Decimal:
    return Decimal(str(value or 0))


def _verify_detail(where: str, saved: list[list[Any]], rows: Sequence[MasterRow], title: str) -> list[str]:
    """A saved Master Data-style sheet vs the transactions it must hold: the date line, the header, row count,
    column م, every cell, and الإجمالي العام - against its table and recomputed from the saved rows."""
    problems = _check_title(where, saved, title)
    table = saved[1:]
    if not table or tuple(table[0]) != (SERIAL_COLUMN, *MASTER_COLUMNS):
        return problems + [f"{where}: header mismatch {table[0] if table else None}"]
    if len(table) - 2 != len(rows):
        return problems + [f"{where}: {len(table) - 2} rows written, {len(rows)} expected (+ الإجمالي العام)"]
    *body, total = table[1:]
    problems += _check_serial(where, table[1:], {len(rows)})
    wrong: list[int] = []
    for line, (got, want) in enumerate(zip(body, (r.values() for r in rows)), start=3):
        day, *texts, sales, net, returns, status = got[1:1 + len(MASTER_COLUMNS)]
        ok = (isinstance(day, datetime) and day.date() == want[0]
              and [(v or "") for v in texts] == want[1:5] and (status or "") == want[8]
              and all(isinstance(v, (int, float)) and abs(Decimal(str(v)) - w) <= _TOLERANCE
                      for v, w in ((sales, want[5]), (net, want[6]), (returns, want[7]))))
        if not ok:
            wrong.append(line)
    if wrong:
        problems.append(f"{where}: {len(wrong)} row(s) differ from the transactions, first at row(s) {wrong[:5]}")
    want = detail_total(rows)
    got = total[1:1 + len(MASTER_COLUMNS)]
    tolerance = _TOLERANCE * max(len(body), 1)
    same = (got[0] == GRAND_TOTAL and got[1] == want[1] and got[3] == want[3]
            and all(abs(_money(g) - w) <= _TOLERANCE for g, w in zip(got[5:8], want[5:8])))
    if not same:
        problems.append(f"{where}: الإجمالي العام {got} differs from {want}")
    names = {r[2] for r in body if r[2] not in (None, "", MASTER_UNREGISTERED, MASTER_UNSPECIFIED)}
    codes = {r[4] for r in body if r[4] not in (None, "")}
    sums = [sum((_money(r[col]) for r in body), ZERO) for col in (6, 7, 8)]
    if (got[1] != len(names) or got[3] != len(codes)
            or any(abs(a - _money(b)) > tolerance for a, b in zip(sums, got[5:8]))
            or abs(_money(got[5]) + _money(got[7]) - _money(got[6])) > 2 * _TOLERANCE):
        problems.append(f"{where}: saved الإجمالي العام does not match its saved rows (distinct representatives "
                        f"{len(names)}, customers {len(codes)}, sales / net / returns {sums}; net = sales + returns)")
    return problems


def _verify_cover(where: str, saved: list[list[Any]], rows: list[list[Any]], title: str) -> list[str]:
    """Saved Cover vs its table: the date line, header, rows, column م, every cell; and on the saved values:
    each group subtotal and الإجمالي العام = the sums of their rows, صافي = sales + returns, and every
    percentage = inside customers / target."""
    problems = _check_title(where, saved, title)
    table = saved[1:]
    if not table or tuple(table[0]) != (SERIAL_COLUMN, *COVER_COLUMNS):
        return problems + [f"{where}: header mismatch {table[0] if table else None}"]
    if len(table) - 1 != len(rows):
        return problems + [f"{where}: {len(table) - 1} rows written, {len(rows)} expected"]
    problems += _check_serial(where, table[1:], {*cover_subtotal_rows(rows), len(rows) - 1})
    body = [row[1:] for row in table[1:]]
    for line, (got, want) in enumerate(zip(body, rows), start=3):
        for g, w, header in zip(got, want, COVER_COLUMNS):
            if isinstance(w, Decimal):
                bad = not isinstance(g, (int, float)) or abs(Decimal(str(g)) - w) > _TOLERANCE
            elif isinstance(w, int):
                bad = g != w
            else:
                bad = (g or "") != w
            if bad:
                problems.append(f"{where}: row {line} '{want[0]}' {header} {g!r} differs from {w}")
    problems += _saved_group_sums(where, body, SUBTOTAL_LABELS, (1, 2, 3))
    for line, row in enumerate(body, start=3):
        sales, returns, net = (_money(row[col]) for col in (1, 2, 3))
        if abs(sales + returns - net) > 2 * _TOLERANCE:
            problems.append(f"{where}: row {line} صافي is not sales + returns")
        if isinstance(row[4], (int, float)) and row[4] and row[0] not in (*SUBTOTAL_LABELS, GRAND_TOTAL):
            if not isinstance(row[7], (int, float)) or abs(Decimal(str(row[7])) - Decimal(row[5]) / Decimal(row[4])) > Decimal("0.0001"):
                problems.append(f"{where}: row {line} نسبة تحقيق هدف العملاء is not inside / target")
    if body[-1][0] != GRAND_TOTAL:
        problems.append(f"{where}: last row is not '{GRAND_TOTAL}'")
    return problems


def _saved_group_sums(where: str, body: list[list[Any]], subtotal_labels: Sequence[str],
                      columns: Sequence[int]) -> list[str]:
    """On saved values: every subtotal row = the sum of the rows of its group (since the previous subtotal),
    and the last row (الإجمالي العام) = the sum of all non-subtotal rows, for the amount ``columns``."""
    problems: list[str] = []
    group, everyone = [ZERO] * len(columns), [ZERO] * len(columns)
    for line, row in enumerate(body[:-1], start=3):
        amounts = [_money(row[col]) for col in columns]
        if row[0] in subtotal_labels:
            if any(abs(a - b) > _TOLERANCE * 20 for a, b in zip(group, amounts)):
                problems.append(f"{where}: row {line} saved '{row[0]}' amounts differ from the sum of its group's rows")
            group = [ZERO] * len(columns)
        else:
            group = [a + b for a, b in zip(group, amounts)]
            everyone = [a + b for a, b in zip(everyone, amounts)]
    total = [_money(body[-1][col]) for col in columns]
    if any(abs(a - b) > _TOLERANCE * max(len(body), 1) for a, b in zip(everyone, total)):
        problems.append(f"{where}: saved grand total amounts differ from the sum of the rows")
    return problems


def _verify_pivot(where: str, saved: list[list[Any]], rows: list[list[Any]], title: str) -> list[str]:
    """Saved Pivot for Last Day: title, header, every cell, and on the saved values: each representative subtotal =
    the sum of its customer rows, الإجمالي العام = the sum of the subtotals, الصافي = sales + returns in every row."""
    problems = _check_title(where, saved, title)
    table = saved[1:]
    if not table or tuple(table[0][:len(PIVOT_COLUMNS)]) != PIVOT_COLUMNS:
        return problems + [f"{where}: header mismatch {table[0] if table else None}"]
    body = [row[:len(PIVOT_COLUMNS)] for row in table[1:]]
    if len(body) != len(rows):
        return problems + [f"{where}: {len(body)} rows written, {len(rows)} expected"]
    for line, (got, want) in enumerate(zip(body, rows), start=3):
        for g, w, header in zip(got, want, PIVOT_COLUMNS):
            if isinstance(w, Decimal):
                bad = not isinstance(g, (int, float)) or abs(Decimal(str(g)) - w) > _TOLERANCE
            elif isinstance(w, int):
                bad = g != w
            else:
                bad = (g or "") != w
            if bad:
                problems.append(f"{where}: row {line} {header} {g!r} differs from {w!r}")
    subtotals = pivot_subtotal_rows(rows)
    group, everyone = [ZERO] * 4, [ZERO] * 4
    for index, row in enumerate(body[:-1]):
        amounts = [_money(row[col]) for col in (2, 3, 4, 5)]
        if index in subtotals:
            if any(abs(a - b) > _TOLERANCE * 20 for a, b in zip(group, amounts)):
                problems.append(f"{where}: row {index + 3} saved '{row[0]}' differs from the sum of its customers {group}")
            everyone = [a + b for a, b in zip(everyone, amounts)]
            group = [ZERO] * 4
        else:
            group = [a + b for a, b in zip(group, amounts)]
    total = body[-1]
    if any(abs(a - _money(b)) > _TOLERANCE * max(len(body), 1) for a, b in zip(everyone, total[2:6])):
        problems.append(f"{where}: saved الإجمالي العام {total[2:6]} differs from the sum of the subtotals {everyone}")
    for line, row in enumerate(body, start=3):
        if abs(_money(row[3]) + _money(row[4]) - _money(row[5])) > 2 * _TOLERANCE:
            problems.append(f"{where}: row {line} الصافي is not sales + returns")
    if total[0] != GRAND_TOTAL:
        problems.append(f"{where}: last row is not '{GRAND_TOTAL}'")
    return problems


def verify_reps_report(result: AnalysisResult) -> list[str]:
    """Re-open Representatives_Analysis: exact sheets, Cover, Master Data and every representative sheet
    cell by cell, and every saved Master Data transaction exactly once across the representative sheets."""
    path = result.reps_report_path
    if path is None or not path.is_file() or path.stat().st_size == 0:
        return [f"missing or empty representatives workbook: {path}"]
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        saved = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
        wb.close()
    except Exception as exc:  # noqa: BLE001 - any failure means "not verified"
        return [f"cannot re-open {path.name}: {type(exc).__name__}: {exc}"]
    names = representative_sheet_names(result)
    expected = [SHEET_COVER, SHEET_MASTER, *names.values(), SHEET_PIVOT]
    if list(saved) != expected:
        return [f"{path.name}: sheets {list(saved)} differ from {expected}"]
    title = report_date_title(result)
    problems = _verify_cover(f"{path.name} / {SHEET_COVER}", saved[SHEET_COVER], cover_table(result), title)
    # expected rows come from the transactions themselves, not from the functions that wrote the sheets
    problems += _verify_detail(f"{path.name} / {SHEET_MASTER}", saved[SHEET_MASTER], result.master, title)
    for label, sheet in names.items():
        problems += _verify_detail(f"{path.name} / {sheet}", saved[sheet],
                                   [r for r in result.master if r.rep_row == label], title)
    problems += _verify_pivot(f"{path.name} / {SHEET_PIVOT}", saved[SHEET_PIVOT], pivot_table(result), pivot_title(result))
    detail = {name: values[2:-1] for name, values in saved.items()
              if name not in (SHEET_COVER, SHEET_PIVOT)}                                 # no date, header, total
    master = detail[SHEET_MASTER]
    saved_sales, saved_net, saved_returns = (sum((_money(r[col]) for r in master), ZERO) for col in (6, 7, 8))
    tolerance = _TOLERANCE * max(len(master), 1)
    if (abs(saved_sales - result.direct.sales) > tolerance or abs(saved_net - result.direct.net) > tolerance
            or abs(saved_returns - result.direct.returns) > tolerance):
        problems.append(f"{path.name} / {SHEET_MASTER}: saved المبيعات {saved_sales} / الصافي {saved_net} / المرتجعات "
                        f"{saved_returns} differ from the report sales {result.direct.sales} / net {result.direct.net} / "
                        f"returns {result.direct.returns}")
    in_sheets = Counter(tuple(r[1:]) for sheet in names.values() for r in detail[sheet])
    in_master = Counter(tuple(r[1:]) for r in master)
    if in_sheets != in_master:
        problems.append(f"{path.name}: the representative sheets miss {sum((in_master - in_sheets).values())} and "
                        f"duplicate {sum((in_sheets - in_master).values())} {SHEET_MASTER} transaction(s)")
    return problems + _excluded_cells(path.name, saved)


def verify_analysis_report(result: AnalysisResult) -> list[str]:
    """Re-open both saved workbooks (every cell), then the reconciliation; any problem: nothing moves."""
    path = result.report_path
    if path is None or not path.is_file() or path.stat().st_size == 0:
        return [f"missing or empty report: {path}"]
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        saved = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
        wb.close()
    except Exception as exc:  # noqa: BLE001 - any failure means "not verified"
        return [f"cannot re-open report: {type(exc).__name__}: {exc}"]
    if list(saved) != list(SHEETS):
        return [f"workbook sheets {list(saved)} differ from {list(SHEETS)}"]
    problems: list[str] = []
    title = report_date_title(result)
    for name, (headers, rows) in sheet_tables(result).items():
        problems += _check_title(name, saved[name], title)
        values = saved[name][1:]
        if not values or tuple(values[0]) != (SERIAL_COLUMN, *headers):
            problems.append(f"{name}: header mismatch {values[0] if values else None}")
            continue
        if len(values) - 1 != len(rows):
            problems.append(f"{name}: {len(values) - 1} rows written, {len(rows)} expected")
            continue
        subtotals = rep_subtotal_rows(rows) if name == SHEET_REPS else set()
        problems += _check_serial(name, values[1:], {*subtotals, len(rows) - 1})
        body = [row[1:] for row in values[1:]]
        for got, want in zip(body, rows):
            for g, w in zip(got, want):
                if isinstance(w, Decimal):
                    if not isinstance(g, (int, float)) or abs(Decimal(str(g)) - w) > _TOLERANCE:
                        problems.append(f"{name}: '{want[0]}' cell {g!r} differs from {w}")
                elif isinstance(w, int):
                    if g != w:
                        problems.append(f"{name}: '{want[0]}' customer count {g!r} differs from {w}")
                elif (g or "") != w:
                    problems.append(f"{name}: cell {g!r} differs from {w!r}")
        first = 2 if name == SHEET_MISSING else 1
        end = len(headers) if name == SHEET_MISSING else 1 + len(result.columns)  # counts are not sums
        problems += _saved_group_sums(name, body, REP_SUBTOTAL_LABELS if name == SHEET_REPS else (),
                                      range(first, end))
        if body[-1][0] != GRAND_TOTAL:
            problems.append(f"{name}: last row is not '{GRAND_TOTAL}'")
    return problems + _excluded_cells(path.name, saved) + verify_reps_report(result) + reconcile(result)


# ------------------------------------------------------------------ run
def run_analysis(config: PipelineConfig, mode: str, reference_date: date | None, main_data: MainData, *,
                 now: datetime | None = None, files: list[Path] | None = None,
                 run_dir: Path | None = None, fixed_codes: frozenset[str] = frozenset()) -> AnalysisResult:
    """Build both workbooks from every CSV in ``config.input_dir`` (nothing is moved).

    ``files``/``run_dir`` let the managed lifecycle pre-select the batch and own the folder.
    No report folder is created when there are no files.
    """
    if mode not in (MODE_YESTERDAY, MODE_ALL):
        raise ValueError(f"unknown report type {mode!r}")
    if (mode == MODE_YESTERDAY) != (reference_date is not None):
        raise ValueError("the yesterday report needs a date; the all-transactions report takes none")
    config = config.validated()
    result = AnalysisResult(mode=mode, reference_date=reference_date, main_data=main_data,
                            input_dir=config.input_dir, started_at=now or datetime.now(), fixed_codes=fixed_codes)
    if files is None:
        files, _ = discover_csv_files(config.input_dir, recursive=config.recursive,
                                      exclude_dirs=[config.output_dir])
    if not files:
        return result
    result.files = [_add_file(path, config, result) for path in files]
    if run_dir is None:
        run_dir = config.output_dir / f"{TASK_PREFIX}_{_free_stamp(config.output_dir, None, TASK_PREFIX, result.started_at)}"
        run_dir.mkdir(parents=True)
    result.run_dir = run_dir
    result.report_path = write_analysis_report(result, run_dir / WORKBOOK_NAME)
    result.reps_report_path = write_reps_report(result, run_dir / reps_workbook_name(result))
    return result


def run_analysis_managed(config: PipelineConfig, mode: str, reference_date: date | None, main_data: MainData,
                         files_config: FileManagementConfig | None = None, *,
                         now: datetime | None = None, fixed_codes: frozenset[str] = frozenset()) -> GroupRunResult:
    """Build, save, re-open and reconcile the workbook; only then move the cleanly read CSV files.

    Same lifecycle as tasks 3 and 4 (``grouped.run_batch_managed``): ``output/Task5_<stamp>``,
    ``processed/Task5_<stamp>``, journal ``state/task5_journal.jsonl``.
    """
    return run_batch_managed(
        TASK_PREFIX, JOURNAL_NAME, config, files_config,
        build=lambda files, run_dir, started: run_analysis(config, mode, reference_date, main_data, now=started,
                                                          files=files, run_dir=run_dir, fixed_codes=fixed_codes),
        verify=lambda result: verify_analysis_report(result), now=now)
