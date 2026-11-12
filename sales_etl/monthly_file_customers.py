"""Task 7: عدد العملاء الشهري لكل ملف - distinct customer codes per CSV file per month.

Files are found, decoded and parsed exactly as in tasks 3-6 (``ingest`` -> ``normalize``). Columns
(found by header):

* A ``التاريخ``  -> the accounting month of the transaction (no date filter; any date counts)
* F ``العميل``   -> customer code, kept exactly as written; the customer name is never used

Every document type counts (not only مبيعات / مرتجعات); amounts are never used. Each file is
counted on its own: a code is counted once per month inside a file, again in another month, and
again in another file. The report has اسم الملف then one column per month from يناير to the last
month found in any file read; months without customers show 0. The last column is the number of
distinct customer codes in the whole file (a code in several months counts once). A file with no transactions
(header only, or no content at all) is not a failure: it gets a row of zeros. No total row.

A transaction row without a customer code is rejected (not counted) and keeps its file in input,
like any other rejected row.

Accounting months (task 7 only): a month runs from the 26th of the previous month to the 25th of
the month itself, so 26-09-2026 .. 25-10-2026 is أكتوبر and 26-12-2025 .. 25-01-2026 is يناير. The file lifecycle is task 3's (``grouped.run_batch_managed``):
``output/عملاء_الملفات_الشهري_<stamp>``, ``processed/عملاء_الملفات_الشهري_<stamp>``, journal
``state/سجل_عملاء_الملفات_الشهري.jsonl``. There is no duplicate detection.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook

from sales_etl.analysis import COUNT_FORMAT
from sales_etl.config import COL_CUSTOMER, FileManagementConfig, PipelineConfig
from sales_etl.grouped import GroupFile, GroupRunResult, _free_stamp, run_batch_managed
from sales_etl.ingest import discover_csv_files, read_raw_table, sha256_file
from sales_etl.models import FileProcessingError
from sales_etl.monthly_customer import MONTHS
from sales_etl.normalize import parse_transactions
from sales_etl.report import _ensure_new, _table

logger = logging.getLogger(__name__)

TASK_PREFIX = "عملاء_الملفات_الشهري"
JOURNAL_NAME = "سجل_عملاء_الملفات_الشهري.jsonl"
REPORT_NAME = "تقرير عدد العملاء الشهري لكل ملف.xlsx"
REPORT_SHEET = "التقرير"
FILE_COLUMN = "اسم الملف"
YEAR_COLUMN = "إجمالي العملاء المميزين خلال السنة"

#: parse_transactions needs a reference date; this task never filters by date.
_NO_REFERENCE_DATE = date(2100, 12, 31)
#: first day of an accounting month, in the previous calendar month
_ACCOUNTING_START_DAY = 26


def accounting_month(day: date) -> int:
    """1..12: from the 26th on, a date belongs to the next month (26-12 -> يناير)."""
    return day.month % 12 + 1 if day.day >= _ACCOUNTING_START_DAY else day.month


def _period_month(day: date) -> int:
    """Independent check of :func:`accounting_month`: find the 26th..25th period holding ``day``."""
    for year in (day.year, day.year + 1):
        for month in range(1, 13):
            end = date(year, month, _ACCOUNTING_START_DAY - 1)
            start = (end.replace(day=1) - timedelta(days=1)).replace(day=_ACCOUNTING_START_DAY)
            if start <= day <= end:
                return month
    raise ValueError(f"no accounting period holds {day}")


@dataclass
class FileCustomers:
    """One report row: the distinct customer codes of one file, per month."""

    item: GroupFile
    months: dict[int, set[str]] = field(default_factory=dict)  # 1..12 -> codes as written
    codes: set[str] = field(default_factory=set)  # every code of the file, for the year column


@dataclass(frozen=True)
class CountedRow:
    """One counted transaction, kept for the independent reconciliation."""

    file_no: int  # position of the file in the report
    file_name: str
    txn_date: date
    code: str


@dataclass
class FileCustomersResult:
    input_dir: Path
    started_at: datetime
    paths: list[Path] = field(default_factory=list)  # input files, in report order
    files: list[GroupFile] = field(default_factory=list)
    rows: list[FileCustomers] = field(default_factory=list)  # files read, in report order
    ledger: list[CountedRow] = field(default_factory=list)
    rejected_lines: list[str] = field(default_factory=list)  # "file line N: reason"
    run_dir: Path | None = None
    report_path: Path | None = None

    @property
    def failed(self) -> list[GroupFile]:
        return [f for f in self.files if f.error is not None]

    @property
    def rejected_rows(self) -> int:
        return sum(f.rejected_rows for f in self.files)

    @property
    def complete(self) -> bool:
        """Every file was read and every transaction row was counted."""
        return bool(self.files) and not self.failed and not self.rejected_rows

    @property
    def last_month(self) -> int:
        return max((m for r in self.rows for m in r.months), default=0)

    @property
    def columns(self) -> tuple[str, ...]:
        return (FILE_COLUMN, *MONTHS[:self.last_month], YEAR_COLUMN)


def _file_key(path: Path) -> list[tuple[int, int | str]]:
    """File names with numbers in numeric order: 3omla29 before 3omla210."""
    return [(0, int(part)) if part.isdigit() else (1, part)
            for part in re.split(r"(\d+)", path.name.casefold()) if part]


def file_table(result: FileCustomersResult) -> list[list[str | int]]:
    """One row per file read: file name, distinct customer codes per month, then in the whole file."""
    months = range(1, result.last_month + 1)
    return [[r.item.file_name, *[len(r.months.get(m, ())) for m in months], len(r.codes)] for r in result.rows]


def _add_file(path: Path, config: PipelineConfig, result: FileCustomersResult) -> GroupFile:
    item = GroupFile(file_name=path.name)
    try:
        raw = read_raw_table(
            path, encodings=config.encodings, delimiters=config.delimiters,
            forced_encoding=config.forced_encoding, forced_delimiter=config.forced_delimiter,
        )
        if COL_CUSTOMER not in raw.column_index:
            raise FileProcessingError("MISSING_COLUMN", f"required column '{COL_CUSTOMER}' not found")
        i_code = raw.column_index[COL_CUSTOMER]
        parsed = parse_transactions(raw, _NO_REFERENCE_DATE)
    except FileProcessingError as exc:
        if exc.code == "EMPTY_FILE":  # no content at all: a file without transactions, not a failure
            item.sha256 = sha256_file(path)
            result.rows.append(FileCustomers(item))
            return item
        item.error = f"{exc.code}: {exc.message}"
        return item
    except Exception as exc:  # noqa: BLE001 - isolate unexpected failures per file
        item.error = f"UNEXPECTED_ERROR: {type(exc).__name__}: {exc}"
        logger.exception("unexpected failure", extra={"event": "file_customers_failed", "file": path.name})
        return item
    item.sha256 = raw.sha256
    item.transactions, item.rejected_rows = len(parsed.transactions), len(parsed.rejected)
    for rejected in parsed.rejected:
        result.rejected_lines.append(f"{path.name} line {rejected.line_no}: {'; '.join(rejected.reasons)}")
    row = FileCustomers(item)
    file_no = len(result.rows)
    result.rows.append(row)
    for t in parsed.transactions:  # every document type counts
        code = t.cells[i_code]
        if not code:
            item.rejected_rows += 1
            result.rejected_lines.append(f"{path.name} line {t.line_no}: {COL_CUSTOMER}: empty")
            continue
        row.months.setdefault(accounting_month(t.txn_date), set()).add(code)
        row.codes.add(code)
        result.ledger.append(CountedRow(file_no, path.name, t.txn_date, code))
    return item


def write_file_customers_report(result: FileCustomersResult, path: Path) -> Path:
    _ensure_new(path)
    wb = Workbook()
    ws = wb.active
    ws.title = REPORT_SHEET
    ws.sheet_view.rightToLeft = True
    _table(ws, result.columns, file_table(result), widths=(30, *[14] * result.last_month, 22))
    for row in ws.iter_rows(min_row=2, min_col=2):
        for cell in row:
            cell.number_format = COUNT_FORMAT
    ws.row_dimensions[1].height = 32
    wb.save(path)
    return path


def reconcile(result: FileCustomersResult) -> list[str]:
    """Independent cross-checks of the table against the ledger of counted rows."""
    problems: list[str] = []
    table = file_table(result)
    read = [p.name for p, f in zip(result.paths, result.files) if f.error is None]
    if len(table) != len(read):
        problems.append(f"{len(table)} file rows, {len(read)} files read")
    if [r[0] for r in table] != read:
        problems.append(f"file names in the report {[r[0] for r in table]} differ from the files read {read}")
    expected: dict[tuple[int, int], set[str]] = {}
    year: dict[int, set[str]] = {}
    for row in result.ledger:
        expected.setdefault((row.file_no, _period_month(row.txn_date)), set()).add(row.code)
        year.setdefault(row.file_no, set()).add(row.code)
        if row.file_no >= len(table) or table[row.file_no][0] != row.file_name:
            problems.append(f"counted row of {row.file_name} is not in its own report row")
    last = max((_period_month(row.txn_date) for row in result.ledger), default=0)
    if result.columns != (FILE_COLUMN, *MONTHS[:last], YEAR_COLUMN):
        problems.append(f"month columns {result.columns[1:-1]} differ from يناير..{MONTHS[last - 1] if last else '-'}")
    for file_no, r in enumerate(table):
        want = len(year.get(file_no, ()))
        if r[-1] != want:
            problems.append(f"{r[0]}: {YEAR_COLUMN} {r[-1]} differs from {want} distinct customer codes in the file")
        for month, value in enumerate(r[1:-1], start=1):
            want = len(expected.get((file_no, month), ()))
            if value != want:
                problems.append(f"{r[0]}: {MONTHS[month - 1]} {value} differs from {want} distinct customer codes")
    return problems


def verify_file_customers_report(result: FileCustomersResult) -> list[str]:
    """Re-open the saved workbook: sheet, headers, row count, every cell; then reconcile."""
    path = result.report_path
    if path is None or not path.is_file() or path.stat().st_size == 0:
        return [f"missing or empty report: {path}"]
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        sheets = wb.sheetnames
        values = [list(r) for r in wb[sheets[0]].iter_rows(values_only=True)]
        wb.close()
    except Exception as exc:  # noqa: BLE001 - any failure means "not verified"
        return [f"cannot re-open report: {type(exc).__name__}: {exc}"]
    if sheets != [REPORT_SHEET]:
        return [f"workbook sheets {sheets} differ from {[REPORT_SHEET]}"]
    if not values or tuple(values[0]) != result.columns:
        return [f"report header mismatch: {values[0] if values else None}"]
    expected = file_table(result)
    if len(values) - 1 != len(expected):
        return [f"{len(values) - 1} rows written, {len(expected)} expected"]
    problems: list[str] = []
    for got, want in zip(values[1:], expected):
        if got[0] != want[0]:
            problems.append(f"file name cell {got[0]!r} differs from {want[0]!r}")
        for g, w in zip(got[1:], want[1:]):
            if type(g) is not int or g != w:
                problems.append(f"{want[0]}: cell {g!r} differs from {w}")
    return problems + reconcile(result)


def run_file_customers(config: PipelineConfig, *, now: datetime | None = None, files: list[Path] | None = None,
                       run_dir: Path | None = None) -> FileCustomersResult:
    """Distinct customers per file per month over every CSV in ``config.input_dir`` (nothing is moved).

    ``files``/``run_dir`` let the managed lifecycle pre-select the batch and own the folder.
    No report folder is created when there are no files.
    """
    config = config.validated()
    result = FileCustomersResult(input_dir=config.input_dir, started_at=now or datetime.now())
    if files is None:
        files, _ = discover_csv_files(config.input_dir, recursive=config.recursive,
                                      exclude_dirs=[config.output_dir])
    if not files:
        return result
    result.paths = sorted(files, key=_file_key)
    result.files = [_add_file(path, config, result) for path in result.paths]
    if run_dir is None:
        run_dir = config.output_dir / f"{TASK_PREFIX}_{_free_stamp(config.output_dir, None, TASK_PREFIX, result.started_at)}"
        run_dir.mkdir(parents=True)
    result.run_dir = run_dir
    result.report_path = write_file_customers_report(result, run_dir / REPORT_NAME)
    return result


def run_file_customers_managed(config: PipelineConfig, files_config: FileManagementConfig | None = None, *,
                               now: datetime | None = None) -> GroupRunResult:
    """Build, save, re-open and reconcile the report; only then move the cleanly read CSV files."""
    return run_batch_managed(
        TASK_PREFIX, JOURNAL_NAME, config, files_config,
        build=lambda files, run_dir, started: run_file_customers(config, now=started, files=files,
                                                                 run_dir=run_dir),
        verify=lambda result: verify_file_customers_report(result), now=now)
