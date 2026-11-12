"""Task 6: صافي مبيعات العملاء الشهري - net per customer code for every month from يناير.

Files are found, decoded and parsed exactly as in task 3 (``ingest`` -> ``normalize``). Columns
(found by header, as in tasks 3 and 4):

* A ``التاريخ``      -> the month of the transaction (no date filter; any date counts)
* B ``نوع المستند``  -> only مبيعات and مرتجعات rows are counted
* F ``العميل``       -> customer code, kept exactly as written (one report row per code)
* G (blank header)  -> customer name: the column right after the code (first name found)
* J (blank header)  -> representative name: the column right after ``المندوب`` (I)
* T ``صافى``         -> amount, summed with its original sign: the monthly net

The report has كود العميل، اسم العميل، المندوب، then one column per month from يناير to the last
month in the data (months without transactions show 0). One year of data is expected: the month
is the calendar month of the date. No sales/returns columns, no total row or column.

A sales/returns row without a customer code is rejected (not counted) and keeps its file in
input, like any other rejected row. The file lifecycle is task 3's (``grouped.run_batch_managed``):
``output/صافي_العملاء_الشهري_<stamp>``, ``processed/صافي_العملاء_الشهري_<stamp>``, journal
``state/سجل_صافي_العملاء_الشهري.jsonl``. There is no duplicate detection.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook

from sales_etl.config import COL_CUSTOMER, COL_REP, FileManagementConfig, PipelineConfig
from sales_etl.grouped import GroupFile, GroupRunResult, _code_key, _free_stamp, run_batch_managed
from sales_etl.ingest import discover_csv_files, read_raw_table
from sales_etl.models import ZERO, Category, FileProcessingError
from sales_etl.normalize import parse_transactions
from sales_etl.report import _ensure_new, _table
from sales_etl.text import canonical

logger = logging.getLogger(__name__)

TASK_PREFIX = "صافي_العملاء_الشهري"
JOURNAL_NAME = "سجل_صافي_العملاء_الشهري.jsonl"
REPORT_NAME = "تقرير صافي مبيعات العملاء الشهري.xlsx"
REPORT_SHEET = "التقرير"
FIXED_COLUMNS: tuple[str, ...] = ("كود العميل", "اسم العميل", "المندوب")
MONTHS: tuple[str, ...] = ("يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
                           "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر")
REP_SEPARATOR = "، "

#: parse_transactions needs a reference date; this task never filters by date.
_NO_REFERENCE_DATE = date(2100, 12, 31)
_TOLERANCE = Decimal("0.005")  # Excel stores IEEE doubles


@dataclass
class MonthlyCustomer:
    code: str
    name: str = ""
    months: dict[int, Decimal] = field(default_factory=dict)  # 1..12 -> net
    #: canonical name -> (first date, read order, name as written); spelling variants are one name
    reps_seen: dict[str, tuple[date, int, str]] = field(default_factory=dict)

    def add_representative(self, name: str, when: date, order: int) -> None:
        key = canonical(name)
        if name and (key not in self.reps_seen or (when, order) < self.reps_seen[key][:2]):
            self.reps_seen[key] = (when, order, name)

    @property
    def representatives(self) -> list[str]:
        """Each name once, in order of first appearance by transaction date (file/row order on ties)."""
        return [name for _, _, name in sorted(self.reps_seen.values())]


@dataclass(frozen=True)
class EligibleRow:
    """One counted transaction, kept for the independent reconciliation."""

    code: str
    txn_date: date
    net: Decimal
    representative: str

    @property
    def month(self) -> int:
        return self.txn_date.month


@dataclass
class MonthlyResult:
    input_dir: Path
    started_at: datetime
    files: list[GroupFile] = field(default_factory=list)
    customers: dict[str, MonthlyCustomer] = field(default_factory=dict)
    names_seen: dict[str, list[str]] = field(default_factory=dict)
    ledger: list[EligibleRow] = field(default_factory=list)
    missing_code_rows: list[str] = field(default_factory=list)  # "file line N"
    other_type_rows: int = 0
    run_dir: Path | None = None
    report_path: Path | None = None

    @property
    def failed(self) -> list[GroupFile]:
        return [f for f in self.files if f.error is not None]

    @property
    def rejected_rows(self) -> int:
        return sum(f.rejected_rows for f in self.files)

    @property
    def name_conflicts(self) -> dict[str, list[str]]:
        """Codes written with more than one name; the report uses the first one found."""
        return {code: names for code, names in self.names_seen.items() if len(names) > 1}

    @property
    def complete(self) -> bool:
        """Every file was read and every transaction row was counted."""
        return bool(self.files) and not self.failed and not self.rejected_rows

    @property
    def last_month(self) -> int:
        return max((m for c in self.customers.values() for m in c.months), default=0)

    @property
    def columns(self) -> tuple[str, ...]:
        return (*FIXED_COLUMNS, *MONTHS[:self.last_month])


def monthly_table(result: MonthlyResult) -> list[list[Any]]:
    """One row per customer code: code, name, representatives, then the net of each month."""
    months = range(1, result.last_month + 1)
    return [[c.code, c.name, REP_SEPARATOR.join(c.representatives), *[c.months.get(m, ZERO) for m in months]]
            for c in sorted(result.customers.values(), key=lambda c: (_code_key(c.code), c.code))]


def _add_file(path: Path, config: PipelineConfig, result: MonthlyResult) -> GroupFile:
    item = GroupFile(file_name=path.name)
    try:
        raw = read_raw_table(
            path, encodings=config.encodings, delimiters=config.delimiters,
            forced_encoding=config.forced_encoding, forced_delimiter=config.forced_delimiter,
        )
        for column in (COL_CUSTOMER, COL_REP):
            if column not in raw.column_index:
                raise FileProcessingError("MISSING_COLUMN", f"required column '{column}' not found")
        i_code = raw.column_index[COL_CUSTOMER]
        i_name = i_code + 1                      # column G: the unlabelled name column after the code
        i_rep = raw.column_index[COL_REP] + 1    # column J: the unlabelled name column after المندوب
        parsed = parse_transactions(raw, _NO_REFERENCE_DATE)
    except FileProcessingError as exc:
        item.error = f"{exc.code}: {exc.message}"
        return item
    except Exception as exc:  # noqa: BLE001 - isolate unexpected failures per file
        item.error = f"UNEXPECTED_ERROR: {type(exc).__name__}: {exc}"
        logger.exception("unexpected failure", extra={"event": "monthly_file_failed", "file": path.name})
        return item
    item.sha256 = raw.sha256
    item.transactions, item.rejected_rows = len(parsed.transactions), len(parsed.rejected)
    for t in parsed.transactions:
        if t.category not in (Category.SALES, Category.RETURNS):
            result.other_type_rows += 1
            continue
        code = t.cells[i_code]
        if not code:
            item.rejected_rows += 1
            result.missing_code_rows.append(f"{path.name} line {t.line_no}")
            continue
        name = t.cells[i_name] if i_name < len(t.cells) else ""
        rep = t.cells[i_rep] if i_rep < len(t.cells) else ""
        customer = result.customers.setdefault(code, MonthlyCustomer(code=code))
        names = result.names_seen.setdefault(code, [])
        if name and name not in names:
            names.append(name)
        if not customer.name and name:
            customer.name = name
        customer.add_representative(rep, t.txn_date, len(result.ledger))
        month = t.txn_date.month
        customer.months[month] = customer.months.get(month, ZERO) + t.net
        result.ledger.append(EligibleRow(code, t.txn_date, t.net, rep))
    return item


def write_monthly_report(result: MonthlyResult, path: Path) -> Path:
    _ensure_new(path)
    wb = Workbook()
    ws = wb.active
    ws.title = REPORT_SHEET
    ws.sheet_view.rightToLeft = True
    _table(ws, result.columns, monthly_table(result), widths=(14, 40, 30, *[16] * result.last_month))
    ws.row_dimensions[1].height = 32
    wb.save(path)
    return path


def reconcile(result: MonthlyResult) -> list[str]:
    """Independent cross-checks of the table against the ledger of counted rows."""
    problems: list[str] = []
    table = monthly_table(result)
    expected: dict[tuple[str, int], Decimal] = {}
    reps: dict[str, list[str]] = {}
    for row in result.ledger:
        expected[(row.code, row.month)] = expected.get((row.code, row.month), ZERO) + row.net
    # representatives: walk the counted rows in date order (stable: read order on equal dates)
    for row in sorted(result.ledger, key=lambda r: r.txn_date):
        seen = reps.setdefault(row.code, [])
        if row.representative and all(canonical(r) != canonical(row.representative) for r in seen):
            seen.append(row.representative)
    last = max((row.month for row in result.ledger), default=0)
    if result.columns != (*FIXED_COLUMNS, *MONTHS[:last]):
        problems.append(f"month columns {result.columns[3:]} differ from يناير..{MONTHS[last - 1] if last else '-'}")
    codes = {row.code for row in result.ledger}
    if len(table) != len(codes) or {r[0] for r in table} != codes:
        problems.append(f"{len(table)} customer rows, {len(codes)} distinct customer codes expected")
    for r in table:
        code = r[0]
        for month, value in enumerate(r[3:], start=1):
            if value != expected.get((code, month), ZERO):
                problems.append(f"code {code}: {MONTHS[month - 1]} {value} differs from the sum of its rows "
                                f"{expected.get((code, month), ZERO)}")
        names = reps.get(code, [])
        if r[2] != REP_SEPARATOR.join(names) or len({canonical(n) for n in names}) != len(names):
            problems.append(f"code {code}: representatives '{r[2]}' differ from {names}")
    table_sum = sum((v for r in table for v in r[3:]), ZERO)
    ledger_sum = sum((row.net for row in result.ledger), ZERO)
    if table_sum != ledger_sum:
        problems.append(f"sum of all monthly nets {table_sum} differs from the net of all counted rows {ledger_sum}")
    return problems


def verify_monthly_report(result: MonthlyResult) -> list[str]:
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
    expected = monthly_table(result)
    if len(values) - 1 != len(expected):
        return [f"{len(values) - 1} rows written, {len(expected)} expected"]
    problems: list[str] = []
    for got, want in zip(values[1:], expected):
        for g, w in zip(got, want):
            if isinstance(w, Decimal):
                if not isinstance(g, (int, float)) or abs(Decimal(str(g)) - w) > _TOLERANCE:
                    problems.append(f"code {want[0]}: cell {g!r} differs from {w}")
            elif (g or "") != w:
                problems.append(f"code {want[0]}: cell {g!r} differs from {w!r}")
    return problems + reconcile(result)


def run_monthly(config: PipelineConfig, *, now: datetime | None = None, files: list[Path] | None = None,
                run_dir: Path | None = None) -> MonthlyResult:
    """Monthly net per customer over every CSV in ``config.input_dir`` (nothing is moved).

    ``files``/``run_dir`` let the managed lifecycle pre-select the batch and own the folder.
    No report folder is created when there are no files.
    """
    config = config.validated()
    result = MonthlyResult(input_dir=config.input_dir, started_at=now or datetime.now())
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
    result.report_path = write_monthly_report(result, run_dir / REPORT_NAME)
    return result


def run_monthly_managed(config: PipelineConfig, files_config: FileManagementConfig | None = None, *,
                        now: datetime | None = None) -> GroupRunResult:
    """Build, save, re-open and reconcile the report; only then move the cleanly read CSV files."""
    return run_batch_managed(
        TASK_PREFIX, JOURNAL_NAME, config, files_config,
        build=lambda files, run_dir, started: run_monthly(config, now=started, files=files, run_dir=run_dir),
        verify=lambda result: verify_monthly_report(result), now=now)
