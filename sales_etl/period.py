"""Task 2: total sales and returns for a user-chosen period (start and end days included).

Files are found, decoded and parsed exactly as in task 1 (``ingest`` -> ``normalize``);
only the ``صافى`` column is summed, with its original sign.

``run_period_managed`` files each input away with task 1's lifecycle code (run lock,
write-ahead journal, ``move_file``, ``recover``), after the report is written and verified:

* PROCESSED -> ``processed/Period_<execution date>[_HHMMSS]/`` (same stamp as ``output/Period_...``)
* FAILED    -> ``failures/``   : unreadable file or rejected rows
* RETAINED  -> stays in ``input``: report not verified, or the file changed during the run

Task 2 keeps its own journal (``state/period_journal.jsonl``). There is no duplicate
detection: every CSV in ``input`` is counted, including content counted before.
``run_period_totals`` alone is the read-only path (``--no-move``).
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Font, Side

from sales_etl.config import (
    COL_DATE,
    COL_NET,
    GRAND_TOTAL_LABEL,
    REPORT_FILE_COLUMN,
    TOTAL_COLUMNS,
    FileManagementConfig,
    PipelineConfig,
)
from sales_etl.dates import format_dmy
from sales_etl.filestore import (
    FAILED,
    LOCK_NAME,
    PROCESSED,
    RETAINED,
    Journal,
    RecoveryAction,
    RunLock,
    recover,
)
from sales_etl.ingest import discover_csv_files, read_raw_table, sha256_file
from sales_etl.managed import FileDisposition, ManagedRunResult, _commit_and_move
from sales_etl.models import ZERO, Category, FileProcessingError, Transaction
from sales_etl.normalize import parse_transactions, total
from sales_etl.pipeline import run_stamps
from sales_etl.report import _THIN, _TOTAL_FILL, MONEY_FORMAT, _ensure_new, _table

logger = logging.getLogger(__name__)

TASK_NAME = "إجمالي المبيعات والمرتجعات لفترة محددة"
PERIOD_COLUMNS: tuple[str, ...] = (REPORT_FILE_COLUMN, *TOTAL_COLUMNS.values())
PERIOD_SHEET = "التقرير"
PERIOD_METADATA_SHEET = "بيانات التقرير"
PERIOD_JOURNAL_NAME = "period_journal.jsonl"


@dataclass
class PeriodFileTotal:
    """One file's totals for the period; ``error`` is set when the file could not be read."""

    file_name: str
    sales: Decimal = ZERO
    returns: Decimal = ZERO
    rows_in_period: int = 0
    rejected_rows: int = 0
    error: str | None = None
    sha256: str = ""

    @property
    def net(self) -> Decimal:
        return self.sales + self.returns


@dataclass
class PeriodResult:
    start_date: date
    end_date: date
    input_dir: Path
    started_at: datetime
    files: list[PeriodFileTotal] = field(default_factory=list)
    run_dir: Path | None = None
    report_path: Path | None = None

    @property
    def successful(self) -> list[PeriodFileTotal]:
        return [f for f in self.files if f.error is None]

    @property
    def failed(self) -> list[PeriodFileTotal]:
        return [f for f in self.files if f.error is not None]

    @property
    def sales(self) -> Decimal:
        return total([f.sales for f in self.successful])

    @property
    def returns(self) -> Decimal:
        return total([f.returns for f in self.successful])

    @property
    def net(self) -> Decimal:
        return self.sales + self.returns

    @property
    def rejected_rows(self) -> int:
        return sum(f.rejected_rows for f in self.successful)

    @property
    def complete(self) -> bool:
        """Every file was read and every transaction row was parsed."""
        return bool(self.files) and not self.failed and not self.rejected_rows


def period_totals(transactions: list[Transaction], start_date: date, end_date: date) -> tuple[Decimal, Decimal, int]:
    """``(sales, returns, rows in period)`` for ``start_date <= التاريخ <= end_date``."""
    in_period = [t for t in transactions if start_date <= t.txn_date <= end_date]
    sales = total([t.net for t in in_period if t.category is Category.SALES])
    returns = total([t.net for t in in_period if t.category is Category.RETURNS])
    return sales, returns, len(in_period)


def _file_total(path: Path, config: PipelineConfig, start_date: date, end_date: date) -> PeriodFileTotal:
    result = PeriodFileTotal(file_name=path.name)
    try:
        raw = read_raw_table(
            path, encodings=config.encodings, delimiters=config.delimiters,
            forced_encoding=config.forced_encoding, forced_delimiter=config.forced_delimiter,
        )
        parsed = parse_transactions(raw, end_date)
        result.sales, result.returns, result.rows_in_period = period_totals(
            parsed.transactions, start_date, end_date)
        result.rejected_rows = len(parsed.rejected)
        result.sha256 = raw.sha256
    except FileProcessingError as exc:
        result.error = f"{exc.code}: {exc.message}"
    except Exception as exc:  # noqa: BLE001 - isolate unexpected failures per file
        result.error = f"UNEXPECTED_ERROR: {type(exc).__name__}: {exc}"
        logger.exception("unexpected failure", extra={"event": "period_file_failed", "file": path.name})
    return result


def period_table(result: PeriodResult) -> list[list[Any]]:
    """Rows in ``PERIOD_COLUMNS`` order, one per readable file; the last one is the total."""
    lines = [(f.file_name, f.sales, f.returns) for f in result.successful]
    lines.append((GRAND_TOTAL_LABEL, result.sales, result.returns))
    return [[label, sales, returns, sales + returns] for label, sales, returns in lines]


def _create_period_dir(output_dir: Path, now: datetime, *, also_free: Path | None = None) -> Path:
    """``output_dir/Period_<execution date>``, never reusing an existing folder.

    ``also_free``: a folder in which ``Period_<same stamp>`` must not exist either, so the
    report folder and the processed-archive folder of one run share the same stamp.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    for stamp in run_stamps(now):
        if also_free is not None and (also_free / f"Period_{stamp}").exists():
            continue
        candidate = output_dir / f"Period_{stamp}"
        try:
            candidate.mkdir()
            return candidate
        except FileExistsError:
            continue
    raise RuntimeError(f"could not create a unique period folder under {output_dir}")


def write_period_report(result: PeriodResult, path: Path) -> Path:
    _ensure_new(path)
    wb = Workbook()
    ws = wb.active
    ws.title = PERIOD_SHEET
    ws.sheet_view.rightToLeft = True
    *file_rows, total_row = period_table(result)
    last = _table(ws, PERIOD_COLUMNS, file_rows, widths=[34, 20, 20, 20])
    for col, value in enumerate(total_row, start=1):
        cell = ws.cell(row=last + 1, column=col, value=value)
        cell.font, cell.fill = Font(bold=True), _TOTAL_FILL
        cell.border = Border(left=_THIN, right=_THIN, top=Side(style="double"), bottom=_THIN)
        if isinstance(value, Decimal):
            cell.number_format = MONEY_FORMAT
    ws.row_dimensions[1].height = 32
    ws.oddFooter.center.text = f"{format_dmy(result.start_date)} - {format_dmy(result.end_date)}"

    entries: list[tuple[str, Any]] = [
        ("المهمة", TASK_NAME),
        ("تاريخ البداية", result.start_date),
        ("تاريخ النهاية", result.end_date),
        ("الفترة", f"{format_dmy(result.start_date)} <= {COL_DATE} <= {format_dmy(result.end_date)}"
                   " (يشمل يوم البداية ويوم النهاية)"),
        ("أساس الحساب", f"عمود {COL_NET} فقط، بالإشارة الأصلية"),
        ("وقت التشغيل", result.started_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("مجلد الملفات", str(result.input_dir)),
        ("عدد الملفات", len(result.files)),
        ("ملفات تمت قراءتها", len(result.successful)),
        ("ملفات فشلت (غير موجودة في التقرير)", len(result.failed)),
        ("صفوف مرفوضة (غير محسوبة)", result.rejected_rows),
    ]
    entries += [(f"ملف فاشل: {f.file_name}", f.error) for f in result.failed]
    meta = wb.create_sheet(PERIOD_METADATA_SHEET)
    meta.sheet_view.rightToLeft = True
    _table(meta, ["البند", "القيمة"], entries, widths=[40, 90])
    wb.save(path)
    return path


def run_period_totals(config: PipelineConfig, start_date: date, end_date: date, *,
                      now: datetime | None = None, files: list[Path] | None = None,
                      run_dir: Path | None = None) -> PeriodResult:
    """Compute the period totals for every CSV in ``config.input_dir`` and save the report.

    ``files``/``run_dir`` let ``run_period_managed`` pre-select the batch and own the folder.
    ``config.reference_date`` is not used. No report folder is created when there are no files.
    """
    if end_date < start_date:
        raise ValueError(f"end date {format_dmy(end_date)} is before start date {format_dmy(start_date)}")
    config = config.validated()
    result = PeriodResult(start_date=start_date, end_date=end_date, input_dir=config.input_dir,
                          started_at=now or datetime.now())
    if files is None:
        files, _ = discover_csv_files(config.input_dir, recursive=config.recursive,
                                      exclude_dirs=[config.output_dir])
    if not files:
        return result
    result.files = [_file_total(path, config, start_date, end_date) for path in files]
    result.run_dir = run_dir or _create_period_dir(config.output_dir, result.started_at)
    result.report_path = write_period_report(
        result, result.run_dir / f"period_totals_{start_date.isoformat()}_to_{end_date.isoformat()}.xlsx")
    return result


# ------------------------------------------------------------------ file lifecycle
@dataclass
class PeriodRunResult:
    journal_path: Path
    run_id: str | None = None
    period: PeriodResult | None = None
    processed_run_dir: Path | None = None
    dispositions: list[FileDisposition] = field(default_factory=list)
    recovery: list[RecoveryAction] = field(default_factory=list)
    output_problems: list[str] = field(default_factory=list)

    @property
    def nothing_to_do(self) -> bool:
        return self.run_id is None

    @property
    def succeeded(self) -> bool:
        """Report verified and every file filed away."""
        return (not self.output_problems
                and (self.period is None or self.period.complete)
                and all(d.disposition == PROCESSED and d.moved for d in self.dispositions))


def verify_period_report(result: PeriodResult) -> list[str]:
    """Re-open the saved report and confirm it holds every counted file before anything moves."""
    path = result.report_path
    if path is None or not path.is_file() or path.stat().st_size == 0:
        return [f"missing or empty report: {path}"]
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        values = [list(r) for r in wb[PERIOD_SHEET].iter_rows(values_only=True)]
        wb.close()
    except Exception as exc:  # noqa: BLE001 - any failure means "not verified"
        return [f"cannot re-open report: {type(exc).__name__}: {exc}"]
    expected = period_table(result)
    if not values or tuple(values[0]) != PERIOD_COLUMNS:
        return [f"report header mismatch: {values[0] if values else None}"]
    if [r[0] for r in values[1:]] != [r[0] for r in expected]:
        return ["report rows do not match the counted files"]
    return []


def _decide(result: PeriodResult | None, files: list[Path], shas: dict[Path, str],
            output_problems: list[str]) -> list[FileDisposition]:
    totals = {f.file_name: f for f in result.files} if result else {}
    decisions: list[FileDisposition] = []
    for path in files:
        sha, item = shas[path], totals.get(path.name)
        try:
            unchanged = path.exists() and sha256_file(path) == sha
        except OSError:
            unchanged = False
        if output_problems:
            disposition, reason = RETAINED, "nothing committed: " + "; ".join(output_problems)
        elif item is None:
            disposition, reason = RETAINED, "file was not part of the run results"
        elif not unchanged or (item.sha256 and item.sha256 != sha):
            disposition, reason = RETAINED, "file changed or disappeared during the run; process it again"
        elif item.error:
            disposition, reason = FAILED, item.error
        elif item.rejected_rows:
            disposition, reason = FAILED, f"{item.rejected_rows} rejected row(s) not counted"
        else:
            disposition, reason = PROCESSED, f"counted: {item.rows_in_period} row(s) in the period"
        decisions.append(FileDisposition(path.name, sha, disposition, reason, path))
    return decisions


def _free_stamp(output_dir: Path, processed_dir: Path, now: datetime) -> str:
    for stamp in run_stamps(now):
        if not (output_dir / f"Period_{stamp}").exists() and not (processed_dir / f"Period_{stamp}").exists():
            return stamp
    raise RuntimeError(f"could not find a unique period stamp under {output_dir}")


def run_period_managed(config: PipelineConfig, start_date: date, end_date: date,
                       files_config: FileManagementConfig | None = None, *,
                       now: datetime | None = None) -> PeriodRunResult:
    """Count every new CSV in ``input`` for the period, then file each one away, crash-safely."""
    if end_date < start_date:
        raise ValueError(f"end date {format_dmy(end_date)} is before start date {format_dmy(start_date)}")
    config = config.validated()
    fm = (files_config or FileManagementConfig.beside(config.input_dir)).validated(config)
    started = now or datetime.now()
    journal = Journal(fm.state_dir / PERIOD_JOURNAL_NAME)
    run = PeriodRunResult(journal_path=journal.path)

    with RunLock(fm.state_dir / LOCK_NAME):
        run.recovery = recover(journal, started)
        files, _ = discover_csv_files(
            config.input_dir, exclude_dirs=[config.output_dir, fm.processed_dir, fm.failures_dir,
                                            fm.state_dir])
        shas: dict[Path, str] = {}
        for path in files:
            try:
                shas[path] = sha256_file(path)
            except OSError:
                continue  # unreadable now: left in input for the next run
        to_count = [p for p in files if p in shas]
        if not to_count:
            return run  # nothing to do: no folder, no journal entry

        stamp = _free_stamp(config.output_dir, fm.processed_dir, started)
        run.run_id = f"Period_{stamp}"
        run.processed_run_dir = fm.processed_dir / run.run_id
        run_dir = config.output_dir / run.run_id
        journal.append("RUN_STARTED", run_id=run.run_id, run_dir=str(run_dir), task="period",
                       execution_time=started.isoformat(timespec="seconds"),
                       start_date=start_date.isoformat(), end_date=end_date.isoformat(),
                       files=[{"file_name": p.name, "sha256": shas[p]} for p in to_count])
        try:
            run_dir.mkdir(parents=True)
            run.period = run_period_totals(config, start_date, end_date, now=started,
                                           files=to_count, run_dir=run_dir)
            run.output_problems = verify_period_report(run.period)
            run.dispositions = _decide(run.period, to_count, shas, run.output_problems)
            _commit_and_move(
                ManagedRunResult(journal_path=journal.path, run_id=run.run_id, run_dir=run_dir,
                                 dispositions=run.dispositions),
                journal, fm,
                {PROCESSED: run.processed_run_dir, FAILED: fm.failures_dir},
                started)
            journal.append("RUN_FINISHED", run_id=run.run_id,
                           status="SUCCEEDED" if run.succeeded else "COMPLETED_WITH_PROBLEMS")
        except Exception as exc:
            journal.append("RUN_FINISHED", run_id=run.run_id, status="FAILED",
                           error=f"{type(exc).__name__}: {exc}")
            logger.exception("period run failed", extra={"event": "run_error"})
            raise
    return run
