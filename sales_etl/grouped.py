"""Tasks 3 and 4: sales, returns and net per code (customer or representative), no date filter.

Files are found, decoded and parsed exactly as in tasks 1 and 2 (``ingest`` -> ``normalize``).
A :class:`GroupSpec` names the code column (found by its header); the name is the
unlabelled column right after it. ``نوع المستند`` (B) decides مبيعات / مرتجعات and
``صافى`` (T) is summed with its original sign.

``run_group_managed`` files the inputs away with task 1's lifecycle code (run lock,
write-ahead journal, ``move_file``, ``recover``), only after the report is saved and re-read:

* PROCESSED -> ``processed/<Task>_<execution date>[_HHMMSS]/`` (same stamp as ``output/<Task>_...``)
* RETAINED  -> stays in ``input``: the file failed or had rejected rows, the report was not
               verified, or the file changed during the run

Each task keeps its own journal (``state/<task>_journal.jsonl``). There is no duplicate
detection: a file supplied again is counted again. ``run_group_totals`` alone is read-only.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from collections.abc import Callable
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook

from sales_etl.config import FileManagementConfig, PipelineConfig
from sales_etl.filestore import LOCK_NAME, PROCESSED, RETAINED, Journal, RecoveryAction, RunLock, recover
from sales_etl.ingest import discover_csv_files, read_raw_table, sha256_file
from sales_etl.managed import FileDisposition, ManagedRunResult, _commit_and_move
from sales_etl.models import ZERO, Category, FileProcessingError
from sales_etl.normalize import parse_transactions
from sales_etl.pipeline import run_stamps
from sales_etl.report import _ensure_new, _table

logger = logging.getLogger(__name__)

REPORT_SHEET = "التقرير"
#: parse_transactions needs a reference date; these tasks never filter by date.
_NO_REFERENCE_DATE = date(2100, 12, 31)
_TOLERANCE = Decimal("0.005")  # Excel stores IEEE doubles


@dataclass(frozen=True)
class GroupSpec:
    task: str                   # "Customer" -> output/Customer_<stamp>, state/customer_journal.jsonl
    code_column: str            # source header of the code column; the name is the next column
    columns: tuple[str, ...]    # the five report headers
    report_name: str            # e.g. "customer_totals.xlsx"
    widths: tuple[int, ...] = (14, 40, 20, 20, 20)

    @property
    def journal_name(self) -> str:
        return f"{self.task.lower()}_journal.jsonl"


@dataclass
class GroupTotal:
    code: str
    name: str = ""
    sales: Decimal = ZERO
    returns: Decimal = ZERO

    @property
    def net(self) -> Decimal:
        return self.sales + self.returns


@dataclass
class GroupFile:
    file_name: str
    sha256: str = ""
    transactions: int = 0
    rejected_rows: int = 0
    error: str | None = None


@dataclass
class GroupResult:
    spec: GroupSpec
    input_dir: Path
    started_at: datetime
    files: list[GroupFile] = field(default_factory=list)
    groups: dict[str, GroupTotal] = field(default_factory=dict)
    names_seen: dict[str, list[str]] = field(default_factory=dict)
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
        """Every file was read and every transaction row was parsed."""
        return bool(self.files) and not self.failed and not self.rejected_rows


def _code_key(code: str) -> tuple[int, int | str]:
    """Numeric codes in numeric order, then any other codes alphabetically."""
    return (0, int(code)) if code.isdigit() else (1, code)


def group_table(result: GroupResult) -> list[list[Any]]:
    """One row per code, in the spec's column order (code, name, sales, returns, net)."""
    return [[g.code, g.name, g.sales, g.returns, g.net]
            for g in sorted(result.groups.values(), key=lambda g: _code_key(g.code))]


def _add_file(path: Path, config: PipelineConfig, result: GroupResult) -> GroupFile:
    item = GroupFile(file_name=path.name)
    spec = result.spec
    try:
        raw = read_raw_table(
            path, encodings=config.encodings, delimiters=config.delimiters,
            forced_encoding=config.forced_encoding, forced_delimiter=config.forced_delimiter,
        )
        if spec.code_column not in raw.column_index:
            raise FileProcessingError("MISSING_CODE_COLUMN", f"required column '{spec.code_column}' not found")
        i_code = raw.column_index[spec.code_column]
        i_name = i_code + 1  # the unlabelled name column right after the code
        parsed = parse_transactions(raw, _NO_REFERENCE_DATE)
    except FileProcessingError as exc:
        item.error = f"{exc.code}: {exc.message}"
        return item
    except Exception as exc:  # noqa: BLE001 - isolate unexpected failures per file
        item.error = f"UNEXPECTED_ERROR: {type(exc).__name__}: {exc}"
        logger.exception("unexpected failure", extra={"event": "group_file_failed", "file": path.name})
        return item
    item.sha256 = raw.sha256
    item.transactions, item.rejected_rows = len(parsed.transactions), len(parsed.rejected)
    for t in parsed.transactions:
        if t.category not in (Category.SALES, Category.RETURNS):
            continue
        code = t.cells[i_code]
        name = t.cells[i_name] if i_name < len(t.cells) else ""
        group = result.groups.setdefault(code, GroupTotal(code=code))
        names = result.names_seen.setdefault(code, [])
        if name and name not in names:
            names.append(name)
        if not group.name and name:
            group.name = name
        if t.category is Category.SALES:
            group.sales += t.net
        else:
            group.returns += t.net
    return item


def _free_stamp(output_dir: Path, processed_dir: Path | None, prefix: str, now: datetime) -> str:
    for stamp in run_stamps(now):
        if (output_dir / f"{prefix}_{stamp}").exists():
            continue
        if processed_dir is not None and (processed_dir / f"{prefix}_{stamp}").exists():
            continue
        return stamp
    raise RuntimeError(f"could not find a unique {prefix} folder name under {output_dir}")


def write_group_report(result: GroupResult, path: Path) -> Path:
    _ensure_new(path)
    wb = Workbook()
    ws = wb.active
    ws.title = REPORT_SHEET
    ws.sheet_view.rightToLeft = True
    _table(ws, result.spec.columns, group_table(result), widths=result.spec.widths)
    ws.row_dimensions[1].height = 32
    wb.save(path)
    return path


def verify_group_report(result: GroupResult) -> list[str]:
    """Re-open the saved report and compare every cell before anything is moved."""
    path = result.report_path
    if path is None or not path.is_file() or path.stat().st_size == 0:
        return [f"missing or empty report: {path}"]
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        values = [list(r) for r in wb[REPORT_SHEET].iter_rows(values_only=True)]
        wb.close()
    except Exception as exc:  # noqa: BLE001 - any failure means "not verified"
        return [f"cannot re-open report: {type(exc).__name__}: {exc}"]
    expected = group_table(result)
    if not values or tuple(values[0]) != result.spec.columns:
        return [f"report header mismatch: {values[0] if values else None}"]
    if len(values) - 1 != len(expected):
        return [f"{len(values) - 1} rows written, {len(expected)} expected"]
    for got, want in zip(values[1:], expected):
        for g, w in zip(got, want):
            if isinstance(w, Decimal):
                if not isinstance(g, (int, float)) or abs(Decimal(str(g)) - w) > _TOLERANCE:
                    return [f"code {want[0]}: cell {g!r} differs from {w}"]
            elif (g or "") != w:
                return [f"code {want[0]}: cell {g!r} differs from {w!r}"]
    return []


def run_group_totals(spec: GroupSpec, config: PipelineConfig, *, now: datetime | None = None,
                     files: list[Path] | None = None, run_dir: Path | None = None) -> GroupResult:
    """Totals per code over every CSV in ``config.input_dir``; ``config.reference_date`` is not used.

    ``files``/``run_dir`` let ``run_group_managed`` pre-select the batch and own the folder.
    No report folder is created when there are no files. Nothing is moved.
    """
    config = config.validated()
    result = GroupResult(spec=spec, input_dir=config.input_dir, started_at=now or datetime.now())
    if files is None:
        files, _ = discover_csv_files(config.input_dir, recursive=config.recursive,
                                      exclude_dirs=[config.output_dir])
    if not files:
        return result
    result.files = [_add_file(path, config, result) for path in files]
    if run_dir is None:
        run_dir = config.output_dir / f"{spec.task}_{_free_stamp(config.output_dir, None, spec.task, result.started_at)}"
        run_dir.mkdir(parents=True)
    result.run_dir = run_dir
    result.report_path = write_group_report(result, run_dir / spec.report_name)
    return result


# ------------------------------------------------------------------ file lifecycle
@dataclass
class GroupRunResult:
    journal_path: Path
    run_id: str | None = None
    result: Any = None  # GroupResult (tasks 3, 4) or analysis.AnalysisResult (task 5)
    processed_run_dir: Path | None = None
    dispositions: list[FileDisposition] = field(default_factory=list)
    recovery: list[RecoveryAction] = field(default_factory=list)
    output_problems: list[str] = field(default_factory=list)

    @property
    def nothing_to_do(self) -> bool:
        return self.run_id is None

    @property
    def succeeded(self) -> bool:
        """Report verified, every file read cleanly and moved to processed."""
        return (not self.output_problems and self.result is not None and self.result.complete
                and all(d.disposition == PROCESSED and d.moved for d in self.dispositions))


def _decide(result: Any, files: list[Path], shas: dict[Path, str],
            output_problems: list[str]) -> list[FileDisposition]:
    """``result.files`` holds one :class:`GroupFile` per input file."""
    items = {f.file_name: f for f in result.files}
    decisions: list[FileDisposition] = []
    for path in files:
        sha, item = shas[path], items.get(path.name)
        try:
            unchanged = path.exists() and sha256_file(path) == sha
        except OSError:
            unchanged = False
        if output_problems:
            disposition, reason = RETAINED, "nothing moved: " + "; ".join(output_problems)
        elif item is None:
            disposition, reason = RETAINED, "file was not part of the run results"
        elif not unchanged or (item.sha256 and item.sha256 != sha):
            disposition, reason = RETAINED, "file changed or disappeared during the run; process it again"
        elif item.error:
            disposition, reason = RETAINED, f"failed, left in input: {item.error}"
        elif item.rejected_rows:
            disposition, reason = RETAINED, f"left in input: {item.rejected_rows} rejected row(s) not counted"
        else:
            disposition, reason = PROCESSED, f"processed: {item.transactions} transaction row(s)"
        decisions.append(FileDisposition(path.name, sha, disposition, reason, path))
    return decisions


def run_group_managed(spec: GroupSpec, config: PipelineConfig,
                      files_config: FileManagementConfig | None = None, *,
                      now: datetime | None = None) -> GroupRunResult:
    """Count every CSV in ``input``, save and verify the report, then move the clean files."""
    return run_batch_managed(
        spec.task, spec.journal_name, config, files_config,
        build=lambda files, run_dir, started: run_group_totals(spec, config, now=started, files=files,
                                                              run_dir=run_dir),
        verify=verify_group_report, now=now)


def run_batch_managed(prefix: str, journal_name: str, config: PipelineConfig,
                      files_config: FileManagementConfig | None, *,
                      build: Callable[[list[Path], Path, datetime], Any],
                      verify: Callable[[Any], list[str]],
                      now: datetime | None = None) -> GroupRunResult:
    """Shared lifecycle of tasks 3, 4 and 5: lock -> recover -> build the report -> verify -> move.

    ``build(files, run_dir, started)`` writes the report and returns a result whose ``files`` are
    :class:`GroupFile` items and whose ``complete`` says every file was read cleanly;
    ``verify(result)`` returns problems (any problem means nothing is moved).
    Folders: ``output/<prefix>_<stamp>`` and ``processed/<prefix>_<stamp>``.
    """
    config = config.validated()
    fm = (files_config or FileManagementConfig.beside(config.input_dir)).validated(config)
    started = now or datetime.now()
    journal = Journal(fm.state_dir / journal_name)
    run = GroupRunResult(journal_path=journal.path)

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
        to_process = [p for p in files if p in shas]
        if not to_process:
            return run  # nothing to do: no folder, no journal entry

        stamp = _free_stamp(config.output_dir, fm.processed_dir, prefix, started)
        run.run_id = f"{prefix}_{stamp}"
        run.processed_run_dir = fm.processed_dir / run.run_id
        run_dir = config.output_dir / run.run_id
        journal.append("RUN_STARTED", run_id=run.run_id, run_dir=str(run_dir), task=prefix.lower(),
                       execution_time=started.isoformat(timespec="seconds"),
                       files=[{"file_name": p.name, "sha256": shas[p]} for p in to_process])
        try:
            run_dir.mkdir(parents=True)
            run.result = build(to_process, run_dir, started)
            run.output_problems = verify(run.result)
            run.dispositions = _decide(run.result, to_process, shas, run.output_problems)
            _commit_and_move(
                ManagedRunResult(journal_path=journal.path, run_id=run.run_id, run_dir=run_dir,
                                 dispositions=run.dispositions),
                journal, fm, {PROCESSED: run.processed_run_dir}, started)
            journal.append("RUN_FINISHED", run_id=run.run_id,
                           status="SUCCEEDED" if run.succeeded else "COMPLETED_WITH_PROBLEMS")
        except Exception as exc:
            journal.append("RUN_FINISHED", run_id=run.run_id, status="FAILED",
                           error=f"{type(exc).__name__}: {exc}")
            logger.exception("%s run failed", prefix, extra={"event": "run_error"})
            raise
    return run
