"""Managed (automatic) mode: process every CSV in ``input``, then file each source away.

Every CSV directly in ``input`` is processed in the current run. There is no duplicate
detection of any kind: re-supplying a file counts its transactions again, by design.

Disposition rules (decided only AFTER all outputs are written and verified):

* PROCESSED -> ``processed/Run_<execution date>[_HHMMSS]/`` : file processed, none of its
                                 own checks FAIL, outputs verified
* FAILED    -> ``failures/``   : the file itself failed (unreadable, missing column, rejected
                                 rows, ERP control total mismatch, ...)
* RETAINED  -> stays in ``input``: a system-level problem (outputs not verified, grand total
                                 inconsistency, file changed during the run); nothing committed

Run folders are named by the EXECUTION date (``output/Report_2026-10-20``); the
reference ("yesterday") date only drives the calculations.

See ``filestore.py`` for the commit/recovery protocol.
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from sales_etl.config import FileManagementConfig, PipelineConfig
from sales_etl.filestore import (
    FAILED,
    JOURNAL_NAME,
    LOCK_NAME,
    PROCESSED,
    RETAINED,
    Journal,
    MoveError,
    RecoveryAction,
    RunLock,
    move_file,
    recover,
    unique_destination,
)
from sales_etl.ingest import discover_csv_files, sha256_file
from sales_etl.logging_setup import attach_run_log, detach
from sales_etl.models import CheckStatus, PipelineResult
from sales_etl.pipeline import create_run_dir, run_pipeline

logger = logging.getLogger(__name__)

_SYSTEM_CHECKS = ("GRAND_TOTAL", "REPORT_READBACK")

# run statuses
SUCCEEDED = "SUCCEEDED"  # report + audit generated and verified, audit PASS, every file archived
COMPLETED_WITH_PROBLEMS = "COMPLETED_WITH_PROBLEMS"  # report generated, but audit FAIL / failed or unmoved files
FAILED_RUN = "FAILED"  # no verified report was produced
NO_INPUT_FILES = "NO_INPUT_FILES"  # nothing in input: no run folder, no report


@dataclass
class FileDisposition:
    file_name: str
    sha256: str
    disposition: str
    reason: str
    source: Path
    destination: Path | None = None
    moved: bool = False
    error: str = ""

    @property
    def location(self) -> Path:
        """Where the file is now."""
        return self.destination if self.moved and self.destination else self.source


@dataclass
class ManagedRunResult:
    journal_path: Path
    execution_time: datetime | None = None
    run_id: str | None = None
    run_dir: Path | None = None
    processed_run_dir: Path | None = None
    pipeline: PipelineResult | None = None
    dispositions: list[FileDisposition] = field(default_factory=list)
    recovery: list[RecoveryAction] = field(default_factory=list)
    notices: list[str] = field(default_factory=list)
    output_problems: list[str] = field(default_factory=list)
    error: str = ""
    manifest_path: Path | None = None

    @property
    def nothing_to_do(self) -> bool:
        return self.run_dir is None

    def expected_outputs(self) -> dict[str, Path | None]:
        p = self.pipeline
        return {
            "report": p.report_path if p else None,
            "audit_xlsx": p.audit_xlsx_path if p else None,
            "audit_json": p.audit_json_path if p else None,
            "log": self.run_dir / "pipeline.log.jsonl" if self.run_dir else None,
        }

    @property
    def report_generated(self) -> bool:
        """The required outputs exist on disk and passed verification."""
        outputs = self.expected_outputs()
        return (not self.error and not self.output_problems and self.pipeline is not None
                and all(path is not None and path.is_file() for path in outputs.values()))

    @property
    def status(self) -> str:
        if self.nothing_to_do:
            return NO_INPUT_FILES
        if not self.report_generated:
            return FAILED_RUN
        if (self.pipeline.audit_status is CheckStatus.PASS  # type: ignore[union-attr]
                and all(d.disposition == PROCESSED and d.moved for d in self.dispositions)):
            return SUCCEEDED
        return COMPLETED_WITH_PROBLEMS

    @property
    def succeeded(self) -> bool:
        return self.status == SUCCEEDED


def verify_outputs(result: PipelineResult) -> list[str]:
    """Re-open every output and confirm it is complete before anything is committed."""
    problems: list[str] = []
    readback = [c for c in result.checks if c.check == "REPORT_READBACK"]
    if not readback or readback[0].status is not CheckStatus.PASS:
        problems.append("report read-back did not PASS")
    for path in (result.report_path, result.audit_xlsx_path, result.audit_json_path, result.log_path):
        if path is None or not path.is_file() or path.stat().st_size == 0:
            problems.append(f"missing or empty output: {path}")
    if problems:
        return problems
    try:
        wb = load_workbook(result.audit_xlsx_path, read_only=True)
        sheets = wb.sheetnames
        wb.close()
        if "Checks" not in sheets or "Transactions" not in sheets:
            problems.append(f"audit workbook incomplete: {sheets}")
        payload = json.loads(result.audit_json_path.read_text(encoding="utf-8"))  # type: ignore[union-attr]
        listed = {f["file"]: f for f in payload["files"]}
        for outcome in result.outcomes:
            entry = listed.get(outcome.file_name)
            if entry is None:
                problems.append(f"audit JSON has no entry for {outcome.file_name}")
            elif outcome.parsed and entry.get("sha256") != outcome.parsed.sha256:
                problems.append(f"audit JSON SHA-256 mismatch for {outcome.file_name}")
        if payload.get("audit_status") != result.audit_status.value:
            problems.append("audit JSON status differs from the computed status")
    except Exception as exc:  # noqa: BLE001 - any failure means "not verified"
        problems.append(f"cannot re-open audit outputs: {type(exc).__name__}: {exc}")
    return problems


def _decide(result: PipelineResult | None, to_process: list[Path], shas: dict[Path, str],
            output_problems: list[str]) -> list[FileDisposition]:
    system = list(output_problems)
    if result is not None:
        system += [f"{c.check} FAIL" for c in result.checks
                   if c.status is CheckStatus.FAIL and c.check.startswith(_SYSTEM_CHECKS)]
    outcomes = {o.path: o for o in result.outcomes} if result else {}
    decisions: list[FileDisposition] = []
    for path in to_process:
        sha, outcome = shas[path], outcomes.get(path)
        try:
            unchanged = path.exists() and sha256_file(path) == sha
        except OSError:
            unchanged = False
        if system:
            disposition, reason = RETAINED, "nothing committed: " + "; ".join(system)
        elif outcome is None:
            disposition, reason = RETAINED, "file was not part of the run results"
        elif not unchanged or (outcome.parsed and outcome.parsed.sha256 != sha):
            disposition, reason = RETAINED, "file changed or disappeared during the run; process it again"
        elif not outcome.succeeded:
            disposition, reason = FAILED, f"{outcome.error_code}: {outcome.error_message}"
        elif failing := [c.check for c in outcome.checks if c.status is CheckStatus.FAIL]:
            disposition, reason = FAILED, "audit FAIL: " + ", ".join(failing)
        else:
            warns = sum(c.status is CheckStatus.WARN for c in outcome.checks)
            disposition = PROCESSED
            reason = f"processed; all {len(outcome.checks)} file checks without FAIL ({warns} WARN)"
        decisions.append(FileDisposition(path.name, sha, disposition, reason, path))
    return decisions


def run_managed(config: PipelineConfig, files_config: FileManagementConfig | None = None, *,
                now: datetime | None = None) -> ManagedRunResult:
    """Process everything in ``input`` and move each file to its final folder, crash-safely."""
    config = config.validated()
    fm = (files_config or FileManagementConfig.beside(config.input_dir)).validated(config)
    for folder in (fm.processed_dir, fm.failures_dir, fm.state_dir, config.output_dir):
        folder.mkdir(parents=True, exist_ok=True)
    started = now or datetime.now()
    journal = Journal(fm.state_dir / JOURNAL_NAME)
    managed = ManagedRunResult(journal_path=journal.path, execution_time=started)

    with RunLock(fm.state_dir / LOCK_NAME):
        managed.recovery = recover(journal, started)
        state = journal.state()
        if state.torn_tail:
            managed.notices.append("the journal's last line was an interrupted write; it never took effect "
                                   "and was ignored")
        pending: dict[Path, list] = {}
        for commit in state.pending_moves():
            pending.setdefault(Path(commit.source).resolve(), []).append(commit)

        files, skipped = discover_csv_files(
            config.input_dir, exclude_dirs=[config.output_dir, fm.processed_dir, fm.failures_dir,
                                            fm.state_dir])
        to_process: list[Path] = []
        shas: dict[Path, str] = {}
        superseded: list = []
        for path in files:
            try:
                sha = sha256_file(path)
            except OSError as exc:
                skipped.append((path, f"cannot read file: {exc}"))
                continue
            shas[path] = sha
            to_process.append(path)  # always processed - no history/content based skipping
            for commit in pending.get(path.resolve(), []):
                superseded.append(commit)
                managed.notices.append(f"{path.name}: an earlier run ({commit.run_id}) could not archive "
                                       "this file; it is processed again in this run")
        for path, reason in skipped:
            managed.notices.append(f"{path.name}: skipped - {reason}")

        if not to_process:
            return managed  # nothing to do: no run folder is created

        run_dir = create_run_dir(config.output_dir, started, also_free=fm.processed_dir)
        stamp = run_dir.name.removeprefix("Report_")
        managed.run_dir, managed.run_id = run_dir, run_dir.name
        managed.processed_run_dir = fm.processed_dir / f"Run_{stamp}"
        folders = {PROCESSED: managed.processed_run_dir, FAILED: fm.failures_dir}
        handler = attach_run_log(run_dir / "pipeline.log.jsonl")
        try:
            journal.append("RUN_STARTED", run_id=run_dir.name, run_dir=str(run_dir),
                           execution_time=started.isoformat(timespec="seconds"),
                           reference_date=config.reference_date.isoformat(),
                           files=[{"file_name": p.name, "sha256": shas[p]} for p in to_process])
            for commit in superseded:  # stop retrying the old move: the file is re-run now
                journal.append("MOVE_ABANDONED", run_id=commit.run_id, sha256=commit.sha256,
                               source=commit.source, destination=commit.destination,
                               error=f"superseded: file processed again in {run_dir.name}")
            managed.pipeline = run_pipeline(config, now=started, files=to_process, skipped=skipped,
                                            run_dir=run_dir)
            managed.pipeline.log_path = run_dir / "pipeline.log.jsonl"
            managed.output_problems = verify_outputs(managed.pipeline)
            managed.dispositions = _decide(managed.pipeline, to_process, shas, managed.output_problems)
            _commit_and_move(managed, journal, fm, folders, started)
            managed.manifest_path = _write_manifest(managed, config)
            journal.append("RUN_FINISHED", run_id=run_dir.name, status=managed.status)
        except Exception as exc:  # noqa: BLE001 - record, keep files where they are, re-raise
            managed.error = f"{type(exc).__name__}: {exc}"
            journal.append("RUN_FINISHED", run_id=run_dir.name, status=FAILED_RUN, error=managed.error)
            logger.exception("managed run failed", extra={"event": "run_error"})
            if managed.manifest_path is None:
                try:
                    managed.manifest_path = _write_manifest(managed, config)
                except Exception:  # noqa: BLE001 - never mask the original error
                    logger.exception("could not write the failure manifest")
            raise
        finally:
            detach(handler)
    return managed


def _commit_and_move(managed: ManagedRunResult, journal: Journal, fm: FileManagementConfig,
                     folders: dict[str, Path], now: datetime) -> None:
    to_commit = [d for d in managed.dispositions if d.disposition != RETAINED]
    if not to_commit:
        return
    reserved: set[Path] = set()
    for item in to_commit:
        item.destination = unique_destination(folders[item.disposition], item.file_name, item.sha256,
                                              now, reserved)
        reserved.add(item.destination)
    journal.append("COMMIT", run_id=managed.run_id, run_dir=str(managed.run_dir), files=[
        {"file_name": d.file_name, "sha256": d.sha256, "source": str(d.source),
         "destination": str(d.destination), "disposition": d.disposition, "reason": d.reason}
        for d in to_commit])
    for item in to_commit:
        base = {"run_id": managed.run_id, "sha256": item.sha256, "source": str(item.source)}
        try:
            move_file(item.source, item.destination, item.sha256)  # type: ignore[arg-type]
        except MoveError as exc:
            item.error = str(exc)
            journal.append("MOVE_FAILED", **base, destination=str(item.destination), error=str(exc))
            logger.error("move failed: %s", exc, extra={"event": "move_failed", "file": item.file_name})
            continue
        item.moved = True
        journal.append("MOVED", **base, destination=str(item.destination), recovered=False)
        logger.info("file moved", extra={"event": "moved", "file": item.file_name,
                                         "disposition": item.disposition,
                                         "destination": str(item.destination)})


def _write_manifest(managed: ManagedRunResult, config: PipelineConfig) -> Path:
    """Manifest of what is actually on disk (output paths are checked, not assumed)."""
    path = managed.run_dir / "run_manifest.json"  # type: ignore[operator]
    pipeline = managed.pipeline
    when = managed.execution_time
    outputs = {name: {"path": str(p) if p else None, "exists": bool(p and p.is_file())}
               for name, p in managed.expected_outputs().items()}
    payload: dict[str, Any] = {
        "run_id": managed.run_id,
        "status": managed.status,
        "report_generated": managed.report_generated,
        "execution_date": when.date().isoformat() if when else None,
        "execution_time": when.isoformat(timespec="seconds") if when else None,
        "reference_date": config.reference_date.isoformat(),
        "audit_status": pipeline.audit_status.value if pipeline else None,
        "report_dir": str(managed.run_dir),
        "processed_run_dir": str(managed.processed_run_dir),
        "outputs": outputs,
        "output_problems": managed.output_problems,
        "error": managed.error or None,
        "files": [{
            "file_name": d.file_name, "sha256": d.sha256, "disposition": d.disposition,
            "reason": d.reason, "source": str(d.source),
            "destination": str(d.destination) if d.destination else None,
            "moved": d.moved, "current_location": str(d.location), "error": d.error,
        } for d in managed.dispositions],
        "recovery": [vars(a) for a in managed.recovery],
        "notices": managed.notices,
        "journal": str(managed.journal_path),
    }
    with open(path, "x", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    return path
