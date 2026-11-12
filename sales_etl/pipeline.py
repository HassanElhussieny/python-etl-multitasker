"""Batch orchestration: discover -> ingest -> normalise -> compute -> reconcile -> report."""

from __future__ import annotations

import logging
from collections.abc import Iterator
from datetime import datetime
from pathlib import Path

from sales_etl.amounts import fmt_amount
from sales_etl.config import PipelineConfig
from sales_etl.ingest import discover_csv_files, read_raw_table, sha256_file
from sales_etl.logging_setup import attach_run_log, detach
from sales_etl.metrics import build_ledger, compute_customer_sets, compute_file_metrics
from sales_etl.models import (
    CheckResult,
    CheckStatus,
    FileOutcome,
    FileProcessingError,
    PipelineResult,
)
from sales_etl.normalize import parse_transactions
from sales_etl.reconcile import reconcile_customers, reconcile_file, reconcile_grand_customers, reconcile_grand_total
from sales_etl.report import verify_report, write_audit_json, write_audit_workbook, write_report

logger = logging.getLogger(__name__)


def run_stamps(now: datetime) -> Iterator[str]:
    """Candidate run names by EXECUTION date: ``2026-10-20``, then ``2026-10-20_104822``, ``..._2``."""
    yield f"{now:%Y-%m-%d}"
    base = f"{now:%Y-%m-%d_%H%M%S}"
    yield base
    for attempt in range(2, 1000):
        yield f"{base}_{attempt}"


def create_run_dir(output_dir: Path, now: datetime, *, also_free: Path | None = None) -> Path:
    """Create ``output_dir/Report_<execution date>`` (never reuses an existing folder).

    ``also_free``: a folder in which ``Run_<same stamp>`` must not exist either, so the
    report folder and the processed-archive folder of one run share the same stamp.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    for stamp in run_stamps(now):
        if also_free is not None and (also_free / f"Run_{stamp}").exists():
            continue
        candidate = output_dir / f"Report_{stamp}"
        try:
            candidate.mkdir()
            return candidate
        except FileExistsError:
            continue
    raise RuntimeError(f"could not create a unique run folder under {output_dir}")


def process_file(path: Path, config: PipelineConfig) -> FileOutcome:
    """Process one CSV; any failure is captured in the outcome, never raised."""
    outcome = FileOutcome(file_name=path.name, path=path)
    try:
        raw = read_raw_table(
            path, encodings=config.encodings, delimiters=config.delimiters,
            forced_encoding=config.forced_encoding, forced_delimiter=config.forced_delimiter,
        )
        parsed = parse_transactions(raw, config.reference_date)
        ledger = build_ledger(parsed.transactions)
        metrics = compute_file_metrics(ledger, config.reference_date)
        outcome.parsed, outcome.metrics = parsed, metrics
        outcome.checks = reconcile_file(parsed, ledger, metrics, config.reference_date)
        outcome.customers = compute_customer_sets(ledger, config.reference_date)
        outcome.checks.append(reconcile_customers(parsed.file_name, "CUSTOMER_COUNTS", outcome.customers,
                                                  parsed.transactions, config.reference_date))
        logger.info("file processed", extra={
            "event": "file_ok", "file": path.name, "valid_rows": len(parsed.transactions),
            "rejected_rows": len(parsed.rejected),
            **{k: str(getattr(metrics, k)) for k in
               ("sales_yesterday", "sales_previous", "returns_yesterday", "returns_previous")},
        })
    except FileProcessingError as exc:
        outcome.error_code, outcome.error_message = exc.code, exc.message
        logger.warning("file failed: %s", exc, extra={"event": "file_failed", "file": path.name,
                                                      "code": exc.code})
    except Exception as exc:  # noqa: BLE001 - isolate unexpected failures per file
        outcome.error_code, outcome.error_message = "UNEXPECTED_ERROR", f"{type(exc).__name__}: {exc}"
        outcome.parsed = outcome.metrics = outcome.customers = None
        outcome.checks = []
        logger.exception("unexpected failure", extra={"event": "file_failed", "file": path.name})
    return outcome


def run_pipeline(
    config: PipelineConfig,
    *,
    now: datetime | None = None,
    files: list[Path] | None = None,
    skipped: list[tuple[Path, str]] | None = None,
    run_dir: Path | None = None,
) -> PipelineResult:
    """Run the whole batch with ONE reference date shared by every file.

    ``files``/``skipped`` let a caller (managed mode) pre-select the batch;
    when ``run_dir`` is given the caller owns the run folder and the run log.
    """
    config = config.validated()
    started = now or datetime.now()
    owns_run_dir = run_dir is None
    if run_dir is None:
        run_dir = create_run_dir(config.output_dir, started)
    result = PipelineResult(reference_date=config.reference_date, input_dir=config.input_dir,
                            run_dir=run_dir, started_at=started)
    stem = f"ref-{config.reference_date.isoformat()}"
    result.log_path = run_dir / "pipeline.log.jsonl"
    handler = attach_run_log(result.log_path) if owns_run_dir else None
    try:
        logger.info("run started", extra={"event": "run_start", "reference_date": config.reference_date,
                                          "input_dir": config.input_dir, "run_dir": run_dir})
        if files is None:
            files, found_skipped = discover_csv_files(config.input_dir, recursive=config.recursive,
                                                      exclude_dirs=[config.output_dir])
            skipped = [*found_skipped, *(skipped or [])]
        result.discovered, result.skipped = list(files), list(skipped or [])
        hashes_before = {}
        for path in files:
            try:
                hashes_before[path] = sha256_file(path)
            except OSError:
                hashes_before[path] = None

        result.outcomes = [process_file(path, config) for path in files]

        checks: list[CheckResult] = [c for o in result.outcomes for c in o.checks]
        result.grand_total, grand_checks = reconcile_grand_total(result.outcomes, config.reference_date)
        checks += grand_checks
        result.grand_customers, customer_checks = reconcile_grand_customers(result.outcomes, config.reference_date)
        checks += customer_checks
        checks.append(CheckResult(
            "BATCH", "FILES_DISCOVERED", CheckStatus.PASS if files else CheckStatus.FAIL,
            expected=">= 1", actual=str(len(files)),
            detail=f"{len(files)} eligible CSV file(s) in {config.input_dir}",
        ))
        failed = result.failed
        checks.append(CheckResult(
            "BATCH", "ALL_FILES_PROCESSED", CheckStatus.FAIL if failed else CheckStatus.PASS,
            expected="0 failed", actual=f"{len(failed)} failed",
            detail="; ".join(f"{o.file_name}: {o.error_code}" for o in failed)
            or "every discovered file was processed",
        ))
        changed = [p.name for p in files if hashes_before[p] is None or sha256_file(p) != hashes_before[p]]
        checks.append(CheckResult(
            "BATCH", "SOURCE_FILES_UNCHANGED", CheckStatus.FAIL if changed else CheckStatus.PASS,
            expected="all hashes identical", actual=f"{len(changed)} changed",
            detail=", ".join(changed) or "SHA-256 of every input file identical before and after the run",
        ))
        result.checks = checks

        result.report_path = run_dir / f"sales_report_{stem}.xlsx"
        result.audit_xlsx_path = run_dir / f"audit_{stem}.xlsx"
        result.audit_json_path = run_dir / f"audit_{stem}.json"
        write_report(result, result.report_path, result.audit_status)
        result.checks.append(verify_report(result.report_path, result))
        result.finished_at = datetime.now()
        write_audit_workbook(result, result.audit_xlsx_path)
        write_audit_json(result, result.audit_json_path)
        logger.info("run finished", extra={
            "event": "run_end", "audit_status": result.audit_status.value,
            "files_ok": len(result.successful), "files_failed": len(result.failed),
            "grand_total": {k: fmt_amount(getattr(result.grand_total, k)) for k in
                            ("sales_yesterday", "sales_previous", "returns_yesterday", "returns_previous")},
        })
    finally:
        if handler is not None:
            detach(handler)
    return result
