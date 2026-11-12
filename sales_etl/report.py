"""Excel report, audit workbook and machine-readable audit JSON.

All values are computed and validated in Python before being written; the
workbooks contain plain numeric cells, no formulas.
"""

from __future__ import annotations

import json
from collections.abc import Iterable, Sequence
from dataclasses import asdict
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from sales_etl import __version__
from sales_etl.config import (
    COL_CUSTOMER,
    COL_DATE,
    COL_DOC_NO,
    COL_DOC_TYPE,
    COL_NET,
    CUSTOMER_COUNT_KEYS,
    GRAND_TOTAL_LABEL,
    METRIC_COLUMNS,
    METRIC_KEYS,
    REPORT_COLUMNS,
    REPORT_FILE_COLUMN,
    TOTAL_KEYS,
)
from sales_etl.dates import describe_date, format_dmy
from sales_etl.models import Category, CheckResult, CheckStatus, FileMetrics, Period, PipelineResult

REPORT_SHEET = "التقرير"
METADATA_SHEET = "بيانات التقرير"
MONEY_FORMAT = "#,##0.00;[Red]-#,##0.00;0.00"
COUNT_FORMAT = "#,##0"
DATE_FORMAT = "dd/mm/yyyy"
_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_THIN = Side(style="thin", color="A6A6A6")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_STATUS_FILLS = {
    CheckStatus.PASS.value: PatternFill("solid", fgColor="C6EFCE"),
    CheckStatus.WARN.value: PatternFill("solid", fgColor="FFEB9C"),
    CheckStatus.FAIL.value: PatternFill("solid", fgColor="FFC7CE"),
}
_INCLUDED_IN = {
    (Category.SALES, Period.YESTERDAY): METRIC_COLUMNS["sales_yesterday"],
    (Category.SALES, Period.PREVIOUS): METRIC_COLUMNS["sales_previous"],
    (Category.RETURNS, Period.YESTERDAY): METRIC_COLUMNS["returns_yesterday"],
    (Category.RETURNS, Period.PREVIOUS): METRIC_COLUMNS["returns_previous"],
}


def _ensure_new(path: Path) -> Path:
    if path.exists():
        raise FileExistsError(f"refusing to overwrite existing file: {path}")
    return path


def _table(
    ws: Worksheet,
    headers: Sequence[str],
    rows: Iterable[Sequence[Any]],
    *,
    widths: Sequence[int] = (),
    status_col: int | None = None,
) -> int:
    """Write a styled header and rows; Decimal -> money format, date -> date format.

    Returns the index of the last row written.
    """
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.font, cell.fill, cell.border = _HEADER_FONT, _HEADER_FILL, _BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    last = 1
    for r, values in enumerate(rows, start=2):
        last = r
        for c, value in enumerate(values, start=1):
            if isinstance(value, CheckStatus):
                value = value.value
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = _BORDER
            if isinstance(value, Decimal):
                cell.number_format = MONEY_FORMAT
            elif isinstance(value, date):
                cell.number_format = DATE_FORMAT
                cell.alignment = Alignment(horizontal="center")
            if c == status_col and value in _STATUS_FILLS:
                cell.fill = _STATUS_FILLS[value]
                cell.font = Font(bold=True)
    for c, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.freeze_panes = "A2"
    return last


# ------------------------------------------------------------------ business report
def report_rows(result: PipelineResult) -> list[tuple[str, list[Decimal]]]:
    """``(file name, [4 metric values])`` for each successfully processed file."""
    return [(o.file_name, [o.metrics.value(k) for k in METRIC_KEYS])  # type: ignore[union-attr]
            for o in result.successful]


def report_totals(metrics: Sequence[Decimal]) -> list[Decimal]:
    """The three leading report columns, from the four metrics (in ``METRIC_KEYS`` order).

    إجمالي المبيعات   = إجمالي مبيعات أمس + إجمالي مبيعات الأيام السابقة
    إجمالي المرتجعات  = إجمالي مرتجعات أمس + إجمالي مرتجعات الأيام السابقة  (sign kept)
    الصافي            = إجمالي المبيعات + إجمالي المرتجعات
    """
    value = dict(zip(METRIC_KEYS, metrics, strict=True))
    sales_total = value["sales_yesterday"] + value["sales_previous"]
    returns_total = value["returns_yesterday"] + value["returns_previous"]
    return [sales_total, returns_total, sales_total + returns_total]


def report_table(result: PipelineResult) -> list[list[Any]]:
    """Every report row in ``REPORT_COLUMNS`` order; the last one is the Grand Total.

    The customer counts come last; the Grand Total ones are counts of the union of the
    per-file code sets, never the sum of the file rows.
    """
    counts = [o.customers.counts() for o in result.successful]  # type: ignore[union-attr]
    lines = [*report_rows(result),
             (GRAND_TOTAL_LABEL, [result.grand_total.value(k) for k in METRIC_KEYS])]
    counts.append(result.grand_customers.counts())
    return [[*report_totals(metrics), label, *metrics, *n] for (label, metrics), n in zip(lines, counts, strict=True)]


def write_report(result: PipelineResult, path: Path, audit_status: CheckStatus) -> Path:
    """Three total columns + the five business columns; one row per file, then Grand Total."""
    _ensure_new(path)
    wb = Workbook()
    ws = wb.active
    ws.title = REPORT_SHEET
    ws.sheet_view.rightToLeft = True
    *file_rows, grand_row = report_table(result)
    last = _table(ws, REPORT_COLUMNS, file_rows, widths=[20, 20, 20, 34, 22, 28, 22, 28, 16, 22, 22])
    first_count = len(REPORT_COLUMNS) - len(CUSTOMER_COUNT_KEYS) + 1
    for row in ws.iter_rows(min_row=2, max_row=last, min_col=first_count):
        for cell in row:
            cell.number_format = COUNT_FORMAT
    for col, value in enumerate(grand_row, start=1):
        cell = ws.cell(row=last + 1, column=col, value=value)
        cell.font, cell.fill = Font(bold=True), _TOTAL_FILL
        cell.border = Border(left=_THIN, right=_THIN, top=Side(style="double"), bottom=_THIN)
        if isinstance(value, Decimal):
            cell.number_format = MONEY_FORMAT
        elif isinstance(value, int):
            cell.number_format = COUNT_FORMAT
    ws.row_dimensions[1].height = 32
    ws.page_setup.orientation = "landscape"
    ws.oddFooter.center.text = f"Reference Yesterday Date: {format_dmy(result.reference_date)}"

    ref = result.reference_date
    entries: list[tuple[str, Any]] = [
        ("تاريخ الأمس المرجعي (Reference Yesterday Date)", ref),
        ("التاريخ بصيغة واضحة", describe_date(ref)),
        ("مبيعات / مرتجعات أمس", f"{COL_DATE} = {format_dmy(ref)}"),
        ("الأيام السابقة", f"{COL_DATE} < {format_dmy(ref)} (لا يشمل تاريخ الأمس)"),
        ("المستبعد", f"كل الحركات بتاريخ بعد {format_dmy(ref)}"),
        ("أساس الحساب", f"عمود {COL_NET} فقط، بالإشارة الأصلية (المرتجعات سالبة كما في المصدر)"),
        ("وقت التشغيل", result.started_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("مجلد الملفات", str(result.input_dir)),
        ("عدد الملفات المكتشفة", len(result.discovered)),
        ("ملفات تمت معالجتها بنجاح", len(result.successful)),
        ("ملفات فشلت (غير موجودة في التقرير)", len(result.failed)),
        ("حالة المراجعة (Reconciliation audit)", audit_status.value),
        ("ملف المراجعة", result.audit_xlsx_path.name if result.audit_xlsx_path else ""),
        ("إصدار البرنامج", f"sales_etl {__version__}"),
    ]
    entries += [(f"ملف فاشل: {o.file_name}", f"{o.error_code}: {o.error_message}") for o in result.failed]
    meta = wb.create_sheet(METADATA_SHEET)
    meta.sheet_view.rightToLeft = True
    _table(meta, ["البند", "القيمة"], entries, widths=[46, 90], status_col=2)
    wb.save(path)
    return path


def verify_report(path: Path, result: PipelineResult) -> CheckResult:
    """Read the saved workbook back and compare every cell with the computed values."""
    tolerance = Decimal("0.005")  # Excel stores IEEE doubles
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        values = [list(r) for r in wb[REPORT_SHEET].iter_rows(values_only=True)]
        wb.close()
    except Exception as exc:  # noqa: BLE001 - any failure here is a FAIL result
        return CheckResult("REPORT", "REPORT_READBACK", CheckStatus.FAIL, detail=f"cannot read report: {exc}")
    expected = report_table(result)
    problems: list[str] = []
    if not values or tuple(values[0]) != REPORT_COLUMNS:
        problems.append(f"header mismatch: {values[0] if values else None}")
    body = values[1:]
    if len(body) != len(expected):
        problems.append(f"{len(body)} rows written, {len(expected)} expected")
    label_col = REPORT_COLUMNS.index(REPORT_FILE_COLUMN)
    for row, wanted in zip(body, expected):
        name = wanted[label_col]
        if len(row) < len(wanted):
            problems.append(f"{name}: only {len(row)} cells written")
            continue
        for got, want in zip(row, wanted):
            if isinstance(want, Decimal):
                if not isinstance(got, (int, float)) or abs(Decimal(str(got)) - want) > tolerance:
                    problems.append(f"{name}: wrote {got!r}, expected {want}")
            elif isinstance(want, int):  # customer counts: exact
                if type(got) is not int or got != want:
                    problems.append(f"{name}: wrote {got!r}, expected {want} customers")
            elif got != want:
                problems.append(f"row '{got}' != '{want}'")
    return CheckResult(
        "REPORT", "REPORT_READBACK", CheckStatus.FAIL if problems else CheckStatus.PASS,
        expected=f"{len(expected)} rows incl. Grand Total", actual=f"{len(body)} rows",
        detail="; ".join(problems[:10]) or "every report cell matches the computed Decimal values",
    )


# ------------------------------------------------------------------ audit outputs
def write_audit_workbook(result: PipelineResult, path: Path) -> Path:
    _ensure_new(path)
    wb = Workbook()
    gt = result.grand_total
    summary = wb.active
    summary.title = "Summary"
    _table(summary, ["Item", "Value"], [
        ("AUDIT STATUS", result.audit_status),
        ("Reference Yesterday Date", result.reference_date),
        ("Reference date (unambiguous)", describe_date(result.reference_date)),
        ("Input folder", str(result.input_dir)),
        ("Files discovered", len(result.discovered)),
        ("Files processed successfully", len(result.successful)),
        ("Files failed", len(result.failed)),
        ("Files skipped (not eligible)", len(result.skipped)),
        ("Checks PASS", result.count_checks(CheckStatus.PASS)),
        ("Checks WARN", result.count_checks(CheckStatus.WARN)),
        ("Checks FAIL", result.count_checks(CheckStatus.FAIL)),
        ("Checks N/A", result.count_checks(CheckStatus.NOT_APPLICABLE)),
        *[(f"Grand Total - {METRIC_COLUMNS[k]}", gt.value(k)) for k in METRIC_KEYS],
        ("Future-dated transactions excluded (rows)", gt.future_count),
        ("Future-dated صافى excluded", gt.future_net),
        ("Other document types excluded (rows)", gt.other_type_count),
        ("Other document types صافى excluded", gt.other_type_net),
        ("Generated by", f"sales_etl {__version__}"),
    ], widths=[44, 70], status_col=2)

    _table(wb.create_sheet("Checks"),
           ["Scope", "Check", "Status", "Expected", "Actual", "Difference", "Detail"],
           [(c.scope, c.check, c.status, c.expected, c.actual, c.difference, c.detail) for c in result.checks],
           widths=[30, 44, 9, 26, 26, 14, 100], status_col=3)

    file_rows = []
    for o in result.outcomes:
        p, m = o.parsed, o.metrics
        file_rows.append((
            o.file_name, "SUCCESS" if o.succeeded else "FAILED", o.error_code or "", o.error_message or "",
            p and p.encoding, p and repr(p.delimiter), p and p.header_line_no,
            p and p.period_from, p and p.period_to, p and p.candidate_rows,
            p and len(p.transactions), p and len(p.rejected), p and p.blank_rows,
            m and m.sales_yesterday_count, m and m.sales_previous_count,
            m and m.returns_yesterday_count, m and m.returns_previous_count,
            m and m.future_count, m and m.future_net, m and m.other_type_count,
            p and p.control.doc_count, p and p.control.net_total, p and p.sha256,
        ))
    _table(wb.create_sheet("Files"), [
        "File", "Status", "Error code", "Error message", "Encoding", "Delimiter", "Header line",
        "Period from", "Period to", "Transaction rows", "Valid rows", "Rejected rows", "Blank rows",
        "Sales yesterday rows", "Sales previous rows", "Returns yesterday rows", "Returns previous rows",
        "Future rows excluded", "Future صافى excluded", "Other-type rows",
        "ERP doc count", "ERP صافى total", "SHA-256",
    ], file_rows, widths=[30, 10, 26, 50, 10, 9, 9, 12, 12, 11, 10, 10, 9, 11, 11, 11, 11, 11, 16, 10, 10, 16, 66])

    issue_rows = [(i.file_name, i.line_no, i.severity.value, i.code, i.message, i.raw)
                  for o in result.outcomes if o.parsed for i in o.parsed.issues]
    issue_rows += [(o.file_name, None, "ERROR", o.error_code, o.error_message, "") for o in result.failed]
    issue_rows += [(str(p), None, "INFO", "FILE_SKIPPED", reason, "") for p, reason in result.skipped]
    _table(wb.create_sheet("Issues"), ["File", "Line", "Severity", "Code", "Message", "Raw row"],
           issue_rows, widths=[30, 7, 10, 34, 90, 90])

    ledger_rows = []
    for o in result.successful:
        for t in o.parsed.transactions:  # type: ignore[union-attr]
            included = _INCLUDED_IN.get((t.category, t.period)) or (
                "مستبعد: تاريخ لاحق" if t.period is Period.FUTURE else "مستبعد: نوع مستند آخر")
            ledger_rows.append((t.source_file, t.line_no, t.txn_date, t.doc_type, t.doc_no,
                                t.customer, t.net, t.period.value, included))
    ledger = wb.create_sheet("Transactions")
    _table(ledger, ["اسم الملف", "Line", COL_DATE, COL_DOC_TYPE, COL_DOC_NO, COL_CUSTOMER, COL_NET,
                    "Period", "Included in"], ledger_rows, widths=[30, 7, 12, 12, 12, 12, 16, 12, 32])
    if ledger_rows:
        ledger.auto_filter.ref = f"A1:I{len(ledger_rows) + 1}"

    _table(wb.create_sheet("Rejected rows"), ["File", "Line", "Reasons", "Raw row"],
           [(r.source_file, r.line_no, "; ".join(r.reasons), r.raw)
            for o in result.successful for r in o.parsed.rejected],  # type: ignore[union-attr]
           widths=[30, 7, 80, 120])
    wb.save(path)
    return path


def _metrics_dict(m: FileMetrics | None) -> dict[str, Any] | None:
    return None if m is None else asdict(m)


def _json_default(value: Any) -> Any:
    if isinstance(value, Decimal):
        return str(value)  # exact; never a float
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, CheckStatus):
        return value.value
    raise TypeError(f"not serialisable: {type(value)!r}")


def write_audit_json(result: PipelineResult, path: Path) -> Path:
    _ensure_new(path)
    payload = {
        "audit_status": result.audit_status.value,
        "reference_date": result.reference_date,
        "reference_date_description": describe_date(result.reference_date),
        "input_dir": result.input_dir,
        "run_dir": result.run_dir,
        "started_at": result.started_at,
        "finished_at": result.finished_at,
        "files_discovered": [p.name for p in result.discovered],
        "files_skipped": [{"path": p, "reason": r} for p, r in result.skipped],
        "report": {
            "columns": list(REPORT_COLUMNS),
            "rows": [{"file": n, **dict(zip(TOTAL_KEYS, report_totals(vals))), **dict(zip(METRIC_KEYS, vals)),
                      **dict(zip(CUSTOMER_COUNT_KEYS, o.customers.counts()))}  # type: ignore[union-attr]
                     for (n, vals), o in zip(report_rows(result), result.successful, strict=True)],
            "grand_total": _metrics_dict(result.grand_total),
            "grand_total_columns": dict(zip(
                TOTAL_KEYS, report_totals([result.grand_total.value(k) for k in METRIC_KEYS]))),
            "grand_total_customer_counts": dict(zip(CUSTOMER_COUNT_KEYS, result.grand_customers.counts())),
        },
        "files": [{
            "file": o.file_name,
            "status": "SUCCESS" if o.succeeded else "FAILED",
            "error_code": o.error_code,
            "error_message": o.error_message,
            "encoding": o.parsed and o.parsed.encoding,
            "encoding_reason": o.parsed and o.parsed.encoding_reason,
            "delimiter": o.parsed and o.parsed.delimiter,
            "header_line": o.parsed and o.parsed.header_line_no,
            "transaction_rows": o.parsed and o.parsed.candidate_rows,
            "valid_rows": o.parsed and len(o.parsed.transactions),
            "rejected_rows": [asdict(r) for r in o.parsed.rejected] if o.parsed else [],
            "metrics": _metrics_dict(o.metrics),
            "erp_control": asdict(o.parsed.control) if o.parsed else None,
            "sha256": o.parsed and o.parsed.sha256,
            "issues": [{"line": i.line_no, "severity": i.severity.value, "code": i.code,
                        "message": i.message} for i in (o.parsed.issues if o.parsed else [])],
        } for o in result.outcomes],
        "checks": [asdict(c) | {"status": c.status.value} for c in result.checks],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=_json_default),
                    encoding="utf-8")
    return path
