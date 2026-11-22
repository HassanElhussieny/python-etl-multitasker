"""Command-line entry point.

    python -m sales_etl                                                # shows the task menu
    python -m sales_etl --ref-date 18/10/2026                          # task 1
    python -m sales_etl --start-date 07-10-2026 --end-date 12-10-2026  # task 2

Task 1 (the original behaviour) is described below; task 2 (``period.py``) only reads
the input files and never moves them.

By default every CSV directly inside the project's ``input`` folder is processed
and, once its results are saved and verified, moved to ``processed`` (or
``failures``). ``--no-move`` gives the read-only behaviour.
The reference date is asked for exactly ONCE and applied to every file.
"""

from __future__ import annotations

import argparse
import sys
from collections.abc import Callable
from datetime import date
from decimal import Decimal
from pathlib import Path

from sales_etl import __version__
from sales_etl import analysis, fixed_customers
from sales_etl.amounts import fmt_amount
from sales_etl.config import (
    DEFAULT_INPUT_DIR,
    DEFAULT_OUTPUT_DIR,
    METRIC_COLUMNS,
    METRIC_KEYS,
    PROJECT_ROOT,
    ConfigError,
    FileManagementConfig,
    PipelineConfig,
)
from sales_etl.dates import DateParseError, describe_date, parse_reference_date
from sales_etl.filestore import RunLockedError
from sales_etl.logging_setup import configure_console
from sales_etl.managed import (
    COMPLETED_WITH_PROBLEMS,
    FAILED_RUN,
    NO_INPUT_FILES,
    SUCCEEDED,
    ManagedRunResult,
    run_managed,
)
from sales_etl.models import CheckStatus, PipelineResult
from sales_etl import monthly_customer
from sales_etl import monthly_file_customers
from sales_etl.customer import CUSTOMER
from sales_etl.grouped import GroupResult, GroupRunResult, GroupSpec, group_table, run_group_managed, run_group_totals
from sales_etl.period import PeriodResult, PeriodRunResult, run_period_managed, run_period_totals
from sales_etl.pipeline import run_pipeline
from sales_etl.representative import REPRESENTATIVE

EXIT_OK, EXIT_AUDIT_FAILED, EXIT_USAGE, EXIT_NO_INPUT = 0, 1, 2, 3
DATE_PROMPT = "اكتب تاريخ أمس (DD/MM/YYYY): "
TASK_DAILY, TASK_PERIOD, TASK_CUSTOMER, TASK_REPRESENTATIVE, TASK_ANALYSIS = "1", "2", "3", "4", "5"
TASK_MONTHLY = "6"
TASK_FILE_CUSTOMERS = "7"
TASKS: dict[str, str] = {
    TASK_DAILY: "Daily",
    TASK_PERIOD: "Period",
    TASK_CUSTOMER: "Customer",
    TASK_REPRESENTATIVE: "Representative",
    TASK_ANALYSIS: "Sector & Representative Sales Analysis",
    TASK_MONTHLY: "صافي مبيعات العملاء الشهري",
    TASK_FILE_CUSTOMERS: "عدد العملاء الشهري لكل ملف",
}
TASK_PROMPT = "Select task number: "
REPORT_TYPES: dict[str, str] = {"1": "Yesterday & Previous Days", "2": "All Transactions", "3": "Receivables Update"}
REPORT_RECEIVABLES = "3"  # task 5 option 3: an independent workflow (sales_etl.receivables)
REPORT_TYPE_PROMPT = "Select report type: "
YESTERDAY_PROMPT = "Enter yesterday's date (DD/MM/YYYY): "
START_DATE_PROMPT = "اكتب تاريخ البداية (DD-MM-YYYY): "
END_DATE_PROMPT = "اكتب تاريخ النهاية (DD-MM-YYYY): "


#: task 5 option 1: the folder of the current collection workbook (report 3's المستحق columns)
DEFAULT_COLLECTION_DIR = PROJECT_ROOT / fixed_customers.COLLECTION_DIR_NAME

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sales_etl",
        description="Daily sales & returns ETL: one reference date across all CSV files.",
    )
    parser.add_argument("--task", choices=tuple(TASKS),
                        help="1 = yesterday vs previous days, 2 = totals for a period, 3 = totals per customer, 4 = totals per representative, 5 = sector & representative analysis, 6 = monthly net per customer, 7 = monthly distinct customers per file "
                        "(default: ask; implied by --ref-date or --start-date/--end-date)")
    parser.add_argument("--input-dir", type=Path,
                        help=f"folder containing the CSV files (default: {DEFAULT_INPUT_DIR})")
    parser.add_argument("--ref-date", help="task 1: Reference Yesterday Date, DD/MM/YYYY (day first)")
    parser.add_argument("--start-date", help="task 2: first day of the period, DD-MM-YYYY (included)")
    parser.add_argument("--end-date", help="task 2: last day of the period, DD-MM-YYYY (included)")
    parser.add_argument("--report-type", choices=tuple(REPORT_TYPES),
                        help="task 5: 1 = yesterday & previous days (needs --ref-date), 2 = all transactions, "
                             "3 = receivables update (needs --sales-csv and --receivables)")
    parser.add_argument("--sales-csv", type=Path, help="task 5 option 3: the Sales CSV file")
    parser.add_argument("--receivables", type=Path,
                        help="task 5 option 3: the receivables workbook (كود العميل in column B, المستحق 7-10, "
                             "المستحق الجديد)")
    parser.add_argument("--collection-dir", type=Path,
                        help="task 5 option 1: folder of the current collection workbook (default: "
                             f"{fixed_customers.COLLECTION_DIR_NAME} in the project root)")
    parser.add_argument("--main-data", type=Path,
                        help=f"task 5: Main Data workbook (default: the single .xlsx in {analysis.DEFAULT_MAIN_DATA_DIR})")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR,
                        help=f"where run folders are created (default: {DEFAULT_OUTPUT_DIR})")
    parser.add_argument("--no-move", action="store_true",
                        help="read-only mode: never move input files and keep no processing journal")
    for name in ("processed", "failures", "state"):
        parser.add_argument(f"--{name}-dir", type=Path,
                            help=f"override the '{name}' folder (default: next to the input folder)")
    parser.add_argument("--encoding", help="force a source encoding (default: auto-detect)")
    parser.add_argument("--delimiter", help="force a delimiter, e.g. ';' or 'tab' (default: auto-detect)")
    parser.add_argument("--recursive", action="store_true", help="also scan sub-folders (--no-move only)")
    parser.add_argument("--non-interactive", action="store_true",
                        help="never prompt; fail if an argument is missing")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def ask_task(input_fn: Callable[[str], str] = input) -> str:
    """Show the task menu and return the chosen task number; ask again while invalid."""
    print("Available Tasks:")
    for number, name in TASKS.items():
        print(f"{number}. {name}")
    while True:
        try:
            choice = input_fn(TASK_PROMPT).strip()
        except (EOFError, KeyboardInterrupt) as exc:
            raise ConfigError("Cancelled: no task chosen") from exc
        if choice in TASKS:
            print(f"Selected task: {choice}. {TASKS[choice]}")
            return choice
        print(f"  Invalid choice: '{choice}' - enter {' or '.join(TASKS)}.")


def choose_task(args: argparse.Namespace, interactive: bool, input_fn: Callable[[str], str]) -> str:
    """The task from ``--task``, else implied by the date arguments, else the menu."""
    period_args = args.start_date is not None or args.end_date is not None
    if args.ref_date is not None and period_args:
        raise ConfigError("use either --ref-date (task 1) or --start-date/--end-date (task 2), not both")
    if (args.task == TASK_DAILY and period_args) or (args.task == TASK_PERIOD and args.ref_date is not None):
        raise ConfigError("--ref-date belongs to task 1; --start-date/--end-date belong to task 2")
    if args.task in (TASK_CUSTOMER, TASK_REPRESENTATIVE, TASK_MONTHLY, TASK_FILE_CUSTOMERS) and (period_args or args.ref_date is not None):
        raise ConfigError(f"task {args.task} ({TASKS[args.task]}) takes no dates; "
                          "it uses every transaction in the file")
    if args.task == TASK_ANALYSIS and period_args:
        raise ConfigError("task 5 takes --ref-date (report type 1) or no date (report type 2)")
    receivables_args = args.sales_csv is not None or args.receivables is not None
    if (args.report_type is not None or args.main_data is not None or receivables_args) \
            and args.task not in (None, TASK_ANALYSIS):
        raise ConfigError("--report-type, --main-data, --sales-csv and --receivables belong to task 5")
    if args.task is None and (args.report_type is not None or receivables_args):
        return TASK_ANALYSIS
    if args.task:
        return args.task
    if period_args:
        return TASK_PERIOD
    if args.ref_date is not None or not interactive:
        return TASK_DAILY
    return ask_task(input_fn)


def ask_reference_date(input_fn: Callable[[str], str] = input, prompt: str = DATE_PROMPT) -> date:
    """Ask once for the reference date; ask again only while the entry is invalid.

    Ctrl+C or end-of-input cancels the run (nothing is processed).
    """
    while True:
        try:
            text = input_fn(prompt)
        except (EOFError, KeyboardInterrupt) as exc:
            raise ConfigError("تم الإلغاء: لم يتم إدخال تاريخ (no date entered)") from exc
        try:
            return parse_reference_date(text)
        except DateParseError as exc:
            print(f"  التاريخ غير صحيح: '{text.strip()}' ({exc})")
            print("  اكتبه بالشكل يوم/شهر/سنة، مثال: 18/10/2026 - حاول تاني.")


def _delimiter(value: str | None) -> str | None:
    if value is None:
        return None
    return {"tab": "\t", "\\t": "\t", "semicolon": ";", "comma": ","}.get(value.lower(), value)


def print_summary(result: PipelineResult) -> None:
    status = result.audit_status
    print(f"Reference Yesterday Date : {describe_date(result.reference_date)}")
    print(f"Input folder             : {result.input_dir}")
    print(f"Files processed this run : {len(result.discovered)}")
    print(f"Processed successfully   : {len(result.successful)}")
    print(f"Failed                   : {len(result.failed)}")
    for outcome in result.failed:
        print(f"   FAILED  {outcome.file_name}  ->  {outcome.error_code}: {outcome.error_message}")
    print("-" * 72)
    for key in METRIC_KEYS:
        print(f"Grand Total | {METRIC_COLUMNS[key]:<32} {fmt_amount(result.grand_total.value(key)):>18}")
    print("-" * 72)
    print(f"Reconciliation audit     : {status.value}  "
          f"(PASS {result.count_checks(CheckStatus.PASS)}, WARN {result.count_checks(CheckStatus.WARN)}, "
          f"FAIL {result.count_checks(CheckStatus.FAIL)})")
    for check in result.checks:
        if check.status is CheckStatus.FAIL:
            print(f"   FAIL  [{check.scope}] {check.check}: {check.detail or check.actual}")
    print(f"Report                   : {result.report_path}")
    print(f"Audit workbook           : {result.audit_xlsx_path}")
    print(f"Audit JSON               : {result.audit_json_path}")
    print(f"Log                      : {result.log_path}")


_STATUS_TEXT = {
    SUCCEEDED: "SUCCEEDED - report generated, audit PASS, all files archived",
    COMPLETED_WITH_PROBLEMS: "COMPLETED WITH PROBLEMS - report generated, see FAIL lines / file dispositions",
    FAILED_RUN: "FAILED - no verified report was produced",
    NO_INPUT_FILES: "NO REPORT - there were no CSV files in the input folder",
}


def print_managed_summary(managed: ManagedRunResult) -> None:
    line = "=" * 72
    print(line)
    for action in managed.recovery:
        print(f"RECOVERY {action.action}: {action.run_id} {action.file_name} - {action.detail}")
    for notice in managed.notices:
        print(f"NOTICE   {notice}")
    if managed.nothing_to_do:
        print(f"RUN RESULT               : {_STATUS_TEXT[NO_INPUT_FILES]}")
        print(f"Journal                  : {managed.journal_path}")
        print(line)
        return
    if managed.execution_time:
        print(f"Execution date           : {managed.execution_time:%Y-%m-%d %H:%M:%S} (folder names)")
    if managed.pipeline:
        print_summary(managed.pipeline)
    for problem in managed.output_problems:
        print(f"OUTPUT PROBLEM: {problem}")
    if managed.error:
        print(f"ERROR: {managed.error}")
    print("-" * 72)
    print(f"Report folder            : {managed.run_dir}")
    print("Generated files:")
    outputs = managed.expected_outputs()
    for name, path in outputs.items():
        state = "OK" if path and path.is_file() else "NOT GENERATED"
        print(f"   [{state}] {name:<11} {path or '-'}")
    if managed.manifest_path:
        print(f"   [OK] {'manifest':<11} {managed.manifest_path}")
    print("-" * 72)
    print(f"Processed archive folder : {managed.processed_run_dir}")
    print("File dispositions:")
    for d in managed.dispositions:
        state = "moved" if d.moved else ("NOT MOVED: " + d.error if d.error else "left in place")
        print(f"   {d.disposition:<9} {d.file_name}  ->  {d.location}  [{state}]")
        print(f"             {d.reason}")
    print(f"Journal                  : {managed.journal_path}")
    print(f"RUN RESULT               : {_STATUS_TEXT[managed.status]}")
    print(line)


def print_period_run_summary(run: PeriodRunResult, start: date, end: date, input_dir: Path) -> None:
    line = "=" * 72
    for action in run.recovery:
        print(f"RECOVERY {action.action}: {action.run_id} {action.file_name} - {action.detail}")
    if run.period is not None:
        print_period_summary(run.period, moved=True)
    else:
        print(line)
        print(f"Task                     : {TASKS[TASK_PERIOD]}")
        print(f"Start date (included)    : {describe_date(start)}")
        print(f"End date (included)      : {describe_date(end)}")
        print(f"Input folder             : {input_dir}")
    if run.nothing_to_do:
        print("RUN RESULT               : NO REPORT - there were no CSV files in the input folder")
        print(line)
        return
    for problem in run.output_problems:
        print(f"OUTPUT PROBLEM: {problem}")
    print(f"Processed archive folder : {run.processed_run_dir}")
    print("File dispositions:")
    for d in run.dispositions:
        state = "moved" if d.moved else ("NOT MOVED: " + d.error if d.error else "left in place")
        print(f"   {d.disposition:<9} {d.file_name}  ->  {d.location}  [{state}]")
        print(f"             {d.reason}")
    print(f"Journal                  : {run.journal_path}")
    if run.succeeded:
        state = "SUCCEEDED - report generated, all files archived"
    else:
        state = "COMPLETED WITH PROBLEMS - see FAILED / WARNING lines and file dispositions"
    print(f"RUN RESULT               : {state}")
    print(line)


def print_period_summary(result: PeriodResult, *, moved: bool = False) -> None:
    line = "=" * 72
    print(line)
    print(f"Task                     : {TASKS[TASK_PERIOD]}")
    print(f"Start date (included)    : {describe_date(result.start_date)}")
    print(f"End date (included)      : {describe_date(result.end_date)}")
    print(f"Input folder             : {result.input_dir}")
    if not result.files:
        print("RUN RESULT               : NO REPORT - there were no CSV files in the input folder")
        print(line)
        return
    print(f"Files read               : {len(result.successful)} of {len(result.files)}")
    for f in result.failed:
        print(f"   FAILED  {f.file_name}  ->  {f.error}")
    for f in result.successful:
        if f.rejected_rows:
            print(f"   WARNING {f.file_name}: {f.rejected_rows} rejected row(s) not counted")
    print("-" * 72)
    for f in result.successful:
        print(f"{f.file_name:<20} | {fmt_amount(f.sales):>16} | {fmt_amount(f.returns):>16} | "
              f"{fmt_amount(f.net):>16}")
    print("-" * 72)
    print(f"Grand Total | إجمالي المبيعات   {fmt_amount(result.sales):>18}")
    print(f"Grand Total | إجمالي المرتجعات  {fmt_amount(result.returns):>18}")
    print(f"Grand Total | الصافي            {fmt_amount(result.net):>18}")
    print("-" * 72)
    print(f"Report                   : {result.report_path}")
    if moved:
        return
    print("Input files were not moved (--no-move).")
    state = "OK" if result.complete else "COMPLETED WITH PROBLEMS - see FAILED / WARNING lines"
    print(f"RUN RESULT               : {state}")
    print(line)


def main(argv: list[str] | None = None, input_fn: Callable[[str], str] = input) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            pass
    args = build_parser().parse_args(argv)
    configure_console()
    interactive = not args.non_interactive
    try:
        task = choose_task(args, interactive, input_fn)
        input_dir = args.input_dir
        if input_dir is None and DEFAULT_INPUT_DIR.is_dir():
            input_dir = DEFAULT_INPUT_DIR
        if input_dir is None:
            if not interactive:
                raise ConfigError("--input-dir is required in non-interactive mode")
            input_dir = Path(input_fn("Folder containing the CSV files: ").strip().strip('"'))
    except ConfigError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    if task == TASK_PERIOD:
        return run_period_task(args, input_dir, interactive, input_fn)
    if task == TASK_CUSTOMER:
        return run_group_task(CUSTOMER, args, input_dir)
    if task == TASK_REPRESENTATIVE:
        return run_group_task(REPRESENTATIVE, args, input_dir)
    if task == TASK_ANALYSIS:
        return run_analysis_task(args, input_dir, interactive, input_fn)
    if task == TASK_MONTHLY:
        return run_monthly_task(args, input_dir)
    if task == TASK_FILE_CUSTOMERS:
        return run_file_customers_task(args, input_dir)
    return run_daily_task(args, input_dir, interactive, input_fn)


def print_group_summary(result: GroupResult, *, moved: bool = False) -> None:
    """Terminal summary for task 3 (Customer) or task 4 (Representative)."""
    line = "=" * 72
    label = result.spec.task.lower() + "s"
    print(line)
    print(f"Task                     : {result.spec.task}")
    print(f"Input folder             : {result.input_dir}")
    if not result.files:
        print("RUN RESULT               : NO REPORT - there were no CSV files in the input folder")
        print(line)
        return
    print(f"Files read               : {len(result.files) - len(result.failed)} of {len(result.files)}")
    for f in result.failed:
        print(f"   FAILED  {f.file_name}  ->  {f.error}")
    for f in result.files:
        if f.rejected_rows:
            print(f"   WARNING {f.file_name}: {f.rejected_rows} rejected row(s) not counted")
    for code, names in result.name_conflicts.items():
        print(f"   WARNING code {code} has {len(names)} names; the report uses '{names[0]}' "
              f"(also: {', '.join(repr(n) for n in names[1:])})")
    rows = group_table(result)
    print(f"{label.capitalize():<25}: {len(rows)}")
    print("-" * 72)
    print(f"All {label} | إجمالي المبيعات   {fmt_amount(sum((r[2] for r in rows), Decimal(0))):>18}")
    print(f"All {label} | إجمالي المرتجعات  {fmt_amount(sum((r[3] for r in rows), Decimal(0))):>18}")
    print(f"All {label} | الصافي            {fmt_amount(sum((r[4] for r in rows), Decimal(0))):>18}")
    print("-" * 72)
    print(f"Report                   : {result.report_path}")
    if moved:
        return
    print("Input files were not moved (--no-move).")
    state = "OK" if result.complete else "COMPLETED WITH PROBLEMS - see FAILED / WARNING lines"
    print(f"RUN RESULT               : {state}")
    print(line)


def print_group_run_summary(run: GroupRunResult, spec: GroupSpec, input_dir: Path) -> None:
    line = "=" * 72
    for action in run.recovery:
        print(f"RECOVERY {action.action}: {action.run_id} {action.file_name} - {action.detail}")
    if run.result is None:
        print(line)
        print(f"Task                     : {spec.task}")
        print(f"Input folder             : {input_dir}")
        print("RUN RESULT               : NO REPORT - there were no CSV files in the input folder")
        print(line)
        return
    print_group_summary(run.result, moved=True)
    for problem in run.output_problems:
        print(f"OUTPUT PROBLEM: {problem}")
    print(f"Processed archive folder : {run.processed_run_dir}")
    print("File dispositions:")
    for d in run.dispositions:
        state = "moved" if d.moved else ("NOT MOVED: " + d.error if d.error else "left in input")
        print(f"   {d.disposition:<9} {d.file_name}  ->  {d.location}  [{state}]")
        print(f"             {d.reason}")
    print(f"Journal                  : {run.journal_path}")
    state = ("SUCCEEDED - report generated, all files archived" if run.succeeded
             else "COMPLETED WITH PROBLEMS - see FAILED / WARNING lines and file dispositions")
    print(f"RUN RESULT               : {state}")
    print(line)


def run_group_task(spec: GroupSpec, args: argparse.Namespace, input_dir: Path) -> int:
    """Tasks 3 and 4: sales, returns and net per code over every transaction (no date filter).

    The report is saved and verified first; then the cleanly read files move to
    ``processed/<Task>_<stamp>/`` (``--no-move``: read-only).
    """
    try:
        config = PipelineConfig(  # reference_date is not used by tasks 3 and 4
            input_dir=input_dir, output_dir=args.output_dir, reference_date=date.today(),
            forced_encoding=args.encoding, forced_delimiter=_delimiter(args.delimiter),
            recursive=args.recursive,
        ).validated()
        files_config = None
        if not args.no_move:
            if args.recursive:
                raise ConfigError("--recursive is only supported together with --no-move")
            default = FileManagementConfig.beside(config.input_dir)
            files_config = FileManagementConfig(
                processed_dir=args.processed_dir or default.processed_dir,
                failures_dir=args.failures_dir or default.failures_dir,
                state_dir=args.state_dir or default.state_dir,
            ).validated(config)
    except ConfigError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    if files_config is None:
        result = run_group_totals(spec, config)
        print_group_summary(result)
        if not result.files:
            return EXIT_NO_INPUT
        return EXIT_OK if result.complete else EXIT_AUDIT_FAILED

    try:
        run = run_group_managed(spec, config, files_config)
    except RunLockedError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    except Exception as exc:  # noqa: BLE001 - nothing moved: files stay in input
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        print("No file was moved; all input files are still in the input folder.", file=sys.stderr)
        return EXIT_AUDIT_FAILED
    print_group_run_summary(run, spec, config.input_dir)
    if run.nothing_to_do:
        return EXIT_NO_INPUT
    return EXIT_OK if run.succeeded else EXIT_AUDIT_FAILED


def print_monthly_summary(result: monthly_customer.MonthlyResult, problems: list[str], *,
                          moved: bool = False) -> None:
    """Terminal summary for task 6 (monthly net per customer)."""
    line = "=" * 72
    print(line)
    print(f"Task                     : {TASKS[TASK_MONTHLY]}")
    print(f"Input folder             : {result.input_dir}")
    if not result.files:
        print("RUN RESULT               : NO REPORT - there were no CSV files in the input folder")
        print(line)
        return
    print(f"Files read               : {len(result.files) - len(result.failed)} of {len(result.files)}")
    for f in result.failed:
        print(f"   FAILED  {f.file_name}  ->  {f.error}")
    for f in result.files:
        if f.rejected_rows:
            print(f"   WARNING {f.file_name}: {f.rejected_rows} rejected row(s) not counted")
    for where in result.missing_code_rows:
        print(f"   WARNING {where}: sales/returns row without a customer code, not counted")
    if result.other_type_rows:
        print(f"Not sales/returns        : {result.other_type_rows} row(s), not counted")
    for code, names in result.name_conflicts.items():
        print(f"   WARNING code {code} has {len(names)} names; the report uses '{names[0]}' "
              f"(also: {', '.join(repr(n) for n in names[1:])})")
    months = monthly_customer.MONTHS[:result.last_month]
    print(f"Months                   : {months[0]} .. {months[-1]}" if months else "Months                   : -")
    print(f"Customers                : {len(result.customers)}")
    print(f"Transactions counted     : {len(result.ledger)}")
    print("-" * 72)
    print(f"All customers | الصافي    {fmt_amount(sum((r.net for r in result.ledger), Decimal(0))):>18}")
    print("-" * 72)
    print(f"Reconciliation           : {'PASS' if not problems else 'FAIL'}")
    for problem in problems:
        print(f"   FAIL  {problem}")
    print(f"Report                   : {result.report_path}")
    if moved:
        return
    print("Input files were not moved (--no-move).")
    state = "OK" if result.complete and not problems else "COMPLETED WITH PROBLEMS - see FAILED / WARNING lines"
    print(f"RUN RESULT               : {state}")
    print(line)


def run_monthly_task(args: argparse.Namespace, input_dir: Path) -> int:
    """Task 6: net per customer code per month, from يناير to the last month in the data.

    The report is saved and verified first; then the cleanly read files move to
    ``processed/صافي_العملاء_الشهري_<stamp>/`` (``--no-move``: read-only).
    """
    try:
        config = PipelineConfig(  # reference_date is not used by task 6
            input_dir=input_dir, output_dir=args.output_dir, reference_date=date.today(),
            forced_encoding=args.encoding, forced_delimiter=_delimiter(args.delimiter),
            recursive=args.recursive,
        ).validated()
        files_config = None
        if not args.no_move:
            if args.recursive:
                raise ConfigError("--recursive is only supported together with --no-move")
            default = FileManagementConfig.beside(config.input_dir)
            files_config = FileManagementConfig(
                processed_dir=args.processed_dir or default.processed_dir,
                failures_dir=args.failures_dir or default.failures_dir,
                state_dir=args.state_dir or default.state_dir,
            ).validated(config)
    except ConfigError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    if files_config is None:
        result = monthly_customer.run_monthly(config)
        problems = monthly_customer.verify_monthly_report(result) if result.files else []
        print_monthly_summary(result, problems)
        if not result.files:
            return EXIT_NO_INPUT
        return EXIT_OK if result.complete and not problems else EXIT_AUDIT_FAILED

    try:
        run = monthly_customer.run_monthly_managed(config, files_config)
    except RunLockedError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    except Exception as exc:  # noqa: BLE001 - nothing moved: files stay in input
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        print("No file was moved; all input files are still in the input folder.", file=sys.stderr)
        return EXIT_AUDIT_FAILED
    line = "=" * 72
    for action in run.recovery:
        print(f"RECOVERY {action.action}: {action.run_id} {action.file_name} - {action.detail}")
    if run.result is None:
        print(line)
        print(f"Task                     : {TASKS[TASK_MONTHLY]}")
        print(f"Input folder             : {config.input_dir}")
        print("RUN RESULT               : NO REPORT - there were no CSV files in the input folder")
        print(line)
        return EXIT_NO_INPUT
    print_monthly_summary(run.result, run.output_problems, moved=True)
    print(f"Processed archive folder : {run.processed_run_dir}")
    print("File dispositions:")
    for d in run.dispositions:
        state = "moved" if d.moved else ("NOT MOVED: " + d.error if d.error else "left in input")
        print(f"   {d.disposition:<9} {d.file_name}  ->  {d.location}  [{state}]")
        print(f"             {d.reason}")
    print(f"Journal                  : {run.journal_path}")
    state = ("SUCCEEDED - report generated and verified, all files archived" if run.succeeded
             else "COMPLETED WITH PROBLEMS - see FAILED / WARNING lines and file dispositions")
    print(f"RUN RESULT               : {state}")
    print(line)
    return EXIT_OK if run.succeeded else EXIT_AUDIT_FAILED


def print_file_customers_summary(result: monthly_file_customers.FileCustomersResult, problems: list[str], *,
                                 moved: bool = False) -> None:
    """Terminal summary for task 7 (monthly distinct customers per file)."""
    line = "=" * 72
    print(line)
    print(f"Task                     : {TASKS[TASK_FILE_CUSTOMERS]}")
    print(f"Input folder             : {result.input_dir}")
    if not result.files:
        print("RUN RESULT               : NO REPORT - there were no CSV files in the input folder")
        print(line)
        return
    print(f"Files read               : {len(result.files) - len(result.failed)} of {len(result.files)}")
    for f in result.failed:
        print(f"   FAILED  {f.file_name}  ->  {f.error}")
    for f in result.files:
        if f.rejected_rows:
            print(f"   WARNING {f.file_name}: {f.rejected_rows} rejected row(s) not counted")
    for where in result.rejected_lines:
        print(f"   REJECTED {where}")
    months = monthly_file_customers.MONTHS[:result.last_month]
    print(f"Months                   : {months[0]} .. {months[-1]}" if months else "Months                   : -")
    print(f"Transactions counted     : {len(result.ledger)}")
    for row in result.rows:
        if not row.months:
            print(f"   NOTE    {row.item.file_name}: no transactions, all months 0")
    print(f"Reconciliation           : {'PASS' if not problems else 'FAIL'}")
    for problem in problems:
        print(f"   FAIL  {problem}")
    print(f"Report                   : {result.report_path}")
    if moved:
        return
    print("Input files were not moved (--no-move).")
    state = "OK" if result.complete and not problems else "COMPLETED WITH PROBLEMS - see FAILED / WARNING lines"
    print(f"RUN RESULT               : {state}")
    print(line)


def run_file_customers_task(args: argparse.Namespace, input_dir: Path) -> int:
    """Task 7: distinct customer codes per file per month, from يناير to the last month in the data.

    The report is saved and verified first; then the cleanly read files move to
    ``processed/عملاء_الملفات_الشهري_<stamp>/`` (``--no-move``: read-only).
    """
    try:
        config = PipelineConfig(  # reference_date is not used by task 7
            input_dir=input_dir, output_dir=args.output_dir, reference_date=date.today(),
            forced_encoding=args.encoding, forced_delimiter=_delimiter(args.delimiter),
            recursive=args.recursive,
        ).validated()
        files_config = None
        if not args.no_move:
            if args.recursive:
                raise ConfigError("--recursive is only supported together with --no-move")
            default = FileManagementConfig.beside(config.input_dir)
            files_config = FileManagementConfig(
                processed_dir=args.processed_dir or default.processed_dir,
                failures_dir=args.failures_dir or default.failures_dir,
                state_dir=args.state_dir or default.state_dir,
            ).validated(config)
    except ConfigError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    if files_config is None:
        result = monthly_file_customers.run_file_customers(config)
        problems = monthly_file_customers.verify_file_customers_report(result) if result.files else []
        print_file_customers_summary(result, problems)
        if not result.files:
            return EXIT_NO_INPUT
        return EXIT_OK if result.complete and not problems else EXIT_AUDIT_FAILED

    try:
        run = monthly_file_customers.run_file_customers_managed(config, files_config)
    except RunLockedError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    except Exception as exc:  # noqa: BLE001 - nothing moved: files stay in input
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        print("No file was moved; all input files are still in the input folder.", file=sys.stderr)
        return EXIT_AUDIT_FAILED
    line = "=" * 72
    for action in run.recovery:
        print(f"RECOVERY {action.action}: {action.run_id} {action.file_name} - {action.detail}")
    if run.result is None:
        print(line)
        print(f"Task                     : {TASKS[TASK_FILE_CUSTOMERS]}")
        print(f"Input folder             : {config.input_dir}")
        print("RUN RESULT               : NO REPORT - there were no CSV files in the input folder")
        print(line)
        return EXIT_NO_INPUT
    print_file_customers_summary(run.result, run.output_problems, moved=True)
    print(f"Processed archive folder : {run.processed_run_dir}")
    print("File dispositions:")
    for d in run.dispositions:
        state = "moved" if d.moved else ("NOT MOVED: " + d.error if d.error else "left in input")
        print(f"   {d.disposition:<9} {d.file_name}  ->  {d.location}  [{state}]")
        print(f"             {d.reason}")
    print(f"Journal                  : {run.journal_path}")
    state = ("SUCCEEDED - report generated and verified, all files archived" if run.succeeded
             else "COMPLETED WITH PROBLEMS - see FAILED / WARNING lines and file dispositions")
    print(f"RUN RESULT               : {state}")
    print(line)
    return EXIT_OK if run.succeeded else EXIT_AUDIT_FAILED


def ask_report_type(input_fn: Callable[[str], str] = input) -> str:
    """Task 5: choose the report type; ask again while invalid."""
    for number, name in REPORT_TYPES.items():
        print(f"{number}. {name}")
    while True:
        try:
            choice = input_fn(REPORT_TYPE_PROMPT).strip()
        except (EOFError, KeyboardInterrupt) as exc:
            raise ConfigError("Cancelled: no report type chosen") from exc
        if choice in REPORT_TYPES:
            print(f"Selected report type: {choice}. {REPORT_TYPES[choice]}")
            return choice
        print(f"  Invalid choice: '{choice}' - enter {' or '.join(REPORT_TYPES)}.")


def print_analysis_summary(result: analysis.AnalysisResult, problems: list[str]) -> None:
    """Task 5's summary: reports 1-2 (Reconciliation) and report 3 (Customer report check) each pass or fail on
    their own lines; ``problems`` = all of them (as verify_task5 returns)."""
    customer_problems = set(result.customers_problems)
    problems = [p for p in problems if p not in customer_problems]
    line = "=" * 72
    print(line)
    print(f"Task                     : {TASKS[TASK_ANALYSIS]}")
    kind = REPORT_TYPES["1" if result.mode == analysis.MODE_YESTERDAY else "2"]
    print(f"Report type              : {kind}")
    if result.reference_date:
        print(f"Yesterday's date         : {describe_date(result.reference_date)}")
    print(f"Main Data                : {result.main_data.path} ({len(result.main_data.customers)} customers)")
    print(f"Input folder             : {result.input_dir}")
    print(f"Files read               : {len(result.files) - len(result.failed)} of {len(result.files)}")
    for f in result.failed:
        print(f"   FAILED  {f.file_name}  ->  {f.error}")
    for f in result.files:
        if f.rejected_rows:
            print(f"   WARNING {f.file_name}: {f.rejected_rows} rejected row(s) not counted")
    print(f"Transactions counted     : {result.direct.rows}")
    if result.after_date_rows:
        print(f"Excluded (after the date): {result.after_date_rows} row(s)")
    if result.other_type_rows:
        print(f"Not sales/returns        : {result.other_type_rows} row(s), not counted")
    print(f"Unregistered customers   : {len(result.missing)} code(s) not in Main Data "
          f"(row '{analysis.UNREGISTERED}' + sheet '{analysis.SHEET_MISSING}')")
    extra_sectors, extra_reps = result.extra_labels
    for label in extra_sectors:
        print(f"   WARNING Main Data sector '{label}' is not one of the 8 sectors; shown as its own row")
    for label in extra_reps:
        print(f"   WARNING Main Data representative '{label}' is not one of the 17; shown as its own row")
    print("-" * 72)
    print(f"Grand Total | إجمالي المبيعات   {fmt_amount(result.direct.sales):>18}")
    print(f"Grand Total | إجمالي المرتجعات  {fmt_amount(result.direct.returns):>18}")
    print(f"Grand Total | الصافي            {fmt_amount(result.direct.net):>18}")
    print("-" * 72)
    print(f"Reconciliation           : {'PASS' if not problems else 'FAIL'}")
    for problem in problems:
        print(f"   FAIL  {problem}")
    print(f"Report                   : {result.report_path}")
    print(f"Representatives report   : {result.reps_report_path}")
    if result.fixed_customers:
        print("-" * 72)
        print(f"Customer report check    : {'PASS' if not result.customers_problems else 'FAIL'}"
              f"   (report 3, Fixed_Customers_Report{' + report 4, Customers_Not_Worked' if result.mode == analysis.MODE_YESTERDAY else ''})")
        for problem in result.customers_problems:
            print(f"   FAIL  {problem}")
        if result.customers_report_path is not None:
            start, end = fixed_customers.customers_period(result)
            print(f"Customer report          : {result.customers_report_path}")
            print(f"   period {describe_date(start)} - {describe_date(end) if end else 'all of the month'}: "
                  f"{len(result.fixed_customers)} fixed customers + "
                  f"{len(fixed_customers.customers_table(result)) - 1 - len(result.fixed_customers)} outside the list")
        if fixed_customers.collection_enabled(result):
            stats = fixed_customers.collection_stats(result)
            print(f"Collection workbook      : {result.collection.path}  (sheet '{result.collection.file.sheet}', "
                  f"header row {result.collection.file.header_row}, code column B, F = المستحق 7-10, "
                  f"L = المستحق الجديد)")
            print(f"   {stats['matched']} of {stats['customers']} report customers matched by code, "
                  f"{stats['unmatched']} not in the file (cells left empty); {stats['file_customers']} codes in the "
                  f"file ({stats['file_only']} not in the report), duplicate codes {stats['duplicates']}")
        if result.not_worked_path is not None:
            from sales_etl import not_worked

            info = not_worked.counts(result)
            print(f"Not-worked report        : {result.not_worked_path}")
            print(f"   {info['customers']} customers with sales = 0 ({info['with_returns']} with returns, "
                  f"{info['no_returns']} without) on {len(info['per_sheet'])} المسؤول sheets")


def _file_arg_or_ask(value: Path | None, name: str, prompt: str, suffixes: tuple[str, ...], interactive: bool,
                     input_fn: Callable[[str], str]) -> Path:
    """An existing file with one of ``suffixes``: from the argument, else asked (again while invalid)."""
    def check(path: Path) -> str | None:
        if not path.is_file():
            return f"{name}: no such file: {path}"
        if path.suffix.lower() not in suffixes:
            return f"{name}: {path.name} is not a {' / '.join(suffixes)} file"
        return None

    if value is not None:
        problem = check(value)
        if problem:
            raise ConfigError(problem)
        return value
    if not interactive:
        raise ConfigError(f"{name} is required for report type 3 in non-interactive mode")
    while True:
        try:
            text = input_fn(prompt).strip().strip('"').strip("'")
        except (EOFError, KeyboardInterrupt) as exc:
            raise ConfigError(f"Cancelled: no {name} given") from exc
        path = Path(text)
        problem = check(path) if text else f"{name}: enter a file path"
        if problem is None:
            return path
        print(f"  {problem}")


def print_receivables_summary(result: "receivables.ReceivablesResult") -> None:
    """Option 3's own summary (it never shares the lines of report types 1-2)."""
    from sales_etl import receivables

    line = "=" * 72
    st = result.stats
    print(line)
    print(f"Task                     : {TASKS[TASK_ANALYSIS]}")
    print(f"Report type              : {REPORT_RECEIVABLES}. {REPORT_TYPES[REPORT_RECEIVABLES]}")
    print(f"Sales CSV                : {result.sales_csv}")
    print(f"Receivables file         : {result.receivables_path} (sheet '{result.receivables.sheet}', "
          f"header row {result.receivables.header_row})")
    print(f"Main Data                : {result.analysis.main_data.path}")
    if st:
        print("-" * 72)
        print(f"Customer codes (CSV)     : {st['csv_codes']}  | found {st['found']} | not found {st['not_found']} "
              f"| duplicates {st['duplicates']}")
        if st["missing_codes"]:
            shown = ", ".join(st["missing_codes"][:20]) + (" ..." if len(st["missing_codes"]) > 20 else "")
            print(f"   not in the receivables file (left empty): {shown}")
        print(f"Master Data              : {st['customers']} customers (one row each), {st['transactions']} "
              "transactions summed")
        print(f"Sales / returns / net    : {fmt_amount(st['sales'])} / {fmt_amount(st['returns'])} / "
              f"{fmt_amount(st['net'])}")
        print(f"{receivables.COL_OLD:<25}: file {fmt_amount(st['source_old'])} | Master Data "
              f"{fmt_amount(st['master_old'])} | Cover {fmt_amount(st['cover_old'])}")
        print(f"{receivables.COL_NEW:<25}: file {fmt_amount(st['source_new'])} | Master Data "
              f"{fmt_amount(st['master_new'])} | Cover {fmt_amount(st['cover_new'])}")
        for name, count, old, new in st["per_representative"]:
            if count:
                print(f"   {name:<24} {count:>4} customers  {fmt_amount(old):>18}  {fmt_amount(new):>18}")
    print("-" * 72)
    print(f"Receivables check        : {'PASS' if not result.problems else 'FAIL'}")
    for problem in result.problems:
        print(f"   FAIL  {problem}")
    print(f"Receivables report       : {result.report_path}")
    print("Input files were not moved (report type 3 never moves files; reports 1-2 were not run).")
    print(f"RUN RESULT               : {'SUCCEEDED - report written, re-opened and verified' if result.succeeded else 'COMPLETED WITH PROBLEMS - see FAIL lines'}")
    print(line)


class _NoInput(Exception):
    """A needed input file is not in the input folder."""


def _choose_one(kind: str, flag: str, found: list[Path], interactive: bool, input_fn: Callable[[str], str]) -> Path:
    """The single file of ``found``; with several, the user chooses by number (never at random), or - not
    interactive - the run stops listing them."""
    if len(found) == 1:
        return found[0]
    names = "\n".join(f"   {i}. {path.name}" for i, path in enumerate(found, start=1))
    if not interactive:
        raise ConfigError(f"{len(found)} {kind} files in the input folder - choose one with {flag}:\n{names}")
    print(f"{len(found)} {kind} files were found in the input folder:\n{names}")
    while True:
        try:
            choice = input_fn(f"Which {kind} file (1-{len(found)})? ").strip()
        except (EOFError, KeyboardInterrupt) as exc:
            raise ConfigError(f"Cancelled: no {kind} file chosen") from exc
        if choice.isdigit() and 1 <= int(choice) <= len(found):
            return found[int(choice) - 1]
        print(f"  Invalid choice: '{choice}' - enter a number from 1 to {len(found)}.")


def run_receivables_task(args: argparse.Namespace, input_dir: Path, interactive: bool,
                         input_fn: Callable[[str], str]) -> int:
    """Task 5 option 3 (Receivables Update): its two files come from the project input folder (the folder every
    task uses), recognised by their structure - a Task 5 sales CSV and a receivables workbook - unless given with
    --sales-csv / --receivables; no path is asked. Several candidates: the user chooses (never at random). Nothing in
    the folder is renamed, moved or deleted; options 1-2 never run."""
    from sales_etl import receivables

    try:
        main_data = analysis.load_main_data(args.main_data or analysis.find_main_data())
        fixed_codes = analysis.load_fixed_codes(main_data.path.parent / analysis.FIXED_CODES_NAME)
        config = PipelineConfig(input_dir=input_dir, output_dir=args.output_dir, reference_date=date.today(),
                                forced_encoding=args.encoding, forced_delimiter=_delimiter(args.delimiter)).validated()
        given_sales = _file_arg_or_ask(args.sales_csv, "--sales-csv", "", (".csv",), False, input_fn) \
            if args.sales_csv is not None else None
        given_recv = _file_arg_or_ask(args.receivables, "--receivables", "", (".xlsx", ".xlsm"), False, input_fn) \
            if args.receivables is not None else None
        scan = receivables.find_inputs(input_dir, config) if given_sales is None or given_recv is None else None
        if scan is not None:
            print(f"Input folder             : {input_dir}")
            missing = []
            if given_sales is None and not scan.sales:
                missing.append(f"no Sales CSV (a .csv file with the Task 5 columns العميل, نوع المستند, صافى)")
            if given_recv is None and not scan.receivables:
                missing.append(f"no receivables Excel file (a .xlsx with the columns {receivables.COL_CODE}, "
                               f"{receivables.COL_OLD}, {receivables.COL_NEW})")
            if missing:
                ignored = "".join(f"\n   ignored: {path.name} - {reason}" for path, reason in scan.ignored)
                raise _NoInput(f"the input folder {input_dir} has " + " and ".join(missing) + ignored)
        sales_csv = given_sales or _choose_one("Sales CSV", "--sales-csv", scan.sales, interactive, input_fn)
        recv = given_recv or _choose_one("receivables Excel", "--receivables", scan.receivables, interactive, input_fn)
        print(f"Sales CSV (found)        : {sales_csv.name}")
        print(f"Receivables (found)      : {recv.name}")
    except _NoInput as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        print("Put the missing file in the input folder (nothing was written, moved or changed).", file=sys.stderr)
        return EXIT_NO_INPUT
    except (ConfigError, analysis.MainDataError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    try:
        fixed_path = main_data.path.parent / analysis.FIXED_CODES_NAME
        result = receivables.run_receivables_update(config, sales_csv, recv, main_data, fixed_codes=fixed_codes,
                                                    fixed_customers=fixed_customers.load_fixed_customers(fixed_path),
                                                    fixed_path=fixed_path)
    except receivables.ReceivablesError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        print("No receivables report was written; no file was moved or changed.", file=sys.stderr)
        return EXIT_AUDIT_FAILED
    except Exception as exc:  # noqa: BLE001 - report the failure, nothing is moved
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return EXIT_AUDIT_FAILED
    print_receivables_summary(result)
    return EXIT_OK if result.succeeded else EXIT_AUDIT_FAILED


def run_analysis_task(args: argparse.Namespace, input_dir: Path, interactive: bool,
                      input_fn: Callable[[str], str]) -> int:
    """Task 5: sector & representative analysis via the Main Data workbook (three-sheet workbook)."""
    try:
        report_type = args.report_type
        if report_type is None:
            if args.ref_date is not None:
                report_type = "1"
            elif interactive:
                report_type = ask_report_type(input_fn)
            else:
                raise ConfigError("--report-type is required for task 5 in non-interactive mode")
        if report_type == REPORT_RECEIVABLES:
            if args.ref_date is not None:
                raise ConfigError("report type 3 (Receivables Update) takes no date")
            return run_receivables_task(args, input_dir, interactive, input_fn)
        if args.sales_csv is not None or args.receivables is not None:
            raise ConfigError("--sales-csv and --receivables belong to report type 3 (Receivables Update)")
        if args.collection_dir is not None and report_type != "1":
            raise ConfigError("--collection-dir belongs to report type 1 (Yesterday & Previous Days)")
        # option 1 only: report 3 reads the current collection workbook of this folder on every run
        collection_dir = (args.collection_dir or DEFAULT_COLLECTION_DIR) if report_type == "1" else None
        reference_date = None
        if report_type == "1":
            reference_date = _date_arg_or_ask(args.ref_date, "--ref-date", YESTERDAY_PROMPT, interactive, input_fn)
            print(f"Yesterday's date confirmed: {describe_date(reference_date)}")
        elif args.ref_date is not None:
            raise ConfigError("report type 2 (All Transactions) takes no date")
        mode = analysis.MODE_YESTERDAY if report_type == "1" else analysis.MODE_ALL
        main_data = analysis.load_main_data(args.main_data or analysis.find_main_data())
        fixed_codes = analysis.load_fixed_codes(main_data.path.parent / analysis.FIXED_CODES_NAME)
        fixed_list = fixed_customers.load_fixed_customers(main_data.path.parent / analysis.FIXED_CODES_NAME)
        config = PipelineConfig(  # reference_date is only used through ``reference_date`` above
            input_dir=input_dir, output_dir=args.output_dir, reference_date=reference_date or date.today(),
            forced_encoding=args.encoding, forced_delimiter=_delimiter(args.delimiter),
            recursive=args.recursive,
        ).validated()
        files_config = None
        if not args.no_move:
            if args.recursive:
                raise ConfigError("--recursive is only supported together with --no-move")
            default = FileManagementConfig.beside(config.input_dir)
            files_config = FileManagementConfig(
                processed_dir=args.processed_dir or default.processed_dir,
                failures_dir=args.failures_dir or default.failures_dir,
                state_dir=args.state_dir or default.state_dir,
            ).validated(config)
    except (ConfigError, DateParseError, analysis.MainDataError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE

    line = "=" * 72
    if files_config is None:
        result = analysis.run_analysis(config, mode, reference_date, main_data, fixed_codes=fixed_codes,
                                       fixed_customers=fixed_list, collection_dir=collection_dir)
        if not result.files:
            print(line)
            print("RUN RESULT               : NO REPORT - there were no CSV files in the input folder")
            print(line)
            return EXIT_NO_INPUT
        problems = analysis.verify_task5(result)
        print_analysis_summary(result, problems)
        print("Input files were not moved (--no-move).")
        ok = result.complete and not problems
        print(f"RUN RESULT               : {'OK' if ok else 'COMPLETED WITH PROBLEMS - see FAIL / WARNING lines'}")
        print(line)
        return EXIT_OK if ok else EXIT_AUDIT_FAILED

    try:
        run = analysis.run_analysis_managed(config, mode, reference_date, main_data, files_config,
                                            fixed_codes=fixed_codes, fixed_customers=fixed_list,
                                            collection_dir=collection_dir)
    except RunLockedError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    except Exception as exc:  # noqa: BLE001 - nothing moved: files stay in input
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        print("No file was moved; all input files are still in the input folder.", file=sys.stderr)
        return EXIT_AUDIT_FAILED
    for action in run.recovery:
        print(f"RECOVERY {action.action}: {action.run_id} {action.file_name} - {action.detail}")
    if run.nothing_to_do:
        print(line)
        print("RUN RESULT               : NO REPORT - there were no CSV files in the input folder")
        print(line)
        return EXIT_NO_INPUT
    print_analysis_summary(run.result, run.output_problems)
    print(f"Processed archive folder : {run.processed_run_dir}")
    print("File dispositions:")
    for d in run.dispositions:
        state = "moved" if d.moved else ("NOT MOVED: " + d.error if d.error else "left in input")
        print(f"   {d.disposition:<9} {d.file_name}  ->  {d.location}  [{state}]")
        print(f"             {d.reason}")
    print(f"Journal                  : {run.journal_path}")
    state = ("SUCCEEDED - report generated and verified, all files archived" if run.succeeded
             else "COMPLETED WITH PROBLEMS - see FAIL / WARNING lines and file dispositions")
    print(f"RUN RESULT               : {state}")
    print(line)
    return EXIT_OK if run.succeeded else EXIT_AUDIT_FAILED


def _date_arg_or_ask(value: str | None, name: str, prompt: str, interactive: bool,
                     input_fn: Callable[[str], str]) -> date:
    if value is not None:
        return parse_reference_date(value)
    if not interactive:
        raise ConfigError(f"{name} is required in non-interactive mode")
    return ask_reference_date(input_fn, prompt)


def run_period_task(args: argparse.Namespace, input_dir: Path, interactive: bool,
                    input_fn: Callable[[str], str]) -> int:
    """Task 2: totals for start <= التاريخ <= end; files are then filed away (``--no-move``: read-only)."""
    try:
        if args.start_date is None and interactive:
            waiting = sorted(Path(input_dir).glob("*.csv")) if Path(input_dir).is_dir() else []
            print(f"مجلد الملفات: {input_dir}  ({len(waiting)} ملف CSV)")
        start = _date_arg_or_ask(args.start_date, "--start-date", START_DATE_PROMPT, interactive, input_fn)
        while True:
            end = _date_arg_or_ask(args.end_date, "--end-date", END_DATE_PROMPT, interactive, input_fn)
            if end >= start:
                break
            message = f"تاريخ النهاية {end:%d-%m-%Y} قبل تاريخ البداية {start:%d-%m-%Y}"
            if args.end_date is not None or not interactive:
                raise ConfigError(message)
            print(f"  {message} - حاول تاني.")
        print(f"Period confirmed: {describe_date(start)} -> {describe_date(end)} (both days included)")
        config = PipelineConfig(  # reference_date is not used by task 2
            input_dir=input_dir, output_dir=args.output_dir, reference_date=end,
            forced_encoding=args.encoding, forced_delimiter=_delimiter(args.delimiter),
            recursive=args.recursive,
        ).validated()
        files_config = None
        if not args.no_move:
            if args.recursive:
                raise ConfigError("--recursive is only supported together with --no-move")
            default = FileManagementConfig.beside(config.input_dir)
            files_config = FileManagementConfig(
                processed_dir=args.processed_dir or default.processed_dir,
                failures_dir=args.failures_dir or default.failures_dir,
                state_dir=args.state_dir or default.state_dir,
            ).validated(config)
    except (ConfigError, DateParseError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    if files_config is None:
        result = run_period_totals(config, start, end)
        print_period_summary(result)
        if not result.files:
            return EXIT_NO_INPUT
        return EXIT_OK if result.complete else EXIT_AUDIT_FAILED

    try:
        run = run_period_managed(config, start, end, files_config)
    except RunLockedError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    except Exception as exc:  # noqa: BLE001 - nothing committed: files stay in input
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        print("No file was archived; all input files are still in the input folder.", file=sys.stderr)
        return EXIT_AUDIT_FAILED
    print_period_run_summary(run, start, end, config.input_dir)
    if run.nothing_to_do:
        return EXIT_NO_INPUT
    return EXIT_OK if run.succeeded else EXIT_AUDIT_FAILED


def run_daily_task(args: argparse.Namespace, input_dir: Path, interactive: bool,
                   input_fn: Callable[[str], str]) -> int:
    """Task 1: yesterday vs previous days (the original behaviour, unchanged)."""
    try:
        if args.ref_date is not None:
            reference_date = parse_reference_date(args.ref_date)
        elif interactive:
            waiting = sorted(Path(input_dir).glob("*.csv")) if Path(input_dir).is_dir() else []
            print(f"مجلد الملفات: {input_dir}  ({len(waiting)} ملف CSV)")
            reference_date = ask_reference_date(input_fn)
        else:
            raise ConfigError("--ref-date is required in non-interactive mode")
        print(f"Reference Yesterday Date confirmed: {describe_date(reference_date)}")
        print("  Yesterday      = التاريخ equal to this date")
        print("  Previous days  = التاريخ strictly before this date; later dates are excluded")
        config = PipelineConfig(
            input_dir=input_dir, output_dir=args.output_dir, reference_date=reference_date,
            forced_encoding=args.encoding, forced_delimiter=_delimiter(args.delimiter),
            recursive=args.recursive,
        ).validated()
        files_config = None
        if not args.no_move:
            if args.recursive:
                raise ConfigError("--recursive is only supported together with --no-move")
            default = FileManagementConfig.beside(config.input_dir)
            files_config = FileManagementConfig(
                processed_dir=args.processed_dir or default.processed_dir,
                failures_dir=args.failures_dir or default.failures_dir,
                state_dir=args.state_dir or default.state_dir,
            ).validated(config)
    except (ConfigError, DateParseError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE

    if files_config is None:
        result = run_pipeline(config)
        print("=" * 72)
        print_summary(result)
        print("=" * 72)
        return EXIT_OK if result.audit_status is CheckStatus.PASS else EXIT_AUDIT_FAILED

    try:
        managed = run_managed(config, files_config)
    except RunLockedError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return EXIT_USAGE
    except Exception as exc:  # noqa: BLE001 - files stay in input; manifest records the failure
        print(f"RUN RESULT               : {_STATUS_TEXT[FAILED_RUN]}", file=sys.stderr)
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        print("No file was archived; all input files are still in the input folder.", file=sys.stderr)
        return EXIT_AUDIT_FAILED
    print_managed_summary(managed)
    if managed.status == NO_INPUT_FILES:
        return EXIT_NO_INPUT
    return EXIT_OK if managed.succeeded else EXIT_AUDIT_FAILED
