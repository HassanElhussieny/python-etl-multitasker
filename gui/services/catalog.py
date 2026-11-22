"""The catalogue of the engine's tasks and how each one is called on the command line.

One entry per task (Task 5 has one entry per report option). ``build_arguments`` turns the values of a task's
form into the exact ``python -m sales_etl`` arguments - the same ones a user could type in the terminal. Nothing
here calculates or changes a business rule; unknown or missing values are refused instead of guessed.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from gui.services.settings import AppSettings


@dataclass(frozen=True)
class Param:
    key: str                 # value key in the form
    label: str               # Arabic label
    kind: str                # "date" | "dir" | "file"
    flag: str                # command-line flag
    date_format: str = ""    # strftime for kind == "date" (the CLI's own format for that flag)
    required: bool = True
    hint: str = ""


@dataclass(frozen=True)
class TaskDef:
    key: str                 # "1" .. "7", "5-1" / "5-2" / "5-3"
    task: str                # --task value
    title: str               # Arabic title
    subtitle: str            # English name as in the CLI menu
    description: str
    params: tuple[Param, ...] = ()
    report_type: str = ""    # Task 5: --report-type
    read_only_only: bool = False   # Task 5 option 3: never moves files (no lifecycle at all)
    output_prefixes: tuple[str, ...] = ()
    exclude_prefixes: tuple[str, ...] = ()
    journal: str = ""        # state/<journal> of its managed runs
    outputs: tuple[str, ...] = field(default=())
    file_types: tuple[str, ...] = (".csv",)    # what the user may pick for this task

    @property
    def visible_params(self) -> tuple[Param, ...]:
        """The values the simple screen asks for: only the required ones (optional ones use their defaults)."""
        return tuple(p for p in self.params if p.required)

    @property
    def file_hint(self) -> str:
        if ".xlsx" in self.file_types:
            return "اختر ملف المبيعات CSV وملف المستحقات Excel"
        return "اختر ملف CSV واحدًا أو أكثر"


DAY_FMT, DASH_FMT = "%d/%m/%Y", "%d-%m-%Y"
MAIN_DATA = Param("main_data", "ملف Main Data (اختياري)", "file", "--main-data", required=False,
                  hint="فارغ = ملف Excel الوحيد في مجلد data")

TASKS: tuple[TaskDef, ...] = (
    TaskDef("1", "1", "التقرير اليومي", "Daily",
            "مبيعات ومرتجعات يوم التاريخ المرجعي مقابل الأيام السابقة لكل ملف، مع التدقيق ومطابقة إجماليات الـERP.",
            (Param("ref_date", "تاريخ أمس (التاريخ المرجعي)", "date", "--ref-date", DAY_FMT),),
            output_prefixes=("Report_",), journal="processing_journal.jsonl",
            outputs=("sales_report_ref-<date>.xlsx", "audit_ref-<date>.xlsx", "audit_ref-<date>.json",
                     "pipeline.log.jsonl", "run_manifest.json")),
    TaskDef("2", "2", "تقرير الفترة", "Period",
            "إجمالي المبيعات والمرتجعات لكل ملف خلال فترة يحددها المستخدم (يومي البداية والنهاية ضمنها).",
            (Param("start_date", "تاريخ البداية", "date", "--start-date", DASH_FMT),
             Param("end_date", "تاريخ النهاية", "date", "--end-date", DASH_FMT)),
            output_prefixes=("Period_",), journal="period_journal.jsonl"),
    TaskDef("3", "3", "إجمالي العملاء", "Customer",
            "المبيعات والمرتجعات والصافي لكل كود عميل على كل الحركات (بدون فلتر تاريخ).",
            output_prefixes=("Customer_",), journal="customer_journal.jsonl"),
    TaskDef("4", "4", "إجمالي المناديب", "Representative",
            "المبيعات والمرتجعات والصافي لكل مندوب على كل الحركات (بدون فلتر تاريخ).",
            output_prefixes=("Representative_",), journal="representative_journal.jsonl"),
    TaskDef("5-1", "5", "تحليل القطاعات والمناديب — أمس والأيام السابقة", "Sector & Representative Analysis · 1",
            "أربعة تقارير: تحليل المبيعات والمرتجعات، Representatives_Analysis، تقرير عملاء الـ224، "
            "والعملاء الذين لم يحققوا مبيعات في الشهر — مع قيم التحصيل.",
            (Param("ref_date", "تاريخ أمس (التاريخ المرجعي)", "date", "--ref-date", DAY_FMT),
             Param("collection_dir", "مجلد التحصيل", "dir", "--collection-dir", required=False,
                   hint="فارغ = المجلد المحدد في الإعدادات"),
             MAIN_DATA),
            report_type="1", output_prefixes=("Task5_",), exclude_prefixes=("Task5_Receivables",),
            journal="task5_journal.jsonl"),
    TaskDef("5-2", "5", "تحليل القطاعات والمناديب — كل الحركات", "Sector & Representative Analysis · 2",
            "التقارير الثلاثة لكل الحركات بدون تقسيم تاريخ.",
            (MAIN_DATA,), report_type="2", output_prefixes=("Task5_",), exclude_prefixes=("Task5_Receivables",),
            journal="task5_journal.jsonl"),
    TaskDef("5-3", "5", "تحديث المستحقات (Receivables Update)", "Sector & Representative Analysis · 3",
            "يقرأ ملف المبيعات وملف المستحقات ويكتب تقرير المستحقات. لا ينقل أي ملف ولا يكتب سجلًا.",
            (Param("sales_csv", "ملف المبيعات CSV (اختياري)", "file", "--sales-csv", required=False,
                   hint="فارغ = يُكتشف من مجلد الإدخال"),
             Param("receivables", "ملف المستحقات Excel (اختياري)", "file", "--receivables", required=False,
                   hint="فارغ = يُكتشف من مجلد الإدخال"),
             MAIN_DATA),
            report_type="3", read_only_only=True, output_prefixes=("Task5_Receivables_",),
            file_types=(".csv", ".xlsx", ".xlsm")),
    TaskDef("6", "6", "صافي مبيعات العملاء الشهري", "Monthly Customer Net Sales",
            "صافي كل عميل لكل شهر من يناير.",
            output_prefixes=("صافي_العملاء_الشهري_",), journal="سجل_صافي_العملاء_الشهري.jsonl"),
    TaskDef("7", "7", "عدد العملاء الشهري لكل ملف", "Monthly Customers per File",
            "عدد العملاء المميزين لكل ملف لكل شهر محاسبي (26 إلى 25).",
            output_prefixes=("عملاء_الملفات_الشهري_",), journal="سجل_عملاء_الملفات_الشهري.jsonl"),
)
BY_KEY: dict[str, TaskDef] = {t.key: t for t in TASKS}


class ArgumentError(ValueError):
    """A task's form cannot be turned into a command (a required value is missing or invalid)."""


def build_arguments(task_key: str, values: dict[str, object], settings: AppSettings, *, read_only: bool,
                    input_dir: str | None = None) -> list[str]:
    """The ``python -m sales_etl`` arguments (without the interpreter) for one job. Read-only adds --no-move; a
    managed job names its processed / failures / state folders explicitly (the engine's lifecycle uses them).
    ``input_dir``: the folder the engine reads (the desktop app passes its staging folder of copied files)."""
    task = BY_KEY.get(task_key)
    if task is None:
        raise ArgumentError(f"مهمة غير معروفة: {task_key}")
    args = ["-m", "sales_etl", "--task", task.task]
    if task.report_type:
        args += ["--report-type", task.report_type]
    for p in task.params:
        value = values.get(p.key)
        if p.key == "collection_dir" and not value:
            value = settings.collection_dir          # always explicit: never an implicit folder
        if p.key == "main_data" and not value:
            value = settings.main_data               # "" -> the CLI finds the single workbook in data
        if value in (None, ""):
            if p.required:
                raise ArgumentError(f"الحقل مطلوب: {p.label}")
            continue
        if p.kind == "date":
            if not isinstance(value, date):
                raise ArgumentError(f"تاريخ غير صالح: {p.label}")
            args += [p.flag, value.strftime(p.date_format)]
        else:
            path = Path(str(value))
            if p.kind == "file" and not path.is_file():
                raise ArgumentError(f"الملف غير موجود: {path}")
            if p.kind == "dir" and not path.is_dir():
                raise ArgumentError(f"المجلد غير موجود: {path}")
            args += [p.flag, str(path)]
    if task.key == "2":
        if values["start_date"] > values["end_date"]:   # type: ignore[operator]
            raise ArgumentError("تاريخ البداية بعد تاريخ النهاية")
    args += ["--input-dir", input_dir or settings.input_dir, "--output-dir", settings.output_dir]
    if read_only or task.read_only_only:
        if not task.read_only_only:
            args.append("--no-move")
    else:
        args += ["--processed-dir", settings.processed_dir, "--failures-dir", settings.failures_dir,
                 "--state-dir", settings.state_dir]
    args.append("--non-interactive")
    return args


def moves_files(task_key: str, read_only: bool) -> bool:
    """True when the job runs the engine's lifecycle (input files are moved after verification)."""
    return not read_only and not BY_KEY[task_key].read_only_only


def matches_output(task: TaskDef, folder_name: str) -> bool:
    return any(folder_name.startswith(p) for p in task.output_prefixes) and \
        not any(folder_name.startswith(p) for p in task.exclude_prefixes)


def task_for_folder(folder_name: str) -> TaskDef | None:
    """The task that writes an output folder of this name (longest prefix wins: Task5_Receivables before Task5)."""
    best: tuple[int, TaskDef] | None = None
    for t in TASKS:
        for p in t.output_prefixes:
            if folder_name.startswith(p) and not any(folder_name.startswith(x) for x in t.exclude_prefixes):
                if best is None or len(p) > best[0]:
                    best = (len(p), t)
    return best[1] if best else None
