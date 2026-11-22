"""Read-only look at the input folder before a job starts: which files the engine will read, their size and the
dates they contain, and warnings worth a second look. Files are opened for reading only (the engine's own
``ingest``/``normalize`` readers); nothing is moved, written or recalculated - the job itself does the work."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from sales_etl.config import DEFAULT_DELIMITERS, DEFAULT_ENCODINGS
from sales_etl.ingest import discover_csv_files, read_raw_table
from sales_etl.normalize import parse_transactions


@dataclass
class InputFile:
    path: Path
    size: int
    modified: datetime
    rows: int | None = None          # valid transaction rows (after inspection)
    rejected: int | None = None
    first_date: date | None = None
    last_date: date | None = None
    error: str = ""

    @property
    def name(self) -> str:
        return self.path.name


def list_input_files(input_dir: Path) -> list[InputFile]:
    """The CSV files directly in the input folder, exactly as the engine selects them (no sub-folders)."""
    if not input_dir.is_dir():
        return []
    files, _ = discover_csv_files(input_dir)
    out = []
    for p in files:
        try:
            st = p.stat()
        except OSError:
            continue
        out.append(InputFile(p, st.st_size, datetime.fromtimestamp(st.st_mtime)))
    return out


def inspect_file(item: InputFile) -> InputFile:
    """Rows and date range of one file with the engine's readers (read only). A file the engine cannot read gets
    ``error`` - the job will report it the engine's way."""
    try:
        raw = read_raw_table(item.path, encodings=DEFAULT_ENCODINGS, delimiters=DEFAULT_DELIMITERS)
        parsed = parse_transactions(raw, date.today())
    except Exception as exc:  # noqa: BLE001 - shown as a warning; the engine decides
        item.error = f"{type(exc).__name__}: {exc}"
        return item
    days = [t.txn_date for t in parsed.transactions]
    item.rows, item.rejected = len(days), len(parsed.rejected)
    item.first_date, item.last_date = (min(days), max(days)) if days else (None, None)
    return item


@dataclass
class Preflight:
    files: list[InputFile] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def first_date(self) -> date | None:
        days = [f.first_date for f in self.files if f.first_date]
        return min(days) if days else None

    @property
    def last_date(self) -> date | None:
        days = [f.last_date for f in self.files if f.last_date]
        return max(days) if days else None


def file_info(path: Path) -> InputFile:
    st = path.stat()
    return InputFile(path, st.st_size, datetime.fromtimestamp(st.st_mtime))


def preflight(input_dir: Path, task_key: str, values: dict[str, object]) -> Preflight:
    """The files a job will read, inspected, plus warnings. Warnings never block: the user decides."""
    return check_files([inspect_file(f) for f in list_input_files(input_dir)], task_key, values)


def check_files(files: list[InputFile], task_key: str, values: dict[str, object]) -> Preflight:
    """Warnings for already inspected files (CSV files only; a receivables workbook is not a sales file)."""
    result = Preflight([f for f in files if f.path.suffix.lower() == ".csv"])
    w = result.warnings
    if not result.files and task_key != "5-3":
        w.append("لا توجد ملفات CSV في مجلد الإدخال — ستنتهي المهمة بدون تقرير.")
    for f in result.files:
        if f.error:
            w.append(f"{f.name}: تعذّرت قراءته ({f.error}) — المحرك سيسجّله كملف فاشل.")
        elif f.rejected:
            w.append(f"{f.name}: {f.rejected} صف غير صالح — لن تُحسب، والملف لن يُعتبر ناجحًا بالكامل.")
        elif f.rows == 0:
            w.append(f"{f.name}: لا يحتوي على حركات.")
    ref = values.get("ref_date")
    if isinstance(ref, date):
        month_start = ref.replace(day=1)
        for f in result.files:
            if f.first_date and f.first_date < month_start:
                w.append(f"{f.name}: يبدأ في {f.first_date:%d/%m/%Y} قبل بداية شهر التاريخ المرجعي "
                         f"({month_start:%d/%m/%Y}) — كل الأيام السابقة ستُجمع في عمود «الأيام السابقة».")
            if f.last_date and f.last_date > ref:
                w.append(f"{f.name}: يحتوي حركات بعد {ref:%d/%m/%Y} (حتى {f.last_date:%d/%m/%Y}) — ستُستبعد من الحساب.")
        if ref >= date.today():
            w.append(f"التاريخ المرجعي {ref:%d/%m/%Y} هو اليوم أو بعده — هل المقصود «أمس»؟")
    start, end = values.get("start_date"), values.get("end_date")
    if isinstance(start, date) and isinstance(end, date):
        for f in result.files:
            if f.first_date and f.last_date and (f.last_date < start or f.first_date > end):
                w.append(f"{f.name}: كل حركاته ({f.first_date:%d/%m/%Y} – {f.last_date:%d/%m/%Y}) خارج الفترة.")
    return result
