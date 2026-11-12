"""Business constants and the runtime configuration object."""

from __future__ import annotations

import codecs
from dataclasses import dataclass, replace
from datetime import date
from pathlib import Path

# --- Source column names (kept exactly as they appear in the ERP export) ---------
COL_DATE = "التاريخ"
COL_DOC_TYPE = "نوع المستند"
COL_DOC_NO = "رقم المستند"
COL_CUSTOMER = "العميل"
COL_REP = "المندوب"
COL_GROSS = "اجمالى"
COL_NET = "صافى"

#: Columns without which a file cannot be processed at all.
REQUIRED_COLUMNS: tuple[str, ...] = (COL_DATE, COL_DOC_TYPE, COL_NET)

#: Columns normally present; their absence is reported as a warning only.
EXPECTED_COLUMNS: tuple[str, ...] = (
    COL_DOC_NO,
    COL_CUSTOMER,
    COL_REP,
    COL_GROSS,
    "خصم 1",
    "خصم صيدلى",
    "خصم موزع",
    "خصم نقدى",
    "ض.قيمة مضا",
    "ضريبة 2",
    "ض.ا-ت صناع",
    "ضريبة 4",
)

DOC_TYPE_SALES = "مبيعات"
DOC_TYPE_RETURNS = "مرتجعات"

# --- ERP report decorations ------------------------------------------------------
#: "عدد الفواتير" / "إجمالى عام" are what the real ERP export prints (alex8, kenna9).
FOOTER_DOC_COUNT_LABELS: tuple[str, ...] = ("عدد الفواتير", "عدد المستندات")
FOOTER_GRAND_TOTAL_LABELS: tuple[str, ...] = ("إجمالى عام", "الاجمالى العام", "الإجمالي العام")
PREAMBLE_FROM_LABEL = "من"
PREAMBLE_TO_LABEL = "الى"

# --- Output report ---------------------------------------------------------------
REPORT_FILE_COLUMN = "اسم الملف"
#: Internal metric key -> report column header (order = report column order).
METRIC_COLUMNS: dict[str, str] = {
    "sales_yesterday": "إجمالي مبيعات أمس",
    "sales_previous": "إجمالي مبيعات الأيام السابقة",
    "returns_yesterday": "إجمالي مرتجعات أمس",
    "returns_previous": "إجمالي مرتجعات الأيام السابقة",
}
METRIC_KEYS: tuple[str, ...] = tuple(METRIC_COLUMNS)
#: Derived totals shown FIRST in the report (computed in Python from the four metrics):
#:   sales_total   = sales_yesterday + sales_previous
#:   returns_total = returns_yesterday + returns_previous   (original sign kept)
#:   net_total     = sales_total + returns_total
TOTAL_COLUMNS: dict[str, str] = {
    "sales_total": "إجمالي المبيعات",
    "returns_total": "إجمالي المرتجعات",
    "net_total": "الصافي",
}
TOTAL_KEYS: tuple[str, ...] = tuple(TOTAL_COLUMNS)
#: Distinct customer codes (العميل) with at least one valid مبيعات row, shown LAST in the report
#: (returns never qualify a customer; a customer with only مرتجعات is not counted):
#:   customers_yesterday    = codes with a مبيعات row dated the reference date
#:   customers_previous     = codes with a مبيعات row dated before it
#:   customers_total        = customers_yesterday ∪ customers_previous (never a sum)
CUSTOMER_COUNT_COLUMNS: dict[str, str] = {
    "customers_yesterday": "عدد عملاء أمس",
    "customers_previous": "عدد عملاء الأيام السابقة",
    "customers_total": "إجمالي العملاء المميزين",
}
CUSTOMER_COUNT_KEYS: tuple[str, ...] = tuple(CUSTOMER_COUNT_COLUMNS)
REPORT_COLUMNS: tuple[str, ...] = (*TOTAL_COLUMNS.values(), REPORT_FILE_COLUMN, *METRIC_COLUMNS.values(),
                                   *CUSTOMER_COUNT_COLUMNS.values())
GRAND_TOTAL_LABEL = "Grand Total"

# --- Parsing defaults ------------------------------------------------------------
#: Tried in order after BOM detection. cp1256 = Windows Arabic (ERP default).
DEFAULT_ENCODINGS: tuple[str, ...] = ("utf-8", "cp1256")
DEFAULT_DELIMITERS: tuple[str, ...] = (";", ",", "\t", "|")


#: Project root (the folder that contains the ``sales_etl`` package).
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_DIR = PROJECT_ROOT / "input"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"


class ConfigError(ValueError):
    """Raised when the runtime configuration is unusable."""


@dataclass(frozen=True)
class FileManagementConfig:
    """Where input files go after a run, and where the processing journal lives."""

    processed_dir: Path
    failures_dir: Path
    state_dir: Path

    @classmethod
    def beside(cls, input_dir: Path) -> FileManagementConfig:
        """Default layout: sibling folders of the input folder."""
        root = Path(input_dir).expanduser().resolve().parent
        return cls(root / "processed", root / "failures", root / "state")

    def validated(self, config: PipelineConfig) -> FileManagementConfig:
        """Resolve paths and make sure no folder collides with input/output or another."""
        resolved = {name: Path(getattr(self, name)).expanduser().resolve()
                    for name in ("processed_dir", "failures_dir", "state_dir")}
        taken = {config.input_dir: "input", config.output_dir: "output"}
        for name, path in resolved.items():
            if path in taken:
                raise ConfigError(f"{name} must differ from the {taken[path]} folder: {path}")
            if path.exists() and not path.is_dir():
                raise ConfigError(f"{name} exists and is not a folder: {path}")
            taken[path] = name
        return FileManagementConfig(**resolved)


@dataclass(frozen=True)
class PipelineConfig:
    """Everything a pipeline run needs; nothing is read from the system clock."""

    input_dir: Path
    output_dir: Path
    reference_date: date
    encodings: tuple[str, ...] = DEFAULT_ENCODINGS
    delimiters: tuple[str, ...] = DEFAULT_DELIMITERS
    forced_encoding: str | None = None
    forced_delimiter: str | None = None
    recursive: bool = False

    def validated(self) -> PipelineConfig:
        """Return a copy with resolved paths, or raise :class:`ConfigError`."""
        input_dir = Path(self.input_dir).expanduser().resolve()
        output_dir = Path(self.output_dir).expanduser().resolve()
        if not input_dir.is_dir():
            raise ConfigError(f"Input folder does not exist or is not a folder: {input_dir}")
        if output_dir == input_dir:
            raise ConfigError("The output folder must be different from the input folder.")
        if output_dir.exists() and not output_dir.is_dir():
            raise ConfigError(f"Output path exists and is not a folder: {output_dir}")
        if self.forced_encoding:
            try:
                codecs.lookup(self.forced_encoding)
            except LookupError as exc:
                raise ConfigError(f"Unknown encoding: {self.forced_encoding}") from exc
        if self.forced_delimiter is not None and len(self.forced_delimiter) != 1:
            raise ConfigError("The delimiter must be exactly one character.")
        if not isinstance(self.reference_date, date):
            raise ConfigError("reference_date must be a datetime.date")
        return replace(self, input_dir=input_dir, output_dir=output_dir)
