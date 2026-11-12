"""Domain models shared by the pipeline stages."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from pathlib import Path

ZERO = Decimal("0")


class Severity(str, Enum):
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"


class CheckStatus(str, Enum):
    PASS = "PASS"
    WARN = "WARN"
    FAIL = "FAIL"
    NOT_APPLICABLE = "N/A"


class Category(str, Enum):
    SALES = "SALES"
    RETURNS = "RETURNS"
    OTHER = "OTHER"


class Period(str, Enum):
    YESTERDAY = "YESTERDAY"  # التاريخ == reference date
    PREVIOUS = "PREVIOUS"  # التاريخ <  reference date
    FUTURE = "FUTURE"  # التاريخ >  reference date (always excluded)


class FileProcessingError(Exception):
    """A file-level failure: the file cannot be processed at all."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(f"{code}: {message}")
        self.code = code
        self.message = message


@dataclass(frozen=True)
class Issue:
    """A validation finding. ERROR rows are rejected; WARNING/INFO are informative."""

    file_name: str
    line_no: int | None
    severity: Severity
    code: str
    message: str
    raw: str = ""


@dataclass(frozen=True)
class Transaction:
    """One valid source row. ``net`` is the original signed صافى value."""

    source_file: str
    line_no: int
    txn_date: date
    doc_type: str
    category: Category
    period: Period
    net: Decimal
    doc_no: str
    customer: str
    cells: tuple[str, ...]


@dataclass(frozen=True)
class RejectedRow:
    source_file: str
    line_no: int
    reasons: tuple[str, ...]
    raw: str


@dataclass
class ControlTotals:
    """Totals printed by the ERP in the report footer (if present)."""

    doc_count: int | None = None
    doc_count_line: int | None = None
    net_total: Decimal | None = None
    net_total_line: int | None = None


@dataclass
class ParsedFile:
    file_name: str
    path: Path
    sha256: str
    size_bytes: int
    encoding: str
    encoding_reason: str
    delimiter: str
    header_line_no: int
    columns: list[str]
    column_index: dict[str, int]
    period_from: date | None
    period_to: date | None
    transactions: list[Transaction]
    rejected: list[RejectedRow]
    issues: list[Issue]
    control: ControlTotals
    blank_rows: int = 0
    repeated_headers: int = 0

    @property
    def candidate_rows(self) -> int:
        """Rows that looked like transactions (valid + rejected)."""
        return len(self.transactions) + len(self.rejected)


@dataclass(frozen=True)
class FileMetrics:
    """The four report metrics plus supporting counts for one file (or the grand total)."""

    sales_yesterday: Decimal = ZERO
    sales_previous: Decimal = ZERO
    returns_yesterday: Decimal = ZERO
    returns_previous: Decimal = ZERO
    sales_yesterday_count: int = 0
    sales_previous_count: int = 0
    returns_yesterday_count: int = 0
    returns_previous_count: int = 0
    future_count: int = 0
    future_net: Decimal = ZERO
    other_type_count: int = 0
    other_type_net: Decimal = ZERO

    def value(self, key: str) -> Decimal:
        return getattr(self, key)

    def count(self, key: str) -> int:
        return getattr(self, f"{key}_count")


@dataclass(frozen=True)
class CustomerSets:
    """Distinct customer codes of one file (or, as a union, of the grand total).

    Only valid مبيعات rows (never مرتجعات) dated on or before the reference date with a non-blank code.
    Kept as sets, never as counts, so files can be combined without counting a customer twice.
    """

    yesterday: frozenset[str] = frozenset()
    previous: frozenset[str] = frozenset()

    def union(self, other: CustomerSets) -> CustomerSets:
        return CustomerSets(self.yesterday | other.yesterday, self.previous | other.previous)

    def counts(self) -> list[int]:
        """In ``CUSTOMER_COUNT_KEYS`` order: yesterday, previous, total."""
        return [len(self.yesterday), len(self.previous), len(self.yesterday | self.previous)]


@dataclass(frozen=True)
class CheckResult:
    scope: str
    check: str
    status: CheckStatus
    expected: str = ""
    actual: str = ""
    difference: str = ""
    detail: str = ""


@dataclass
class FileOutcome:
    file_name: str
    path: Path
    parsed: ParsedFile | None = None
    metrics: FileMetrics | None = None
    customers: CustomerSets | None = None
    error_code: str | None = None
    error_message: str | None = None
    checks: list[CheckResult] = field(default_factory=list)

    @property
    def succeeded(self) -> bool:
        return self.error_code is None


@dataclass
class PipelineResult:
    reference_date: date
    input_dir: Path
    run_dir: Path
    started_at: datetime
    finished_at: datetime | None = None
    discovered: list[Path] = field(default_factory=list)
    skipped: list[tuple[Path, str]] = field(default_factory=list)
    outcomes: list[FileOutcome] = field(default_factory=list)
    grand_total: FileMetrics = field(default_factory=FileMetrics)
    grand_customers: CustomerSets = field(default_factory=CustomerSets)
    checks: list[CheckResult] = field(default_factory=list)
    report_path: Path | None = None
    audit_xlsx_path: Path | None = None
    audit_json_path: Path | None = None
    log_path: Path | None = None

    @property
    def successful(self) -> list[FileOutcome]:
        return [o for o in self.outcomes if o.succeeded]

    @property
    def failed(self) -> list[FileOutcome]:
        return [o for o in self.outcomes if not o.succeeded]

    @property
    def audit_status(self) -> CheckStatus:
        """FAIL if any check failed (or none ran); otherwise PASS."""
        if not self.checks or any(c.status is CheckStatus.FAIL for c in self.checks):
            return CheckStatus.FAIL
        return CheckStatus.PASS

    def count_checks(self, status: CheckStatus) -> int:
        return sum(1 for c in self.checks if c.status is status)
