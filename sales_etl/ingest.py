"""CSV discovery and raw ingestion.

Source files are opened read-only (``rb``) exactly once. The encoding and the
delimiter are determined per file from evidence:

* encoding: byte-order mark -> strict UTF-8 -> configured fallbacks (cp1256);
* delimiter: an optional Excel ``sep=`` hint, otherwise the candidate delimiter
  under which a row containing *all* required Arabic headers can be found.

ERP exports put a preamble before the header row and control totals after the
data, so the header row is located by content, not by position.
"""

from __future__ import annotations

import codecs
import csv
import hashlib
import io
import logging
import re
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from sales_etl.config import EXPECTED_COLUMNS, REQUIRED_COLUMNS
from sales_etl.models import FileProcessingError, Issue, Severity
from sales_etl.text import canonical, clean_cell

logger = logging.getLogger(__name__)

_BOMS: tuple[tuple[bytes, str], ...] = (
    (codecs.BOM_UTF8, "utf-8-sig"),
    (codecs.BOM_UTF16_LE, "utf-16"),
    (codecs.BOM_UTF16_BE, "utf-16"),
)
_SEP_HINT = re.compile(r"^\s*sep=(.)\s*$", re.IGNORECASE)

Row = tuple[int, list[str]]  # (physical line number, cells)


@dataclass
class RawTable:
    """A decoded CSV split into preamble, header and body rows."""

    path: Path
    file_name: str
    sha256: str
    size_bytes: int
    encoding: str
    encoding_reason: str
    delimiter: str
    header_line_no: int
    header_cells: list[str]
    columns: list[str]
    column_index: dict[str, int]
    preamble: list[Row]
    body: list[Row]
    issues: list[Issue] = field(default_factory=list)


def sha256_file(path: Path) -> str:
    """SHA-256 of a file's bytes (read-only)."""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _is_within(path: Path, folder: Path) -> bool:
    try:
        path.resolve().relative_to(folder.resolve())
        return True
    except ValueError:
        return False


def discover_csv_files(
    input_dir: Path, *, recursive: bool = False, exclude_dirs: Iterable[Path] = ()
) -> tuple[list[Path], list[tuple[Path, str]]]:
    """Return ``(eligible_files, skipped)`` sorted by name; nothing is hardcoded."""
    excluded = [Path(p) for p in exclude_dirs]
    eligible: list[Path] = []
    skipped: list[tuple[Path, str]] = []
    for path in input_dir.rglob("*") if recursive else input_dir.iterdir():
        if not path.is_file() or path.suffix.lower() != ".csv":
            continue
        if any(_is_within(path, folder) for folder in excluded):
            skipped.append((path, "inside the output folder"))
        elif path.name.startswith(("~$", ".")):
            skipped.append((path, "temporary or hidden file"))
        else:
            eligible.append(path)
    eligible.sort(key=lambda p: str(p).casefold())
    return eligible, skipped


def decode_bytes(
    data: bytes, candidates: Sequence[str], forced: str | None = None
) -> tuple[str, str, str]:
    """Decode bytes, returning ``(text, encoding, reason)``."""
    if forced:
        try:
            return data.decode(forced), forced, "forced by configuration"
        except UnicodeDecodeError as exc:
            raise FileProcessingError(
                "ENCODING_ERROR", f"file is not valid {forced}: {exc}"
            ) from exc
    for bom, encoding in _BOMS:
        if data.startswith(bom):
            try:
                return data.decode(encoding), encoding, "byte-order mark"
            except UnicodeDecodeError as exc:
                raise FileProcessingError(
                    "ENCODING_ERROR", f"BOM says {encoding} but content is invalid: {exc}"
                ) from exc
    failures: list[str] = []
    for encoding in candidates:
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError as exc:
            failures.append(f"{encoding}: {exc.reason} at byte {exc.start}")
            continue
        reason = "strict decode succeeded"
        if failures:
            reason += f" (rejected: {'; '.join(failures)})"
        return text, encoding, reason
    raise FileProcessingError("ENCODING_ERROR", "no candidate encoding fits: " + "; ".join(failures))


def _split_rows(text: str, delimiter: str) -> list[Row]:
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter)
    rows: list[Row] = []
    try:
        for cells in reader:
            rows.append((reader.line_num, cells))
    except csv.Error as exc:
        raise FileProcessingError(
            "CSV_PARSE_ERROR", f"line {reader.line_num}: {exc}"
        ) from exc
    return rows


_REQUIRED_KEYS = {canonical(name): name for name in REQUIRED_COLUMNS}


def _locate_header(rows: list[Row]) -> tuple[int | None, set[str]]:
    """Index of the first row holding every required header, plus the best partial match."""
    best: set[str] = set()
    for index, (_, cells) in enumerate(rows):
        keys = {canonical(c) for c in cells}
        found = {name for key, name in _REQUIRED_KEYS.items() if key in keys}
        if len(found) == len(_REQUIRED_KEYS):
            return index, found
        if len(found) > len(best):
            best = found
    return None, best


def resolve_columns(
    header_cells: list[str], file_name: str, header_line: int
) -> tuple[list[str], dict[str, int], list[Issue]]:
    """Map business columns to positions; give blank/duplicate headers safe names.

    A *required* column appearing more than once is a file-level failure: we refuse
    to guess which one holds the real ``صافى`` / ``التاريخ`` data.
    """
    issues: list[Issue] = []
    keys = [canonical(c) for c in header_cells]
    column_index: dict[str, int] = {}

    for name in REQUIRED_COLUMNS:
        positions = [i for i, key in enumerate(keys) if key == canonical(name)]
        if len(positions) > 1:
            cols = ", ".join(str(p + 1) for p in positions)
            raise FileProcessingError(
                "AMBIGUOUS_REQUIRED_COLUMN",
                f"required column '{name}' appears {len(positions)} times (columns {cols}); "
                "refusing to guess which one holds the data",
            )
        column_index[name] = positions[0]

    for name in EXPECTED_COLUMNS:
        positions = [i for i, key in enumerate(keys) if key == canonical(name)]
        if not positions:
            issues.append(Issue(file_name, header_line, Severity.WARNING, "MISSING_EXPECTED_COLUMN",
                                f"expected column '{name}' not found (not needed for the totals)"))
        elif len(positions) > 1:
            issues.append(Issue(file_name, header_line, Severity.WARNING, "DUPLICATE_EXPECTED_COLUMN",
                                f"column '{name}' appears {len(positions)} times; it will not be used"))
        else:
            column_index[name] = positions[0]

    names: list[str] = []
    seen: dict[str, int] = {}
    blanks: list[int] = []
    for position, cell in enumerate(header_cells):
        label = clean_cell(cell)
        if not label:
            blanks.append(position + 1)
            label = f"[عمود بدون عنوان {position + 1}]"
        key = canonical(label)
        seen[key] = seen.get(key, 0) + 1
        if seen[key] > 1:
            issues.append(Issue(file_name, header_line, Severity.WARNING, "DUPLICATE_HEADER",
                                f"header '{label}' repeated at column {position + 1}; renamed"))
            label = f"{label} ({seen[key]})"
        names.append(label)
    if blanks:
        issues.append(Issue(file_name, header_line, Severity.INFO, "BLANK_HEADERS",
                            f"blank header cells at columns {', '.join(map(str, blanks))}"))
    return names, column_index, issues


def read_raw_table(
    path: Path,
    *,
    encodings: Sequence[str],
    delimiters: Sequence[str],
    forced_encoding: str | None = None,
    forced_delimiter: str | None = None,
) -> RawTable:
    """Read and split one CSV file. Raises :class:`FileProcessingError` on file-level problems."""
    file_name = path.name
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise FileProcessingError("READ_ERROR", str(exc)) from exc
    sha256 = hashlib.sha256(data).hexdigest()
    if not data.strip():
        raise FileProcessingError("EMPTY_FILE", "file has no content")

    text, encoding, reason = decode_bytes(data, encodings, forced_encoding)
    if "\x00" in text:
        raise FileProcessingError("ENCODING_ERROR", f"NUL characters after decoding as {encoding}")
    if not clean_cell(text):
        raise FileProcessingError("EMPTY_FILE", "file contains only whitespace")

    issues: list[Issue] = []
    candidates = list(delimiters)
    first_line = text.split("\n", 1)[0]
    hint = _SEP_HINT.match(first_line)
    if forced_delimiter:
        candidates = [forced_delimiter]
    elif hint:
        candidates = [hint.group(1)]
        issues.append(Issue(file_name, 1, Severity.INFO, "SEP_HINT",
                            f"delimiter {hint.group(1)!r} taken from 'sep=' line"))

    matches: list[tuple[str, list[Row], int]] = []
    best_partial: set[str] = set()
    for delimiter in candidates:
        rows = _split_rows(text, delimiter)
        header_index, partial = _locate_header(rows)
        if header_index is not None:
            matches.append((delimiter, rows, header_index))
        elif len(partial) > len(best_partial):
            best_partial = partial

    if not matches:
        missing = [name for name in REQUIRED_COLUMNS if name not in best_partial]
        raise FileProcessingError(
            "MISSING_REQUIRED_COLUMNS",
            f"no header row containing all required columns; missing: {', '.join(missing)} "
            f"(decoded as {encoding}, delimiters tried: {', '.join(map(repr, candidates))})",
        )
    if len(matches) > 1:
        issues.append(Issue(file_name, None, Severity.WARNING, "AMBIGUOUS_DELIMITER",
                            f"header found with delimiters {[m[0] for m in matches]}; "
                            f"using {matches[0][0]!r}"))
    delimiter, rows, header_index = matches[0]
    header_line, header_cells = rows[header_index]
    columns, column_index, column_issues = resolve_columns(header_cells, file_name, header_line)
    issues.extend(column_issues)

    preamble = rows[:header_index]
    if hint and not forced_delimiter:
        preamble = [r for r in preamble if not _SEP_HINT.match(delimiter.join(r[1]))]

    logger.info(
        "file ingested",
        extra={"event": "ingest", "file": file_name, "encoding": encoding,
               "delimiter": delimiter, "header_line": header_line,
               "body_rows": len(rows) - header_index - 1},
    )
    return RawTable(
        path=path, file_name=file_name, sha256=sha256, size_bytes=len(data),
        encoding=encoding, encoding_reason=reason, delimiter=delimiter,
        header_line_no=header_line, header_cells=header_cells, columns=columns,
        column_index=column_index, preamble=preamble, body=rows[header_index + 1:],
        issues=issues,
    )
