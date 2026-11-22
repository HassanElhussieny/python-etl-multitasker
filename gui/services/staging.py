"""Isolated staging folders for desktop runs.

The engine reads one input folder. For a desktop run the user's chosen files (from any folder) are COPIED into a
fresh folder of their own under ``%TEMP%\\PythonETLMultitasker\\staging\\<run>\\``; that folder is the run's
``--input-dir`` and the run uses ``--no-move``, so the engine never moves anything and the project's ``input``,
``processed``, ``failures`` and ``state`` folders are untouched. The originals are only read (SHA-256 recorded
before, checked after). Only folders created here, inside the staging root, are ever deleted - never a user file.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


def staging_root() -> Path:
    override = os.environ.get("ETL_GUI_STAGING")
    return Path(override) if override else Path(tempfile.gettempdir()) / "PythonETLMultitasker" / "staging"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


@dataclass
class StagedFile:
    source: Path
    staged: Path
    sha: str


@dataclass
class StagedRun:
    folder: Path
    files: list[StagedFile] = field(default_factory=list)

    @property
    def renamed(self) -> list[StagedFile]:
        return [f for f in self.files if f.staged.name != f.source.name]


class StagingError(RuntimeError):
    """A chosen file cannot be copied (missing, unreadable): nothing runs."""


def stage_files(sources: list[Path]) -> StagedRun:
    """Copy the chosen files into a new staging folder (the same name kept; two files of the same name from
    different folders get " (2)"). The originals are opened for reading only."""
    if not sources:
        raise StagingError("لم يتم اختيار أي ملف")
    root = staging_root()
    root.mkdir(parents=True, exist_ok=True)
    folder = Path(tempfile.mkdtemp(prefix=datetime.now().strftime("run_%Y%m%d_%H%M%S_"), dir=root))
    run = StagedRun(folder)
    try:
        for src in sources:
            src = Path(src)
            if not src.is_file():
                raise StagingError(f"الملف غير موجود: {src}")
            target = folder / src.name
            n = 1
            while target.exists():
                n += 1
                target = folder / f"{src.stem} ({n}){src.suffix}"
            digest = sha256(src)
            shutil.copy2(src, target)
            if sha256(target) != digest:
                raise StagingError(f"النسخة المؤقتة لا تطابق الأصل: {src}")
            run.files.append(StagedFile(src, target, digest))
    except (OSError, StagingError) as exc:
        remove_staging(folder)
        raise exc if isinstance(exc, StagingError) else StagingError(f"تعذّر نسخ الملف: {exc}") from exc
    return run


def changed_sources(run_files: list[tuple[str, str]]) -> list[str]:
    """The originals that are missing or differ from their SHA-256 at selection time (expected: none)."""
    changed = []
    for path, digest in run_files:
        p = Path(path)
        try:
            if sha256(p) != digest:
                changed.append(path)
        except OSError:
            changed.append(path)
    return changed


def _inside_root(folder: Path) -> bool:
    try:
        return folder.resolve().parent == staging_root().resolve()
    except OSError:
        return False


def remove_staging(folder: Path | str | None) -> bool:
    """Delete one staging folder - only a direct child of the staging root (never anything else)."""
    if not folder:
        return False
    folder = Path(folder)
    if not folder.is_dir() or not _inside_root(folder):
        return False
    shutil.rmtree(folder, ignore_errors=True)
    return not folder.exists()


def cleanup_stale(max_age_hours: float = 24.0) -> int:
    """Remove staging folders left by an earlier session (older than ``max_age_hours``)."""
    root = staging_root()
    if not root.is_dir():
        return 0
    removed = 0
    limit = time.time() - max_age_hours * 3600
    for folder in root.iterdir():
        try:
            if folder.is_dir() and folder.stat().st_mtime < limit and remove_staging(folder):
                removed += 1
        except OSError:
            continue
    return removed


@dataclass
class FolderCopy:
    """Result of «تحميل كل التقارير»: the files written, the renamed ones and the failures."""

    copied: list[Path] = field(default_factory=list)
    renamed: list[tuple[str, str]] = field(default_factory=list)   # (original name, name used in the folder)
    failed: list[tuple[str, str]] = field(default_factory=list)    # (report, reason)


def copy_reports_to_folder(reports: list[Path | str], folder: Path) -> FolderCopy:
    """«تحميل كل التقارير»: copy every given report into ``folder`` as its own .xlsx file, under its original name.
    A name already taken in the folder (or by an earlier report of this run) gets " (2)", " (3)"... - an existing
    file is NEVER overwritten (written to a .part file, then renamed only onto a name that does not exist).
    The reports themselves are only read. Every report is tried; failures are listed, never hidden."""
    result = FolderCopy()
    folder = Path(folder)
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        result.failed = [(str(r), f"تعذّر الوصول إلى المجلد: {exc}") for r in reports]
        return result
    for report in (Path(r) for r in reports):
        if not report.is_file():
            result.failed.append((str(report), "الملف غير موجود"))
            continue
        tmp = folder / f".{report.name}.{os.getpid()}.part"
        try:
            shutil.copy2(report, tmp)
            if sha256(tmp) != sha256(report):
                raise OSError("النسخة لا تطابق الأصل")
            n, target = 1, folder / report.name
            while True:
                if not target.exists():
                    try:
                        os.rename(tmp, target)        # Windows: fails if the target appeared meanwhile
                        break
                    except FileExistsError:
                        pass
                n += 1
                target = folder / f"{report.stem} ({n}){report.suffix}"
            result.copied.append(target)
            if target.name != report.name:
                result.renamed.append((report.name, target.name))
        except OSError as exc:
            result.failed.append((str(report), str(exc)))
        finally:
            if tmp.exists():
                tmp.unlink()
    return result


def save_copy(report: Path, destination: Path) -> Path:
    """"تحميل التقرير": copy a generated report to the place the user chose; the report stays in output."""
    report, destination = Path(report), Path(destination)
    if not report.is_file():
        raise FileNotFoundError(report)
    if destination.resolve() == report.resolve():
        raise ValueError("المكان المختار هو نفس ملف التقرير الأصلي")
    destination.parent.mkdir(parents=True, exist_ok=True)
    tmp = destination.with_name(destination.name + ".part")
    shutil.copy2(report, tmp)
    tmp.replace(destination)
    return destination
