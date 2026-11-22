"""Application settings: the folders and the Python interpreter the engine runs with.

Stored as JSON in ``%APPDATA%\\PythonETLMultitasker\\settings.json`` (or the file named by the environment variable
``ETL_GUI_SETTINGS``, used by the tests). The defaults are exactly the command line's defaults, so the GUI and
``python -m sales_etl`` work on the same folders unless the user changes them here.
"""

from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass, fields
from pathlib import Path

def _find_project_root() -> Path:
    """The folder that contains ``sales_etl``: beside this package normally; for the packaged .exe, the first
    parent of the executable that has it (dist\\Python ETL Multitasker\\ inside the project)."""
    if getattr(sys, "frozen", False):
        for parent in Path(sys.executable).resolve().parents:
            if (parent / "sales_etl" / "cli.py").is_file():
                return parent
    return Path(__file__).resolve().parents[2]


PROJECT_ROOT = _find_project_root()


def app_data_dir() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home())
    return Path(base) / "PythonETLMultitasker"


def settings_path() -> Path:
    override = os.environ.get("ETL_GUI_SETTINGS")
    return Path(override) if override else app_data_dir() / "settings.json"


def default_python() -> str:
    """The interpreter for the engine: this one (python.exe, not pythonw.exe), when not frozen."""
    exe = Path(sys.executable)
    if getattr(sys, "frozen", False):                  # the .exe has no engine interpreter of its own
        import shutil

        found = shutil.which("python")
        known = Path.home() / "AppData" / "Local" / "Programs" / "Python" / "Python314" / "python.exe"
        return str(known) if known.is_file() else (found or "python")
    if exe.name.lower() == "pythonw.exe" and (exe.parent / "python.exe").is_file():
        return str(exe.parent / "python.exe")
    return str(exe)


@dataclass
class AppSettings:
    project_root: str = str(PROJECT_ROOT)        # where ``sales_etl`` lives: the working folder of every run
    python: str = ""                             # interpreter that runs ``-m sales_etl``
    input_dir: str = ""
    output_dir: str = ""
    processed_dir: str = ""
    failures_dir: str = ""
    state_dir: str = ""
    collection_dir: str = ""                     # Task 5 option 1: the التحصيل folder
    main_data: str = ""                          # Task 5: Main Data workbook ("" = the CLI's own default in data)
    max_parallel_readonly: int = 2               # read-only jobs that may run together
    refresh_seconds: int = 5                     # kept for settings-file compatibility

    def __post_init__(self) -> None:
        root = Path(self.project_root)
        defaults = {"python": default_python(), "input_dir": root / "input", "output_dir": root / "output",
                    "processed_dir": root / "processed", "failures_dir": root / "failures",
                    "state_dir": root / "state", "collection_dir": root / "التحصيل"}
        for name, value in defaults.items():
            if not getattr(self, name):
                setattr(self, name, str(value))
        self.max_parallel_readonly = max(1, int(self.max_parallel_readonly))
        self.refresh_seconds = max(2, int(self.refresh_seconds))

    # convenient Path views
    @property
    def root(self) -> Path:
        return Path(self.project_root)

    def path(self, name: str) -> Path:
        return Path(getattr(self, name))

    @property
    def lock_file(self) -> Path:
        return self.path("state_dir") / "etl.lock"

    def is_default_layout(self) -> bool:
        """True when processed / failures / state are the CLI's defaults (beside the input folder)."""
        base = self.path("input_dir").parent
        return all(self.path(n) == base / d for n, d in
                   (("processed_dir", "processed"), ("failures_dir", "failures"), ("state_dir", "state")))


def load_settings(path: Path | None = None) -> AppSettings:
    path = path or settings_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        known = {f.name for f in fields(AppSettings)}
        return AppSettings(**{k: v for k, v in data.items() if k in known})
    except (OSError, ValueError, TypeError):
        return AppSettings()


def save_settings(settings: AppSettings, path: Path | None = None) -> Path:
    path = path or settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(asdict(settings), ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)
    return path
