"""Desktop-app tests: offscreen Qt, every folder and setting in the test's temporary folder."""

from __future__ import annotations

import os
import time

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication  # noqa: E402

from gui.services.settings import PROJECT_ROOT, AppSettings  # noqa: E402


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance() or QApplication([])
    yield app


@pytest.fixture
def gui_env(tmp_path, monkeypatch, qapp):
    """Temporary project folders + settings; the GUI's own files (settings, job log) in tmp too."""
    monkeypatch.setenv("ETL_GUI_SETTINGS", str(tmp_path / "gui_settings.json"))
    monkeypatch.setenv("ETL_GUI_HISTORY", str(tmp_path / "gui_jobs.jsonl"))
    monkeypatch.setenv("ETL_GUI_STAGING", str(tmp_path / "staging"))
    for d in ("input", "output", "processed", "failures", "state", "التحصيل", "data"):
        (tmp_path / d).mkdir()
    settings = AppSettings(project_root=str(PROJECT_ROOT), input_dir=str(tmp_path / "input"),
                           output_dir=str(tmp_path / "output"), processed_dir=str(tmp_path / "processed"),
                           failures_dir=str(tmp_path / "failures"), state_dir=str(tmp_path / "state"),
                           collection_dir=str(tmp_path / "التحصيل"))
    return tmp_path, settings


def wait_until(app, condition, timeout=120.0):
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        app.processEvents()
        if condition():
            return True
        time.sleep(0.02)
    app.processEvents()
    return condition()
