"""The application's pages."""

from __future__ import annotations

from PySide6.QtCore import QObject, Signal

from gui.services.jobs import JobQueue
from gui.services.settings import AppSettings


class AppContext(QObject):
    """What every page shares: the settings and the job queue."""

    settings_changed = Signal()
    navigate = Signal(str)          # page key

    def __init__(self, settings: AppSettings, queue: JobQueue) -> None:
        super().__init__()
        self.settings = settings
        self.queue = queue
