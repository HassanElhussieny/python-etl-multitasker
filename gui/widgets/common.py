"""Small building blocks shared by the pages."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtWidgets import (QAbstractItemView, QFrame, QHBoxLayout, QHeaderView, QLabel, QMessageBox,
                               QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from gui.services.jobs import Status

STATUS_COLORS = {Status.PENDING: "#6b7280", Status.RUNNING: "#2563eb", Status.COMPLETED: "#059669",
                 Status.FAILED: "#dc2626", Status.CANCELLED: "#d97706"}
RUN_STATUS_COLORS = {"SUCCEEDED": "#059669", "COMPLETED_WITH_PROBLEMS": "#d97706", "FAILED": "#dc2626",
                     "INTERRUPTED": "#dc2626", "READ_ONLY": "#2563eb"}


def card() -> QFrame:
    frame = QFrame()
    frame.setObjectName("Card")
    return frame


def label(text: str, name: str = "", wrap: bool = False) -> QLabel:
    lab = QLabel(text)
    if name:
        lab.setObjectName(name)
    lab.setWordWrap(wrap)
    return lab


def page_header(title: str, subtitle: str) -> QWidget:
    w = QWidget()
    lay = QVBoxLayout(w)
    lay.setContentsMargins(0, 0, 0, 6)
    lay.setSpacing(2)
    lay.addWidget(label(title, "PageTitle"))
    lay.addWidget(label(subtitle, "PageSubtitle", wrap=True))
    return w


class StatCard(QFrame):
    def __init__(self, title: str, color: str = "#2563eb") -> None:
        super().__init__()
        self.setObjectName("Card")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 14)
        bar = QFrame()
        bar.setFixedHeight(4)
        bar.setStyleSheet(f"background:{color}; border-radius:2px;")
        self.value = label("—", "StatValue")
        lay.addWidget(bar)
        lay.addWidget(self.value)
        lay.addWidget(label(title, "StatLabel"))

    def set(self, value: object) -> None:
        self.value.setText(str(value))


class Banner(QLabel):
    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("Banner")
        self.setWordWrap(True)
        self.hide()

    def show_message(self, text: str, kind: str = "info") -> None:
        self.setProperty("kind", kind)
        self.style().unpolish(self)
        self.style().polish(self)
        self.setText(text)
        self.setVisible(bool(text))


def make_table(headers: list[str]) -> QTableWidget:
    t = QTableWidget(0, len(headers))
    t.setHorizontalHeaderLabels(headers)
    t.verticalHeader().setVisible(False)
    t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    t.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    t.setAlternatingRowColors(False)
    t.setShowGrid(False)
    t.horizontalHeader().setStretchLastSection(True)
    t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    return t


def fill_table(table: QTableWidget, rows: list[list[QTableWidgetItem]]) -> None:
    """Replace a table's rows in one pass. While filling, painting is paused and the columns are NOT resized to
    their contents after every cell (that made a refill of a few hundred rows take 10-40 s); they are sized once
    at the end."""
    header = table.horizontalHeader()
    table.setUpdatesEnabled(False)
    table.setSortingEnabled(False)
    for c in range(table.columnCount()):
        header.setSectionResizeMode(c, QHeaderView.ResizeMode.Interactive)
    try:
        table.clearContents()
        table.setRowCount(len(rows))
        for r, items in enumerate(rows):
            for c, item in enumerate(items):
                table.setItem(r, c, item)
        table.resizeColumnsToContents()        # sized once; the columns stay user-resizable (Interactive)
    finally:
        table.setUpdatesEnabled(True)


def folder_signature(*folders: Path, patterns: tuple[str, ...] = ("*",)) -> tuple:
    """A cheap fingerprint of what a page shows (names + modification times of the folders' direct entries): a
    page refreshes only when it changed, not on every visit."""
    sig = []
    for folder in folders:
        try:
            for pattern in patterns:
                for p in folder.glob(pattern):
                    sig.append((str(p), p.stat().st_mtime_ns))
        except OSError:
            sig.append((str(folder), None))
    return tuple(sorted(sig, key=lambda x: x[0]))


def cell(text: object, color: str | None = None, data: object = None, bold: bool = False) -> QTableWidgetItem:
    item = QTableWidgetItem("" if text is None else str(text))
    if color:
        item.setForeground(QColor(color))
    if bold:
        f = item.font()
        f.setBold(True)
        item.setFont(f)
    if data is not None:
        item.setData(Qt.ItemDataRole.UserRole, data)
    return item


def fmt_time(value: datetime | None) -> str:
    return value.strftime("%d/%m/%Y %H:%M:%S") if value else "—"


def fmt_elapsed(seconds: float) -> str:
    s = int(seconds)
    return f"{s // 3600:02d}:{s % 3600 // 60:02d}:{s % 60:02d}"


def fmt_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:,.0f} {unit}" if unit == "B" else f"{n:,.1f} {unit}"
        n /= 1024
    return str(n)


def open_path(path: Path | str, parent: QWidget | None = None) -> bool:
    """Open a file with its Windows application, or a folder in Explorer."""
    p = Path(path)
    if not p.exists():
        QMessageBox.warning(parent, "غير موجود", f"المسار غير موجود:\n{p}")
        return False
    return QDesktopServices.openUrl(QUrl.fromLocalFile(str(p)))


def hbox(*widgets: QWidget, stretch_end: bool = False) -> QHBoxLayout:
    lay = QHBoxLayout()
    for w in widgets:
        lay.addWidget(w)
    if stretch_end:
        lay.addStretch(1)
    return lay
