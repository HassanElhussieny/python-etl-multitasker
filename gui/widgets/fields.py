"""Input fields: a path with a browse button, a day-first date."""

from __future__ import annotations

from datetime import date

from PySide6.QtCore import QDate, QLocale, Qt
from PySide6.QtWidgets import QDateEdit, QFileDialog, QHBoxLayout, QLineEdit, QPushButton, QWidget


class PathField(QWidget):
    def __init__(self, kind: str, hint: str = "") -> None:
        super().__init__()
        self.kind = kind
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.edit = QLineEdit()
        self.edit.setPlaceholderText(hint)
        self.edit.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        browse = QPushButton("استعراض…")
        browse.clicked.connect(self._browse)
        lay.addWidget(self.edit, 1)
        lay.addWidget(browse)

    def _browse(self) -> None:
        if self.kind == "dir":
            path = QFileDialog.getExistingDirectory(self, "اختر مجلدًا", self.edit.text())
        else:
            path, _ = QFileDialog.getOpenFileName(self, "اختر ملفًا", self.edit.text())
        if path:
            self.edit.setText(path)

    def value(self) -> str:
        return self.edit.text().strip()


def date_field(day: date) -> QDateEdit:
    """dd/MM/yyyy with Latin digits. Direction and locale are set BEFORE the format: in the RTL application Qt
    would otherwise show the date sections in reverse order."""
    w = QDateEdit()
    w.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
    w.setLocale(QLocale(QLocale.Language.English, QLocale.Country.UnitedKingdom))
    w.setDisplayFormat("dd/MM/yyyy")
    w.setCalendarPopup(True)
    w.setDate(QDate(day.year, day.month, day.day))
    return w


def date_value(w: QDateEdit) -> date:
    q = w.date()
    return date(q.year(), q.month(), q.day())
