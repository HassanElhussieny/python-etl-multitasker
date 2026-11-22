"""التقارير: كل ملفات Excel الناتجة في مجلد المخرجات، مع الفتح بالبرنامج الافتراضي وفتح المجلد."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHBoxLayout, QLineEdit, QPushButton, QVBoxLayout, QWidget

from gui.pages import AppContext
from gui.services.history import load_reports
from gui.widgets.common import card, cell, fill_table, fmt_size, folder_signature, label, make_table, open_path, page_header


class ReportsPage(QWidget):
    def __init__(self, ctx: AppContext) -> None:
        super().__init__()
        self.ctx = ctx
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(12)
        root.addWidget(page_header("التقارير", "ملفات Excel التي أنشأها المحرك في مجلد المخرجات (الأحدث أولًا)."))
        box = card()
        bl = QVBoxLayout(box)
        bl.setContentsMargins(14, 12, 14, 12)
        top = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("بحث باسم التقرير أو المجلد أو المهمة…")
        self.search.textChanged.connect(self._filter)
        refresh = QPushButton("تحديث")
        refresh.clicked.connect(lambda: self.refresh())
        top.addWidget(self.search, 1)
        top.addWidget(refresh)
        bl.addLayout(top)
        self.count = label("", "Muted")
        bl.addWidget(self.count)
        self.table = make_table(["التقرير", "المهمة", "مجلد التشغيل", "تاريخ الإنشاء", "الحجم"])
        self.table.doubleClicked.connect(lambda _: self._open(False))
        bl.addWidget(self.table, 1)
        buttons = QHBoxLayout()
        o = QPushButton("فتح التقرير")
        o.setObjectName("Primary")
        o.clicked.connect(lambda: self._open(False))
        d = QPushButton("فتح المجلد")
        d.clicked.connect(lambda: self._open(True))
        buttons.addWidget(o)
        buttons.addWidget(d)
        buttons.addStretch(1)
        bl.addLayout(buttons)
        root.addWidget(box, 1)
        self._signature: tuple | None = None          # what the table shows; None = must reload
        ctx.settings_changed.connect(self._changed)
        ctx.queue.job_changed.connect(lambda _: self._changed())   # mark stale only: never reload while hidden

    def _changed(self) -> None:
        self._signature = None
        if self.isVisible():
            self.refresh_if_changed()

    def showEvent(self, event) -> None:  # noqa: N802 - Qt
        super().showEvent(event)
        self.refresh_if_changed()

    def _current_signature(self) -> tuple:
        return folder_signature(self.ctx.settings.path("output_dir"))

    def refresh_if_changed(self) -> bool:
        """Reload only when the output folder changed since the table was filled (a visit costs one listing)."""
        sig = self._current_signature()
        if sig == self._signature:
            return False
        self.refresh(sig)
        return True

    def refresh(self, signature: tuple | None = None) -> None:
        reports = load_reports(self.ctx.settings.path("output_dir"))
        fill_table(self.table, [[cell(rep.path.name, data=str(rep.path), bold=True), cell(rep.task_title),
                                 cell(rep.run_folder), cell(rep.modified.strftime("%d/%m/%Y %H:%M")),
                                 cell(fmt_size(rep.size))] for rep in reports])
        self._signature = signature if signature is not None else self._current_signature()
        self.count.setText(f"{len(reports)} تقرير")
        self._filter(self.search.text())

    def _filter(self, text: str) -> None:
        text = text.strip().lower()
        for r in range(self.table.rowCount()):
            row = " ".join(self.table.item(r, c).text().lower() for c in range(3))
            self.table.setRowHidden(r, bool(text) and text not in row)

    def _open(self, folder: bool) -> None:
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return
        path = self.table.item(rows[0].row(), 0).data(Qt.ItemDataRole.UserRole)
        from pathlib import Path
        open_path(Path(path).parent if folder else path, self)
