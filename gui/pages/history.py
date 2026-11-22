"""السجل: التشغيلات السابقة كما في مجلد المخرجات وسجلات state وملفات manifest/audit."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHBoxLayout, QListWidget, QListWidgetItem, QPushButton, QSplitter, QVBoxLayout, QWidget

from gui.pages import AppContext
from gui.services.history import STATUS_AR, Run, load_runs
from gui.widgets.common import RUN_STATUS_COLORS, card, cell, fill_table, folder_signature, label, make_table, open_path, page_header

SOURCE_AR = {"journal": "سجل state", "manifest": "run_manifest.json", "gui": "سجل البرنامج", "folder": "اسم المجلد"}
DISPOSITION_AR = {"PROCESSED": "نُقل إلى processed", "FAILED": "نُقل إلى failures", "RETAINED": "بقي في input"}


class HistoryPage(QWidget):
    def __init__(self, ctx: AppContext) -> None:
        super().__init__()
        self.ctx = ctx
        self.runs: list[Run] = []
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(12)
        root.addWidget(page_header("السجل", "كل التشغيلات السابقة — من سطر الأوامر أو من هذا البرنامج — كما سجّلها المحرك."))
        split = QSplitter(Qt.Orientation.Horizontal)
        left = card()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(14, 12, 14, 12)
        top = QHBoxLayout()
        self.count = label("", "Muted")
        refresh = QPushButton("تحديث")
        refresh.clicked.connect(lambda: self.refresh())
        top.addWidget(self.count, 1)
        top.addWidget(refresh)
        ll.addLayout(top)
        self.table = make_table(["المهمة", "مجلد التشغيل", "التاريخ", "الحالة", "التدقيق", "المصدر"])
        self.table.itemSelectionChanged.connect(self._select)
        ll.addWidget(self.table, 1)
        split.addWidget(left)

        right = card()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(14, 12, 14, 12)
        self.title = label("اختر تشغيلًا", "SectionTitle")
        rl.addWidget(self.title)
        rl.addWidget(label("ملفات الإدخال", "SectionTitle"))
        self.files = QListWidget()
        rl.addWidget(self.files, 1)
        rl.addWidget(label("التقارير", "SectionTitle"))
        self.reports = QListWidget()
        self.reports.itemDoubleClicked.connect(lambda it: open_path(it.data(Qt.ItemDataRole.UserRole), self))
        rl.addWidget(self.reports, 1)
        b = QHBoxLayout()
        o = QPushButton("فتح التقرير")
        o.clicked.connect(self._open_report)
        d = QPushButton("فتح مجلد التشغيل")
        d.clicked.connect(self._open_dir)
        b.addWidget(o)
        b.addWidget(d)
        rl.addLayout(b)
        split.addWidget(right)
        split.setSizes([760, 420])
        root.addWidget(split, 1)
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
        from gui.services.jobs import history_file
        s = self.ctx.settings
        return (folder_signature(s.path("output_dir")), folder_signature(s.path("state_dir"), patterns=("*.jsonl",)),
                folder_signature(history_file().parent, patterns=(history_file().name,)))

    def refresh_if_changed(self) -> bool:
        """Reload only when output, the journals or the GUI job log changed since the table was filled."""
        sig = self._current_signature()
        if sig == self._signature:
            return False
        self.refresh(sig)
        return True

    def refresh(self, signature: tuple | None = None) -> None:
        s = self.ctx.settings
        self.runs = load_runs(s.path("output_dir"), s.path("state_dir"))
        fill_table(self.table, [[cell(run.task_title), cell(run.folder.name, data=r),
                                 cell(run.when.strftime("%d/%m/%Y %H:%M")),
                                 cell(STATUS_AR.get(run.status, run.status), RUN_STATUS_COLORS.get(run.status), bold=True),
                                 cell(run.audit or "—", "#059669" if run.audit == "PASS" else
                                      ("#dc2626" if run.audit == "FAIL" else None)),
                                 cell(SOURCE_AR.get(run.source, run.source))] for r, run in enumerate(self.runs)])
        self._signature = signature if signature is not None else self._current_signature()
        self.count.setText(f"{len(self.runs)} تشغيل")

    def _current(self) -> Run | None:
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return None
        return self.runs[self.table.item(rows[0].row(), 1).data(Qt.ItemDataRole.UserRole)]

    def _select(self) -> None:
        run = self._current()
        if run is None:
            return
        self.title.setText(f"{run.task_title} — {run.folder.name}")
        self.files.clear()
        for f in run.files:
            disp = DISPOSITION_AR.get(f.get("disposition") or "", f.get("disposition") or "")
            self.files.addItem(f"{f.get('name')}   {('— ' + disp) if disp else ''}")
        if not run.files:
            self.files.addItem("لا توجد بيانات ملفات (تشغيل قراءة فقط أو مخرجات قديمة)")
        self.reports.clear()
        for p in run.reports:
            item = QListWidgetItem(p.name)
            item.setData(Qt.ItemDataRole.UserRole, str(p))
            self.reports.addItem(item)

    def _open_report(self) -> None:
        item = self.reports.currentItem() or (self.reports.item(0) if self.reports.count() else None)
        if item:
            open_path(item.data(Qt.ItemDataRole.UserRole), self)

    def _open_dir(self) -> None:
        run = self._current()
        if run:
            open_path(run.folder, self)
