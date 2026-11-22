"""مراقبة التنفيذ: قائمة الانتظار، الحالة، الوقت المنقضي، المخرجات الحية للمحرك، والتقارير الناتجة."""

from __future__ import annotations

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (QHBoxLayout, QListWidget, QListWidgetItem, QMessageBox, QPlainTextEdit, QPushButton,
                               QSplitter, QVBoxLayout, QWidget)

from gui.pages import AppContext
from gui.services.jobs import STATUS_AR, Job, Status
from gui.widgets.common import (STATUS_COLORS, Banner, card, cell, fmt_elapsed, fmt_time, label, make_table,
                                open_path, page_header)

COLUMNS = ["رقم التنفيذ", "المهمة", "الطريقة", "الحالة", "البدء", "الوقت المنقضي", "رمز الخروج", "ملاحظة"]


class MonitorPage(QWidget):
    def __init__(self, ctx: AppContext) -> None:
        super().__init__()
        self.ctx = ctx
        self.current: str | None = None
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(12)
        root.addWidget(page_header("مراقبة التنفيذ",
                                   "كل مهمة تعمل كعملية مستقلة من سطر الأوامر الحالي. لا توجد نسبة تقدم لأن المحرك لا "
                                   "يوفّرها؛ يُعرض الوقت المنقضي والمخرجات الحية."))
        split = QSplitter(Qt.Orientation.Vertical)
        top = card()
        tl = QVBoxLayout(top)
        tl.setContentsMargins(14, 12, 14, 12)
        self.table = make_table(COLUMNS)
        self.table.itemSelectionChanged.connect(self._select)
        tl.addWidget(self.table)
        bar = QHBoxLayout()
        self.cancel_btn = QPushButton("إلغاء / إيقاف المهمة المحددة")
        self.cancel_btn.setObjectName("Danger")
        self.cancel_btn.clicked.connect(self._cancel)
        clear = QPushButton("مسح المهام المنتهية من القائمة")
        clear.clicked.connect(self.ctx.queue.clear_finished)
        bar.addWidget(self.cancel_btn)
        bar.addWidget(clear)
        bar.addStretch(1)
        tl.addLayout(bar)
        split.addWidget(top)

        bottom = card()
        bl = QVBoxLayout(bottom)
        bl.setContentsMargins(14, 12, 14, 12)
        self.detail_title = label("اختر مهمة لعرض تفاصيلها", "SectionTitle")
        bl.addWidget(self.detail_title)
        self.banner = Banner()
        bl.addWidget(self.banner)
        inner = QHBoxLayout()
        self.log = QPlainTextEdit()
        self.log.setObjectName("Log")
        self.log.setReadOnly(True)
        self.log.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.log.setMaximumBlockCount(20000)
        inner.addWidget(self.log, 3)
        side = QVBoxLayout()
        side.addWidget(label("التقارير الناتجة", "SectionTitle"))
        self.reports = QListWidget()
        self.reports.itemDoubleClicked.connect(lambda it: open_path(it.data(Qt.ItemDataRole.UserRole), self))
        side.addWidget(self.reports, 1)
        open_file = QPushButton("فتح التقرير")
        open_file.clicked.connect(self._open_report)
        open_dir = QPushButton("فتح مجلد التشغيل")
        open_dir.clicked.connect(self._open_dir)
        side.addWidget(open_file)
        side.addWidget(open_dir)
        inner.addLayout(side, 1)
        bl.addLayout(inner, 1)
        split.addWidget(bottom)
        split.setSizes([300, 420])
        root.addWidget(split, 1)

        q = ctx.queue
        q.job_added.connect(self._job_added)
        q.job_changed.connect(self._job_changed)
        q.output.connect(self._output)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)
        self.timer.start(1000)

    # ------------------------------------------------------------------ table
    def _row_of(self, job_id: str) -> int:
        for r in range(self.table.rowCount()):
            if self.table.item(r, 0) and self.table.item(r, 0).text() == job_id:
                return r
        return -1

    def _fill(self, r: int, job: Job) -> None:
        color = STATUS_COLORS[job.status]
        values = [job.id, job.title, job.mode_text, STATUS_AR[job.status], fmt_time(job.started),
                  fmt_elapsed(job.elapsed()) if job.started else "—",
                  "—" if job.exit_code is None else job.exit_code, job.note]
        for c, v in enumerate(values):
            self.table.setItem(r, c, cell(v, color if c == 3 else None, bold=c == 3))

    def _job_added(self, job_id: str) -> None:
        job = self.ctx.queue.get(job_id)
        if job is None:
            return
        r = self.table.rowCount()
        self.table.insertRow(r)
        self._fill(r, job)
        if self.current is None:
            self.table.selectRow(r)

    def _job_changed(self, job_id: str) -> None:
        if not job_id:                                 # list cleared
            self.table.setRowCount(0)
            for job in self.ctx.queue.jobs:
                self._job_added(job.id)
            if self.current and self.ctx.queue.get(self.current) is None:
                self.current = None
                self.log.clear()
                self.reports.clear()
            return
        job = self.ctx.queue.get(job_id)
        r = self._row_of(job_id)
        if job is not None and r >= 0:
            self._fill(r, job)
            if job_id == self.current:
                self._show_detail(job, reload_log=False)

    def _tick(self) -> None:
        for job in self.ctx.queue.running():
            r = self._row_of(job.id)
            if r >= 0:
                self.table.setItem(r, 5, cell(fmt_elapsed(job.elapsed())))

    def _select(self) -> None:
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return
        job = self.ctx.queue.get(self.table.item(rows[0].row(), 0).text())
        if job is not None:
            self.current = job.id
            self._show_detail(job, reload_log=True)

    # ------------------------------------------------------------------ detail
    def _show_detail(self, job: Job, reload_log: bool) -> None:
        self.detail_title.setText(f"{job.id} — {job.title}")
        kind = {Status.COMPLETED: "ok", Status.FAILED: "error", Status.CANCELLED: "warn"}.get(job.status, "info")
        self.banner.show_message(f"{STATUS_AR[job.status]} — {job.note}" if job.note else STATUS_AR[job.status], kind)
        if reload_log:
            self.log.setPlainText(job.text)
            self.log.moveCursor(QTextCursor.MoveOperation.End)
        self.reports.clear()
        for path in job.reports:
            item = QListWidgetItem(path.replace("\\", "/").split("/")[-1])
            item.setData(Qt.ItemDataRole.UserRole, path)
            item.setToolTip(path)
            self.reports.addItem(item)
        self.cancel_btn.setEnabled(job.status in (Status.PENDING, Status.RUNNING))

    def _output(self, job_id: str, text: str) -> None:
        if job_id == self.current:
            self.log.moveCursor(QTextCursor.MoveOperation.End)
            self.log.insertPlainText(text)
            self.log.moveCursor(QTextCursor.MoveOperation.End)

    def _cancel(self) -> None:
        job = self.ctx.queue.get(self.current or "")
        if job is None:
            return
        if job.status is Status.RUNNING:
            extra = ("\n\nهذه المهمة تنقل الملفات. الإيقاف آمن بحكم سجل المحرك: قبل COMMIT تبقى الملفات في input، "
                     "وبعده يُكمل التشغيل التالي النقل تلقائيًا.") if job.moves_files else ""
            if QMessageBox.question(self, "إيقاف المهمة", f"إيقاف «{job.title}» أثناء التشغيل؟{extra}") \
                    != QMessageBox.StandardButton.Yes:
                return
        self.ctx.queue.cancel(job.id)

    def _open_report(self) -> None:
        item = self.reports.currentItem() or (self.reports.item(0) if self.reports.count() else None)
        if item is not None:
            open_path(item.data(Qt.ItemDataRole.UserRole), self)

    def _open_dir(self) -> None:
        job = self.ctx.queue.get(self.current or "")
        if job and job.output_dirs:
            open_path(job.output_dirs[0], self)
        elif job:
            QMessageBox.information(self, "لا يوجد مجلد", "لم يُنشئ هذا التشغيل مجلد مخرجات.")
