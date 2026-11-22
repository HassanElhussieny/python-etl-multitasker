"""تشغيل تقرير: اختر المهمة ← اختر الملف (من أي مجلد) ← تشغيل ← تحميل التقرير.

The chosen files are copied into an isolated staging folder (gui.services.staging) and the existing engine runs on
that folder with ``--no-move``: the originals and the project's input / processed / failures / state folders are
never touched. Reports are written to the normal output folder; "تحميل التقرير" saves a copy where the user wants.
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from PySide6.QtCore import QStandardPaths, Qt, QTimer
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (QDateEdit, QFileDialog, QFormLayout, QFrame, QGridLayout, QHBoxLayout, QLabel,
                               QListWidget, QListWidgetItem, QMessageBox, QPlainTextEdit, QProgressBar, QPushButton,
                               QScrollArea, QVBoxLayout, QWidget)

from gui.pages import AppContext
from gui.services.catalog import TASKS, ArgumentError, TaskDef, build_arguments
from gui.services.jobs import STATUS_AR, Job, Status
from gui.services.scanner import InputFile, check_files, file_info, inspect_file
from gui.services.staging import StagingError, copy_reports_to_folder, remove_staging, save_copy, stage_files
from gui.widgets.common import Banner, card, cell, fmt_elapsed, fmt_size, label, make_table, open_path, page_header
from gui.widgets.fields import PathField, date_field, date_value


class RunPage(QWidget):
    def __init__(self, ctx: AppContext) -> None:
        super().__init__()
        self.ctx = ctx
        self.files: list[Path] = []
        self._info: dict[Path, InputFile] = {}
        self.fields: dict[str, QWidget] = {}
        self.job: Job | None = None
        self.last_dir = str(Path.home())

        root = QVBoxLayout(self)
        root.setContentsMargins(28, 22, 28, 22)
        root.setSpacing(12)
        root.addWidget(page_header("تشغيل تقرير", "اختر المهمة ← اختر الملف ← تشغيل التقرير ← تحميل التقرير"))
        body = QHBoxLayout()
        body.setSpacing(16)

        # ---- the task list (one task at a time)
        self.task_list = QListWidget()
        self.task_list.setObjectName("TaskList")
        self.task_list.setFixedWidth(330)
        self.task_list.setWordWrap(True)
        self.task_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        for t in TASKS:
            item = QListWidgetItem(f"{t.title}\n{t.subtitle}")
            item.setData(Qt.ItemDataRole.UserRole, t.key)
            self.task_list.addItem(item)
        body.addWidget(self.task_list)

        # ---- the selected task: description, parameters, files, run, results
        panel = QWidget()
        pl = QVBoxLayout(panel)
        pl.setContentsMargins(0, 0, 0, 0)
        pl.setSpacing(12)

        info = card()
        il = QVBoxLayout(info)
        il.setContentsMargins(20, 16, 20, 16)
        self.title = label("", "SectionTitle")
        self.subtitle = label("", "Muted")
        self.description = label("", wrap=True)
        il.addWidget(self.title)
        il.addWidget(self.subtitle)
        il.addWidget(self.description)
        pl.addWidget(info)

        self.params_card = card()
        self.params_form = QFormLayout(self.params_card)
        self.params_form.setContentsMargins(20, 14, 20, 14)
        self.params_form.setSpacing(10)
        pl.addWidget(self.params_card)

        files = card()
        fl = QVBoxLayout(files)
        fl.setContentsMargins(20, 14, 20, 14)
        top = QHBoxLayout()
        self.pick_btn = QPushButton("اختيار الملف")
        self.pick_btn.setObjectName("Primary")
        self.pick_btn.setMinimumHeight(38)
        self.pick_btn.clicked.connect(self.pick_files)
        self.file_hint = label("", "Muted")
        top.addWidget(self.pick_btn)
        top.addWidget(self.file_hint, 1)
        fl.addLayout(top)
        self.file_list = make_table(["الملف", "الصفوف", "من", "إلى", "الحجم", "المجلد"])
        self.file_list.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self.file_list.setMinimumHeight(110)
        self.file_list.setMaximumHeight(170)
        fl.addWidget(self.file_list)
        fb = QHBoxLayout()
        self.remove_btn = QPushButton("إزالة الملف المحدد")
        self.remove_btn.clicked.connect(self.remove_selected)
        self.clear_btn = QPushButton("مسح كل الملفات")
        self.clear_btn.clicked.connect(self.clear_files)
        fb.addWidget(self.remove_btn)
        fb.addWidget(self.clear_btn)
        fb.addStretch(1)
        fl.addLayout(fb)
        self.warnings = Banner()
        fl.addWidget(self.warnings)
        pl.addWidget(files)

        run = card()
        rl = QVBoxLayout(run)
        rl.setContentsMargins(20, 14, 20, 14)
        rb = QHBoxLayout()
        self.run_btn = QPushButton("تشغيل التقرير")
        self.run_btn.setObjectName("Primary")
        self.run_btn.setMinimumHeight(42)
        self.run_btn.setMinimumWidth(200)
        self.run_btn.clicked.connect(self.run)
        self.stop_btn = QPushButton("إيقاف")
        self.stop_btn.setObjectName("Danger")
        self.stop_btn.clicked.connect(self.stop)
        self.stop_btn.hide()
        self.log_btn = QPushButton("عرض السجل")
        self.log_btn.setCheckable(True)
        rb.addWidget(self.run_btn)
        rb.addWidget(self.stop_btn)
        rb.addStretch(1)
        rb.addWidget(self.log_btn)
        rl.addLayout(rb)
        self.status = Banner()
        rl.addWidget(self.status)
        self.busy = QProgressBar()
        self.busy.setRange(0, 0)                      # "working" indicator only: the engine gives no percentage
        self.busy.setTextVisible(False)
        self.busy.setMaximumHeight(8)
        self.busy.hide()
        rl.addWidget(self.busy)
        self.log = QPlainTextEdit()
        self.log.setObjectName("Log")
        self.log.setReadOnly(True)
        self.log.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.log.setMinimumHeight(180)
        self.log.hide()
        self.log_btn.toggled.connect(self.log.setVisible)
        rl.addWidget(self.log)
        pl.addWidget(run)

        self.results = card()
        self.results_lay = QVBoxLayout(self.results)
        self.results_lay.setContentsMargins(20, 14, 20, 14)
        self.results.hide()
        pl.addWidget(self.results)
        pl.addStretch(1)

        self.scroll = scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(panel)
        body.addWidget(scroll, 1)
        root.addLayout(body, 1)

        self.task_list.currentRowChanged.connect(self.show_task)
        ctx.queue.job_changed.connect(self._job_changed)
        ctx.queue.output.connect(self._output)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)
        self.task_list.setCurrentRow(0)

    # ------------------------------------------------------------------ task
    @property
    def task(self) -> TaskDef:
        return TASKS[max(self.task_list.currentRow(), 0)]

    def show_task(self, row: int) -> None:
        """The selected task and the details shown always match (one task at a time)."""
        if row < 0:
            return
        t = TASKS[row]
        self.title.setText(t.title)
        self.subtitle.setText(f"{t.subtitle}  ·  المهمة {t.task}" + (f"  ·  الخيار {t.report_type}" if t.report_type else ""))
        self.description.setText(t.description)
        self.file_hint.setText(t.file_hint)
        while self.params_form.rowCount():
            self.params_form.removeRow(0)
        self.fields = {}
        yesterday = date.today() - timedelta(days=1)
        for p in t.visible_params:
            if p.kind == "date":
                w = date_field(yesterday)
                w.dateChanged.connect(lambda _=None: self._refresh_warnings())
            else:
                w = PathField(p.kind, p.hint)
            self.fields[p.key] = w
            self.params_form.addRow(p.label, w)
        self.params_card.setVisible(bool(t.visible_params))
        self.results.hide()
        self.status.show_message("", "info")
        self._refresh_warnings()

    def values(self) -> dict[str, object]:
        out: dict[str, object] = {}
        for key, w in self.fields.items():
            out[key] = date_value(w) if isinstance(w, QDateEdit) else w.value()
        return out

    # ------------------------------------------------------------------ files
    def pick_files(self) -> None:
        t = self.task
        patterns = " ".join(f"*{x}" for x in t.file_types)
        title = "ملفات المبيعات والمستحقات" if len(t.file_types) > 1 else "ملفات CSV"
        paths, _ = QFileDialog.getOpenFileNames(self, "اختيار الملف", self.last_dir, f"{title} ({patterns})")
        if paths:
            self.last_dir = str(Path(paths[0]).parent)
            self.add_files([Path(p) for p in paths])

    def add_files(self, paths: list[Path]) -> None:
        for p in paths:
            p = Path(p)
            if p not in self.files and p.is_file():
                self.files.append(p)
        self._render_files()

    def remove_selected(self) -> None:
        rows = sorted({i.row() for i in self.file_list.selectionModel().selectedRows()}, reverse=True)
        if not rows and 0 <= self.file_list.currentRow() < len(self.files):
            rows = [self.file_list.currentRow()]
        for row in rows:
            if 0 <= row < len(self.files):
                self.files.pop(row)              # only removed from the list: the file itself is untouched
        if rows:
            self._render_files()

    def clear_files(self) -> None:
        self.files = []
        self._render_files()

    def _inspect(self, p: Path) -> InputFile:
        if p not in self._info:
            info = file_info(p)
            self._info[p] = inspect_file(info) if p.suffix.lower() == ".csv" else info
        return self._info[p]

    def _render_files(self) -> None:
        self.file_list.setRowCount(len(self.files))
        for r, p in enumerate(self.files):
            f = self._inspect(p)
            values = [p.name, "تعذّرت قراءته" if f.error else ("—" if f.rows is None else f"{f.rows:,}"),
                      f.first_date.strftime("%d/%m/%Y") if f.first_date else "—",
                      f.last_date.strftime("%d/%m/%Y") if f.last_date else "—", fmt_size(f.size), str(p.parent)]
            for c, v in enumerate(values):
                item = cell(v, "#dc2626" if (c == 1 and f.error) else None, bold=c == 0)
                item.setToolTip(str(p))
                self.file_list.setItem(r, c, item)
        self.remove_btn.setEnabled(bool(self.files))
        self.clear_btn.setEnabled(bool(self.files))
        self._refresh_warnings()

    def _refresh_warnings(self) -> None:
        if not self.files:
            self.warnings.show_message("", "info")
            return
        t = self.task
        notes = check_files([self._inspect(p) for p in self.files], t.key, self.values()).warnings
        wrong = [p.name for p in self.files if p.suffix.lower() not in t.file_types]
        if wrong:
            notes.insert(0, "ملفات لا تناسب هذه المهمة: " + "، ".join(wrong))
        names = [p.name for p in self.files]
        if len(set(names)) != len(names):
            notes.append("ملفان بنفس الاسم من مجلدين مختلفين: سيُضاف (2) لاسم النسخة المؤقتة للثاني.")
        self.warnings.show_message("\n".join("• " + n for n in notes), "warn" if notes else "info")

    # ------------------------------------------------------------------ run
    def run(self) -> Job | None:
        t = self.task
        if self.job is not None and self.job.status in (Status.PENDING, Status.RUNNING):
            return None
        if not self.files:
            QMessageBox.information(self, "اختيار الملف", "اختر ملفًا واحدًا على الأقل بزر «اختيار الملف».")
            return None
        wrong = [p.name for p in self.files if p.suffix.lower() not in t.file_types]
        if wrong:
            QMessageBox.warning(self, "ملفات غير مناسبة", "هذه الملفات لا تناسب المهمة المختارة:\n" + "\n".join(wrong))
            return None
        values = self.values()
        try:
            build_arguments(t.key, values, self.ctx.settings, read_only=True, input_dir=str(Path.home()))  # validate
            staged = stage_files(self.files)
        except (ArgumentError, StagingError) as exc:
            QMessageBox.warning(self, "تعذّر التشغيل", str(exc))
            return None
        try:
            args = build_arguments(t.key, values, self.ctx.settings, read_only=True, input_dir=str(staged.folder))
        except ArgumentError as exc:
            remove_staging(staged.folder)
            QMessageBox.warning(self, "تعذّر التشغيل", str(exc))
            return None
        job = Job(t.key, args, True, files_at_enqueue=[p.name for p in self.files],
                  staging_dir=str(staged.folder), sources=[(str(f.source), f.sha) for f in staged.files])
        self.job = job
        self.log.clear()
        self.results.hide()
        self._set_running(True)
        self.status.show_message(f"جارٍ التشغيل: {t.title} — {len(self.files)} ملف", "info")
        self.ctx.queue.add(job)
        return job

    def stop(self) -> None:
        if self.job and self.job.status in (Status.PENDING, Status.RUNNING):
            if QMessageBox.question(self, "إيقاف", "إيقاف التقرير الجاري؟ الملفات الأصلية لا تتأثر.") \
                    == QMessageBox.StandardButton.Yes:
                self.ctx.queue.cancel(self.job.id)

    def _set_running(self, running: bool) -> None:
        for w in (self.run_btn, self.pick_btn, self.remove_btn, self.clear_btn, self.task_list):
            w.setEnabled(not running)
        if not running:
            self.remove_btn.setEnabled(bool(self.files))
            self.clear_btn.setEnabled(bool(self.files))
        self.stop_btn.setVisible(running)
        self.busy.setVisible(running)
        if running:
            self.timer.start(1000)
        else:
            self.timer.stop()

    def _tick(self) -> None:
        if self.job and self.job.status is Status.RUNNING:
            self.status.show_message(f"جارٍ التشغيل: {self.job.title} — الوقت المنقضي {fmt_elapsed(self.job.elapsed())}",
                                     "info")

    def _output(self, job_id: str, text: str) -> None:
        if self.job and job_id == self.job.id:
            self.log.moveCursor(QTextCursor.MoveOperation.End)
            self.log.insertPlainText(text)

    def _job_changed(self, job_id: str) -> None:
        job = self.job
        if job is None or job_id != job.id:
            return
        if job.status is Status.PENDING and job.note:
            self.status.show_message(f"بالانتظار: {job.note}", "warn")
            return
        if job.status in (Status.PENDING, Status.RUNNING):
            return
        self._set_running(False)
        self.log.setPlainText(job.text)
        self.show_results(job)

    # ------------------------------------------------------------------ results
    def show_results(self, job: Job) -> None:
        while self.results_lay.count():
            item = self.results_lay.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        ok = job.status is Status.COMPLETED
        if ok:
            self.status.show_message(f"تم إنشاء التقرير بنجاح ({fmt_elapsed(job.elapsed())}).", "ok")
        elif job.status is Status.CANCELLED:
            self.status.show_message("أُوقف التشغيل. الملفات الأصلية لم تتأثر.", "warn")
        else:
            self.status.show_message(f"{STATUS_AR[job.status]}: {job.note} — اضغط «عرض السجل» للتفاصيل.", "error")
        self.all_btn = QPushButton("تحميل كل التقارير")
        self.all_btn.setObjectName("Primary")
        self.all_btn.setMinimumHeight(38)
        self.all_btn.clicked.connect(self.download_all)
        if not job.reports:
            if job.status is not Status.CANCELLED:
                self.results_lay.addWidget(label("لم يُنشأ أي تقرير Excel.", "Muted"))
                self.all_btn.setEnabled(False)
                self.all_btn.setToolTip("لا توجد تقارير لهذا التشغيل")
                self.results_lay.addWidget(self.all_btn, 0, Qt.AlignmentFlag.AlignLeading)
                self.results.show()
            return
        head = QHBoxLayout()
        head.addWidget(label(f"التقارير الناتجة ({len(job.reports)})", "SectionTitle"))
        head.addStretch(1)
        head.addWidget(self.all_btn)                 # every report of this run into one chosen folder
        head_w = QWidget()
        head_w.setLayout(head)
        self.results_lay.addWidget(head_w)
        if not ok:
            self.results_lay.addWidget(label("التدقيق أظهر مشاكل — راجع السجل قبل استخدام التقارير.", "Muted", wrap=True))
        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)
        for r, path in enumerate(job.reports):
            p = Path(path)
            name = QLabel(p.name)
            name.setToolTip(str(p))
            name.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            save = QPushButton("تحميل التقرير")
            save.setObjectName("Primary")
            save.clicked.connect(lambda _=False, x=p: self.download(x))
            open_btn = QPushButton("فتح")
            open_btn.clicked.connect(lambda _=False, x=p: open_path(x, self))
            grid.addWidget(name, r, 0)
            grid.addWidget(save, r, 1)
            grid.addWidget(open_btn, r, 2)
        grid.setColumnStretch(0, 1)
        holder = QWidget()
        holder.setLayout(grid)
        self.results_lay.addWidget(holder)
        folder = QPushButton("فتح مجلد المخرجات")
        target = job.output_dirs[0] if job.output_dirs else self.ctx.settings.output_dir
        folder.clicked.connect(lambda _=False, x=target: open_path(x, self))
        row = QHBoxLayout()
        row.addWidget(folder)
        row.addStretch(1)
        wrap = QWidget()
        wrap.setLayout(row)
        self.results_lay.addWidget(wrap)
        self.results.show()
        QTimer.singleShot(50, lambda: self.scroll.ensureWidgetVisible(self.results))   # the reports in view

    def download_all(self) -> Path | None:
        """«تحميل كل التقارير»: every report of the CURRENT run copied, each as its own .xlsx file, into a folder
        the user selects. The reports stay in output untouched; cancelling the dialog does nothing."""
        job = self.job
        reports = list(job.reports) if job else []
        if not reports:
            QMessageBox.information(self, "تحميل كل التقارير", "لا توجد تقارير لهذا التشغيل.")
            return None
        start = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DownloadLocation) or str(Path.home())
        folder = QFileDialog.getExistingDirectory(self, "تحميل كل التقارير — اختر مجلد الحفظ", start)
        if not folder:
            return None                              # cancelled: nothing written, no message
        result = copy_reports_to_folder(reports, Path(folder))
        renamed = "".join(f"\n• {a} ← حُفظ باسم {b} (يوجد ملف بنفس الاسم)" for a, b in result.renamed)
        if result.failed:
            failed = "\n".join(f"• {Path(r).name}: {why}" for r, why in result.failed)
            self.status.show_message(f"تم تحميل {len(result.copied)} من {len(reports)} تقرير فقط — فشل تحميل:\n"
                                     f"{failed}", "error")
            QMessageBox.warning(self, "لم تُحمَّل كل التقارير", f"فشل تحميل:\n{failed}")
        else:
            self.status.show_message(f"تم تحميل جميع التقارير بنجاح — {len(result.copied)} ملفات Excel.\n"
                                     f"المجلد: {folder}{renamed}", "ok")
        return Path(folder)

    def download(self, report: Path) -> Path | None:
        """Save As: a copy of the report where the user chooses; the report stays in output."""
        start = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DownloadLocation) or str(Path.home())
        dest, _ = QFileDialog.getSaveFileName(self, "تحميل التقرير", str(Path(start) / report.name),
                                              "ملفات Excel (*.xlsx)")
        if not dest:
            return None
        if not dest.lower().endswith(".xlsx"):
            dest += ".xlsx"
        try:
            saved = save_copy(report, Path(dest))
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "تعذّر الحفظ", str(exc))
            return None
        self.status.show_message(f"تم حفظ نسخة من التقرير في: {saved}", "ok")
        return saved
