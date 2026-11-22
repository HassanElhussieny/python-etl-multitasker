"""الإعدادات: المجلدات ومفسّر Python وتفضيلات التطبيق. الثوابت التجارية تبقى في المحرك كما هي."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFormLayout, QHBoxLayout, QMessageBox, QPushButton, QVBoxLayout, QWidget

from gui.pages import AppContext
from gui.widgets.fields import PathField
from gui.services.settings import AppSettings, save_settings, settings_path
from gui.widgets.common import Banner, card, label, page_header

FIELDS = (("output_dir", "مجلد المخرجات (تُحفظ فيه التقارير)", "dir"),
          ("collection_dir", "مجلد التحصيل (المهمة 5 — الخيار 1)", "dir"),
          ("state_dir", "مجلد state (القفل العام: لا يبدأ التشغيل أثناء تشغيل آخر من سطر الأوامر)", "dir"),
          ("main_data", "ملف Main Data للمهمة 5 (اختياري — فارغ = ملف مجلد data)", "file"),
          ("project_root", "مجلد المشروع (sales_etl)", "dir"), ("python", "مفسّر Python", "file"))


class SettingsPage(QWidget):
    def __init__(self, ctx: AppContext) -> None:
        super().__init__()
        self.ctx = ctx
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(12)
        root.addWidget(page_header("الإعدادات", "القيم الافتراضية هي مجلدات المشروع نفسها التي يستخدمها "
                                                 "python -m sales_etl."))
        self.banner = Banner()
        root.addWidget(self.banner)
        box = card()
        form = QFormLayout(box)
        form.setContentsMargins(20, 18, 20, 18)
        form.setSpacing(12)
        self.fields: dict[str, PathField] = {}
        for key, text, kind in FIELDS:
            f = PathField(kind)
            self.fields[key] = f
            form.addRow(text, f)
        root.addWidget(box)
        root.addWidget(label(f"ملف الإعدادات: {settings_path()}", "Muted"))
        b = QHBoxLayout()
        save = QPushButton("حفظ")
        save.setObjectName("Primary")
        save.clicked.connect(self.save)
        reset = QPushButton("استعادة الافتراضي (مجلدات سطر الأوامر)")
        reset.clicked.connect(self.reset)
        b.addWidget(save)
        b.addWidget(reset)
        b.addStretch(1)
        root.addLayout(b)
        root.addStretch(1)
        self.load(ctx.settings)

    def load(self, s: AppSettings) -> None:
        for key, f in self.fields.items():
            f.edit.setText(getattr(s, key))
        self._check(s)

    def _check(self, s: AppSettings) -> None:
        missing = [k for k, _, kind in FIELDS if kind == "dir" and k != "state_dir" and not Path(getattr(s, k)).is_dir()]
        if s.main_data and not Path(s.main_data).is_file():
            missing.append("main_data")
        if not (Path(s.project_root) / "sales_etl").is_dir():
            self.banner.show_message("مجلد المشروع لا يحتوي على sales_etl.", "error")
        elif missing:
            self.banner.show_message("مجلدات غير موجودة: " + "، ".join(missing), "warn")
        else:
            self.banner.show_message("الإعدادات صالحة. الملفات المختارة تُنسخ مؤقتًا ولا يُستخدم مجلد input.", "ok")

    def save(self) -> None:
        if self.ctx.queue.running():
            QMessageBox.warning(self, "مهام قيد التشغيل", "لا يمكن تغيير الإعدادات أثناء تشغيل مهام.")
            return
        values = {k: f.value() for k, f in self.fields.items()}
        new = AppSettings(**values)
        for key in values:
            setattr(self.ctx.settings, key, getattr(new, key))
        save_settings(self.ctx.settings)
        self._check(self.ctx.settings)
        self.ctx.settings_changed.emit()

    def reset(self) -> None:
        self.load(AppSettings(project_root=self.ctx.settings.project_root))
