"""The main window: sidebar navigation + pages."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (QFrame, QHBoxLayout, QListWidget, QListWidgetItem, QMainWindow, QMessageBox,
                               QStackedWidget, QVBoxLayout, QWidget)

from gui import APP_NAME, APP_NAME_AR, __version__
from gui.pages import AppContext
from gui.pages.history import HistoryPage
from gui.pages.monitor import MonitorPage
from gui.pages.reports import ReportsPage
from gui.pages.run import RunPage
from gui.pages.settings import SettingsPage
from gui.services.jobs import JobQueue
from gui.services.settings import AppSettings
from gui.widgets.common import label

PAGES = (("run", "تشغيل تقرير"), ("reports", "التقارير"), ("history", "السجل"),
         ("monitor", "سجل التنفيذ (Logs)"), ("settings", "الإعدادات"))


class MainWindow(QMainWindow):
    def __init__(self, settings: AppSettings) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} — {APP_NAME_AR}")
        self.resize(1360, 860)
        self.setMinimumSize(1100, 700)
        self.queue = JobQueue(settings, parent=self)
        self.ctx = AppContext(settings, self.queue)

        central = QWidget()
        lay = QHBoxLayout(central)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(240)
        sl = QVBoxLayout(sidebar)
        sl.setContentsMargins(0, 0, 0, 0)
        sl.addWidget(label(APP_NAME, "Brand"))
        sl.addWidget(label(APP_NAME_AR, "BrandSub"))
        self.nav = QListWidget()
        self.nav.setObjectName("NavList")
        for key, text in PAGES:
            item = QListWidgetItem(text)
            item.setData(Qt.ItemDataRole.UserRole, key)
            self.nav.addItem(item)
        sl.addWidget(self.nav, 1)
        sl.addWidget(label(f"الإصدار {__version__}\nالمحرك: sales_etl (سطر الأوامر)", "SidebarFooter"))
        lay.addWidget(sidebar)

        content = QWidget()
        content.setObjectName("Content")
        cl = QVBoxLayout(content)
        cl.setContentsMargins(0, 0, 0, 0)
        self.stack = QStackedWidget()
        self.pages = {"run": RunPage(self.ctx), "reports": ReportsPage(self.ctx), "history": HistoryPage(self.ctx),
                      "monitor": MonitorPage(self.ctx), "settings": SettingsPage(self.ctx)}
        for key, _ in PAGES:
            self.stack.addWidget(self.pages[key])
        cl.addWidget(self.stack)
        lay.addWidget(content, 1)
        self.setCentralWidget(central)

        self.nav.currentRowChanged.connect(self.stack.setCurrentIndex)
        self.ctx.navigate.connect(self.go)
        self.nav.setCurrentRow(0)
        self.statusBar().showMessage("جاهز — كل تقرير يُشغَّل بمحرك سطر الأوامر الحالي على نسخ مؤقتة من الملفات المختارة.")

    def go(self, key: str) -> None:
        for i, (k, _) in enumerate(PAGES):
            if k == key:
                self.nav.setCurrentRow(i)

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802 - Qt
        running = self.queue.running()
        if running:
            movers = any(j.moves_files for j in running)
            text = (f"توجد {len(running)} مهمة قيد التشغيل. الانتظار حتى تنتهي ثم الإغلاق؟\n\n"
                    "«نعم»: ينتظر البرنامج انتهاءها.  «لا»: تُوقف الآن" +
                    (" (آمن بحكم سجل المحرك)." if movers else ".") + "  «إلغاء»: يبقى البرنامج مفتوحًا.")
            answer = QMessageBox.question(self, "مهام قيد التشغيل", text,
                                          QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No |
                                          QMessageBox.StandardButton.Cancel)
            if answer == QMessageBox.StandardButton.Cancel:
                event.ignore()
                return
            self.queue.shutdown(kill=answer == QMessageBox.StandardButton.No)
        event.accept()
