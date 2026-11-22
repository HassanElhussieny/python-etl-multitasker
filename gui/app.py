"""QApplication set-up: Arabic RTL, font, theme, icon."""

from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPainterPath, QPixmap
from PySide6.QtWidgets import QApplication

from gui import APP_NAME, __version__

RESOURCES = Path(__file__).resolve().parent / "resources"
if getattr(sys, "frozen", False):                      # PyInstaller: resources unpacked beside the app
    RESOURCES = Path(getattr(sys, "_MEIPASS", RESOURCES)) / "gui" / "resources"


def draw_icon(size: int = 256) -> QPixmap:
    """The application icon, drawn (no external image needed): a blue rounded square with rising bars."""
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    path = QPainterPath()
    path.addRoundedRect(0, 0, size, size, size * 0.22, size * 0.22)
    p.fillPath(path, QColor("#2563eb"))
    p.setPen(Qt.PenStyle.NoPen)
    p.setBrush(QColor("#ffffff"))
    w = size * 0.13
    for i, h in enumerate((0.30, 0.48, 0.66)):
        x = size * 0.20 + i * size * 0.22
        p.drawRoundedRect(int(x), int(size * (0.80 - h)), int(w), int(size * h), w / 3, w / 3)
    p.setBrush(QColor("#93c5fd"))
    p.drawEllipse(int(size * 0.66), int(size * 0.14), int(size * 0.16), int(size * 0.16))
    p.end()
    return pm


def app_icon() -> QIcon:
    ico = RESOURCES / "app.ico"
    return QIcon(str(ico)) if ico.is_file() else QIcon(draw_icon())


def create_application(argv: list[str] | None = None) -> QApplication:
    app = QApplication.instance() or QApplication(argv if argv is not None else sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(__version__)
    app.setOrganizationName("PythonETLMultitasker")
    app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    app.setFont(QFont("Segoe UI", 10))
    app.setStyle("Fusion")
    qss = RESOURCES / "style.qss"
    if qss.is_file():
        app.setStyleSheet(qss.read_text(encoding="utf-8"))
    app.setWindowIcon(app_icon())
    return app
