# PyInstaller spec: the desktop GUI only. The engine is NOT frozen into the .exe: every task still runs the
# project's own ``python -m sales_etl`` (the interpreter set in Settings), so there is one engine for both
# interfaces. Build:  powershell -ExecutionPolicy Bypass -File tools\build_exe.ps1
from pathlib import Path

ROOT = Path(SPECPATH).parent

a = Analysis(
    [str(ROOT / "gui" / "main.py")],
    pathex=[str(ROOT)],
    datas=[(str(ROOT / "gui" / "resources"), "gui/resources")],
    hiddenimports=["gui.pages.run", "gui.pages.monitor", "gui.pages.reports", "gui.services.staging", "gui.selftest",
                   "gui.pages.history", "gui.pages.settings", "sales_etl.ingest", "sales_etl.normalize"],
    excludes=["PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtPdf", "PySide6.Qt3DCore",
              "PySide6.QtMultimedia", "PySide6.QtQuick", "PySide6.QtQml", "tkinter", "pytest"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="Python ETL Multitasker", console=False,
          icon=str(ROOT / "gui" / "resources" / "app.ico"))
coll = COLLECT(exe, a.binaries, a.datas, name="Python ETL Multitasker")
