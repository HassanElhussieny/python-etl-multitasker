"""Write gui/resources/app.ico and app.png from the drawn application icon (run once; files are kept)."""

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication  # noqa: E402

from gui.app import draw_icon  # noqa: E402

app = QApplication([])
res = ROOT / "gui" / "resources"
pm = draw_icon(256)
assert pm.save(str(res / "app.png"), "PNG")
assert pm.save(str(res / "app.ico"), "ICO")
print("written:", res / "app.ico", res / "app.png")
