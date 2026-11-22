"""Double-click launcher for the desktop app (pythonw: no console window).

The command line is unchanged and still works: ``python -m sales_etl``.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from gui.main import main  # noqa: E402

sys.exit(main())
