"""Entry point: ``python -m sales_etl`` or VS Code "Run Python File" on this file.

Running this file directly puts ``sales_etl/`` (not the project folder) at the
front of ``sys.path``, so ``import sales_etl`` would fail. In that case the
package folder is replaced by the project folder before importing.
"""

import sys
from pathlib import Path

if not __package__:  # executed as a script, not as part of the package
    _package_dir = Path(__file__).resolve().parent
    sys.path = [p for p in sys.path if Path(p or ".").resolve() != _package_dir]
    sys.path.insert(0, str(_package_dir.parent))

from sales_etl.cli import main  # noqa: E402 - must follow the path fix above

if __name__ == "__main__":
    sys.exit(main())
