"""Start the desktop application: ``python -m gui`` (or the launcher / desktop shortcut)."""

from __future__ import annotations

import sys
from pathlib import Path

if not __package__:                      # run as a file: make the project folder importable
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else list(argv)
    if "--self-test" in args:                    # practical GUI self-test (launcher and packaged .exe)
        from gui.selftest import run_selftest

        return run_selftest(args[args.index("--self-test"):])

    from gui.app import create_application
    from gui.services.settings import load_settings
    from gui.windows.main_window import MainWindow

    from gui.services.staging import cleanup_stale

    cleanup_stale()                      # staging copies left by an earlier session (only our own folders)
    app = create_application(argv)
    window = MainWindow(load_settings())
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
