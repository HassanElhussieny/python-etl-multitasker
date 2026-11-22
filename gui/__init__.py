"""Python ETL Multitasker - the desktop interface (PySide6) on top of the existing ``sales_etl`` engine.

The GUI never calculates anything itself: every task runs as a separate process of the existing command line
(``python -m sales_etl ... --non-interactive``), exactly as from VS Code or the terminal. The GUI only builds the
command, queues it, shows its live output and reads the folders the engine writes (output, journals, manifests).
"""

APP_NAME = "Python ETL Multitasker"
APP_NAME_AR = "منصة مهام المبيعات"
__version__ = "1.0.0"
