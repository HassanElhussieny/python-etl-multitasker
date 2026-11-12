"""Structured logging: JSON-lines file per run plus a concise console stream."""

from __future__ import annotations

import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path

LOGGER_NAME = "sales_etl"
_STANDARD_ATTRS = set(vars(logging.makeLogRecord({}))) | {"message", "asctime"}


class JsonLinesFormatter(logging.Formatter):
    """One JSON object per line; ``extra={...}`` fields become top-level keys."""

    def format(self, record: logging.LogRecord) -> str:
        payload = {
            "ts": datetime.fromtimestamp(record.created, tz=timezone.utc).isoformat(),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
        }
        for key, value in vars(record).items():
            if key not in _STANDARD_ATTRS and not key.startswith("_"):
                payload[key] = value
        if record.exc_info:
            payload["exception"] = self.formatException(record.exc_info)
        return json.dumps(payload, ensure_ascii=False, default=str)


def configure_console(level: int = logging.WARNING) -> None:
    """Console handler for CLI use (library code never adds handlers by itself)."""
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)
    if not any(getattr(h, "_sales_etl_console", False) for h in logger.handlers):
        handler = logging.StreamHandler(sys.stderr)
        handler.setLevel(level)
        handler.setFormatter(logging.Formatter("%(levelname)s %(message)s"))
        handler._sales_etl_console = True  # type: ignore[attr-defined]
        logger.addHandler(handler)


def attach_run_log(path: Path) -> logging.Handler:
    """Add a JSON-lines handler for one run; caller must remove it afterwards."""
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)
    handler = logging.FileHandler(path, mode="x", encoding="utf-8")
    handler.setLevel(logging.DEBUG)
    handler.setFormatter(JsonLinesFormatter())
    logger.addHandler(handler)
    return handler


def detach(handler: logging.Handler) -> None:
    logging.getLogger(LOGGER_NAME).removeHandler(handler)
    handler.close()
