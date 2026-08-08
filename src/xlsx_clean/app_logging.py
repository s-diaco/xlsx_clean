"""Append-only file logging for QC datasheet creation."""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

_LOGGER_NAME = "xlsx_clean.create"
_configured = False


def default_log_path() -> Path:
    """Resolve the log file path.

    Order:
    1. ``XLSX_CLEAN_LOG`` env var (file path)
    2. Next to the frozen executable (PyInstaller)
    3. ``~/.xlsx-clean/xlsx-clean.log``
    """
    env = os.environ.get("XLSX_CLEAN_LOG")
    if env:
        return Path(env)

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "xlsx-clean.log"

    return Path.home() / ".xlsx-clean" / "xlsx-clean.log"


def get_logger() -> logging.Logger:
    """Return a logger that appends timestamped lines to the log file."""
    global _configured
    logger = logging.getLogger(_LOGGER_NAME)
    if not _configured:
        log_path = default_log_path()
        log_path.parent.mkdir(parents=True, exist_ok=True)
        handler = logging.FileHandler(log_path, encoding="utf-8")
        handler.setFormatter(
            logging.Formatter(
                fmt="%(asctime)s %(levelname)s %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
        )
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        logger.propagate = False
        _configured = True
    return logger
