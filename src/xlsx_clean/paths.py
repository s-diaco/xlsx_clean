"""Config and path helpers for cross-platform workbook roots."""

from __future__ import annotations

import os
import sys
from pathlib import Path


def _package_dir() -> Path:
    """Directory containing package data (CSV/txt), including PyInstaller bundles."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bundled = Path(sys._MEIPASS) / "xlsx_clean"
        if bundled.is_dir():
            return bundled
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


PACKAGE_DIR = _package_dir()

def package_file(name: str) -> Path:
    return PACKAGE_DIR / name

def default_config_path() -> Path:
    """Resolve the default location for the TOML configuration file."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "config.toml"
    return Path.home() / ".xlsx-clean" / "config.toml"


def default_backend() -> str:
    return "com" if sys.platform == "win32" else "ooxml"


def resolve_backend(requested: str | None, config_backend: str | None = None) -> str:
    """Resolve backend using explicitly requested, config file, or system default."""
    if requested:
        backend = requested.lower()
    elif config_backend:
        backend = config_backend.lower()
    else:
        backend = default_backend().lower()
        
    if backend not in {"com", "ooxml"}:
        raise ValueError(f"Unknown backend {backend!r}; use 'com' or 'ooxml'")
    if backend == "com" and sys.platform != "win32":
        raise RuntimeError(
            "COM backend requires Windows with Microsoft Excel and pywin32"
        )
    return backend
