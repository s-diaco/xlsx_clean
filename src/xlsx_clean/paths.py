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

# Historical Windows root used in file_data.csv
DEFAULT_OPENCLOUD_ROOT = Path(r"D:\OpenCloud")
OPENCLOUD_MARKER = "OpenCloud"


def package_file(name: str) -> Path:
    return PACKAGE_DIR / name


def get_opencloud_root() -> Path:
    """Workbook tree root from XLSX_CLEAN_ROOT, else D:\\OpenCloud on Windows."""
    env = os.environ.get("XLSX_CLEAN_ROOT")
    if env:
        return Path(env)
    return DEFAULT_OPENCLOUD_ROOT


def remap_opencloud_path(path_str: str, root: Path | None = None) -> Path:
    """Map a CSV path under OpenCloud onto the configured root.

    Examples:
      D:\\OpenCloud\\inco\\... + root=/data  ->  /data/inco/...
      Relative paths without OpenCloud are returned as-is (Path).
    """
    if root is None:
        root = get_opencloud_root()

    normalized = path_str.replace("\\", "/")
    # Case-insensitive marker search
    lower = normalized.lower()
    marker = f"{OPENCLOUD_MARKER.lower()}/"
    idx = lower.find(marker)
    if idx != -1:
        rel = normalized[idx + len(marker) :]
        return root / rel

    # Already relative to root-style path without drive
    return Path(path_str)


def default_backend() -> str:
    return "com" if sys.platform == "win32" else "ooxml"


def resolve_backend(requested: str | None) -> str:
    backend = (requested or default_backend()).lower()
    if backend not in {"com", "ooxml"}:
        raise ValueError(f"Unknown backend {backend!r}; use 'com' or 'ooxml'")
    if backend == "com" and sys.platform != "win32":
        raise RuntimeError(
            "COM backend requires Windows with Microsoft Excel and pywin32"
        )
    return backend
