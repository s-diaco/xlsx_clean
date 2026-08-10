#!/usr/bin/env bash
# Operator setup for New QC Sheet on Linux desktops.
# Installs deps with uv and creates a Desktop shortcut.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if ! command -v uv >/dev/null 2>&1; then
  echo "uv not found on PATH. Install uv first: https://docs.astral.sh/uv/" >&2
  exit 1
fi

echo "Syncing dependencies..."
uv sync

echo "Installing Desktop shortcut..."
.venv/bin/python desktop/install_desktop_shortcut.py

echo
echo "Done. Launch with:"
echo "  .venv/bin/python -m xlsx_clean.gui_app"
echo "or use the Desktop shortcut."
