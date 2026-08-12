#!/usr/bin/env bash
# Operator launcher for New QC Sheet on macOS.
# Double-click in Finder (or run from Terminal). Opens the GUI in a new Terminal window.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if [[ ! -x .venv/bin/python ]]; then
  echo "Virtual environment not found. Run 'uv sync' first, or use desktop/setup_macos.sh." >&2
  exit 1
fi

exec ./.venv/bin/python -m xlsx_clean.gui_app