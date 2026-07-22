#!/usr/bin/env bash
# Operator setup for New QC Sheet on Ubuntu/Debian desktops.
# Installs system GTK/WebKit bindings, creates a uv venv that can see them,
# syncs deps, and installs a Desktop shortcut.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

APT_PACKAGES=(python3-gi gir1.2-gtk-3.0 gir1.2-webkit2-4.1)

need_apt=()
for pkg in "${APT_PACKAGES[@]}"; do
  if ! dpkg -s "$pkg" >/dev/null 2>&1; then
    need_apt+=("$pkg")
  fi
done

if ((${#need_apt[@]})); then
  echo "Missing system packages for the native window: ${need_apt[*]}"
  echo "Installing with apt (sudo)..."
  sudo apt-get update
  sudo apt-get install -y "${need_apt[@]}"
else
  echo "System GTK/WebKit packages already installed."
fi

if ! command -v uv >/dev/null 2>&1; then
  echo "uv not found on PATH. Install uv first: https://docs.astral.sh/uv/" >&2
  exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 not found on PATH." >&2
  exit 1
fi

echo "Creating .venv with system-site-packages (so apt python3-gi is visible)..."
rm -rf .venv
uv venv --python python3 --system-site-packages
uv sync

echo "Checking native-window bindings..."
if ! .venv/bin/python - <<'PY'
import os
import sys

# Match web_app detection without starting NiceGUI.
sys.path.insert(0, "src")
from xlsx_clean.web_app import _linux_webview_libs_available, _linux_has_display

ok_libs = _linux_webview_libs_available()
ok_display = _linux_has_display()
print(f"Python: {sys.version.split()[0]}")
print(f"GTK/WebKit usable: {ok_libs}")
print(f"Display available: {ok_display}")
if not ok_libs:
    sys.exit(2)
PY
then
  echo "Warning: GTK/WebKit still not importable from the venv." >&2
  echo "Try: sudo apt install ${APT_PACKAGES[*]}" >&2
else
  echo "Native window libraries look OK."
fi

echo "Installing Desktop shortcut..."
.venv/bin/python desktop/install_desktop_shortcut.py

echo
echo "Done. Launch with:"
echo "  .venv/bin/python -m xlsx_clean.web_app"
echo "or use the Desktop shortcut."
