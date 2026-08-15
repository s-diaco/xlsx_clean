#!/usr/bin/env bash
# Operator setup for New QC Sheet on Debian/Ubuntu Linux desktops.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

SYS_PY=/usr/bin/python3

if ! command -v uv >/dev/null 2>&1; then
  echo "uv not found on PATH. Install uv first: https://docs.astral.sh/uv/" >&2
  exit 1
fi

if ! command -v apt-get >/dev/null 2>&1; then
  echo "This setup script requires apt-get (Debian/Ubuntu). Install python3-gi,"
  echo "gir1.2-gtk-3.0, and gir1.2-webkit2-4.1 (or WebKit2 4.0) manually." >&2
  exit 1
fi

if [[ ! -x "$SYS_PY" ]]; then
  echo "System Python not found at $SYS_PY" >&2
  exit 1
fi

packages=(python3-gi gir1.2-gtk-3.0)
if apt-cache show gir1.2-webkit2-4.1 >/dev/null 2>&1; then
  packages+=(gir1.2-webkit2-4.1)
elif apt-cache show gir1.2-webkit2-4.0 >/dev/null 2>&1; then
  packages+=(gir1.2-webkit2-4.0)
else
  echo "Neither gir1.2-webkit2-4.1 nor gir1.2-webkit2-4.0 is available." >&2
  exit 1
fi

missing=()
for package in "${packages[@]}"; do
  if ! dpkg-query -W -f='${db:Status-Status}' "$package" 2>/dev/null | grep -qx installed; then
    missing+=("$package")
  fi
done

if ((${#missing[@]})); then
  echo "Installing GTK/WebKitGTK requirements: ${missing[*]}"
  sudo apt-get update
  sudo apt-get install -y "${missing[@]}"
else
  echo "GTK/WebKitGTK requirements already installed."
fi

if ! "$SYS_PY" -c 'import gi' >/dev/null 2>&1; then
  echo "System $SYS_PY cannot import gi. Install python3-gi." >&2
  exit 1
fi

enable_system_site_packages() {
  local cfg=.venv/pyvenv.cfg
  if [[ ! -f "$cfg" ]]; then
    return 1
  fi
  if grep -q '^include-system-site-packages' "$cfg"; then
    sed -i 's/^include-system-site-packages = .*/include-system-site-packages = true/' "$cfg"
  else
    printf '\ninclude-system-site-packages = true\n' >> "$cfg"
  fi
}

if [[ ! -x .venv/bin/python ]] || ! .venv/bin/python -c 'import gi' >/dev/null 2>&1; then
  echo "Creating venv with the system Python and system site packages..."
  rm -rf .venv
  uv venv --python /usr/bin/python3 --no-managed-python --system-site-packages
fi

echo "Syncing dependencies..."
# Pin the same interpreter so uv does not honor .python-version (3.12) and
# recreate a managed venv without apt gi (python3-gi).
uv sync --python /usr/bin/python3 --no-managed-python

if ! .venv/bin/python -c 'import gi' >/dev/null 2>&1; then
  echo "Restoring system site packages on .venv so apt gi is importable..."
  enable_system_site_packages
fi

venv_ver="$(.venv/bin/python -c 'import sys; print("%d.%d" % sys.version_info[:2])')"
sys_ver="$("$SYS_PY" -c 'import sys; print("%d.%d" % sys.version_info[:2])')"
if [[ "$venv_ver" != "$sys_ver" ]]; then
  echo "venv Python $venv_ver does not match system Python $sys_ver at $SYS_PY." >&2
  echo "uv likely honored .python-version instead of the system interpreter." >&2
  exit 1
fi

echo "Verifying GTK and WebKit2 bindings..."
.venv/bin/python -c '
import gi
gi.require_version("Gtk", "3.0")
try:
    gi.require_version("WebKit2", "4.1")
except ValueError:
    gi.require_version("WebKit2", "4.0")
from gi.repository import Gtk, WebKit2
print("GTK/WebKit2 bindings available")
'

echo
echo "Done. Launch with:"
echo "  .venv/bin/python -m xlsx_clean.gui_app"
echo "Create your own shortcut to that command."
