#!/usr/bin/env bash
# Operator setup for New QC Sheet on Ubuntu/Debian desktops.
# Installs system GTK/WebKit bindings, creates a uv venv that uses the OS
# Python (so apt python3-gi/_gi loads), syncs deps, and installs a Desktop shortcut.
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

# Must use the OS interpreter that apt's python3-gi/_gi was built for.
# Never use `command -v python3`: with an activated .venv or uv on PATH it often
# resolves to ~/.local/share/uv/python/... which cannot import apt gi.
resolve_os_python() {
  local candidate resolved
  for candidate in /usr/bin/python3 /usr/local/bin/python3; do
    if [[ -x "$candidate" ]]; then
      resolved="$(readlink -f "$candidate")"
      # Reject uv-managed installs even if somehow symlinked here.
      if [[ "$resolved" == *"/uv/python/"* ]] || [[ "$resolved" == *".local/share/uv/"* ]]; then
        continue
      fi
      printf '%s\n' "$resolved"
      return 0
    fi
  done
  return 1
}

if ! SYSTEM_PYTHON="$(resolve_os_python)"; then
  echo "OS Python not found at /usr/bin/python3." >&2
  echo "Install python3 from apt, then re-run. Do not rely on uv-managed Python." >&2
  exit 1
fi

SYSTEM_VERSION="$("$SYSTEM_PYTHON" -c 'import sys; print(sys.version.split()[0])')"
echo "Using OS Python: ${SYSTEM_PYTHON} (${SYSTEM_VERSION})"
if PATH_PYTHON="$(command -v python3 2>/dev/null || true)"; then
  PATH_PYTHON_RESOLVED="$(readlink -f "$PATH_PYTHON" 2>/dev/null || true)"
  if [[ -n "$PATH_PYTHON_RESOLVED" && "$PATH_PYTHON_RESOLVED" != "$SYSTEM_PYTHON" ]]; then
    echo "Note: PATH python3 is ${PATH_PYTHON_RESOLVED} (ignored; apt gi needs OS Python)."
  fi
fi

echo "Preflight: system Python must import Gtk/WebKit..."
if ! "$SYSTEM_PYTHON" - <<'PY'
import gi

gi.require_version("Gtk", "3.0")
gi.require_version("Gdk", "3.0")
try:
    gi.require_version("WebKit2", "4.1")
except ValueError:
    gi.require_version("WebKit2", "4.0")
from gi.repository import Gtk, WebKit2  # noqa: F401

print("System gi OK")
PY
then
  echo "System Python cannot import Gtk/WebKit. Install:" >&2
  echo "  sudo apt install ${APT_PACKAGES[*]}" >&2
  exit 1
fi

echo "Creating .venv with system-site-packages (OS Python only)..."
rm -rf .venv
# --no-managed-python: refuse uv-managed builds that break apt _gi.
# Also pass it to sync — otherwise uv may recreate .venv from a managed
# download (see .python-version) and undo the system interpreter.
uv venv --python "$SYSTEM_PYTHON" --no-managed-python --system-site-packages
uv sync --python "$SYSTEM_PYTHON" --no-managed-python

echo "Checking native-window bindings in .venv..."
if ! .venv/bin/python - <<'PY'
import sys

sys.path.insert(0, "src")
from xlsx_clean.web_app import (
    _linux_has_display,
    _linux_webview_libs_available,
    _NATIVE_DIAG,
)

ok_libs = _linux_webview_libs_available()
ok_display = _linux_has_display()
print(f"Python: {sys.version.split()[0]} ({sys.executable})")
print(f"GTK/WebKit usable: {ok_libs}")
print(f"Display available: {ok_display}")
if not ok_libs:
    if _NATIVE_DIAG:
        print(f"Detail: {_NATIVE_DIAG}", file=sys.stderr)
    sys.exit(2)
PY
then
  echo "ERROR: GTK/WebKit still not usable from .venv." >&2
  echo "The venv Python must be the same build as apt python3-gi." >&2
  echo "Retry: rm -rf .venv && uv venv --python /usr/bin/python3 --no-managed-python --system-site-packages && uv sync" >&2
  exit 1
fi

echo "Native window libraries look OK."

echo "Installing Desktop shortcut..."
.venv/bin/python desktop/install_desktop_shortcut.py

echo
echo "Done. Launch with:"
echo "  .venv/bin/python -m xlsx_clean.web_app"
echo "or use the Desktop shortcut."
