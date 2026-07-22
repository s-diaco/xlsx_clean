"""Install a Desktop shortcut that launches the xlsx-clean web UI."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
DESKTOP_DIR = Path(__file__).resolve().parent


def _user_desktop() -> Path:
    if sys.platform == "win32":
        return Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop"
    # Linux / macOS
    xdg = os.environ.get("XDG_DESKTOP_DIR")
    if xdg:
        return Path(xdg)
    return Path.home() / "Desktop"


def install_windows() -> Path:
    ps1 = DESKTOP_DIR / "Install-DesktopShortcut.ps1"
    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(ps1),
        ],
        check=True,
    )
    return _user_desktop() / "XlsxClean.lnk"


def install_linux() -> Path:
    desktop = _user_desktop()
    desktop.mkdir(parents=True, exist_ok=True)
    python = REPO_ROOT / ".venv" / "bin" / "python"
    if not python.is_file():
        python = Path(sys.executable)
    out = desktop / "XlsxClean.desktop"
    exec_line = (
        f'{python} -m xlsx_clean.web_app --host 127.0.0.1 --port 8080'
    )
    content = f"""[Desktop Entry]
Type=Application
Version=1.0
Name=XlsxClean
Comment=Create a new QC datasheet (xlsx-clean web UI)
Exec={exec_line}
Path={REPO_ROOT}
Icon=applications-office
Terminal=false
Categories=Office;
StartupNotify=true
"""
    out.write_text(content, encoding="utf-8")
    out.chmod(out.stat().st_mode | 0o111)
    return out


def main() -> None:
    if sys.platform == "win32":
        path = install_windows()
    else:
        path = install_linux()
    print(f"Desktop shortcut created: {path}")


if __name__ == "__main__":
    main()
