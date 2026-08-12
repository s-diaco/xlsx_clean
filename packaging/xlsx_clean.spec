# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for a double-clickable New QC Sheet executable.

Build on Windows (recommended for shop-floor PCs):
  desktop\\build_windows_exe.bat

Or manually:
  pyinstaller packaging/xlsx_clean.spec
"""

import os
import sys
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# SPECPATH is set by PyInstaller to the folder containing this spec file
ROOT_DIR = os.path.abspath(os.path.join(SPECPATH, ".."))

datas = [
    (os.path.join(ROOT_DIR, "src", "xlsx_clean", "file_data.csv"), "xlsx_clean"),
    (os.path.join(ROOT_DIR, "src", "xlsx_clean", "strings.txt"), "xlsx_clean"),
    (os.path.join(ROOT_DIR, "src", "xlsx_clean", "fonts"), "xlsx_clean/fonts"),
    (os.path.join(ROOT_DIR, "src", "xlsx_clean", "ui"), "xlsx_clean/ui"),
]
binaries = []
hiddenimports = [
    "xlsx_clean",
    "xlsx_clean.core",
    "xlsx_clean.gui_app",
    "xlsx_clean.clean_cells",
    "xlsx_clean.ooxml_backend",
    "xlsx_clean.com_backend",
    "xlsx_clean.paths",
]

webview_ret = collect_all("webview")
datas += webview_ret[0]
binaries += webview_ret[1]
hiddenimports += webview_ret[2]

if sys.platform == "win32":
    clr_ret = collect_all("clr_loader")
    datas += clr_ret[0]
    binaries += clr_ret[1]
    hiddenimports += clr_ret[2]
    hiddenimports += ["clr", "clr_loader", "pythonnet"]

a = Analysis(
    [os.path.join(ROOT_DIR, "src", "xlsx_clean", "gui_app.py")],
    pathex=[ROOT_DIR],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="New QC Sheet",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(ROOT_DIR, "desktop", "xlsx-clean.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="New QC Sheet",
)
