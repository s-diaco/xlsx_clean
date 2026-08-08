# xlsx-clean

Create a new QC datasheet from the latest matching workbook: clear configured cells,
blank notes, write the batch serial, and save under a new `[SERIAL]` name.

Everyday use is a **Desktop icon** that opens a browser form (Set → Ink → Serial → Create).
No terminal commands are required for operators.

## Quick start (Windows operators)

1. One-time setup (IT / first install): install Python 3.12+, clone/copy this project, create a venv, install deps, then double-click  
   **`desktop\Install-DesktopShortcut.bat`**
2. Every day: double-click the **XlsxClean** icon on the Desktop  
3. In the browser form: choose **Set**, **Ink color**, enter **Serial**, click **Create datasheet**

The Desktop shortcut runs `desktop\XlsxClean.vbs` (no console window) and opens
`http://127.0.0.1:8080`.

If something fails, use **`desktop\XlsxClean.bat`** instead — it shows a console with errors.

## Standalone Windows exe (optional)

To put a self-contained app on other PCs (build **on Windows**):

1. Double-click **`desktop\build_windows_exe.bat`**
2. It creates `dist\XlsxClean\XlsxClean.exe` and a Desktop shortcut
3. Copy the whole **`dist\XlsxClean\`** folder when moving to another machine (keep files together)

## What the app does

| Step | Action |
|------|--------|
| 1 | Find the latest matching workbook in the product folder |
| 2 | Clear configured cells / blank notes |
| 3 | Write the batch serial |
| 4 | Save as a new file using the `[SERIAL]` name pattern |

### Excel backends

| Platform | Default | Notes |
|----------|---------|--------|
| Windows | `com` | Uses Microsoft Excel + add-ins from `strings.txt`; leaves Excel open |
| Linux / macOS | `ooxml` | Surgical `.xlsx` edit; keeps formulas, formatting, and queries on disk |

Choose the backend in the web UI dropdown, or with CLI `--backend com\|ooxml`.

**Limits of `ooxml`:** does not refresh Power Query, calculate formulas, or load `.xlam`/`.xla`
add-ins. Open the file in Excel afterward for refresh/calc. Prefer **`com` on Windows** when
you need add-ins and a live Excel session.

## First-time install (admins)

```bash
# from the project root
python -m venv .venv

# Windows
.venv\Scripts\pip install -e .

# Linux / macOS
.venv/bin/pip install -e .
# or: uv pip install -e . -p .venv/bin/python
```

Then create the Desktop shortcut:

- **Windows:** double-click `desktop\Install-DesktopShortcut.bat`
- **Linux:** `python desktop/install_desktop_shortcut.py`

Optional: set `XLSX_CLEAN_ROOT` if workbooks are not under `D:\OpenCloud`
(e.g. `set XLSX_CLEAN_ROOT=D:\OpenCloud` or `export XLSX_CLEAN_ROOT=/mnt/opencloud`).

## Developer interfaces

These are optional; operators should use the Desktop icon.

**Web UI**

```bash
xlsx-clean-web
# or
python -m xlsx_clean.web_app
```

**CLI**

```bash
xlsx-clean
# or
python -m xlsx_clean.clean_cells --backend ooxml
```

Windows CLI helper: `src\xlsx_clean\new_xslx.bat`

## Config files

| File | Purpose |
|------|---------|
| `src/xlsx_clean/file_data.csv` | Product folders, filename patterns, cells to clear/set |
| `src/xlsx_clean/strings.txt` | UI prompts + Windows Excel add-in paths (`com` only) |
| `XLSX_CLEAN_ROOT` (env) | Replaces the `D:\OpenCloud\...` root in CSV paths |

## Project layout (launchers)

| Path | Purpose |
|------|---------|
| `desktop/Install-DesktopShortcut.bat` | Create Desktop icon (Windows) |
| `desktop/XlsxClean.vbs` | Silent double-click launcher |
| `desktop/XlsxClean.bat` | Launcher with visible console |
| `desktop/build_windows_exe.bat` | Build standalone `XlsxClean.exe` |
| `desktop/install_desktop_shortcut.py` | Create Desktop shortcut (Linux) |
| `packaging/xlsx_clean.spec` | PyInstaller spec |
