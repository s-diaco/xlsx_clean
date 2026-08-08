# xlsx-clean

Operator-facing name: **New QC Sheet**.

Create a new QC datasheet from the latest matching workbook: clear configured cells,
blank notes, write the batch serial, and save under a new `[SERIAL]` name.

Everyday use is a **Desktop icon** that opens a **single app window**
(Set → Ink → Serial → Create) — not your full browser with all tabs.
No terminal commands are required for operators.

## Quick start (Windows operators)

1. One-time setup (IT / first install): install [uv](https://docs.astral.sh/uv/getting-started/installation/), clone/copy this project, then double-click  
   **`desktop\Setup.bat`**
2. Every day: double-click the **New QC Sheet** icon on the Desktop  
3. In the app window: choose **Set**, **Ink color**, enter **Serial**, click **Create datasheet**

`Setup.bat` creates `.venv` with uv, installs the package, and adds the Desktop shortcut
(using the modern icon in `desktop\xlsx-clean.ico`).
The shortcut runs `desktop\XlsxClean.vbs` and opens one dedicated window (via pywebview).
Closing that window stops the app.

When you **Create** with the **com** backend, Excel is maximized and brought to the front
after the new workbook is saved.

If something fails, use **`desktop\XlsxClean.bat`** instead — it shows a console with errors.
To force the old “open in default browser” behavior for debugging:

```bat
.venv\Scripts\python.exe -m xlsx_clean.web_app --browser
```

## Standalone Windows exe (optional)

To put a self-contained app on other PCs (build **on Windows**):

1. Install [uv](https://docs.astral.sh/uv/getting-started/installation/) if needed
2. Double-click **`desktop\build_windows_exe.bat`**
3. It creates `dist\New QC Sheet\New QC Sheet.exe` and a Desktop shortcut
4. Copy the whole **`dist\New QC Sheet\`** folder when moving to another machine (keep files together)

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

Uses **uv** only (no pip / no rye). Install uv once:
https://docs.astral.sh/uv/getting-started/installation/

**Windows (recommended):** double-click **`desktop\Setup.bat`**

**Linux (recommended):** from the project root:

```bash
chmod +x desktop/setup_linux.sh
./desktop/setup_linux.sh
```

That installs apt GTK/WebKit packages if needed, recreates `.venv` with the
**OS** `/usr/bin/python3` (`--no-managed-python --system-site-packages` so apt
`python3-gi` / `_gi` loads), runs `uv sync`, and creates a Desktop shortcut.
Do not use a uv-managed CPython for the Linux native window — `_gi` will fail.

**Manual / macOS** (deps only; on Linux this alone is not enough for a native window):

```bash
uv sync
```

That creates `.venv` and installs locked deps from `uv.lock`.

### Linux single app window (native mode)

Prefer **`./desktop/setup_linux.sh`** (above). Native mode uses **pywebview**,
which needs **system** GTK/WebKit packages (`uv sync` cannot install these).

Manual equivalent:

```bash
# Ubuntu / Debian (names may vary by release: webkit2-4.0 vs 4.1)
sudo apt install python3-gi gir1.2-gtk-3.0 gir1.2-webkit2-4.1

rm -rf .venv
# Must be OS python3 — `uv venv --python python3` may pick a managed build
# that cannot load apt _gi.
uv venv --python /usr/bin/python3 --no-managed-python --system-site-packages
uv sync

.venv/bin/python -m xlsx_clean.web_app
```

If you already have a normal isolated `.venv`, the app will still try to load
system `python3-gi` from `/usr/lib/python3/dist-packages` when those apt packages
are installed. Without GTK/WebKit (or on a headless machine with no display), it
falls back to opening the default browser, or you can force:

```bash
.venv/bin/python -m xlsx_clean.web_app --browser
.venv/bin/python -m xlsx_clean.web_app --no-browser
```

**Windows** native mode uses Edge/WebView2 and does **not** need the apt packages above.

Before building the standalone exe:

```bash
uv sync --extra desktop
```

Then create the Desktop shortcut (if you did not use `Setup.bat`):

- **Windows:** double-click `desktop\Install-DesktopShortcut.bat`
- **Linux:** `./desktop/setup_linux.sh` (preferred), or
  `.venv/bin/python desktop/install_desktop_shortcut.py`

Optional: set `XLSX_CLEAN_ROOT` if workbooks are not under `D:\OpenCloud`
(e.g. `set XLSX_CLEAN_ROOT=D:\OpenCloud` or `export XLSX_CLEAN_ROOT=/mnt/opencloud`).

## Developer interfaces

These are optional; operators should use the Desktop icon.

**App window (default: native single window)**

```bash
.venv/bin/python -m xlsx_clean.web_app
# Windows: .venv\Scripts\python.exe -m xlsx_clean.web_app

# Optional: open in the default browser instead
.venv/bin/python -m xlsx_clean.web_app --browser
```

**CLI**

```bash
.venv/bin/python -m xlsx_clean.clean_cells --backend ooxml
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
| `desktop/xlsx-clean.ico` | Modern Windows Desktop / exe icon |
| `desktop/Setup.bat` | Windows one-time setup via uv + Desktop shortcut |
| `desktop/setup_linux.sh` | Linux one-time setup (apt + OS Python venv + shortcut) |
| `desktop/Install-DesktopShortcut.bat` | Create Desktop icon only (Windows) |
| `desktop/XlsxClean.vbs` | Silent double-click launcher |
| `desktop/XlsxClean.bat` | Launcher with visible console |
| `desktop/build_windows_exe.bat` | Build standalone `New QC Sheet.exe` with uv |
| `desktop/install_desktop_shortcut.py` | Create Desktop shortcut (Linux) |
| `packaging/xlsx_clean.spec` | PyInstaller spec |
