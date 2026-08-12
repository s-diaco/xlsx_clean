# xlsx-clean

Operator-facing name: **New QC Sheet**.

Create a new QC datasheet from the latest matching workbook: clear configured cells,
blank notes, write the batch serial, and save under a new `[SERIAL]` name.

Everyday use opens a **single app window** (Set -> Ink -> Serial -> Create).
No terminal commands are required for operators.

## Quick start (Windows operators)

1. One-time setup (IT / first install): install [uv](https://docs.astral.sh/uv/getting-started/installation/), clone/copy this project, then double-click
   **`desktop\Setup.bat`**
2. Every day: launch the app via the **New QC Sheet** shortcut
3. In the app window: choose **Set**, **Ink color**, enter **Serial**, click **Create datasheet**

`Setup.bat` creates `.venv` with uv and installs the package. Operators create their own
shortcut to `desktop\XlsxClean.vbs` (uses the modern icon in `desktop\xlsx-clean.ico`).
On Windows, install the Microsoft Edge WebView2 Runtime separately before opening the
PyWebView GUI; `pythonnet` does not install it. The shortcut runs `desktop\XlsxClean.vbs`
which launches the GUI window. Closing that window stops the app.

When you **Create** with the **com** backend, Excel is maximized and brought to the front
after the new workbook is saved.

If something fails, use **`desktop\XlsxClean.bat`** instead - it shows a console with errors.

## Standalone Windows exe (optional)

To put a self-contained app on other PCs (build **on Windows**):

1. Install [uv](https://docs.astral.sh/uv/getting-started/installation/) if needed
2. Double-click **`desktop\build_windows_exe.bat`**
3. It creates `dist\New QC Sheet\New QC Sheet.exe`
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

Choose the backend in the `config.toml` file, or override with CLI `--backend com|ooxml`.

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

That installs the required GTK/WebKitGTK packages, creates a venv that can import the
apt-provided `gi` bindings, and runs `uv sync`. It uses WebKit2 4.1 when available and
supports distro packages that provide WebKit2 4.0. Operators create their own shortcut
to `.venv/bin/python -m xlsx_clean.gui_app`.

**Manual / macOS:**

The macOS backend defaults to **`ooxml`** (no Excel/COM on macOS). WKWebView is built into
macOS, so no extra system packages are needed beyond Python 3.12+ and `uv`.

From the project root:

```bash
uv sync
```

That creates `.venv` and installs locked deps from `uv.lock` (including the `pyobjc`
bindings used by PyWebView).

To create a double-clickable Finder launcher that runs the GUI:

```bash
chmod +x desktop/XlsxClean.command
```

Then double-click **`desktop/XlsxClean.command`** in Finder. The first launch may require
right-click → Open (Gatekeeper) once. The script opens the app window in a new Terminal
window; closing that window stops the app. Operators create their own shortcut to this
executable.

Before building the standalone exe:

```bash
uv sync --extra desktop
```

Optional: Edit `~/.xlsx-clean/config.toml` (created automatically on first run) if workbooks are not under the default `D:\\OpenCloud` paths.

## Developer interfaces

These are optional; operators use the shortcut to the app executable.

**GUI (desktop window)**

```bash
.venv/bin/python -m xlsx_clean.gui_app --width 800 --height 600
# Windows: .venv\Scripts\python.exe -m xlsx_clean.gui_app --width 800 --height 600
```

**CLI**

```bash
.venv/bin/python -m xlsx_clean.clean_cells --backend ooxml
```

Windows CLI helper: `src\xlsx_clean\new_xslx.bat`

## Config files

| File | Purpose |
|------|---------|
| `config.toml` | Backend choice and root directory path for each set (auto-generated in `~/.xlsx-clean/` on first run) |
| `src/xlsx_clean/file_data.csv` | Relative product folders, filename patterns, cells to clear/set |
| `src/xlsx_clean/strings.txt` | UI prompts + Windows Excel add-in paths (`com` only) |

## Project layout (launchers)

| Path | Purpose |
|------|---------|
| `desktop/xlsx-clean.ico` | Modern Windows Desktop / exe icon |
| `desktop/Setup.bat` | Windows one-time setup via uv |
| `desktop/setup_linux.sh` | Linux one-time setup (uv sync) |
| `desktop/XlsxClean.command` | macOS double-click launcher (opens GUI in Terminal) |
| `desktop/XlsxClean.vbs` | Silent Windows double-click launcher |
| `desktop/XlsxClean.bat` | Windows launcher with visible console |
| `desktop/build_windows_exe.bat` | Build standalone `New QC Sheet.exe` with uv |
| `packaging/xlsx_clean.spec` | PyInstaller spec |
