# xlsx-clean

Create a new QC datasheet from the latest matching workbook: clear configured
cells, blank notes, write the batch serial, and save under a new `[SERIAL]` name.

## Platforms

| Platform | Default backend | Behavior |
|----------|-----------------|----------|
| Windows | `com` | Microsoft Excel via `pywin32`; loads add-ins from `strings.txt`; leaves Excel open |
| Linux / macOS | `ooxml` | Surgical OOXML zip edit; preserves formulas, formatting, queries/connections on disk |

Force a backend with `--backend com` or `--backend ooxml` (CLI) or the Backend
dropdown in the web UI.

The OOXML path does **not** refresh Power Query, evaluate formulas, or load `.xlam`/`.xla`
add-ins. Open the written file in Excel for refresh/calc. Prefer `com` on Windows when
you need add-ins and a live Excel session.

## Desktop icon (no commands)

### Option A — shortcut to the installed project (fast)

1. Install the project once (venv + `pip/uv install`).
2. Double-click **`desktop\Install-DesktopShortcut.bat`** (Windows).
3. Use the **XlsxClean** icon on your Desktop.

That shortcut runs `desktop\XlsxClean.vbs`, which starts the web UI and opens the
browser with no console window.

Linux:

```bash
python desktop/install_desktop_shortcut.py
```

### Option B — standalone Windows `.exe` (best for other PCs)

On a Windows machine with Python:

1. Double-click **`desktop\build_windows_exe.bat`**
2. It builds `dist\XlsxClean\XlsxClean.exe` and places an **XlsxClean** Desktop shortcut
3. Copy the whole `dist\XlsxClean\` folder to other PCs if needed (keep the folder together)

> The Windows `.exe` must be built **on Windows**. This Linux/cloud environment cannot
> produce a native Windows executable.

Troubleshooting launcher (shows a console): `desktop\XlsxClean.bat`

## Interfaces

### Web UI (NiceGUI)

```bash
xlsx-clean-web
# or
python -m xlsx_clean.web_app
```

Opens `http://127.0.0.1:8080` (bind address/port: `--host`, `--port`).

### CLI

```bash
xlsx-clean
# or
python -m xlsx_clean.clean_cells --backend ooxml
```

On Windows, `src/xlsx_clean/new_xslx.bat` runs the CLI module entry point.

## Config

- `src/xlsx_clean/file_data.csv` — product dirs, filename patterns, cells to clear/set
- `src/xlsx_clean/strings.txt` — prompts + Windows add-in paths (COM only)
- `XLSX_CLEAN_ROOT` — optional root that replaces `D:\OpenCloud` in CSV paths
  (example: `export XLSX_CLEAN_ROOT=/mnt/opencloud`)
