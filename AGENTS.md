# xlsx-clean

`xlsx-clean` generates a new Excel QC datasheet from the most recent workbook in a
configured directory, clearing/resetting specific cells. Interfaces:

- CLI: `src/xlsx_clean/clean_cells.py` (`beaupy`)
- Web UI: `src/xlsx_clean/web_app.py` (NiceGUI; default **native** single window via pywebview)

## Backends

- **`com` (Windows default):** Microsoft Excel COM via `pywin32`. Loads add-ins listed
  in `strings.txt`, edits cells, `SaveAs`, leaves Excel visible/maximized and brought
  to the front.
- **`ooxml` (Linux/macOS default):** Surgical edit of the `.xlsx` zip (worksheet XML
  only). Preserves non-worksheet parts (styles, connections/Power Query definitions,
  drawings, etc.). Does not calculate formulas or refresh queries.

Select with `--backend com|ooxml` (CLI) or the Backend dropdown (web UI).

## Cursor Cloud specific instructions

### Platform notes
- `pywin32` has no Linux wheels and is declared as Windows-only
  (`pywin32>=312; sys_platform == 'win32'`). It is intentionally **not** installed here.
- End-to-end COM automation needs Windows + desktop Excel and does not run on this VM.
- On Linux, the **ooxml** backend can clear/set cells and write a new workbook when
  sample files and `XLSX_CLEAN_ROOT` point at a real tree. Without those data files,
  interactive selection still works for config loading, but globbing finds no templates.
- NiceGUI defaults to a native window (`pywebview`). On Linux that needs **system**
  GTK/WebKit packages (`python3-gi`, `gir1.2-gtk-3.0`, `gir1.2-webkit2-4.1` via apt) —
  not installable with `uv`. On this headless VM native mode is unavailable; the app
  falls back to the browser, or use `--no-browser` / `--browser`.

### Environment
- Prefer **uv** for all installs (`uv` in `~/.local/bin`, on PATH via `~/.bashrc`).
- Rye is **not** used (`requirements.lock` removed). Use `uv.lock` + `uv sync`.

  ```bash
  uv sync
  ```

- Run Python via the venv interpreter: `.venv/bin/python`.
- Optional: `XLSX_CLEAN_ROOT` remaps `D:\OpenCloud\...` paths from `file_data.csv`.
- Note: the `uv`-created `.venv` does not include `pip`; use `uv pip ...` / `uv sync`
  for package operations.

### Lint / test / build
- There is **no lint config, no test suite, and no build step** in this repo (no ruff/flake8,
  no pytest, no CI). Don't invent one unless asked.
- `pyproject.toml` uses `hatchling` as the build backend; there is nothing to "build" for dev.

### Running / verifying
- CLI: `.venv/bin/python -m xlsx_clean.clean_cells` (defaults to `ooxml` on Linux).
- Web UI: `.venv/bin/python -m xlsx_clean.web_app` (native window) or
  `--no-browser` / `--browser` on headless / for default browser.
- Desktop shortcut (Linux): `.venv/bin/python desktop/install_desktop_shortcut.py`
- Windows shop-floor: double-click `desktop/Setup.bat` (uv sync + shortcut), or build
  an exe with `desktop/build_windows_exe.bat` **on a Windows PC** (cannot build Windows
  exe here).
- Sanity-check imports:
  `.venv/bin/python -c "import pandas, beaupy, nicegui; from xlsx_clean import hello; from xlsx_clean.ooxml_backend import expand_a1_range; print(hello(), expand_a1_range('A1:B2'))"`.
