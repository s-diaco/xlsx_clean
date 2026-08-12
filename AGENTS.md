# xlsx-clean

`xlsx-clean` generates a new Excel QC datasheet from the most recent workbook in a
configured directory, clearing/resetting specific cells. Interfaces:

- CLI: `src/xlsx_clean/clean_cells.py` (`beaupy`)
- GUI: `src/xlsx_clean/gui_app.py` (PyWebView cross-platform desktop window with a local HTML/CSS/JS frontend)

## Backends

- **`com` (Windows default):** Microsoft Excel COM via `pywin32`. Loads add-ins listed
  in `strings.txt`, edits cells, `SaveAs`, leaves Excel visible/maximized and brought
  to the front.
- **`ooxml` (Linux/macOS default):** Surgical edit of the `.xlsx` zip (worksheet XML
  only). Preserves non-worksheet parts (styles, connections/Power Query definitions,
  drawings, etc.). Does not calculate formulas or refresh queries.

Select with `--backend com|ooxml` (CLI) or `config.toml`.

## Cursor Cloud Specific Instructions

### Platform Notes
- `pywin32` has no Linux wheels and is declared as Windows-only
  (`pywin32>=312; sys_platform == 'win32'`). It is intentionally **not** installed here.
- End-to-end COM automation needs Windows + desktop Excel and does not run on this VM.
- On Linux, the **ooxml** backend can clear/set cells and write a new workbook when
  sample files and `XLSX_CLEAN_ROOT` point at a real tree. Without those data files,
  interactive selection still works for config loading, but globbing finds no templates.
- PyWebView supports Windows (Edge WebView2), Linux (WebKitGTK), and macOS (WKWebView).
  Windows needs the separately-installed Microsoft Edge WebView2 Runtime; `pythonnet`
  does not install it. Linux needs `python3-gi`, GTK, and WebKitGTK typelibs; use the
  system Python with `--system-site-packages` so apt-provided `gi` remains importable.
- On headless machines without a display server the GUI cannot open a native window.

### Environment
- Prefer **uv** for all installs (`uv` in `~/.local/bin`, on PATH via `~/.bashrc`).
- Rye is **not** used (`requirements.lock` removed). Use `uv.lock` + `uv sync`.

  ```bash
  # Linux desktop (system WebKitGTK deps + venv + Desktop shortcut):
  ./desktop/setup_linux.sh

  # Dependencies and test tooling (any OS):
  uv sync
  ```

- Run Python via the venv interpreter: `.venv/bin/python`.
- Optional: `XLSX_CLEAN_ROOT` remaps `D:\OpenCloud\...` paths from `file_data.csv`.
- Note: the `uv`-created `.venv` does not include `pip`; use `uv pip ...` / `uv sync`
  for package operations.

### Lint / Test / Build
- Tests use `pytest`; run `uv run pytest -q`.
- `pyproject.toml` uses `hatchling` as the build backend; build a wheel with `uv build`.

### Running / Verifying
- CLI: `.venv/bin/python -m xlsx_clean.clean_cells` (defaults to `ooxml` on Linux).
- GUI: `.venv/bin/python -m xlsx_clean.gui_app [--width 800 --height 600]`.
- Linux setup: `./desktop/setup_linux.sh` (GTK/WebKitGTK check, uv sync).
- Windows shop-floor: double-click `desktop/Setup.bat` (uv sync), or build
  an exe with `desktop/build_windows_exe.bat` **on a Windows PC** (cannot build Windows
  exe here).
- Operators create their own shortcut to the GUI launcher (no installer script).
- Sanity-check imports:
  `.venv/bin/python -c "import webview, beaupy; from xlsx_clean import hello; from xlsx_clean.core import load_config, list_sets; from xlsx_clean.ooxml_backend import expand_a1_range; print(webview, hello(), expand_a1_range('A1:B2'))"`.
