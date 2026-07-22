# xlsx-clean

`xlsx-clean` is a small **Windows-only** Python CLI (`src/xlsx_clean/clean_cells.py`) that
generates a new Excel QC datasheet from the most recent workbook in a configured directory,
clearing/resetting specific cells via Microsoft Excel COM automation (`pywin32`).

## Cursor Cloud specific instructions

### Platform limitation (important)
- The app is **Windows-only** and **cannot run end-to-end on this Linux VM**:
  - `pywin32` has no Linux wheels, so it is intentionally **not** installed here.
  - `src/xlsx_clean/clean_cells.py` imports `win32com.client` at module top level and drives
    the desktop **Microsoft Excel** app via COM, plus reads `D:\OpenCloud\...` Windows paths
    from `src/xlsx_clean/file_data.csv`. None of that exists on Linux.
- What **does** run on Linux: everything except the Excel COM layer, i.e. the config/data
  logic (reading `strings.txt` + `file_data.csv` with `pandas`, resolving the set/ink-color
  selection, and computing the new workbook filename from the `[SERIAL]` pattern).

### Environment
- Dependencies are installed into a `uv`-managed virtualenv at `.venv` by the startup update
  script (`uv` itself is installed via `pip install --break-system-packages uv` and lives in
  `~/.local/bin`, which is on PATH for interactive shells via `~/.bashrc`).
- The install faithfully uses pinned versions from `requirements.lock`, excluding the `-e file:.`
  editable line and `pywin32`; the package is then installed editable with `--no-deps`.
- Run Python via the venv interpreter: `.venv/bin/python`.

### Lint / test / build
- There is **no lint config, no test suite, and no build step** in this repo (no ruff/flake8,
  no pytest, no CI). Don't invent one unless asked.
- `pyproject.toml` uses `hatchling` as the build backend; there is nothing to "build" for dev.

### Running / verifying
- The interactive tool (`cd src/xlsx_clean && python clean_cells.py`) needs Windows + Excel and
  will fail on import here — do not expect it to run on the VM.
- To sanity-check the environment on Linux, verify imports and the package:
  `.venv/bin/python -c "import pandas, beaupy; from xlsx_clean import hello; print(hello())"`.
- Note: the `uv`-created `.venv` does not include `pip`; use `uv pip ...` for package operations.
