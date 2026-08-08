# xlsx-clean

`xlsx-clean` is a Python CLI (`src/xlsx_clean/clean_cells.py`) that generates a new
Excel QC datasheet from the most recent workbook in a configured directory,
clearing/resetting specific cells.

## Backends

- **`com` (Windows default):** Microsoft Excel COM via `pywin32`. Loads add-ins listed
  in `strings.txt`, edits cells, `SaveAs`, leaves Excel visible/maximized.
- **`ooxml` (Linux/macOS default):** Surgical edit of the `.xlsx` zip (worksheet XML
  only). Preserves non-worksheet parts (styles, connections/Power Query definitions,
  drawings, etc.). Does not calculate formulas or refresh queries.

Select with `--backend com|ooxml`.

## Cursor Cloud specific instructions

### Platform notes
- `pywin32` has no Linux wheels and is declared as Windows-only
  (`pywin32>=306; sys_platform == 'win32'`). It is intentionally **not** installed here.
- End-to-end COM automation needs Windows + desktop Excel and does not run on this VM.
- On Linux, the **ooxml** backend can clear/set cells and write a new workbook when
  sample files and `XLSX_CLEAN_ROOT` point at a real tree. Without those data files,
  interactive selection still works for config loading, but globbing finds no templates.

### Environment
- Dependencies are installed into a `uv`-managed virtualenv at `.venv` by the startup
  update script (`uv` itself is installed via `pip install --break-system-packages uv`
  and lives in `~/.local/bin`, which is on PATH for interactive shells via `~/.bashrc`).
- The install uses pinned versions from `requirements.lock`, excluding the `-e file:.`
  editable line and `pywin32`; the package is then installed editable with `--no-deps`.
- Run Python via the venv interpreter: `.venv/bin/python`.
- Optional: `XLSX_CLEAN_ROOT` remaps `D:\OpenCloud\...` paths from `file_data.csv`.

### Lint / test / build
- There is **no lint config, no test suite, and no build step** in this repo (no ruff/flake8,
  no pytest, no CI). Don't invent one unless asked.
- `pyproject.toml` uses `hatchling` as the build backend; there is nothing to "build" for dev.

### Running / verifying
- Interactive tool: `.venv/bin/python -m xlsx_clean.clean_cells` (defaults to `ooxml` on Linux).
- Sanity-check imports:
  `.venv/bin/python -c "import pandas, beaupy; from xlsx_clean import hello; from xlsx_clean.ooxml_backend import expand_a1_range; print(hello(), expand_a1_range('A1:B2'))"`.
- Note: the `uv`-created `.venv` does not include `pip`; use `uv pip ...` for package operations.
