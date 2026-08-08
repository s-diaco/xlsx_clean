# xlsx-clean

Create a new QC datasheet from the latest matching workbook: clear configured
cells, blank notes, write the batch serial, and save under a new `[SERIAL]` name.

## Platforms

| Platform | Default backend | Behavior |
|----------|-----------------|----------|
| Windows | `com` | Microsoft Excel via `pywin32`; loads add-ins from `strings.txt`; leaves Excel open |
| Linux / macOS | `ooxml` | Surgical OOXML zip edit; preserves formulas, formatting, queries/connections on disk |

Force a backend with `--backend com` or `--backend ooxml`.

The OOXML path does **not** refresh Power Query, evaluate formulas, or load `.xlam`/`.xla`
add-ins. Open the written file in Excel for refresh/calc. Prefer `com` on Windows when
you need add-ins and a live Excel session.

## Config

- `src/xlsx_clean/file_data.csv` — product dirs, filename patterns, cells to clear/set
- `src/xlsx_clean/strings.txt` — prompts + Windows add-in paths (COM only)
- `XLSX_CLEAN_ROOT` — optional root that replaces `D:\OpenCloud` in CSV paths
  (example: `export XLSX_CLEAN_ROOT=/mnt/opencloud`)

## Run

```bash
# installed entry point
xlsx-clean

# or module
python -m xlsx_clean.clean_cells --backend ooxml
```

On Windows, `src/xlsx_clean/new_xslx.bat` runs the same module entry point.
