"""Core business logic for xlsx_clean."""

from __future__ import annotations

import csv
import shutil
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import tomllib

from xlsx_clean.app_logging import get_logger
from xlsx_clean.paths import (
    default_config_path,
    package_file,
    resolve_backend,
)

DEFAULT_CONFIG_CONTENT = """\
# Configuration for xlsx_clean

# backend options: 
#   - "com": Use Microsoft Excel COM (Windows only, keeps Excel visible)
#   - "ooxml": Surgical zip edit (Cross-platform, very fast, preserves non-worksheet data)
backend = "com"

[sets]
"Kimia Razi" = "D:\\\\OpenCloud\\\\Kimia Razi"
"inco" = "D:\\\\OpenCloud\\\\inco"
"""


NEW_VALUE = ""


@dataclass(frozen=True)
class CreateResult:
    ok: bool
    message: str
    template: Path | None = None
    destination: Path | None = None
    backend: str | None = None
    skipped: bool = False


def find_last_workbook(files: list[str]) -> str:
    return sorted(files)[-1]


def get_workbook_names(
    files: list[str],
    batch_serial: str,
    source_dir: Path,
    pattern: str,
) -> tuple[str, Path]:
    last_workbook = find_last_workbook(files)
    path_of_the_year = Path(str(source_dir).replace("2025", str(datetime.now().year)))
    return last_workbook, path_of_the_year / pattern.replace(
        "[SERIAL]", batch_serial.split("/")[0]
    )


def load_config() -> tuple[list[str], list[dict], dict]:
    config_path = default_config_path()
    if not config_path.exists():
        config_path.parent.mkdir(parents=True, exist_ok=True)
        config_path.write_text(DEFAULT_CONFIG_CONTENT, encoding="utf-8")
        
    with open(config_path, "rb") as f:
        config = tomllib.load(f)

    strings_path = package_file("strings.txt")
    with open(strings_path, encoding="utf-8") as f:
        content = [line.strip() for line in f.readlines()]

    csv_path = package_file("file_data.csv")
    rows = []
    set_roots = config.get("sets", {})
    with open(csv_path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            set_name = row["set_name"]
            rel_dir = row["dir"].replace("\\", "/")
            row["rel_dir"] = rel_dir
            root_str = set_roots.get(set_name, "")
            if root_str:
                row["dir"] = str(Path(root_str) / rel_dir)
            else:
                row["dir"] = str(Path(rel_dir))
            rows.append(row)
            
    return content, rows, config


def list_sets(rows: list[dict]) -> list[str]:
    return list(dict.fromkeys(r["set_name"] for r in rows))


import re

def _sort_key(name: str):
    match = re.search(r'\d+', name)
    num = int(match.group()) if match else 0
    return (num, name)

def list_ink_colors(rows: list[dict], selected_set: str) -> list[str]:
    inks = (Path(r["dir"]).stem for r in rows if r["set_name"] == selected_set)
    return sorted(inks, key=_sort_key)


def create_datasheet(
    selected_set: str,
    selected_dir: str,
    batch_serial: str,
    backend: str | None = None,
    rows: list[dict] | None = None,
    addin_paths: list[str] | None = None,
    config: dict | None = None,
) -> CreateResult:
    log = get_logger()
    ctx = f"set={selected_set!r} ink={selected_dir!r} serial={batch_serial!r}"
    
    if not selected_set:
        message = "Select a set."
        log.error("Create failed: %s (%s)", message, ctx)
        return CreateResult(ok=False, message=message)
    if not selected_dir:
        message = "Select an ink color."
        log.error("Create failed: %s (%s)", message, ctx)
        return CreateResult(ok=False, message=message)
    if not batch_serial or not str(batch_serial).strip():
        message = "Enter a serial."
        log.error("Create failed: %s (%s)", message, ctx)
        return CreateResult(ok=False, message=message)

    batch_serial = str(batch_serial).strip()
    ctx = f"set={selected_set!r} ink={selected_dir!r} serial={batch_serial!r}"
    if config is None:
        content, loaded_rows, config = load_config()
        if rows is None:
            rows = loaded_rows
        if addin_paths is None:
            addin_paths = content[3:5] if len(content) >= 5 else []
            
    try:
        resolved_backend = resolve_backend(backend, config.get("backend"))
    except (ValueError, RuntimeError) as exc:
        log.error("Create failed: %s (%s)", exc, ctx)
        return CreateResult(ok=False, message=str(exc))

    ctx = f"{ctx} backend={resolved_backend}"

    matches = [r for r in rows if r["set_name"] == selected_set and r["dir"].endswith(selected_dir)]
    if not matches:
        message = f"No config row for set={selected_set!r} ink={selected_dir!r}."
        log.error("Create failed: %s (%s)", message, ctx)
        return CreateResult(ok=False, message=message)

    row = matches[0]
    path_ = Path(row["dir"])
    pattern = row["pattern"]
    search_pattern = pattern.replace("[SERIAL]", "*")
    files = [str(x) for x in path_.glob(search_pattern) if x.is_file()]
    
    if not files:
        message = f"No workbooks matching {search_pattern!r} in {path_}"
        log.error("Create failed: %s (%s)", message, ctx)
        return CreateResult(
            ok=False,
            message=message,
            backend=resolved_backend,
        )

    ref_workbook_name, new_workbook_name = get_workbook_names(
        files, batch_serial, path_, pattern
    )
    template = Path(ref_workbook_name)

    if new_workbook_name.is_file():
        message = f"Destination already exists, skipped: {new_workbook_name}"
        log.info(
            "Create skipped: %s (template=%s destination=%s %s)",
            message,
            template,
            new_workbook_name,
            ctx,
        )
        return CreateResult(
            ok=True,
            skipped=True,
            message=message,
            template=template,
            destination=new_workbook_name,
            backend=resolved_backend,
        )

    cells_to_clear = row["cells_to_clear"]
    notes_cell = row["notes_cell"]
    serial_cell = row["serial_cell"]

    try:
        if resolved_backend == "com":
            from xlsx_clean.com_backend import clean_workbook_com

            with tempfile.TemporaryDirectory(prefix="xlsx_clean_") as tmp:
                temp_workbook = Path(tmp) / "temp_workbook.xlsx"
                shutil.copyfile(ref_workbook_name, temp_workbook)
                clean_workbook_com(
                    source=temp_workbook,
                    destination=new_workbook_name,
                    cells_to_clear=cells_to_clear,
                    notes_cell=notes_cell,
                    serial_cell=serial_cell,
                    batch_serial=batch_serial,
                    addin_paths=addin_paths or [],
                    notes_value=NEW_VALUE,
                )
        else:
            from xlsx_clean.ooxml_backend import clean_workbook_ooxml

            clean_workbook_ooxml(
                source=ref_workbook_name,
                destination=new_workbook_name,
                cells_to_clear=cells_to_clear,
                notes_cell=notes_cell,
                serial_cell=serial_cell,
                batch_serial=batch_serial,
                notes_value=NEW_VALUE,
            )
    except Exception as exc:  # noqa: BLE001 - surface to CLI/UI
        message = f"Failed to create workbook: {exc}"
        log.exception(
            "Create failed: %s (template=%s destination=%s %s)",
            message,
            template,
            new_workbook_name,
            ctx,
        )
        return CreateResult(
            ok=False,
            message=message,
            template=template,
            destination=new_workbook_name,
            backend=resolved_backend,
        )

    message = f"Wrote {new_workbook_name}"
    log.info(
        "Create succeeded: %s (template=%s destination=%s %s)",
        message,
        template,
        new_workbook_name,
        ctx,
    )
    return CreateResult(
        ok=True,
        message=message,
        template=template,
        destination=new_workbook_name,
        backend=resolved_backend,
    )
