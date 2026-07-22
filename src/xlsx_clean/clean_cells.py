"""Interactive CLI: copy latest QC workbook, clear cells, stamp serial."""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas
from beaupy import prompt, select
from rich.console import Console

from xlsx_clean.paths import (
    default_backend,
    get_opencloud_root,
    package_file,
    remap_opencloud_path,
    resolve_backend,
)

NEW_VALUE = ""


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


def load_config() -> tuple[list[str], pandas.DataFrame]:
    strings_path = package_file("strings.txt")
    with open(strings_path, encoding="utf-8") as f:
        content = [line.strip() for line in f.readlines()]

    path_df = pandas.read_csv(package_file("file_data.csv"))
    root = get_opencloud_root()
    path_df["dir"] = [
        str(remap_opencloud_path(path, root)) for path in path_df["dir"]
    ]
    return content, path_df


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Create a new QC datasheet from the latest matching workbook, "
            "clearing configured cells and writing the batch serial."
        )
    )
    parser.add_argument(
        "--backend",
        choices=("com", "ooxml"),
        default=None,
        help=(
            "Excel edit backend: 'com' (Windows Excel + add-ins) or "
            f"'ooxml' (surgical zip edit). Default: {default_backend()}."
        ),
    )
    return parser.parse_args(argv)


def run(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    backend = resolve_backend(args.backend)
    content, path_df = load_config()
    q1, q2, q3 = content[0], content[1], content[2]
    addin_paths = content[3:5] if len(content) >= 5 else []

    parent_path = list(
        path_df.drop_duplicates(subset=["set_name"], keep="first")["set_name"]
    )

    console = Console()
    console.print(q1)
    selected_set = select(parent_path, cursor_style="cyan")

    console.print(q2)
    selected_dir = select(
        [
            Path(path).stem
            for path in list(path_df[path_df["set_name"] == selected_set]["dir"])
        ]
    )
    console.print(f"Selected: {selected_dir}")

    row = path_df[
        (path_df["set_name"] == selected_set)
        & (path_df["dir"].str.endswith(selected_dir))
    ].iloc[0]
    path_ = Path(row["dir"])
    pattern = row["pattern"]
    search_pattern = pattern.replace("[SERIAL]", "*")
    files = [str(x) for x in path_.glob(search_pattern) if x.is_file()]
    if not files:
        console.print(f"[red]No workbooks matching {search_pattern!r} in {path_}[/red]")
        return 1

    batch_serial = prompt(q3, target_type=str)
    ref_workbook_name, new_workbook_name = get_workbook_names(
        files, batch_serial, path_, pattern
    )
    console.print(f"Selected: {Path(ref_workbook_name).stem}")
    console.print(f"Backend: {backend}")

    if new_workbook_name.is_file():
        console.print(f"[yellow]Destination already exists, skipping: {new_workbook_name}[/yellow]")
        return 0

    cells_to_clear = row["cells_to_clear"]
    notes_cell = row["notes_cell"]
    serial_cell = row["serial_cell"]

    if backend == "com":
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
                addin_paths=addin_paths,
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

    console.print(f"[green]Wrote {new_workbook_name}[/green]")
    return 0


def main() -> None:
    raise SystemExit(run())


if __name__ == "__main__":
    main()
