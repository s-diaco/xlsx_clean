"""Interactive CLI: copy latest QC workbook, clear cells, stamp serial."""

from __future__ import annotations

import argparse
import shutil
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas
from beaupy import prompt, select
from rich.console import Console

from xlsx_clean.app_logging import get_logger
from xlsx_clean.paths import (
    default_backend,
    get_opencloud_root,
    package_file,
    remap_opencloud_path,
    resolve_backend,
)

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


def list_sets(path_df: pandas.DataFrame) -> list[str]:
    return list(
        path_df.drop_duplicates(subset=["set_name"], keep="first")["set_name"]
    )


def list_ink_colors(path_df: pandas.DataFrame, selected_set: str) -> list[str]:
    return [
        Path(path).stem
        for path in list(path_df[path_df["set_name"] == selected_set]["dir"])
    ]


def create_datasheet(
    selected_set: str,
    selected_dir: str,
    batch_serial: str,
    backend: str | None = None,
    path_df: pandas.DataFrame | None = None,
    addin_paths: list[str] | None = None,
) -> CreateResult:
    """Create a cleaned QC workbook for the given selection.

    Shared by the CLI and the NiceGUI web UI.
    """
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
    try:
        resolved_backend = resolve_backend(backend)
    except (ValueError, RuntimeError) as exc:
        log.error("Create failed: %s (%s)", exc, ctx)
        return CreateResult(ok=False, message=str(exc))

    ctx = f"{ctx} backend={resolved_backend}"

    if path_df is None or addin_paths is None:
        content, loaded_df = load_config()
        if path_df is None:
            path_df = loaded_df
        if addin_paths is None:
            addin_paths = content[3:5] if len(content) >= 5 else []

    matches = path_df[
        (path_df["set_name"] == selected_set)
        & (path_df["dir"].str.endswith(selected_dir))
    ]
    if matches.empty:
        message = f"No config row for set={selected_set!r} ink={selected_dir!r}."
        log.error("Create failed: %s (%s)", message, ctx)
        return CreateResult(ok=False, message=message)

    row = matches.iloc[0]
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
    # Validate early so CLI fails fast on unsupported COM.
    try:
        backend = resolve_backend(args.backend)
    except (ValueError, RuntimeError) as exc:
        Console().print(f"[red]{exc}[/red]")
        return 1

    content, path_df = load_config()
    q1, q2, q3 = content[0], content[1], content[2]
    addin_paths = content[3:5] if len(content) >= 5 else []

    console = Console()
    console.print(q1)
    selected_set = select(list_sets(path_df), cursor_style="cyan")

    console.print(q2)
    selected_dir = select(list_ink_colors(path_df, selected_set))
    console.print(f"Selected: {selected_dir}")

    batch_serial = prompt(q3, target_type=str)
    result = create_datasheet(
        selected_set=selected_set,
        selected_dir=selected_dir,
        batch_serial=batch_serial,
        backend=backend,
        path_df=path_df,
        addin_paths=addin_paths,
    )

    if result.template is not None:
        console.print(f"Selected: {result.template.stem}")
    if result.backend is not None:
        console.print(f"Backend: {result.backend}")

    if result.skipped:
        console.print(f"[yellow]{result.message}[/yellow]")
        return 0
    if not result.ok:
        console.print(f"[red]{result.message}[/red]")
        return 1

    console.print(f"[green]{result.message}[/green]")
    return 0


def main() -> None:
    raise SystemExit(run())


if __name__ == "__main__":
    main()
