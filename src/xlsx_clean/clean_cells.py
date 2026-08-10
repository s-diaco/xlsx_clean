"""Interactive CLI: select set/ink/serial, create a cleaned QC workbook."""

from __future__ import annotations

import argparse

from beaupy import prompt, select
from rich.console import Console

from xlsx_clean.core import (
    create_datasheet,
    list_ink_colors,
    list_sets,
    load_config,
)
from xlsx_clean.paths import default_backend, resolve_backend


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
            f"'ooxml' (surgical zip edit). Default: {default_backend()} (Overrides config.toml)."
        ),
    )
    parser.add_argument(
        "--root",
        type=str,
        default=None,
        help="Override the root directory for the selected set.",
    )
    return parser.parse_args(argv)



def run(argv: list[str] | None = None) -> int:
    from pathlib import Path
    args = parse_args(argv)
    
    content, rows, config = load_config()
    
    # Validate early so CLI fails fast on unsupported COM.
    try:
        backend = resolve_backend(args.backend, config.get("backend"))
    except (ValueError, RuntimeError) as exc:
        Console().print(f"[red]{exc}[/red]")
        return 1


    q1, q2, q3 = content[0], content[1], content[2]
    addin_paths = content[3:5] if len(content) >= 5 else []

    console = Console()
    console.print(q1)
    selected_set = select(list_sets(rows), cursor_style="cyan")

    console.print(q2)
    selected_dir = select(list_ink_colors(rows, selected_set))
    console.print(f"Selected: {selected_dir}")

    if args.root:
        for r in rows:
            if r["set_name"] == selected_set:
                r["dir"] = str(Path(args.root) / r["rel_dir"])

    batch_serial = prompt(q3, target_type=str)
    result = create_datasheet(
        selected_set=selected_set,
        selected_dir=selected_dir,
        batch_serial=batch_serial,
        backend=backend,
        rows=rows,
        addin_paths=addin_paths,
        config=config,
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
