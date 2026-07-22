"""NiceGUI web UI for selecting set / ink / serial and creating a datasheet."""

from __future__ import annotations

import argparse
import sys

from nicegui import ui

from xlsx_clean.clean_cells import (
    create_datasheet,
    list_ink_colors,
    list_sets,
    load_config,
)
from xlsx_clean.paths import default_backend


def _build_page() -> None:
    content, path_df = load_config()
    q1, q2, q3 = content[0], content[1], content[2]
    sets = list_sets(path_df)
    initial_set = sets[0] if sets else None
    initial_inks = list_ink_colors(path_df, initial_set) if initial_set else []

    ui.page_title("xlsx-clean")
    ui.colors(primary="#0f766e", secondary="#115e59", accent="#14b8a6")

    with ui.column().classes("w-full max-w-xl mx-auto p-6 gap-4"):
        ui.label("xlsx-clean").classes("text-3xl font-bold text-teal-800")
        ui.label(
            "Create a new QC datasheet from the latest matching workbook."
        ).classes("text-gray-600")

        ink_ref: dict = {}

        def on_set_change(e) -> None:
            selected = e.value
            inks = list_ink_colors(path_df, selected) if selected else []
            ink = ink_ref["select"]
            ink.options = inks
            ink.value = inks[0] if inks else None
            ink.update()

        set_select = ui.select(
            options=sets,
            value=initial_set,
            label=q1,
            with_input=True,
            on_change=on_set_change,
        ).classes("w-full")

        ink_select = ui.select(
            options=initial_inks,
            value=initial_inks[0] if initial_inks else None,
            label=q2,
            with_input=True,
        ).classes("w-full")
        ink_ref["select"] = ink_select

        serial_input = ui.input(label=q3, placeholder="e.g. 1234/A").classes("w-full")

        backend_select = ui.select(
            options=["ooxml", "com"],
            value=default_backend(),
            label="Backend",
        ).classes("w-full")
        if sys.platform != "win32":
            ui.label(
                "COM backend requires Windows with Microsoft Excel; use ooxml here."
            ).classes("text-sm text-amber-700")

        status = ui.markdown("Ready.").classes("w-full p-3 rounded bg-gray-50")

        def on_create() -> None:
            result = create_datasheet(
                selected_set=set_select.value or "",
                selected_dir=ink_select.value or "",
                batch_serial=serial_input.value or "",
                backend=backend_select.value,
                path_df=path_df,
                addin_paths=content[3:5] if len(content) >= 5 else [],
            )
            lines = [result.message]
            if result.template is not None:
                lines.append(f"**Template:** `{result.template.name}`")
            if result.destination is not None:
                lines.append(f"**Output:** `{result.destination}`")
            if result.backend is not None:
                lines.append(f"**Backend:** `{result.backend}`")
            status.set_content("\n\n".join(lines))
            if result.ok and not result.skipped:
                status.classes(replace="w-full p-3 rounded bg-teal-50")
            elif result.skipped:
                status.classes(replace="w-full p-3 rounded bg-amber-50")
            else:
                status.classes(replace="w-full p-3 rounded bg-red-50")

        ui.button("Create datasheet", on_click=on_create).props("unelevated").classes(
            "w-full"
        )


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Start the xlsx-clean NiceGUI interface."
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Bind address (default: 127.0.0.1).",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8080,
        help="Port (default: 8080; native mode may pick a free port).",
    )
    parser.add_argument(
        "--reload",
        action="store_true",
        help="Enable NiceGUI auto-reload (dev only; ignored in native mode).",
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--browser",
        action="store_true",
        help="Open in the default web browser instead of a native app window.",
    )
    mode.add_argument(
        "--no-browser",
        action="store_true",
        help="Run the server only; do not open a window or browser.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    _build_page()

    frozen = getattr(sys, "frozen", False)
    use_native = not args.browser and not args.no_browser
    run_kwargs: dict = {
        "host": args.host,
        "port": args.port,
        "reload": False if frozen or use_native else args.reload,
        "title": "xlsx-clean",
    }
    if args.no_browser:
        run_kwargs["show"] = False
        run_kwargs["native"] = False
    elif args.browser:
        run_kwargs["show"] = True
        run_kwargs["native"] = False
    else:
        # Single dedicated window (pywebview), not the default browser with tabs.
        run_kwargs["native"] = True
        run_kwargs["window_size"] = (900, 700)
        run_kwargs["show"] = False

    ui.run(**run_kwargs)


if __name__ in {"__main__", "__mp_main__"}:
    # NiceGUI may re-import as __mp_main__ when reload is enabled.
    main()
