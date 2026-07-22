"""NiceGUI web UI for selecting set / ink / serial and creating a datasheet."""

from __future__ import annotations

import argparse
import os
import sys

from nicegui import ui

from xlsx_clean.clean_cells import (
    create_datasheet,
    list_ink_colors,
    list_sets,
    load_config,
)
from xlsx_clean.paths import default_backend

APP_DISPLAY_NAME = "New QC Sheet"

_LINUX_NATIVE_HINT = (
    "Linux native window needs system GTK/WebKit packages "
    "(uv cannot install them), e.g.:\n"
    "  sudo apt install python3-gi gir1.2-gtk-3.0 gir1.2-webkit2-4.1\n"
    "Falling back to the default browser. "
    "Use --no-browser for server-only, or install the packages above for a "
    "single app window."
)


def _linux_has_display() -> bool:
    return bool(os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"))


def _linux_webview_libs_available() -> bool:
    """True if GTK (gi) or Qt (qtpy) Python bindings are importable."""
    try:
        import gi  # noqa: F401

        return True
    except ImportError:
        pass
    try:
        import qtpy  # noqa: F401

        return True
    except ImportError:
        return False


def _native_mode_available() -> bool:
    """Whether pywebview is likely able to open a native window."""
    if sys.platform == "win32" or sys.platform == "darwin":
        return True
    if sys.platform.startswith("linux"):
        return _linux_has_display() and _linux_webview_libs_available()
    return _linux_webview_libs_available()


def _build_page() -> None:
    content, path_df = load_config()
    q1, q2, q3 = content[0], content[1], content[2]
    sets = list_sets(path_df)
    initial_set = sets[0] if sets else None
    initial_inks = list_ink_colors(path_df, initial_set) if initial_set else []

    ui.page_title(APP_DISPLAY_NAME)
    ui.colors(primary="#0f766e", secondary="#115e59", accent="#14b8a6")

    with ui.column().classes("w-full max-w-xl mx-auto p-6 gap-4"):
        ui.label(APP_DISPLAY_NAME).classes("text-3xl font-bold text-teal-800")
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
        description=f"Start the {APP_DISPLAY_NAME} NiceGUI interface."
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
    if use_native and not _native_mode_available():
        if os.environ.get("XLSX_CLEAN_NATIVE_FALLBACK") != "1":
            print(_LINUX_NATIVE_HINT, file=sys.stderr)
            os.environ["XLSX_CLEAN_NATIVE_FALLBACK"] = "1"
        use_native = False
        args.browser = True

    run_kwargs: dict = {
        "host": args.host,
        "port": args.port,
        "reload": False if frozen or use_native else args.reload,
        "title": APP_DISPLAY_NAME,
    }
    if args.no_browser:
        run_kwargs["show"] = False
        run_kwargs["native"] = False
    elif use_native:
        # Single dedicated window (pywebview), not the default browser with tabs.
        run_kwargs["native"] = True
        run_kwargs["window_size"] = (900, 700)
        run_kwargs["show"] = False
    else:
        run_kwargs["show"] = True
        run_kwargs["native"] = False

    try:
        ui.run(**run_kwargs)
    except Exception as exc:
        # Native mode can still fail at runtime (e.g. missing WebKit gir).
        msg = str(exc).lower()
        if run_kwargs.get("native") and (
            "gtk" in msg or "qt" in msg or "webview" in msg or "gi" in msg
        ):
            print(_LINUX_NATIVE_HINT, file=sys.stderr)
            print(f"(native start failed: {exc})", file=sys.stderr)
            run_kwargs["native"] = False
            run_kwargs.pop("window_size", None)
            run_kwargs["show"] = True
            run_kwargs["reload"] = False if frozen else args.reload
            ui.run(**run_kwargs)
        else:
            raise


if __name__ in {"__main__", "__mp_main__"}:
    # NiceGUI may re-import as __mp_main__ when reload is enabled.
    main()
