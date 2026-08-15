"""PyWebView desktop GUI for New QC Sheet."""

from __future__ import annotations

import argparse
from importlib.metadata import PackageNotFoundError, version as package_version

import webview

from xlsx_clean.core import create_datasheet, list_ink_colors, list_sets, load_config
from xlsx_clean.app_logging import get_logger
from xlsx_clean.paths import package_file

try:
    _VERSION = package_version("xlsx-clean")
except PackageNotFoundError:
    _VERSION = "dev"


class Api:
    """Python functions exposed to JavaScript via ``window.pywebview.api``."""

    def __init__(self) -> None:
        self.rows: list[dict] = []
        self.addin_paths: list[str] = []
        self.config: dict = {}
        self._initial_data: dict | None = None
        self._initialization_error: Exception | None = None

    def initialize(self) -> None:
        """Load configuration before the local frontend is opened."""
        if self._initial_data is not None or self._initialization_error is not None:
            return
        try:
            content, self.rows, self.config = load_config()
            sets = list_sets(self.rows)
            self.addin_paths = content[3:5] if len(content) >= 5 else []
            self._initial_data = {
                "sets": sets,
                "inks": list_ink_colors(self.rows, sets[0]) if sets else [],
                "labels": {
                    "set": content[0] if len(content) > 0 else "Select Set",
                    "ink": content[1] if len(content) > 1 else "Select Ink Color",
                    "serial": content[2] if len(content) > 2 else "Serial:",
                },
                "version": _VERSION,
            }
        except Exception as exc:  # noqa: BLE001 - serialize bridge failures to JS
            self._initialization_error = exc

    def get_initial_data(self) -> dict:
        """Return configuration-derived values for the initial page render."""
        self.initialize()
        if self._initialization_error is not None:
            raise self._initialization_error
        assert self._initial_data is not None
        return self._initial_data

    def get_inks(self, selected_set: str) -> list[str]:
        """Return ink colors for a selected set."""
        return list_ink_colors(self.rows, selected_set)

    def create(
        self, selected_set: str, selected_dir: str, batch_serial: str
    ) -> dict:
        """Create a datasheet and return a JSON-safe result.

        Unexpected exceptions are serialized into the result contract so the
        frontend always sees the same shape, regardless of platform-specific
        bridge behavior.
        """
        log = get_logger()
        try:
            result = create_datasheet(
                selected_set=selected_set,
                selected_dir=selected_dir,
                batch_serial=batch_serial,
                backend=None,
                rows=self.rows,
                addin_paths=self.addin_paths,
                config=self.config,
            )
        except Exception as exc:  # noqa: BLE001 - serialize bridge failures to JS
            log.exception(
                "Unexpected error in Api.create(set=%r, dir=%r, serial=%r)",
                selected_set,
                selected_dir,
                batch_serial,
            )
            return {
                "ok": False,
                "skipped": False,
                "message": f"Unexpected error: {exc}",
                "template": None,
                "destination": None,
                "backend": None,
            }
        return {
            "ok": result.ok,
            "skipped": result.skipped,
            "message": result.message,
            "template": str(result.template) if result.template is not None else None,
            "destination": (
                str(result.destination) if result.destination is not None else None
            ),
            "backend": result.backend,
        }


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse GUI dimensions without changing process-global arguments."""
    parser = argparse.ArgumentParser(
        description="Start the New QC Sheet PyWebView interface."
    )
    parser.add_argument("--width", type=int, default=800, help="Window width")
    parser.add_argument("--height", type=int, default=600, help="Window height")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    """Open the native PyWebView window."""
    args = parse_args(argv)
    api = Api()
    api.initialize()
    index_html = package_file("ui/index.html")
    webview.create_window(
        title=f"New QC Sheet · v{_VERSION}",
        url=index_html.as_uri(),
        js_api=api,
        width=args.width,
        height=args.height,
        min_size=(450, 400),
        background_color="#0b0e14",
    )
    webview.start()


if __name__ == "__main__":
    main()
