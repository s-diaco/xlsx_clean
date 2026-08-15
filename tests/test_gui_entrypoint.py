from pathlib import Path
from unittest.mock import MagicMock, patch

from xlsx_clean import gui_app


def test_parse_args_uses_default_dimensions():
    args = gui_app.parse_args([])
    assert (args.width, args.height) == (800, 600)


def test_main_initializes_api_before_opening_window():
    index_html = Path("/tmp/index.html")
    events = []
    api = MagicMock()

    def create_window(**kwargs):
        events.append(("create_window", kwargs))

    with (
        patch("xlsx_clean.gui_app.Api", return_value=api),
        patch("xlsx_clean.gui_app.package_file", return_value=index_html) as package_file,
        patch("xlsx_clean.gui_app.webview.create_window", side_effect=create_window),
        patch("xlsx_clean.gui_app.webview.start", side_effect=lambda: events.append(("start",))),
    ):
        gui_app.main(["--width", "900", "--height", "700"])

    package_file.assert_called_once_with("ui/index.html")
    api.initialize.assert_called_once_with()
    assert events == [
        (
            "create_window",
            {
                "title": f"New QC Sheet · v{gui_app._VERSION}",
                "url": index_html.as_uri(),
                "js_api": api,
                "width": 900,
                "height": 700,
                "min_size": (450, 400),
                "background_color": "#0b0e14",
            },
        ),
        ("start",),
    ]
