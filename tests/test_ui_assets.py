import re
from html.parser import HTMLParser

from xlsx_clean.paths import package_file


class StrictHtmlParser(HTMLParser):
    def error(self, message):
        raise AssertionError(message)


def test_ui_asset_exists_and_has_required_bridge_content():
    html_path = package_file("ui/index.html")
    assert html_path.is_file()
    content = html_path.read_text(encoding="utf-8")

    for expected in (
        "New QC Sheet",
        "Create a new QC datasheet from the latest matching workbook.",
        'id="set-label"',
        'id="ink-label"',
        'id="serial-label"',
        "Create datasheet",
        "status-ready",
        "status-success",
        "status-skipped",
        "status-error",
        'id="version"',
        "pywebviewready",
        "get_initial_data()",
        "get_inks(setDropdown.value)",
        "api.create(",
        "finally",
    ):
        assert expected in content

    assert "tailwind" not in content.lower()
    assert "quasar" not in content.lower()
    assert "cdn" not in content.lower()
    assert not re.search(r"https?://", content)
    assert "#0b0e14" in content
    assert "#e54d5e" in content
    assert "background-size: var(--grid-size) var(--grid-size)" in content
    StrictHtmlParser().feed(content)


def test_ui_font_sources_are_bundled_locally():
    content = package_file("ui/index.html").read_text(encoding="utf-8")
    font_urls = re.findall(r'url\("([^\"]+\.ttf)"\)', content)

    assert font_urls == ["../fonts/Inter-Regular.ttf", "../fonts/Inter-Bold.ttf"]
    for font_url in font_urls:
        assert (package_file("ui") / font_url).resolve().is_file()
