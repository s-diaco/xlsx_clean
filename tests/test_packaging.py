from pathlib import Path


def test_pyinstaller_spec_collects_webview_and_local_assets():
    spec = Path("packaging/xlsx_clean.spec").read_text(encoding="utf-8")

    assert 'collect_all("webview")' in spec
    assert 'collect_all("dearpygui")' not in spec
    assert '"ui"), "xlsx_clean/ui")' in spec
    assert '"fonts"), "xlsx_clean/fonts")' in spec
    assert 'if sys.platform == "win32":' in spec
    assert 'collect_all("clr_loader")' in spec
    for hidden_import in ('"clr"', '"clr_loader"', '"pythonnet"'):
        assert hidden_import in spec
