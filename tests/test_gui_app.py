from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from xlsx_clean.gui_app import Api


def test_initial_data_stores_state_and_uses_first_set_inks():
    rows = [{"set_name": "Set A"}]
    config = {"backend": "ooxml"}
    with patch(
        "xlsx_clean.gui_app.load_config",
        return_value=(["Set", "Ink", "Serial", "addin-a", "addin-b"], rows, config),
    ), patch("xlsx_clean.gui_app.list_sets", return_value=["Set A"]), patch(
        "xlsx_clean.gui_app.list_ink_colors", return_value=["Blue"]
    ) as inks:
        api = Api()
        data = api.get_initial_data()

    assert data["sets"] == ["Set A"]
    assert data["inks"] == ["Blue"]
    assert data["labels"] == {"set": "Set", "ink": "Ink", "serial": "Serial"}
    assert api.rows == rows
    assert api.addin_paths == ["addin-a", "addin-b"]
    assert api.config == config
    inks.assert_called_once_with(rows, "Set A")


def test_initial_data_uses_safe_label_fallbacks_and_empty_inks():
    with (
        patch("xlsx_clean.gui_app.load_config", return_value=([], [], {})),
        patch("xlsx_clean.gui_app.list_sets", return_value=[]),
        patch("xlsx_clean.gui_app.list_ink_colors") as inks,
    ):
        data = Api().get_initial_data()

    assert data["labels"] == {
        "set": "Select Set",
        "ink": "Select Ink Color",
        "serial": "Serial:",
    }
    assert data["inks"] == []
    inks.assert_not_called()


def test_get_inks_uses_initialized_rows():
    api = Api()
    api.rows = [{"set_name": "Set A"}]
    with patch("xlsx_clean.gui_app.list_ink_colors", return_value=["Blue"]) as inks:
        assert api.get_inks("Set A") == ["Blue"]

    inks.assert_called_once_with(api.rows, "Set A")


def test_create_passes_core_arguments_and_serializes_paths():
    api = Api()
    api.rows = [{"set_name": "Set A"}]
    api.addin_paths = ["addin-a"]
    api.config = {"backend": "ooxml"}
    result = SimpleNamespace(
        ok=True,
        skipped=False,
        message="Wrote workbook",
        template=Path("template.xlsx"),
        destination=Path("output.xlsx"),
        backend="ooxml",
    )
    with patch("xlsx_clean.gui_app.create_datasheet", return_value=result) as create:
        response = api.create("Set A", "Blue", "1234/A")

    create.assert_called_once_with(
        selected_set="Set A",
        selected_dir="Blue",
        batch_serial="1234/A",
        backend=None,
        rows=api.rows,
        addin_paths=api.addin_paths,
        config=api.config,
    )
    assert response == {
        "ok": True,
        "skipped": False,
        "message": "Wrote workbook",
        "template": "template.xlsx",
        "destination": "output.xlsx",
        "backend": "ooxml",
    }


def test_create_serializes_missing_paths_for_error_or_skipped_result():
    api = Api()
    result = SimpleNamespace(
        ok=False,
        skipped=False,
        message="Select a set.",
        template=None,
        destination=None,
        backend=None,
    )
    with patch("xlsx_clean.gui_app.create_datasheet", return_value=result):
        response = api.create("", "", "")

    assert response["template"] is None
    assert response["destination"] is None
    assert response["ok"] is False


def test_create_serializes_unexpected_exceptions_into_result_contract():
    api = Api()
    api.rows = [{"set_name": "Set A"}]
    api.addin_paths = []
    api.config = {"backend": "ooxml"}

    def boom(**kwargs):
        raise RuntimeError("backend exploded")

    with patch("xlsx_clean.gui_app.create_datasheet", side_effect=boom):
        response = api.create("Set A", "Blue", "1234/A")

    assert response == {
        "ok": False,
        "skipped": False,
        "message": "Unexpected error: backend exploded",
        "template": None,
        "destination": None,
        "backend": None,
    }
