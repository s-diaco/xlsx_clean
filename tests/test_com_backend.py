"""Tests for win32com gen_py cache recovery (no real Excel required)."""

from __future__ import annotations

import sys
from types import ModuleType, SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest

from xlsx_clean.com_backend import (
    _clear_gen_py_cache,
    _excel_application,
    _is_gen_py_cache_error,
)


def _install_fake_win32com(*, ensure, dispatch=None, gen_path="/tmp/gen_py"):
    """Put fake win32com + win32com.client modules into sys.modules."""
    win32com = ModuleType("win32com")
    win32com.__gen_path__ = gen_path
    client = ModuleType("win32com.client")
    client.gencache = SimpleNamespace(EnsureDispatch=ensure)
    client.Dispatch = dispatch if dispatch is not None else MagicMock()
    return patch.dict(sys.modules, {"win32com": win32com, "win32com.client": client})


def test_is_gen_py_cache_error_matches_known_messages():
    assert _is_gen_py_cache_error(
        AttributeError(
            "module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9' "
            "has no attribute 'CLSIDToPackageMap'"
        )
    )
    assert _is_gen_py_cache_error(
        AttributeError("... has no attribute 'CLSIDToClassMap'")
    )
    assert not _is_gen_py_cache_error(AttributeError("something else"))
    assert not _is_gen_py_cache_error(RuntimeError("CLSIDToPackageMap"))


def test_clear_gen_py_cache_removes_modules_and_directory(tmp_path):
    gen_root = tmp_path / "gen_py"
    bad = gen_root / "00020813-0000-0000-C000-000000000046x0x1x9"
    bad.mkdir(parents=True)
    (bad / "__init__.py").write_text("# corrupt", encoding="utf-8")

    mod_name = "win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9"
    fake_mod = ModuleType(mod_name)
    fake_win32com = ModuleType("win32com")
    fake_win32com.__gen_path__ = str(gen_root)

    with patch.dict(sys.modules, {"win32com": fake_win32com, mod_name: fake_mod}):
        cleared = _clear_gen_py_cache()
        assert cleared == gen_root
        assert not gen_root.exists()
        assert mod_name not in sys.modules


def test_excel_application_clears_cache_and_retries_ensure_dispatch():
    excel = object()
    ensure = MagicMock(
        side_effect=[
            AttributeError(
                "module 'win32com.gen_py.00020813x0x1x9' has no attribute 'CLSIDToPackageMap'"
            ),
            excel,
        ]
    )

    with (
        _install_fake_win32com(ensure=ensure),
        patch("xlsx_clean.com_backend._clear_gen_py_cache", return_value=None) as clear,
    ):
        result = _excel_application()

    assert result is excel
    assert ensure.call_count == 2
    clear.assert_called_once_with()


def test_excel_application_falls_back_to_dispatch_after_second_failure():
    excel = object()
    err = AttributeError("has no attribute 'CLSIDToClassMap'")
    ensure = MagicMock(side_effect=[err, err])
    dispatch = MagicMock(return_value=excel)

    with (
        _install_fake_win32com(ensure=ensure, dispatch=dispatch),
        patch("xlsx_clean.com_backend._clear_gen_py_cache", return_value=None),
    ):
        result = _excel_application()

    assert result is excel
    dispatch.assert_called_once_with("Excel.Application")


def test_excel_application_rethrows_unrelated_attribute_error():
    ensure = MagicMock(side_effect=AttributeError("unrelated"))

    with (
        _install_fake_win32com(ensure=ensure),
        pytest.raises(AttributeError, match="unrelated"),
    ):
        _excel_application()
