"""Windows Excel COM backend (pywin32). Imported only when selected."""

from __future__ import annotations

import logging
import shutil
import sys
from pathlib import Path

# Excel xlMaximized — avoid win32.constants when gen_py early-binding cache is corrupt.
_XL_MAXIMIZED = -4137

_log = logging.getLogger("xlsx_clean.create")


def _is_gen_py_cache_error(exc: BaseException) -> bool:
    """True for the known corrupt gen_py AttributeError (CLSIDTo*Map missing)."""
    if not isinstance(exc, AttributeError):
        return False
    msg = str(exc)
    return "CLSIDToClassMap" in msg or "CLSIDToPackageMap" in msg


def _clear_gen_py_cache() -> Path | None:
    """Remove win32com's generated typelib cache and drop it from sys.modules."""
    import win32com

    gen_path = Path(win32com.__gen_path__)
    for name in [m for m in sys.modules if m.startswith("win32com.gen_py")]:
        del sys.modules[name]
    if gen_path.is_dir():
        shutil.rmtree(gen_path, ignore_errors=True)
        return gen_path
    return None


def _excel_application():
    """Return Excel.Application, recovering from a corrupt win32com gen_py cache.

    ``gencache.EnsureDispatch`` writes early-binding wrappers under ``%TEMP%\\gen_py``.
    Interrupted or partial writes leave modules without ``CLSIDToClassMap`` /
    ``CLSIDToPackageMap``; deleting that folder and retrying fixes it (same as the
    manual workaround). Falls back to late-bound ``Dispatch`` if rebuild still fails.
    """
    import win32com.client as win32

    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except AttributeError as exc:
        if not _is_gen_py_cache_error(exc):
            raise
        cleared = _clear_gen_py_cache()
        _log.warning(
            "Corrupt win32com gen_py cache (%s); cleared %s and retrying",
            exc,
            cleared or "(missing)",
        )
        try:
            return win32.gencache.EnsureDispatch("Excel.Application")
        except AttributeError as retry_exc:
            if not _is_gen_py_cache_error(retry_exc):
                raise
            _log.warning(
                "EnsureDispatch still failing after cache clear (%s); using late-bound Dispatch",
                retry_exc,
            )
            return win32.Dispatch("Excel.Application")


def _bring_excel_to_front(excel, workbook) -> None:
    """Maximize and force Excel ahead of other windows (e.g. NiceGUI)."""
    import win32con
    import win32gui

    excel.Visible = True
    excel.WindowState = _XL_MAXIMIZED
    try:
        workbook.Activate()
    except Exception:
        pass
    try:
        hwnd = int(excel.Hwnd)
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        # Allow SetForegroundWindow to succeed when another app has focus.
        try:
            import win32api
            import win32process

            fg = win32gui.GetForegroundWindow()
            fg_tid, _ = win32process.GetWindowThreadProcessId(fg)
            our_tid = win32api.GetCurrentThreadId()
            if fg_tid != our_tid:
                win32process.AttachThreadInput(our_tid, fg_tid, True)
                try:
                    win32gui.SetForegroundWindow(hwnd)
                finally:
                    win32process.AttachThreadInput(our_tid, fg_tid, False)
            else:
                win32gui.SetForegroundWindow(hwnd)
        except Exception:
            win32gui.SetForegroundWindow(hwnd)
    except Exception:
        # Focus stealing can fail under some Windows policies; create still succeeded.
        pass


def clean_workbook_com(
    source: Path | str,
    destination: Path | str,
    cells_to_clear: str,
    notes_cell: str,
    serial_cell: str,
    batch_serial: str,
    addin_paths: list[str] | None = None,
    notes_value: str = "",
    visible: bool = True,
    maximize: bool = True,
) -> None:
    """Clear/set cells via Excel COM, SaveAs destination, leave Excel open."""
    # Lazy import so Linux never loads pywin32 (only needed if callers import helpers).
    source = Path(source)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)

    excel = _excel_application()
    excel.Visible = visible

    workbook = excel.Workbooks.Open(str(source.resolve()))
    for addin_path in addin_paths or []:
        if addin_path:
            excel.Workbooks.Open(addin_path)

    for workbook_data in cells_to_clear.split(","):
        workbook_data = workbook_data.strip()
        if not workbook_data:
            continue
        sheet_name, a1 = workbook_data.split("!", 1)
        worksheet = workbook.Worksheets(sheet_name.replace("'", ""))
        worksheet.Range(a1).ClearContents()

    for workbook_data in notes_cell.split(","):
        workbook_data = workbook_data.strip()
        if not workbook_data:
            continue
        sheet_name, a1 = workbook_data.split("!", 1)
        worksheet = workbook.Worksheets(sheet_name.replace("'", ""))
        worksheet.Range(a1).Value = notes_value

    for workbook_data in serial_cell.split(","):
        workbook_data = workbook_data.strip()
        if not workbook_data:
            continue
        sheet_name, a1 = workbook_data.split("!", 1)
        worksheet = workbook.Worksheets(sheet_name.replace("'", ""))
        worksheet.Range(a1).Value = batch_serial

    workbook.SaveAs(str(destination.resolve()))
    if maximize or visible:
        _bring_excel_to_front(excel, workbook)
