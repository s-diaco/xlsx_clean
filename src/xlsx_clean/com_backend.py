"""Windows Excel COM backend (pywin32). Imported only when selected."""

from __future__ import annotations

from pathlib import Path


def _bring_excel_to_front(excel, workbook) -> None:
    """Maximize and force Excel ahead of other windows (e.g. NiceGUI)."""
    import win32com.client as win32
    import win32con
    import win32gui

    excel.Visible = True
    excel.WindowState = win32.constants.xlMaximized
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
    # Lazy import so Linux never loads pywin32.
    import win32com.client as win32

    source = Path(source)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)

    # AttributeError: ... has no attribute 'CLSIDToClassMap'
    # https://stackoverflow.com/questions/52889704
    excel = win32.gencache.EnsureDispatch("Excel.Application")
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
