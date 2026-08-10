@echo off
REM Double-click launcher with visible console (useful for troubleshooting).
setlocal
cd /d "%~dp0\.."

if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -m xlsx_clean.gui_app %*
) else (
  py -3 -m xlsx_clean.gui_app %*
)
