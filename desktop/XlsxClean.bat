@echo off
REM Double-click launcher with visible console (useful for troubleshooting).
setlocal
cd /d "%~dp0\.."
set PORT=8080

if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -m xlsx_clean.web_app --host 127.0.0.1 --port %PORT%
) else (
  py -3 -m xlsx_clean.web_app --host 127.0.0.1 --port %PORT%
)
