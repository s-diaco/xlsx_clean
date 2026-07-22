@echo off
REM One-time Windows setup with uv: sync deps from uv.lock, Desktop shortcut.
setlocal
cd /d "%~dp0\.."

where uv >nul 2>&1
if errorlevel 1 (
  echo uv was not found on PATH.
  echo Install uv from https://docs.astral.sh/uv/getting-started/installation/
  echo then re-run this script.
  pause
  exit /b 1
)

echo Syncing project with uv (creates .venv if needed)...
uv sync
if errorlevel 1 (
  echo uv sync failed.
  pause
  exit /b 1
)

echo Creating Desktop shortcut...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-DesktopShortcut.ps1"
if errorlevel 1 (
  echo Desktop shortcut step failed.
  pause
  exit /b 1
)

echo.
echo Setup complete. Double-click the XlsxClean icon on your Desktop.
pause
