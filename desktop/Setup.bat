@echo off
REM One-time Windows setup with uv: create .venv, install package, Desktop shortcut.
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

if not exist ".venv\Scripts\python.exe" (
  echo Creating .venv with uv...
  uv venv .venv
  if errorlevel 1 (
    echo uv venv failed.
    pause
    exit /b 1
  )
) else (
  echo Using existing .venv
)

echo Installing xlsx-clean into .venv...
uv pip install -e . -p .venv\Scripts\python.exe
if errorlevel 1 (
  echo uv pip install failed.
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
