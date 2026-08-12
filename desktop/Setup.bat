@echo off
REM One-time Windows setup with uv: sync deps from uv.lock.
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

echo.
echo Setup complete. Create your own shortcut to desktop\XlsxClean.vbs.
pause
