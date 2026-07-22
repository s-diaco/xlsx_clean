@echo off
REM Build dist\XlsxClean\XlsxClean.exe on Windows with uv, then Desktop shortcut.
setlocal
cd /d "%~dp0\.."

where uv >nul 2>&1
if errorlevel 1 (
  echo uv was not found on PATH.
  echo Install uv from https://docs.astral.sh/uv/getting-started/installation/
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
)

echo Installing packaging deps with uv...
uv pip install -e ".[desktop]" -p .venv\Scripts\python.exe
if errorlevel 1 (
  echo uv pip install failed.
  pause
  exit /b 1
)

echo Building executable with PyInstaller...
uv run --python .venv\Scripts\python.exe pyinstaller --noconfirm packaging\xlsx_clean.spec
if errorlevel 1 (
  echo PyInstaller build failed.
  pause
  exit /b 1
)

set EXE=%CD%\dist\XlsxClean\XlsxClean.exe
if not exist "%EXE%" (
  echo Expected exe not found: %EXE%
  pause
  exit /b 1
)

echo Creating Desktop shortcut to the exe...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$desk=[Environment]::GetFolderPath('Desktop');" ^
  "$lnk=Join-Path $desk 'XlsxClean.lnk';" ^
  "$ico=Join-Path '%CD%' 'desktop\xlsx-clean.ico';" ^
  "$w=New-Object -ComObject WScript.Shell;" ^
  "$s=$w.CreateShortcut($lnk);" ^
  "$s.TargetPath='%EXE%';" ^
  "$s.WorkingDirectory='%~dp0..\dist\XlsxClean';" ^
  "$s.Description='Open xlsx-clean';" ^
  "if (Test-Path $ico) { $s.IconLocation = $ico };" ^
  "$s.Save();" ^
  "Write-Host Created $lnk"

echo.
echo Done. Double-click the XlsxClean icon on your Desktop.
echo Exe folder: dist\XlsxClean\
pause
