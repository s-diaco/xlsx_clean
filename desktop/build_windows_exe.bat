@echo off
REM Build dist\New QC Sheet\New QC Sheet.exe on Windows with uv, then Desktop shortcut.
setlocal
cd /d "%~dp0\.."

where uv >nul 2>&1
if errorlevel 1 (
  echo uv was not found on PATH.
  echo Install uv from https://docs.astral.sh/uv/getting-started/installation/
  pause
  exit /b 1
)

echo Syncing project + desktop extras with uv...
uv sync --extra desktop
if errorlevel 1 (
  echo uv sync failed.
  pause
  exit /b 1
)

echo Building executable with PyInstaller...
uv run pyinstaller --noconfirm packaging\xlsx_clean.spec
if errorlevel 1 (
  echo PyInstaller build failed.
  pause
  exit /b 1
)

set EXE=%CD%\dist\New QC Sheet\New QC Sheet.exe
if not exist "%EXE%" (
  echo Expected exe not found: %EXE%
  pause
  exit /b 1
)

echo Creating Desktop shortcut to the exe...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$desk=[Environment]::GetFolderPath('Desktop');" ^
  "$lnk=Join-Path $desk 'New QC Sheet.lnk';" ^
  "$legacy=Join-Path $desk 'XlsxClean.lnk';" ^
  "if (Test-Path $legacy) { Remove-Item $legacy -Force };" ^
  "$ico=Join-Path '%CD%' 'desktop\xlsx-clean.ico';" ^
  "$w=New-Object -ComObject WScript.Shell;" ^
  "$s=$w.CreateShortcut($lnk);" ^
  "$s.TargetPath='%EXE%';" ^
  "$s.WorkingDirectory='%~dp0..\dist\New QC Sheet';" ^
  "$s.Description='Open New QC Sheet';" ^
  "if (Test-Path $ico) { $s.IconLocation = $ico };" ^
  "$s.Save();" ^
  "Write-Host Created $lnk"

echo.
echo Done. Double-click the New QC Sheet icon on your Desktop.
echo Exe folder: dist\New QC Sheet\
pause
