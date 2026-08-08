@echo off
REM Build dist\XlsxClean\XlsxClean.exe on Windows, then copy a Desktop shortcut.
setlocal
cd /d "%~dp0\.."

echo Installing packaging deps...
py -3 -m pip install -e ".[desktop]"
if errorlevel 1 (
  echo pip install failed. Ensure Python 3.12+ is installed.
  pause
  exit /b 1
)

echo Building executable with PyInstaller...
py -3 -m PyInstaller --noconfirm packaging\xlsx_clean.spec
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
  "$w=New-Object -ComObject WScript.Shell;" ^
  "$s=$w.CreateShortcut($lnk);" ^
  "$s.TargetPath='%EXE%';" ^
  "$s.WorkingDirectory='%~dp0..\dist\XlsxClean';" ^
  "$s.Description='Open xlsx-clean';" ^
  "$s.Save();" ^
  "Write-Host Created $lnk"

echo.
echo Done. Double-click the XlsxClean icon on your Desktop.
echo Exe folder: dist\XlsxClean\
pause
