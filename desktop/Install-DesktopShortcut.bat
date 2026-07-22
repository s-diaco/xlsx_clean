@echo off
REM Create a Desktop shortcut that points at desktop\XlsxClean.vbs
setlocal
cd /d "%~dp0\.."

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-DesktopShortcut.ps1"
if errorlevel 1 (
  echo Failed to create desktop shortcut.
  pause
  exit /b 1
)
echo Desktop shortcut created: XlsxClean
pause
