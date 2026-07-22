# Creates %USERPROFILE%\Desktop\New QC Sheet.lnk -> desktop\XlsxClean.vbs
$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$target = Join-Path $PSScriptRoot "XlsxClean.vbs"
$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "New QC Sheet.lnk"

if (-not (Test-Path $target)) {
    throw "Launcher not found: $target"
}

# Remove older shortcut name if present
$legacy = Join-Path $desktop "XlsxClean.lnk"
if (Test-Path $legacy) {
    Remove-Item $legacy -Force
}

$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $target
$shortcut.WorkingDirectory = $repoRoot
$shortcut.WindowStyle = 7
$shortcut.Description = "Open New QC Sheet"
$iconCandidate = Join-Path $PSScriptRoot "xlsx-clean.ico"
if (Test-Path $iconCandidate) {
    $shortcut.IconLocation = "$iconCandidate,0"
}
$shortcut.Save()
Write-Host "Created $shortcutPath"
