param(
  [string]$Name = "DS_Dashboard"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
  Write-Host "PyInstaller is not installed. Installing with pip..."
  python -m pip install pyinstaller
}

python -m PyInstaller `
  --clean `
  --noconfirm `
  --onefile `
  --name $Name `
  --add-data "dashboard.html;." `
  dashboard_server.py

Write-Host ""
Write-Host "Build complete: $root\dist\$Name.exe"
Write-Host "Put YYMM data folders, such as 2605, next to the exe before running it."
