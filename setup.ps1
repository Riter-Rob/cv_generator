# One-time setup for the CV Generator (Windows PowerShell).
# Creates the virtual environment, installs dependencies, and builds the
# starter template + Excel input sheet.
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Write-Host "Creating virtual environment..." -ForegroundColor Cyan
python -m venv "$root\.venv"
$py = "$root\.venv\Scripts\python.exe"
Write-Host "Upgrading pip..." -ForegroundColor Cyan
& $py -m pip install --upgrade pip
Write-Host "Installing dependencies (this can take several minutes)..." -ForegroundColor Cyan
& $py -m pip install -r "$root\requirements.txt"
Write-Host "Building template + Excel input sheet..." -ForegroundColor Cyan
& $py "$root\src\build_template.py"
if (-not (Test-Path "$root\input\applicants.xlsx")) {
    & $py "$root\src\excel_io.py"
} else {
    Write-Host "input\applicants.xlsx already exists - keeping it." -ForegroundColor Yellow
}
Write-Host "`nSetup complete." -ForegroundColor Green
Write-Host "Next: put passport images in input\passports\, fill input\applicants.xlsx, then run  .\run.ps1"
