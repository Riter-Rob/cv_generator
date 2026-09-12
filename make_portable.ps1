# Build a portable .zip of this tool to copy to another Windows PC.
# Excludes the virtual environment, outputs and caches (those are rebuilt by
# setup.ps1 on the target machine).
$root = $PSScriptRoot
$stamp = Get-Date -Format "yyyyMMdd"
$zip = Join-Path (Split-Path $root -Parent) "cvgen_portable_$stamp.zip"

$staging = Join-Path $env:TEMP "cvgen_pkg"
if (Test-Path $staging) { Remove-Item -Recurse -Force $staging }
New-Item -ItemType Directory -Path $staging | Out-Null

# copy everything except the non-portable / generated folders
robocopy $root $staging /E /XD ".venv" "output" "__pycache__" /XF "*.pyc" | Out-Null

# keep an empty output structure so the tool has somewhere to write
New-Item -ItemType Directory -Force -Path (Join-Path $staging "output\docx") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $staging "output\pdf") | Out-Null

if (Test-Path $zip) { Remove-Item $zip }
Compress-Archive -Path (Join-Path $staging "*") -DestinationPath $zip
Remove-Item -Recurse -Force $staging

Write-Host "Portable package created:" -ForegroundColor Green
Write-Host "  $zip"
Write-Host "On the other PC: install Python 3.10+ and MS Word, unzip, then run  .\setup.ps1  and  .\run_app.ps1"
