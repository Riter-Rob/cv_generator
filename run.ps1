# Run the CV generator. Any arguments are forwarded to generate.py.
# Examples:
#   .\run.ps1                 # generate DOCX + PDF for every row
#   .\run.ps1 --no-pdf        # DOCX only
#   .\run.ps1 --row 3         # only Excel row 3
#   .\run.ps1 --list          # list rows, generate nothing
$root = $PSScriptRoot
$py = "$root\.venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Host "Virtual environment not found. Run .\setup.ps1 first." -ForegroundColor Red
    exit 1
}
& $py "$root\src\generate.py" @args
