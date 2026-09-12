# Launch the CV Generator desktop app (opens in your browser).
# Also reachable from phones/PCs on the same Wi-Fi via the Network URL it prints.
$root = $PSScriptRoot
$py = "$root\.venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Host "Virtual environment not found. Run .\setup.ps1 first." -ForegroundColor Red
    exit 1
}
& $py -m streamlit run "$root\src\app.py" --server.port 8531 --browser.gatherUsageStats false
