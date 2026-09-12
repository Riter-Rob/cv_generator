@echo off
REM Double-click to open the CV Generator desktop app in your browser.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Virtual environment not found. Run setup.ps1 first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m streamlit run "src\app.py" --server.port 8531 --browser.gatherUsageStats false
