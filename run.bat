@echo off
REM Double-click to generate CVs for every row in input\applicants.xlsx
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Virtual environment not found. Run setup.ps1 first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" "src\generate.py" %*
echo.
pause
