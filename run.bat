@echo off
cd /d "%~dp0backend"
where python >nul 2>nul || (echo Python is not installed. Get it from https://www.python.org/downloads/ ^(tick "Add Python to PATH"^) & pause & exit /b 1)
if not exist "..\venv" (
  echo First run: installing requirements...
  python -m venv ..\venv
  ..\venv\Scripts\pip install -r requirements.txt
)
start "" http://localhost:5000
..\venv\Scripts\python app.py
pause
