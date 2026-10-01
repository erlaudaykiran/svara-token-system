@echo off
title SVARA 2026 Token System (Local Laptop Server)
cd /d "%~dp0"

echo ======================================================================
echo   SVARA 2026 LUCKY DRAW - PERMANENT LAPTOP SERVER
echo ======================================================================
echo   * All token data is saved PERMANENTLY on this laptop in:
echo     - database\svara.db (Full SQLite Database)
echo     - exports\SVARA_2026_Tokens.xlsx (Full Excel File)
echo   * Deploys to GitHub/Render will NEVER overwrite your laptop files!
echo ======================================================================
echo.

set PYTHON_CMD=python
if exist "C:\Users\%USERNAME%\anaconda3\python.exe" (
    set PYTHON_CMD="C:\Users\%USERNAME%\anaconda3\python.exe"
)

echo Starting SVARA Server on your laptop...
echo Opening http://localhost:5000 in your browser...
start "" http://localhost:5000

%PYTHON_CMD% backend/app.py
pause
