@echo off
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on this computer.
    echo Please install Python from https://www.python.org/downloads/
    echo During setup, check the box "Add python.exe to PATH"
    pause
    exit /b
)

if not exist ".venv" (
    echo First-time setup, please wait...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    pip install -r requirements.txt
) else (
    call .venv\Scripts\activate.bat
)

python app.py
pause
