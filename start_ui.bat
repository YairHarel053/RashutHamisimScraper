@echo off
REM Double-click launcher for the Nadlan scraper web UI.
REM Creates a virtual environment and installs dependencies on first run,
REM then starts the local server.
cd /d "%~dp0"

REM Create the virtual environment if it does not exist yet.
if not exist "venv\Scripts\python.exe" (
    echo [setup] Creating virtual environment...
    python -m venv venv
)

REM Make sure dependencies are installed (no-op if already present).
"venv\Scripts\python.exe" -c "import openpyxl, playwright" 2>nul
if errorlevel 1 (
    echo [setup] Installing dependencies...
    "venv\Scripts\python.exe" -m pip install --upgrade pip
    "venv\Scripts\python.exe" -m pip install -r requirements.txt
)

REM Start the web UI.
"venv\Scripts\python.exe" nadlan_ui.py
pause
