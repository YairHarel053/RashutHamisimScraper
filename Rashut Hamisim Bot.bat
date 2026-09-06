@echo off
setlocal
REM Double-click launcher for the Nadlan scraper web UI.
REM Creates a virtual environment and installs dependencies on first run,
REM then starts the local server.
cd /d "%~dp0"

set "PYTHON_CMD="
where py >nul 2>nul
if not errorlevel 1 (
    py -3.14 -c "import sys" >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=py -3.14"
)

if not defined PYTHON_CMD (
    where python >nul 2>nul
    if not errorlevel 1 (
        python -c "import sys" >nul 2>&1
        if not errorlevel 1 set "PYTHON_CMD=python"
    )
)

if not defined PYTHON_CMD (
    echo [setup] No usable Python interpreter was found.
    echo Please install Python 3.10+ from https://www.python.org/downloads/windows/.
    pause
    exit /b 1
)

REM Create or repair the virtual environment.
if not exist "venv\Scripts\python.exe" (
    echo [setup] Creating virtual environment...
    %PYTHON_CMD% -m venv venv
) else (
    REM Recreate the environment if it points to an unavailable launcher.
    "venv\Scripts\python.exe" -c "import sys; print(sys.executable)" >nul 2>&1
    if errorlevel 1 (
        echo [setup] Recreating virtual environment...
        rmdir /s /q venv
        %PYTHON_CMD% -m venv venv
    )
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
