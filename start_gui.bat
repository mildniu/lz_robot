@echo off
cd /d "%~dp0"
echo ============================================
echo Quantum Bot - Tk Desktop App v5.2
echo ============================================
echo.

REM Check Python environment
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found, please install Python first
    pause
    exit /b 1
)

echo [1/3] Checking dependencies...
pip show customtkinter >nul 2>&1
if errorlevel 1 (
    echo [INFO] GUI dependencies not installed, installing...
    pip install -r requirements_gui.txt
)

pip show watchdog >nul 2>&1
if errorlevel 1 (
    echo [INFO] Watchdog not installed, installing...
    pip install watchdog
)

pip show pystray >nul 2>&1
if errorlevel 1 (
    echo [INFO] pystray not installed, installing...
    pip install pystray
)

echo [2/3] Starting Tk desktop application v5.2...
echo.

python gui_app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Application error occurred
    pause
)
