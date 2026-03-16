@echo off
cd /d "%~dp0"
echo ============================================
echo Mail Attachment Bot - Modern Desktop App v5.0
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
    echo [INFO] CustomTkinter not installed, installing...
    pip install -r requirements_gui.txt
)

pip show watchdog >nul 2>&1
if errorlevel 1 (
    echo [INFO] Watchdog not installed, installing...
    pip install watchdog
)

echo [2/3] Starting desktop application v5.0...
echo.

python gui_app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Application error occurred
    pause
)
