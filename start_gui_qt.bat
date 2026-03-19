@echo off
cd /d "%~dp0"
echo ============================================
echo QuantumBot PySide6 v5.1
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found, please install Python first
    pause
    exit /b 1
)

pip show PySide6 >nul 2>&1
if errorlevel 1 (
    echo [INFO] PySide6 not installed, installing...
    pip install -r requirements_gui_qt.txt
)

echo [INFO] Starting PySide6 app...
python gui_qt_app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Application error occurred
    pause
)
