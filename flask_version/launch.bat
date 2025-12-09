@echo off
title Measurement Data Analyzer
echo ========================================
echo   Measurement Data Analyzer
echo ========================================
echo.
echo Starting server...
echo.

:: Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/
    pause
    exit /b 1
)

:: Check if required packages are installed, install if missing
pip show flask >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    pip install -r requirements.txt
)

:: Start the browser after a short delay (in background)
start "" cmd /c "timeout /t 2 >nul && start http://localhost:5000"

:: Start the Flask server
echo.
echo Server running at http://localhost:5000
echo Press Ctrl+C to stop the server
echo.
python app.py
