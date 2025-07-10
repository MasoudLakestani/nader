@echo off
echo Starting Gold Transaction Manager...

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please run setup.bat first
    pause
    exit /b 1
)

REM Check if gold_app.py exists
if not exist "gold_app.py" (
    echo ERROR: gold_app.py not found
    echo Make sure all files are in the same folder
    pause
    exit /b 1
)

REM Run the application
echo Running application...
python gold_app.py

REM If the application closes, pause to see any error messages
if %errorlevel% neq 0 (
    echo.
    echo Application closed with error code: %errorlevel%
    echo Press any key to exit...
    pause >nul
)