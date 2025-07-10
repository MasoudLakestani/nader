@echo off
echo Gold Transaction Manager - Windows Setup
echo =====================================

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed!
    echo.
    echo Please follow these steps:
    echo 1. Go to https://www.python.org/downloads/
    echo 2. Download Python 3.11 or newer
    echo 3. During installation, make sure to check "Add Python to PATH"
    echo 4. Run this setup script again
    echo.
    pause
    exit /b 1
)

echo Python is installed ✅
python --version

REM Check if pip is working
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: pip is not working
    pause
    exit /b 1
)

echo pip is working ✅
pip --version

REM Install required packages
echo.
echo Installing required packages...
echo Installing openpyxl...
pip install openpyxl==3.1.2
if %errorlevel% neq 0 (
    echo ERROR: Failed to install openpyxl
    pause
    exit /b 1
)

echo Installing pyinstaller...
pip install pyinstaller==6.2.0
if %errorlevel% neq 0 (
    echo ERROR: Failed to install pyinstaller
    pause
    exit /b 1
)

echo.
echo ✅ Setup completed successfully!
echo.
echo You can now:
echo 1. Run the application directly: python gold_app.py
echo 2. Build an EXE file: run build.bat
echo.
echo Press any key to exit...
pause >nul