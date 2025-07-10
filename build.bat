@echo off
echo Building Gold Transaction Manager for Windows...

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Install required packages
echo Installing required packages...
pip install openpyxl==3.1.2
pip install pyinstaller==6.2.0

REM Check if installation was successful
if %errorlevel% neq 0 (
    echo ERROR: Failed to install packages
    pause
    exit /b 1
)

REM Build the executable
echo Building executable...
pyinstaller --onefile --windowed --name "GoldTransactionManager" gold_app.py

REM Check if build was successful
if exist "dist\GoldTransactionManager.exe" (
    echo.
    echo ✅ SUCCESS! EXE file created successfully!
    echo Location: %cd%\dist\GoldTransactionManager.exe
    echo.
    echo You can now run the application by double-clicking the EXE file.
    echo The application will create an Excel file in the same directory.
    echo.
    
    REM Copy the exe to current directory for easy access
    copy "dist\GoldTransactionManager.exe" "GoldTransactionManager.exe"
    echo EXE file copied to current directory for easy access.
    
) else (
    echo ERROR: Build failed. Check the output above for errors.
)

echo.
echo Press any key to exit...
pause >nul