@echo off
echo ========================================
echo Building Form Filler Executable...
echo ========================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python first to build the executable
    pause
    exit /b 1
)

REM Install PyInstaller if not present
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt

REM Run the build
echo.
echo Building executable...
python build.py

echo.
echo ========================================
if exist "dist\FormFiller.exe" (
    echo SUCCESS! Executable created at:
    echo %~dp0dist\FormFiller.exe
    echo.
    echo You can now distribute this file to users.
    echo They do NOT need Python installed.
) else (
    echo BUILD FAILED - Check errors above
)
echo ========================================
pause
