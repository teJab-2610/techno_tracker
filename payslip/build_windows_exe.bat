@echo off
REM ============================================================
REM  Build payslip_gui.py into a standalone Windows .exe
REM
REM  Prerequisites:
REM    1. Python 3.8+ installed on this Windows machine
REM    2. Run this script from the project directory
REM
REM  Output:
REM    dist\PayslipGenerator.exe  (single standalone file, no console)
REM ============================================================

echo.
echo === Payslip Generator - Windows Build Script ===
echo.

REM Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Download from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/5] Creating virtual environment...
if exist venv (
    echo Virtual environment already exists, reusing it.
) else (
    python -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment.
        pause
        exit /b 1
    )
)

echo.
echo [2/5] Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment.
    pause
    exit /b 1
)

echo.
echo [3/5] Upgrading pip and installing dependencies...
python -m pip install --upgrade pip
python -m pip install openpyxl reportlab qrcode pillow pyinstaller
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo [4/5] Building executable...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name PayslipGenerator ^
    --hidden-import=qrcode.image.pil ^
    --hidden-import=reportlab.graphics ^
    --hidden-import=PIL ^
    --add-data "generate_payslips.py;." ^
    payslip_gui.py

if errorlevel 1 (
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo.
echo [5/5] Cleaning up build artifacts...
rmdir /s /q build 2>nul
del /q PayslipGenerator.spec 2>nul

echo.
echo ============================================================
echo  BUILD SUCCESSFUL!
echo  Executable: dist\PayslipGenerator.exe
echo ============================================================
echo.
echo Double-click PayslipGenerator.exe to launch the GUI.
echo.
echo NOTE: The "venv" folder can be deleted to save space.
echo.
pause
