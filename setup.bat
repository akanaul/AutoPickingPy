@echo off
REM Quick setup script for AutoPickingPy

setlocal enabledelayedexpansion

color 0B
cls

echo.
echo ================================================================================
echo                     AutoPickingPy - Quick Setup Wizard
echo ================================================================================
echo.

REM Check Python
echo [1/3] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.8+ and add to PATH
    pause
    exit /b 1
)
echo [OK] Python installed

REM Check Git
echo [2/3] Checking Git installation...
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git not found. Please install Git
    pause
    exit /b 1
)
echo [OK] Git installed

REM Install dependencies
echo [3/3] Installing Python dependencies...
pip install -q pyinstaller openpyxl requests

if errorlevel 1 (
    echo [WARNING] Some dependencies may not have installed properly
) else (
    echo [OK] Dependencies installed
)

echo.
echo ================================================================================
echo                           Setup Complete!
echo ================================================================================
echo.
echo Next steps:
echo.
echo 1. Build the executable:
echo    .\build.bat
echo.
echo 2. Test the executable:
echo    .\dist\AutoPickingPy.exe
echo.
pause
