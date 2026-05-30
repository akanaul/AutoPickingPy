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
set "PY_CMD=python"
%PY_CMD% --version >nul 2>&1
if errorlevel 1 (
    rem try the Python launcher
    py -3 --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python not found. Please install Python 3.8+ and add to PATH.
        pause
        exit /b 1
    ) else (
        set "PY_CMD=py -3"
    )
)
echo [OK] Python found

REM Create virtual environment
echo [2/3] Creating virtual environment (.venv)...
if not exist ".venv" (
    %PY_CMD% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
)
echo [OK] Virtual environment prepared.

REM Install dependencies
echo [3/3] Installing Python dependencies...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if exist "requirements.txt" (
    python -m pip install -r requirements.txt
) else (
    python -m pip install openpyxl odfpy pyinstaller
)

if errorlevel 1 (
    echo [WARNING] Some dependencies may not have installed properly.
) else (
    echo [OK] Dependencies installed successfully.
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
echo 2. Run the application:
echo    .\run.bat
echo.
pause
