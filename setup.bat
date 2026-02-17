@echo off
REM Quick setup script for AutoPickingPy GitHub License System

setlocal enabledelayedexpansion

color 0B
cls

echo.
echo ================================================================================
echo                     AutoPickingPy - Quick Setup Wizard
echo ================================================================================
echo.

REM Check Python
echo [1/5] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.8+ and add to PATH
    pause
    exit /b 1
)
echo [OK] Python installed

REM Check Git
echo [2/5] Checking Git installation...
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git not found. Please install Git
    pause
    exit /b 1
)
echo [OK] Git installed

REM Get GitHub username
echo.
echo [3/5] GitHub Configuration
echo.
set /p GITHUB_USER="Enter your GitHub username: "

if "!GITHUB_USER!"=="" (
    echo [ERROR] GitHub username required
    pause
    exit /b 1
)

REM Update license_manager.py
echo [4/5] Configuring license manager...

REM Use findstr to locate and prepare replacement
findstr /n "GITHUB_REPO = " license_manager.py > nul
if errorlevel 1 (
    echo [ERROR] Could not find GITHUB_REPO in license_manager.py
    pause
    exit /b 1
)

REM Create temp file with replacement (PowerShell is easier)
powershell -Command ^
  "(Get-Content 'license_manager.py') -replace 'GITHUB_REPO = \"[^\"]+\"', ('GITHUB_REPO = \"' + '!GITHUB_USER!' + '/autopickingpy\"') | Set-Content 'license_manager.py'"

echo [OK] License manager configured for user: !GITHUB_USER!

REM Install dependencies
echo [5/5] Installing Python dependencies...
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
echo 1. Verify GitHub repository exists: https://github.com/!GITHUB_USER!/autopickingpy
echo.
echo 2. Commit changes to GitHub:
echo    git add license_manager.py
echo    git commit -m "Configure GitHub license system"
echo    git push origin main
echo.
echo 3. Build the executable:
echo    .\build.bat
echo.
echo 4. Test the executable:
echo    .\dist\AutoPickingPy.exe
echo.
echo 5. Read GITHUB_LICENSE_SYSTEM.txt for detailed documentation
echo.
echo For more info, see: GITHUB_LICENSE_SYSTEM.txt
echo.
pause
