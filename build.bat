@echo off
REM Build script for AutoPickingPy
REM This script compiles the Python application into a Windows .exe file

setlocal enabledelayedexpansion

echo.
echo ================================================================================
echo                        AutoPickingPy - Build Script
echo ================================================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo Please install Python 3.8+ and add it to your PATH
    pause
    exit /b 1
)

echo [INFO] Python found
python --version

REM Check if venv exists
if not exist ".venv" (
    echo [INFO] Creating virtual environment...
    python -m venv .venv
)

REM Activate venv
call .venv\Scripts\activate.bat

echo [INFO] Installing dependencies...
pip install pyinstaller openpyxl 

REM Create build directory
if not exist "build" mkdir build
if not exist "dist" mkdir dist

echo [INFO] Building executable...
REM PyInstaller command with icon and metadata
pyinstaller.exe ^
    --name="AutoPickingPy" ^
    --onefile ^
    --windowed ^
    --add-data "LICENSE:." ^
    --add-data "Pasta de Viagens Itu.xlsx:." ^
    --hidden-import=openpyxl ^
    --hidden-import=odf ^
    --version-file=version_info.txt ^
    --output=./dist ^
    gerar_solicitacao_pickings.py

echo.
echo ================================================================================
echo [SUCCESS] Build complete!
echo ================================================================================
echo.
echo Executable location: dist\AutoPickingPy.exe
echo.
echo [INFO] Testing the build...
dist\AutoPickingPy.exe

pause
