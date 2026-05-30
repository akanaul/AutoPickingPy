@echo off
REM Runner script for AutoPickingPy - ensure venv and launch application

setlocal enabledelayedexpansion

color 0B
cls

echo.
echo ================================================================================
echo                     AutoPickingPy - Launcher
echo ================================================================================
echo.

REM Locate Python interpreter (try 'py' launcher first for corporate setups)
set "PY_CMD=python"
%PY_CMD% --version >nul 2>&1
if errorlevel 1 (
    rem try the Python launcher
    py -3 --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python not found. Please install Python 3.8+ or ensure the
        echo        Python launcher ('py') is available.
        pause
        exit /b 1
    ) else (
        set "PY_CMD=py -3"
    )
)

echo [INFO] Python found (%PY_CMD%)

REM Create virtual environment if it doesn't exist
if not exist ".venv" (
    echo [INFO] Creating virtual environment...
    %PY_CMD% -m venv .venv
    call .venv\Scripts\activate.bat
    echo [INFO] Installing dependencies...
    %PY_CMD% -m pip install -q pyinstaller openpyxl
) else (
    call .venv\Scripts\activate.bat
)

echo.
echo [INFO] Running AutoPickingPy...
%PY_CMD% gerar_solicitacao_pickings.py %*

endlocal
