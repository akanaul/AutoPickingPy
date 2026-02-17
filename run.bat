@echo off
setlocal

set VENV_DIR=%~dp0.venv

if not exist "%VENV_DIR%" (
  python -m venv "%VENV_DIR%"
)

call "%VENV_DIR%\Scripts\activate.bat"

python -m pip install --upgrade pip
python -m pip install openpyxl odfpy

python "%~dp0gerar_solicitacao_pickings.py"

endlocal
