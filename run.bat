@echo off
REM Script de Execucao para AutoPickingPy - Menu Interativo Guiado (Producao)
REM Layout premium, 100% resiliente, sem blocos parentetizados para total compatibilidade no Windows.
REM Calcula a data-alvo automaticamente de forma dinamica com base no horario de envio (D+0 ou D+1).
REM Incorpora busca profunda de Python fora do PATH em caminhos corporativos e de usuario comuns.

setlocal enabledelayedexpansion

color 0B
title AutoPickingPy - Painel de Controle Interativo

REM ================================================================================
REM VERIFICACAO GLOBAL DO PYTHON (COM BUSCA PROFUNDA EM PASTAS PADRAO SE FORA DO PATH)
REM ================================================================================
set "PY_CMD=python"
python --version >nul 2>&1
if not errorlevel 1 goto PYTHON_OK

set "PY_CMD=py -3"
py -3 --version >nul 2>&1
if not errorlevel 1 goto PYTHON_OK

REM Se nao encontrado no PATH, varrer locais comuns de instalacao do Windows
echo [INFO] Python nao encontrado no PATH. Procurando em pastas de instalacao padrao...

REM 1. Varrer AppData do usuario (instalacao de usuario padrao)
for /d %%d in ("%LOCALAPPDATA%\Programs\Python\Python*") do (
    if exist "%%d\python.exe" (
        set "PY_CMD=%%d\python.exe"
        goto PYTHON_OK
    )
)

REM 2. Varrer Program Files (instalacao global/corporativa)
for /d %%d in ("%ProgramFiles%\Python*") do (
    if exist "%%d\python.exe" (
        set "PY_CMD=%%d\python.exe"
        goto PYTHON_OK
    )
)

REM 3. Varrer Program Files 32-bit
for /d %%d in ("%ProgramFiles(x86)%\Python*") do (
    if exist "%%d\python.exe" (
        set "PY_CMD=%%d\python.exe"
        goto PYTHON_OK
    )
)

REM 4. Caminhos de Anaconda / Miniconda
if exist "%USERPROFILE%\anaconda3\python.exe" (
    set "PY_CMD=%USERPROFILE%\anaconda3\python.exe"
    goto PYTHON_OK
)
if exist "%USERPROFILE%\miniconda3\python.exe" (
    set "PY_CMD=%USERPROFILE%\miniconda3\python.exe"
    goto PYTHON_OK
)

echo.
echo ================================================================================
echo  [ERRO] Python nao encontrado no sistema (nem no PATH e nem nos locais comuns).
echo  Por favor, instale o Python 3.8+ para que a automacao possa funcionar.
echo ================================================================================
echo.
pause
exit /b 1

:PYTHON_OK

:MAIN_MENU
cls
echo.
echo  ================================================================================
echo                      AUTOPICKINGPY - PAINEL INTERATIVO
echo  ================================================================================
echo.
echo   [1] Iniciar Automacao Completa (Passo a Passo Guiado)
echo   [2] Executar Apenas Extracao de Dados (Sem Criar E-mail)
echo   [3] Sair
echo.
echo  ================================================================================
set /p "OPCAO=Escolha uma opcao (1-3): "

if "!OPCAO!"=="1" goto AUTOMATION_FLOW
if "!OPCAO!"=="2" goto EXTRACTION_ONLY
if "!OPCAO!"=="3" goto EXIT_PROG
goto MAIN_MENU


:AUTOMATION_FLOW
cls
echo.
echo  ================================================================================
echo   [ETAPA 1/4] Verificacao e Preparacao do Ambiente Virtual (venv)
echo  ================================================================================
echo.

if exist ".venv" goto VENV_OK
echo [INFO] Criando Ambiente Virtual (.venv) para isolamento de dependencias...
!PY_CMD! -m venv .venv
if not exist ".venv" goto VENV_FAILED
:VENV_OK

echo [INFO] Ativando Ambiente Virtual...
call .venv\Scripts\activate.bat

echo [INFO] Atualizando gerenciador de pacotes pip...
python -m pip install --upgrade pip -q

echo [INFO] Verificando e instalando bibliotecas necessarias...
if not exist requirements.txt goto NO_REQUIREMENTS
python -m pip install -r requirements.txt -q
goto REQ_DONE

:NO_REQUIREMENTS
python -m pip install openpyxl pywin32 pyinstaller -q

:REQ_DONE
if errorlevel 1 goto DEPS_ERROR
echo [OK] Ambiente Virtual e dependencias validados com sucesso!
timeout /t 2 >nul
goto DEPS_OK

:DEPS_ERROR
echo.
echo [AVISO] Houve um problema na instalacao automatica de algumas dependencias.
echo Tentaremos prosseguir mesmo assim...
timeout /t 3 >nul

:DEPS_OK
goto STAGE_2

:VENV_FAILED
echo.
echo [ERRO] Falha critica ao tentar criar o ambiente virtual (.venv).
echo Verifique se voce possui permissoes de escrita nesta pasta.
pause
goto MAIN_MENU


:STAGE_2
cls
echo.
echo  ================================================================================
echo   [ETAPA 2/4] Verificacao dos Dados de E-mail (config_email.txt)
echo  ================================================================================
echo.
echo [INFO] Validando informacoes de destinatarios e copias...

python gerar_solicitacao_pickings.py --check-config
set "EXIT_CODE=%errorlevel%"

if "!EXIT_CODE!"=="2" goto CONFIG_SUSPENDED
if not "!EXIT_CODE!"=="0" goto CONFIG_FAILED

echo.
echo [OK] Configuracoes validadas com sucesso!
echo Prosseguindo para a proxima etapa...
timeout /t 2 >nul
goto STAGE_3

:CONFIG_SUSPENDED
echo.
echo ================================================================================
echo  [INFO] Automacao suspensa para que voce ajuste o arquivo 'config_email.txt'.
echo  Por favor, abra o arquivo na pasta do projeto, insira os e-mails reais e salve.
echo ================================================================================
pause
goto MAIN_MENU

:CONFIG_FAILED
echo.
echo [ERRO] Falha ao tentar verificar o arquivo de configuracoes (Codigo: !EXIT_CODE!).
pause
goto MAIN_MENU


:STAGE_3
cls
echo.
echo  ================================================================================
echo   [ETAPA 3/4] Extracao, Filtragem e Transformacao de Dados
echo  ================================================================================
echo.
echo [INFO] Lendo a planilha "Pasta de Viagens Itu.xlsx"...
echo [INFO] Filtrando dados de pickings ITU (Origem: FABRICA ITU)...
echo Aguarde a conclusao do processamento...
echo.

python gerar_solicitacao_pickings.py --extract-only
set "EXIT_CODE=%errorlevel%"

if not "!EXIT_CODE!"=="0" goto EXTRACT_FAILED

echo.
echo  ================================================================================
echo                       ALERTA DE SEGURANCA CRITICO (PROXIMA ETAPA)
echo  ================================================================================
echo   [RECOMENDACAO IMPORTANTE]:
echo   Na proxima etapa, o Outlook sera acionado para formatacao do e-mail.
echo   E EXTREMAMENTE RECOMENDADO que voce espere a conclusao da formatacao
echo   SEM MOVIMENTAR O MOUSE ou APERTAR NENHUMA TECLA, para garantir a correta
echo   insercao da tabela e evitar perda de foco.
echo  ================================================================================
echo.
set /p "PROCEED=Pressione a tecla [ENTER] para prosseguir com a automacao de e-mail..."
goto STAGE_4

:EXTRACT_FAILED
echo.
echo [ERRO] Falha durante o processo de extracao de dados (Codigo: !EXIT_CODE!).
echo Verifique se a planilha 'Pasta de Viagens Itu.xlsx' esta na raiz e fechada.
pause
goto MAIN_MENU


:STAGE_4
cls
echo.
echo  ================================================================================
echo   [ETAPA 4/4] Geracao Automatica e Formatacao do E-mail
echo  ================================================================================
echo.
echo [INFO] Disparando rascunho de e-mail e executando colagem inteligente...
echo [INFO] Aguarde ate que a automacao termine de digitar e cole os dados...

python gerar_solicitacao_pickings.py --email-only
set "EXIT_CODE=%errorlevel%"

if not "!EXIT_CODE!"=="0" goto EMAIL_FAILED
goto SUCCESS_SCREEN

:EMAIL_FAILED
echo.
echo [ERRO] Falha ao tentar gerar ou colar a tabela no e-mail (Codigo: !EXIT_CODE!).
pause
goto MAIN_MENU


:EXTRACTION_ONLY
cls
echo.
echo  ================================================================================
echo   Executando Apenas Extracao de Dados (Sem Criacao de E-mail)
echo  ================================================================================
echo.

if exist ".venv" goto VENV_OPT2_OK
echo [INFO] Criando Ambiente Virtual (.venv)...
!PY_CMD! -m venv .venv
:VENV_OPT2_OK

call .venv\Scripts\activate.bat

echo [INFO] Executando rotina de extracao e salvamento de arquivos...
python gerar_solicitacao_pickings.py --extract-only
set "EXIT_CODE=%errorlevel%"

if not "!EXIT_CODE!"=="0" goto EXTRACT_OPT2_ERR
echo.
echo ================================================================================
echo  [SUCESSO] Extracao executada com sucesso!
echo  Arquivos formatados salvos na raiz do projeto.
echo ================================================================================
goto EXTRACT_OPT2_DONE

:EXTRACT_OPT2_ERR
echo.
echo [ERRO] Falha durante a extracao de dados (Codigo: !EXIT_CODE!).

:EXTRACT_OPT2_DONE
pause
goto MAIN_MENU


:SUCCESS_SCREEN
cls
echo.
echo  ================================================================================
echo                        AUTOMACAO CONCLUIDA COM SUCESSO!
echo  ================================================================================
echo.
echo   Os arquivos Excel e CSV foram gerados e salvos com sucesso na raiz.
echo   O historico de execucoes anteriores foi atualizado na pasta 'historico_gerado/'.
echo   O e-mail foi gerado e a tabela formatada foi colada perfeitamente!
echo.
echo   [ARQUIVOS ATUALIZADOS]:
echo     - Pasta de Viagens Itu.xlsx (Arquivo original de entrada)
echo     - solicitacao de pickings [Data].xlsx (Planilha formatada final)
echo     - solicitacao de pickings [Data].csv (CSV final delimitado por ";")
echo.
echo  ================================================================================
echo   [CONCLUIDO] A execucao diaria unica foi realizada.
echo   Esta janela de comando sera encerrada automaticamente em 5 segundos...
echo  ================================================================================
timeout /t 5 >nul
goto EXIT_PROG


:EXIT_PROG
cls
echo.
echo Obrigado por utilizar o AutoPickingPy!
echo Finalizando sessao...
echo.
timeout /t 1 >nul
endlocal
exit
