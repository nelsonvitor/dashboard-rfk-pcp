@echo off
setlocal

REM ===================================================================
REM  IMPORTANTE: troque a linha do SET PYTHON_EXE abaixo pelo caminho
REM  real do seu python.exe. Descubra rodando "where python" no cmd.
REM ===================================================================
set "PYTHON_EXE=C:\Users\jaildo.junior\AppData\Local\Python\pythoncore-3.14-64\python.exe"
set "PASTA_PROJETO=C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK"
set "SCRIPT=%PASTA_PROJETO%\atualizar_dashboard.py"
set "LOG_BAT=%PASTA_PROJETO%\logs\bat_%date:~-4,4%-%date:~-7,2%-%date:~-10,2%.log"

REM Garante que a pasta de logs existe antes de escrever nela
if not exist "%PASTA_PROJETO%\logs" mkdir "%PASTA_PROJETO%\logs"

echo ===== Execucao do .bat iniciada em %date% %time% ===== >> "%LOG_BAT%"

REM Confere se o python.exe configurado realmente existe
if not exist "%PYTHON_EXE%" (
    echo ERRO: python.exe nao encontrado em "%PYTHON_EXE%". Rode "where python" e corrija o caminho neste .bat. >> "%LOG_BAT%"
    exit /b 1
)

cd /d "%PASTA_PROJETO%"

echo Rodando: "%PYTHON_EXE%" "%SCRIPT%" >> "%LOG_BAT%"
"%PYTHON_EXE%" "%SCRIPT%" >> "%LOG_BAT%" 2>&1

echo Codigo de saida do Python: %errorlevel% >> "%LOG_BAT%"
echo ===== Execucao do .bat finalizada em %date% %time% ===== >> "%LOG_BAT%"

endlocal
