@echo off
chcp 65001 > nul
title Atualizar Dashboards de Inventário
cd /d "%~dp0"

echo ============================================================
echo   ATUALIZAR DASHBOARDS DE INVENTÁRIO
echo ============================================================
echo.
echo Verificando Python...

python --version > nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado no PATH.
    echo Instale Python em https://www.python.org/downloads/
    echo Marque a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

echo Verificando dependencias...
python -c "import pandas, openpyxl, numpy" > nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias necessarias...
    pip install pandas openpyxl numpy --quiet
    if errorlevel 1 (
        echo ERRO ao instalar dependencias. Execute manualmente:
        echo   pip install pandas openpyxl numpy
        pause
        exit /b 1
    )
    echo Dependencias instaladas com sucesso!
    echo.
)

echo Iniciando analise e geracao dos dashboards...
echo.
python atualizar_dashboards.py

if errorlevel 1 (
    echo.
    echo ERRO durante a execucao. Verifique as mensagens acima.
    pause
)
