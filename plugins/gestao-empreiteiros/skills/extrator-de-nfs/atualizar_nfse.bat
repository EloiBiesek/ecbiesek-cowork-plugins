@echo off
chcp 65001 >nul 2>&1
title Atualizar NFSe
set PYTHONIOENCODING=utf-8

echo.
echo ============================================================
echo   ATUALIZAR NFSe - Controle de Notas Fiscais
echo   Clique aqui e aguarde o processamento...
echo ============================================================
echo.

cd /d "%~dp0"

REM Verificar se Python esta disponivel
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado!
    echo Instale Python em https://python.org
    echo.
    pause
    exit /b 1
)

REM Instalar dependencias se necessario
python -c "import pdfplumber, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias...
    pip install pdfplumber openpyxl >nul 2>&1
)

REM Determinar pasta da obra (2 niveis acima da skill: .skills/nfse-extractor -> obra)
set "OBRA_DIR=%~dp0..\.."
for %%i in ("%OBRA_DIR%") do set "OBRA_DIR=%%~fi"

REM Executar pipeline com a pasta da obra
python scripts/atualizar_nfse.py --obra-dir "%OBRA_DIR%" %*

echo.
echo Pressione qualquer tecla para fechar...
pause >nul
