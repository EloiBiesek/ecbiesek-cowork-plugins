@echo off
chcp 65001 >nul 2>&1
title Atualizar SEFIP
set PYTHONIOENCODING=utf-8

echo.
echo ============================================================
echo   ATUALIZAR SEFIP
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
python -c "import pdfplumber, openpyxl, fitz, rapidocr_onnxruntime" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias...
    pip install pdfplumber openpyxl pymupdf rapidocr-onnxruntime >nul 2>&1
)

REM Determinar pasta da obra (2 niveis acima da skill: .skills/sefip-extractor -> obra)
set "OBRA_DIR=%~dp0..\.."
for %%i in ("%OBRA_DIR%") do set "OBRA_DIR=%%~fi"

REM Configurar obra na primeira execucao (se .sefip-state/obra.json nao existe)
if not exist "%OBRA_DIR%\.sefip-state\obra.json" (
    if not exist "obra.json" (
        echo Primeira execucao - configurando obra...
        echo.
        python scripts/configurar_obra.py --obra-dir "%OBRA_DIR%"
        if errorlevel 1 (
            echo.
            echo Configuracao cancelada.
            pause
            exit /b 1
        )
        echo.
    )
)

REM Executar pipeline com a pasta da obra
python scripts/atualizar_sefip.py --obra-dir "%OBRA_DIR%" %*

echo.
echo Pressione qualquer tecla para fechar...
pause >nul
