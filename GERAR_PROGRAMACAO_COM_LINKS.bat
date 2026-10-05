@echo off
chcp 65001 >nul
cd /d "%~dp0"
title HAGAP - Programacao com Links Drive

echo ============================================================
echo HAGAP - GERAR PROGRAMACAO COM LINKS DO DRIVE
echo.
echo A PROGRAMACAO ORIGINAL NAO SERA ALTERADA.
echo Sera criado: PROGRAMAÇÃO UMU - LINKS DRIVE.xlsx
echo ============================================================
echo.

python -m pip install --quiet openpyxl playwright
python GERAR_PROGRAMACAO_COM_LINKS.py

echo.
pause
