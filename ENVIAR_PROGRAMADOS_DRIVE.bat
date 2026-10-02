@echo off
chcp 65001 >nul
cd /d "%~dp0"
title HAGAP - Enviar Projetos Programados para o Drive

echo ============================================================
echo HAGAP - ENVIAR PROGRAMADOS PARA O DRIVE
echo Fonte: Teams - 3- Obras para Execucao
echo Destino: pasta projetos do Google Drive
echo ============================================================
echo.
echo Este programa e independente do Controle de Obras.
echo.

python -m pip install --quiet openpyxl playwright
python enviar_programados_drive.py

echo.
pause
