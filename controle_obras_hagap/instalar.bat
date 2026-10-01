@echo off
cd /d "%~dp0"
title HAGAP - Instalacao Controle de Obras
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo [ERRO] Falha ao instalar dependencias.
  pause
  exit /b 1
)
echo.
echo [OK] Dependencias instaladas.
echo Proximo passo: conectar_gmail.bat
pause
