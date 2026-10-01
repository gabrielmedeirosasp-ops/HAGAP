@echo off
cd /d "%~dp0"
title HAGAP - Controle de Obras
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://127.0.0.1:5055"
python app.py
pause
