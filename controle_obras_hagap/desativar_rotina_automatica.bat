@echo off
set "ATALHO=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\HAGAP Controle de Obras.lnk"
if exist "%ATALHO%" del /q "%ATALHO%"
echo [CONFIRMADO] Inicio automatico removido.
pause
