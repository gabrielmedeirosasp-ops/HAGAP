@echo off
set "ARQ=%USERPROFILE%\Desktop\HAGAP - Controle de Obras.url"
(
  echo [InternetShortcut]
  echo URL=http://127.0.0.1:5055
) > "%ARQ%"
echo [CONFIRMADO] Atalho criado na Area de Trabalho.
pause
