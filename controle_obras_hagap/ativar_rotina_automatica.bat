@echo off
cd /d "%~dp0"
set "ATALHO=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\HAGAP Controle de Obras.lnk"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut('%ATALHO%'); $s.TargetPath='%~dp0iniciar_oculto.vbs'; $s.WorkingDirectory='%~dp0'; $s.Save()"
if errorlevel 1 (
  echo [ERRO] Nao foi possivel ativar a inicializacao automatica.
  pause
  exit /b 1
)
echo.
echo [CONFIRMADO] Rotina automatica ativada.
echo O Controle de Obras iniciara junto com o Windows e sincronizara o Gmail no intervalo configurado.
pause
