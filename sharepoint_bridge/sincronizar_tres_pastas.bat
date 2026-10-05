@echo off
cd /d "%~dp0"
python sync_sharepoint_multiplas.py --pastas obras rmds
if errorlevel 1 goto erro
python sync_bmd2026.py
if errorlevel 1 goto erro
echo.
echo SINCRONIZACAO CONCLUIDA.
pause
exit /b 0
:erro
echo.
echo ERRO NA SINCRONIZACAO.
pause
exit /b 1
