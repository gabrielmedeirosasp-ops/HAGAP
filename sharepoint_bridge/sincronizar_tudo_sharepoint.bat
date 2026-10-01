@echo off
cd /d "%~dp0"

echo ============================================================
echo HAGAP - SINCRONIZACAO COMPLETA SHAREPOINT
echo 1. AES Emitida
echo 2. Obras para Execucao
echo 3. RMDs
echo 4. BMD 2026
echo ============================================================
echo.

python sync_sharepoint.py --modo sync
if errorlevel 1 goto erro

python sync_sharepoint_multiplas.py --pastas obras rmds
if errorlevel 1 goto erro

python sync_bmd2026.py
if errorlevel 1 goto erro

echo.
echo [CONFIRMADO] SINCRONIZACAO COMPLETA CONCLUIDA.
pause
exit /b 0

:erro
echo.
echo [ERRO] SINCRONIZACAO INTERROMPIDA. NENHUMA ETAPA SEGUINTE FOI EXECUTADA.
pause
exit /b 1
