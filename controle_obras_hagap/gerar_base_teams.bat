@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo HAGAP - GERAR BASE LIMPA DO TEAMS
echo Fonte: C:\HAGAP\TEAMS_SYNC
echo Site/db.json antigo: NAO UTILIZADO
echo ============================================================
echo.
python teams_base_clean.py
echo.
pause
