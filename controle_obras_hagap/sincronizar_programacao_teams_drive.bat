@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo HAGAP - PROGRAMACAO ^> TEAMS ^> GOOGLE DRIVE
echo Fonte Teams: 3- Obras para Execucao
echo Destino Drive: projetos
echo ============================================================
echo.
python -m pip install --quiet openpyxl playwright google-api-python-client google-auth
python sincronizar_programacao_teams_drive.py
echo.
pause
