@echo off
cd /d "%~dp0"
python app.py >> rotina.log 2>&1
