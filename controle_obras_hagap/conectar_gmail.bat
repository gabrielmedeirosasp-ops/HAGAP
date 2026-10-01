@echo off
cd /d "%~dp0"
title HAGAP - Conectar Gmail
python gmail_sync.py --auth-only
pause
