@echo off
cd /d "%~dp0"
title HAGAP - Sincronizar Gmail
python gmail_sync.py --sync
pause
