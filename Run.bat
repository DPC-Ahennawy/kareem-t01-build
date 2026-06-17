@echo off
chcp 65001 >nul
title Kareem T-01 - EDECS Contract Processor
cd /d "%~dp0web\backend"
echo Starting Kareem T-01 (local, offline) ...
echo Open http://127.0.0.1:8000/  in your browser.
start "" http://127.0.0.1:8000/
py -m uvicorn api:app --host 127.0.0.1 --port 8000
if errorlevel 1 pause
