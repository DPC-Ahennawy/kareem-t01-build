@echo off
chcp 65001 >nul
title Kareem T-01 - Setup
cd /d "%~dp0"
echo ============================================================
echo   Kareem T-01 - Installing Python libraries
echo ============================================================
where py >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python is not installed. Install Python 3.10+ from python.org, then re-run.
  pause & exit /b 1
)
py -m pip install --upgrade pip
py -m pip install fastapi "uvicorn[standard]" python-multipart PyMuPDF pytesseract Pillow openpyxl numpy pywin32
echo.
echo Next: run fetch_ocr.bat once to get bundled OCR, then Run.bat to start the app.
pause
