@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title Kareem T-01 - Build installer (PyInstaller + Inno Setup)
cd /d "%~dp0\.."
set "APP=%CD%"
set "OCR=%APP%\ocr\tesseract"

echo ============================================================
echo   Kareem T-01 - One-click installer build (Route B)
echo   App root: %APP%
echo ============================================================

REM ---------- 0) prerequisites ----------
where py >nul 2>nul || (echo [ERROR] Python 3.10+ required on the BUILD machine. & goto :fail)
echo [1/5] Installing Python build + runtime libraries...
py -m pip install --upgrade pip >nul
py -m pip install pyinstaller fastapi "uvicorn[standard]" python-multipart PyMuPDF pytesseract Pillow openpyxl numpy pywin32 || goto :fail

REM ---------- 1) fetch REAL OCR into ocr\tesseract BEFORE packaging ----------
echo [2/5] Downloading bundled Tesseract engine + ara/eng/fra data...
powershell -NoProfile -ExecutionPolicy Bypass -File "%APP%\packaging\fetch_ocr_full.ps1" || goto :ocrfail

REM ---------- 1b) HARD GATE: refuse to build without real OCR files ----------
set "MISSING="
if not exist "%OCR%\tesseract.exe"               set "MISSING=!MISSING! tesseract.exe"
if not exist "%OCR%\tessdata\eng.traineddata"    set "MISSING=!MISSING! eng.traineddata"
if not exist "%OCR%\tessdata\ara.traineddata"    set "MISSING=!MISSING! ara.traineddata"
if not exist "%OCR%\tessdata\fra.traineddata"    set "MISSING=!MISSING! fra.traineddata"
if not "!MISSING!"=="" (
  echo [ERROR] OCR bundle incomplete. Missing:!MISSING!
  echo         Build aborted - will NOT produce an installer with placeholders.
  goto :fail
)
echo        OCR bundle verified (engine + DLLs + 3 languages).

REM ---------- 2) PyInstaller -> dist\KareemT01\KareemT01.exe ----------
echo [3/5] Building KareemT01.exe with PyInstaller...
rmdir /s /q build dist >nul 2>nul
py -m PyInstaller --noconfirm packaging\KareemT01.spec || goto :fail
if not exist "dist\KareemT01\KareemT01.exe" (echo [ERROR] PyInstaller did not produce KareemT01.exe & goto :fail)

REM ---------- 3) Inno Setup -> packaging\Output\Kareem_T01_Setup.exe ----------
echo [4/5] Compiling Kareem_T01_Setup.exe with Inno Setup...
set "ISCC="
for %%P in ("%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" "%ProgramFiles%\Inno Setup 6\ISCC.exe") do if exist "%%~P" set "ISCC=%%~P"
if "%ISCC%"=="" (echo [ERROR] Inno Setup 6 not found. Install from https://jrsoftware.org/isdl.php & goto :fail)
"%ISCC%" "packaging\KareemT01.iss" || goto :fail

echo [5/5] DONE.
echo ============================================================
echo   Installer: %APP%\packaging\Output\Kareem_T01_Setup.exe
echo ============================================================
pause
exit /b 0

:ocrfail
echo [ERROR] OCR download failed (no internet, or source URL changed).
echo         See packaging\fetch_ocr_full.ps1 to update the engine URL.
:fail
echo BUILD FAILED - no installer produced.
pause
exit /b 1
