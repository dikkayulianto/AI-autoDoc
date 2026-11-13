@echo off
title AI-autoDoc - Generator Bukti Fisik Jabar Beraksi (Local & Remote Engine)
chcp 65001 >nul
cd /d "%~dp0"

echo ====================================================================
echo   MEMERIKSA LINGKUNGAN PYTHON & DEPENDENSI...
echo ====================================================================

REM Cek executable Python
set "PY_CMD=python"
python --version >nul 2>&1
if %errorlevel% neq 0 (
    if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
        set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
    ) else if exist "C:\Python314\python.exe" (
        set "PY_CMD=C:\Python314\python.exe"
    ) else if exist "C:\Python313\python.exe" (
        set "PY_CMD=C:\Python313\python.exe"
    ) else if exist "C:\Python312\python.exe" (
        set "PY_CMD=C:\Python312\python.exe"
    ) else (
        echo [ERROR] Python tidak ditemukan di sistem ini!
        echo Silakan instal Python 3.10+ dari https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

echo Python terdeteksi: %PY_CMD%

REM Jalankan launcher.py
"%PY_CMD%" launcher.py

pause
