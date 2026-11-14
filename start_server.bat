@echo off
setlocal
title AI-autoDoc - Generator Bukti Fisik Jabar Beraksi (1-Click System)
cd /d "%~dp0"

set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"

echo ====================================================================
echo   MENYIAPKAN SISTEM AI-AUTODOC (1-KLIK)
echo ====================================================================

set "PY_CMD="

REM 1. Cek python di PATH
where python >nul 2>&1
if %errorlevel% equ 0 set "PY_CMD=python"

REM 2. Cek path instalasi standar jika belum ketemu
if not defined PY_CMD if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
if not defined PY_CMD if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PY_CMD if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY_CMD if exist "C:\Python314\python.exe" set "PY_CMD=C:\Python314\python.exe"
if not defined PY_CMD if exist "C:\Python313\python.exe" set "PY_CMD=C:\Python313\python.exe"
if not defined PY_CMD if exist "C:\Python312\python.exe" set "PY_CMD=C:\Python312\python.exe"

if not defined PY_CMD goto :no_python

echo [OK] Python terdeteksi: %PY_CMD%

REM 3. Pastikan dependensi terpasang
"%PY_CMD%" -c "import flask, docx, requests" >nul 2>&1
if %errorlevel% neq 0 (
    echo [i] Memasang paket yang dibutuhkan: flask, python-docx, requests
    "%PY_CMD%" -m pip install flask flask-cors python-docx requests
)

REM 4. Jalankan launcher.py
echo [i] Menjalankan launcher server dan Cloudflare tunnel...
"%PY_CMD%" launcher.py
goto :end

:no_python
echo [ERROR] Python tidak ditemukan di sistem ini!
echo Silakan unduh dan pasang Python dari https://www.python.org/downloads/
echo Pastikan centang Add Python to PATH saat instalasi.
pause
exit /b 1

:end
echo.
echo [i] Sesi server telah selesai.
pause