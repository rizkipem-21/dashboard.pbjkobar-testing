:: ini file run_pengadaan.bat   ::

@echo off
cd /d "%~dp0..\.."
set ERR=0
set DL=0

echo ==================================================
echo MEMULAI PROSES UPDATE PENGADAAN...  [%date% %time%]
echo ==================================================

del /f /q .git\index.lock >nul 2>&1

:: Download hanya bila dijalankan sendiri (dari induk: parameter nodownload)
if /i not "%~1"=="nodownload" (
    python scripts\induk\download_json.py
    if errorlevel 1 set DL=2
)

if "%DL%"=="2" (
    echo Semua unduhan gagal, generate dilewati.
    set ERR=1
    goto selesai
)

echo [%time%] Generate Pengadaan...
python scripts\pengadaan\generate_pengadaan.py
if errorlevel 1 set ERR=1
if "%ERR%"=="1" goto selesai

echo [%time%] Generate Tepra...
python scripts\pengadaan\generate_tepra.py
if errorlevel 1 set ERR=1

echo [%time%] Arsip lokal Pengadaan...
python scripts\pengadaan\arsip_pengadaan.py

if "%ERR%"=="0" (
    echo [%time%] Arsip lokal Tepra...
    python scripts\pengadaan\arsip_tepra.py
)

:selesai

echo.
echo PENGADAAN SELESAI, kode: %ERR%  [%date% %time%]
if not defined DARI_INDUK mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses Pengadaan Selesai!"", 5, ""Update Selesai"", 4160)(window.close)")
exit /b %ERR%