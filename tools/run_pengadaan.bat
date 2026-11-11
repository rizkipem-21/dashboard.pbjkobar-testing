@echo off
cd /d "%~dp0.."
set ERR=0

echo ==================================================
echo MEMULAI PROSES UPDATE PENGADAAN...  [%date% %time%]
echo ==================================================

del /f /q .git\index.lock >nul 2>&1

:: Download hanya bila dijalankan sendiri (dari induk: parameter nodownload)
if /i not "%~1"=="nodownload" python scripts\download_json\download_json.py

echo [%time%] Generate data...
python scripts\pengadaan\generate_pengadaan.py
if errorlevel 1 set ERR=1

echo.
echo PENGADAAN SELESAI, kode: %ERR%  [%date% %time%]
if not defined DARI_INDUK mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses Pengadaan Selesai!"", 5, ""Update Selesai"", 4160)(window.close)")
exit /b %ERR%