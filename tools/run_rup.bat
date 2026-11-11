@echo off
cd /d "%~dp0.."
set ERR=0

for /f %%i in ('powershell -NoProfile -Command "[DateTimeOffset]::Now.ToUnixTimeSeconds()"') do set RUP_START=%%i

echo ==================================================
echo MEMULAI PROSES UPDATE RUP...  [%date% %time%]
echo ==================================================

del /f /q .git\index.lock >nul 2>&1

:: Download hanya bila dijalankan sendiri (dari induk: parameter nodownload)
if /i not "%~1"=="nodownload" python scripts\download_json\download_json.py

python scripts\rup\detail_rup.py
if errorlevel 1 set ERR=1

python scripts\rup\konsolidasi_rup.py
if errorlevel 1 set ERR=1

python scripts\rup\generate_rup.py
if errorlevel 1 set ERR=1

echo.
echo RUP SELESAI, kode: %ERR%
if not defined DARI_INDUK mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses RUP Selesai!"", 5, ""Update Selesai"", 4160)(window.close)")
exit /b %ERR%