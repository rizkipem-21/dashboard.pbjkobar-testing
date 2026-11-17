@echo off
setlocal
cd /d "%~dp0..\.."
set LOCK=running.lock
set DL=scripts\induk\download_json.py

set DARI_INDUK=1
set ERR_GIT=0

:: Kunci anti-ganda (dianggap basi setelah 2 jam)
if exist "%LOCK%" (
  powershell -NoProfile -Command "if (((Get-Date) - (Get-Item '%LOCK%').LastWriteTime).TotalHours -lt 2) { exit 1 } else { exit 0 }"
  if errorlevel 1 (
    echo Proses sebelumnya masih berjalan, dibatalkan.
    exit /b 0
  )
  echo Kunci lama dianggap basi, dilanjutkan.
)
echo %date% %time% > "%LOCK%"
del /f /q .git\index.lock >nul 2>&1

echo ==================================================
echo UPDATE SEMUA  [%date% %time%]
echo ==================================================

:: 1. Download terpusat
echo [%time%] TAHAP 1: Download JSON...
python %DL%
set ERR_DL=%ERRORLEVEL%
echo [%time%] TAHAP 1 selesai, kode: %ERR_DL%

if not "%ERR_DL%"=="0" (
  echo Download gagal total atau crash, generate dilewati.
  set ERR_RUP=skip
  set ERR_PENG=skip
  goto AKHIR
)

:: 2. Generate RUP lalu Pengadaan (berurutan)
echo [%time%] TAHAP 2: RUP...
call "%~dp0..\rup\run_rup.bat" nodownload
set ERR_RUP=%ERRORLEVEL%

echo [%time%] TAHAP 3: Pengadaan...
call "%~dp0..\pengadaan\run_pengadaan.bat" nodownload
set ERR_PENG=%ERRORLEVEL%

echo [%time%] TAHAP 4: Push GitHub...
git config user.name "rizkipem-21"
git config user.email "rizki.pem@gmail.com"
git add .
git commit -m "Auto update %date% %time%" >nul 2>&1
if not exist "logs\induk" mkdir "logs\induk"
git push > logs\induk\log_git.txt 2>&1
set ERR_GIT=%ERRORLEVEL%
if not "%ERR_GIT%"=="0" python -c "import config_rahasia as c,requests;e=open('logs/induk/log_git.txt',encoding='utf-8',errors='ignore').read()[-800:];requests.post('https://api.telegram.org/bot'+c.BOT_TOKEN+'/sendMessage',data={'chat_id':c.CHAT_ID,'text':'PUSH GITHUB GAGAL\n\n'+e},timeout=30)"

:AKHIR
del /f /q "%LOCK%" >nul 2>&1
echo.
echo SELESAI  download=%ERR_DL%  rup=%ERR_RUP%  pengadaan=%ERR_PENG%  git=%ERR_GIT%  [%date% %time%]

if "%ERR_DL%"=="0" if "%ERR_RUP%"=="0" if "%ERR_PENG%"=="0" if "%ERR_GIT%"=="0" goto SUKSES

mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Update SELESAI DENGAN ERROR. Cek log di folder logs."", 120, ""Update Gagal"", 4112)(window.close)")
exit /b 1

:SUKSES
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Update Selesai!"", 5, ""Update Selesai"", 4160)(window.close)")
taskkill /f /im mshta.exe >nul 2>&1
exit /b 0