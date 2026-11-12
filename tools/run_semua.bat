@echo off
cd /d "%~dp0.."
set LOCK=running.lock
set DL=scripts\download_json\download_json.py

set DARI_INDUK=1

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
call "%~dp0run_rup.bat" nodownload
set ERR_RUP=%ERRORLEVEL%

echo [%time%] TAHAP 3: Pengadaan...
call "%~dp0run_pengadaan.bat" nodownload
set ERR_PENG=%ERRORLEVEL%

:AKHIR
del /f /q "%LOCK%" >nul 2>&1
echo.
echo SELESAI  download=%ERR_DL%  rup=%ERR_RUP%  pengadaan=%ERR_PENG%  [%date% %time%]

if "%ERR_DL%"=="0" if "%ERR_RUP%"=="0" if "%ERR_PENG%"=="0" goto SUKSES

mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Update SELESAI DENGAN ERROR. Cek log di folder tools."", 120, ""Update Gagal"", 4112)(window.close)")
exit 1

:SUKSES
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Update Selesai!"", 5, ""Update Selesai"", 4160)(window.close)")
taskkill /f /im mshta.exe >nul 2>&1
exit 0