@echo off  
setlocal enabledelayedexpansion
cd /d D:\rup-2026-inaproc

echo ========================= >> tools\log_rup.txt
echo START %date% %time% >> tools\log_rup.txt

:: 1. PROSES DATA (Otomatis Download, JSON, dan Excel)
echo RUN GENERATE RUP MULTI-TAHUN >> tools\log_rup.txt
python scripts\rup\generate_rup.py >> tools\log_rup.txt 2>&1

:: 2. FORMAT TANGGAL DAN UPDATE LAST-UPDATE (Full Bahasa Indonesia)
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do (
    set dd=%%a
    set mm=%%b
    set yyyy=%%c
)

set bulan=
if "%mm%"=="01" set bulan=Januari
if "%mm%"=="02" set bulan=Februari
if "%mm%"=="03" set bulan=Maret
if "%mm%"=="04" set bulan=April
if "%mm%"=="05" set bulan=Mei
if "%mm%"=="06" set bulan=Juni
if "%mm%"=="07" set bulan=Juli
if "%mm%"=="08" set bulan=Agustus
if "%mm%"=="09" set bulan=September
if "%mm%"=="10" set bulan=Oktober
if "%mm%"=="11" set bulan=November
if "%mm%"=="12" set bulan=Desember

for /f "tokens=1-2 delims=:." %%a in ("%time%") do (
    set hh=%%a
    set mn=%%b
)

:: Menghilangkan spasi pada jam jika 1 digit (misal " 8" menjadi "8")
set hh=!hh: =!

echo UPDATE LAST-UPDATE >> tools\log_rup.txt
echo !dd! !bulan! !yyyy! ^| !hh!.!mn! WIB > data\last-update-rup.txt

:: 3. PROSES UPLOAD KE GITHUB
echo GIT CONFIG >> tools\log_rup.txt
git config user.name "rizkipem-21"
git config user.email "rizki.pem@gmail.com"

echo GIT STATUS >> tools\log_rup.txt
git status >> tools\log_rup.txt 2>&1

:: FIX LOCK (Mencegah error Git nyangkut)
del /f /q .git\index.lock >nul 2>&1

echo GIT ADD >> tools\log_rup.txt
git add . >> tools\log_rup.txt 2>&1

echo GIT COMMIT >> tools\log_rup.txt
git commit -m "Auto update RUP %date% %time%" >> tools\log_rup.txt 2>&1

echo GIT PUSH >> tools\log_rup.txt
git push origin main >> tools\log_rup.txt 2>&1

echo PUSH STATUS: %ERRORLEVEL% >> tools\log_rup.txt

echo ========================= >> tools\log_rup.txt
echo SELESAI %date% %time% >> tools\log_rup.txt

:: 4. NOTIFIKASI SELESAI (POP-UP ALWAYS ON TOP SELAMA 5 DETIK)
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses update RUP Multi-Tahun telah SELESAI!"", 5, ""Update Selesai"", 4160)(window.close)")