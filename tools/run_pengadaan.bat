@echo off
setlocal enabledelayedexpansion
cd /d D:\rup-2026-inaproc

echo ========================= >> tools\log_pengadaan.txt
echo START %date% %time% >> tools\log_pengadaan.txt

echo DOWNLOAD DATA >> tools\log_pengadaan.txt
powershell -ExecutionPolicy Bypass -File scripts\pengadaan\download_pengadaan.ps1 >> tools\log_pengadaan.txt 2>&1

echo GENERATE REKAP >> tools\log_pengadaan.txt
python scripts\pengadaan\generate_pengadaan.py >> tools\log_pengadaan.txt 2>&1

:: -------------------------------------------------------
:: COPY EXCEL TERBARU KE data\master_pengadaan.xlsx
:: -------------------------------------------------------
echo COPY EXCEL TERBARU >> tools\log_pengadaan.txt

set "src_dir=D:\rup-2026-inaproc\output\pengadaan"
set "dst=D:\rup-2026-inaproc\data\master_pengadaan.xlsx"
set "file_terbaru="

:: Cari file excel terbaru berdasarkan urutan nama (nama sudah mengandung tanggal YYYY-MM-DD)
for /f "delims=" %%f in ('dir /b /o-n /a-d "!src_dir!\Paket Pengadaan Tahun *.xlsx" 2^>nul') do (
    if not defined file_terbaru set "file_terbaru=%%f"
)

if not defined file_terbaru (
    echo ERROR: Tidak ada file Excel ditemukan di !src_dir! >> tools\log_pengadaan.txt
) else (
    echo File terbaru: !file_terbaru! >> tools\log_pengadaan.txt
    copy /y "!src_dir!\!file_terbaru!" "!dst!" >> tools\log_pengadaan.txt 2>&1
    echo Copy selesai ke: !dst! >> tools\log_pengadaan.txt
)

:: FORMAT TANGGAL
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
set hh=!hh: =!

echo UPDATE LAST-UPDATE >> tools\log_pengadaan.txt
echo !dd! !bulan! !yyyy! ^| !hh!.!mn! WIB > data\last-update-pengadaan.txt

echo GIT CONFIG >> tools\log_pengadaan.txt
git config user.name "rizkipem-21"
git config user.email "rizki.pem@gmail.com"

echo GIT STATUS >> tools\log_pengadaan.txt
git status >> tools\log_pengadaan.txt 2>&1

:: FIX LOCK
del /f /q .git\index.lock >nul 2>&1

echo GIT ADD >> tools\log_pengadaan.txt
git add . >> tools\log_pengadaan.txt 2>&1

echo GIT COMMIT >> tools\log_pengadaan.txt
git commit -m "auto update pengadaan %date% %time%" >> tools\log_pengadaan.txt 2>&1

echo GIT PUSH >> tools\log_pengadaan.txt
git push origin main >> tools\log_pengadaan.txt 2>&1
echo PUSH STATUS: %ERRORLEVEL% >> tools\log_pengadaan.txt

echo ========================= >> tools\log_pengadaan.txt
echo SELESAI %date% %time% >> tools\log_pengadaan.txt

mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses update data Pengadaan telah SELESAI!"", 5, ""Notifikasi Sistem"", 4160)(window.close)")