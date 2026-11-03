@echo off
REM ============================================================================
REM  JUDUL  : run_pengadaan.bat
REM  FUNGSI : Update data Pengadaan: 1) download data, 2) generate + push Git + Telegram
REM  CATATAN: Menampilkan jam tiap tahap dan kode error, sehingga terlihat tahap mana
REM           yang lambat/macet. Popup sukses tutup sendiri 5 detik lalu jendela
REM           menutup. Jika ada error: popup error + jendela TETAP terbuka.
REM ============================================================================

:: Otomatis mendeteksi lokasi file .bat, lalu naik 1 tingkat (..) ke folder utama project
cd /d "%~dp0.."

echo ==================================================
echo MEMULAI PROSES UPDATE PENGADAAN...  [%date% %time%]
echo ==================================================

:: Menghapus file lock Git jika sebelumnya sempat macet
del /f /q .git\index.lock >nul 2>&1

:: 1. MENGUNDUH data (Download)
echo.
echo [%time%] TAHAP 1: Download data...
python scripts\pengadaan\download_pengadaan.py
set ERR_DOWNLOAD=%ERRORLEVEL%
echo [%time%] TAHAP 1 selesai (kode: %ERR_DOWNLOAD%)

echo.
echo ==================================================
echo MELANJUTKAN KE PROSES GENERATE DATA...
echo ==================================================

:: 2. MENGOLAH data, Push Git, dan Kirim Telegram
echo [%time%] TAHAP 2: Generate data...
python scripts\pengadaan\generate_pengadaan.py
set ERR_GENERATE=%ERRORLEVEL%
echo [%time%] TAHAP 2 selesai (kode: %ERR_GENERATE%)

if not "%ERR_DOWNLOAD%"=="0" goto GAGAL
if not "%ERR_GENERATE%"=="0" goto GAGAL
goto SUKSES

:GAGAL
echo.
echo ==================================================
echo PROSES SELESAI DENGAN ERROR! (download=%ERR_DOWNLOAD%, generate=%ERR_GENERATE%)
echo Periksa pesan di atas dan tools\log_pengadaan.txt
echo ==================================================
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses Pengadaan SELESAI DENGAN ERROR! Periksa jendela dan log."", 0, ""Update Gagal"", 4112)(window.close)")
pause
exit /b 1

:SUKSES
echo.
echo PROSES SELESAI!  [%date% %time%]
:: Menampilkan Pop-Up sukses di layar Windows selama 5 detik
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses Pengadaan Selesai!"", 5, ""Update Selesai"", 4160)(window.close)")
exit