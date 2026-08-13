@echo off
:: Otomatis mendeteksi lokasi file .bat, lalu naik 1 tingkat (..) ke folder utama project
cd /d "%~dp0.."

echo ==================================================
echo MEMULAI PROSES UPDATE RUP...
echo ==================================================

:: Menghapus file lock Git jika sebelumnya sempat macet
del /f /q .git\index.lock >nul 2>&1

:: 1. Memanggil Python untuk mengunduh data API Utama
python scripts\rup\download_rup.py

:: 2. Memanggil Python untuk mengolah Excel Detail RUP
python scripts\rup\detail_rup.py

:: 3. Memanggil Python untuk mengolah Excel Konsolidasi RUP (Scraping Anak)
python scripts\rup\konsolidasi_rup.py

:: 4. Memanggil Python untuk mengolah Rekap, Git Push, dan kirim Telegram (FINAL)
python scripts\rup\generate_rup.py

echo.
echo PROSES SELESAI!
:: Menampilkan Pop-Up sukses di layar Windows selama 5 detik
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses RUP Selesai!"", 5, ""Update Selesai"", 4160)(window.close)")