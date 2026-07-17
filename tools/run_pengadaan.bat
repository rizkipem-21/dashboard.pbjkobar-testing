@echo off
:: Otomatis mendeteksi lokasi file .bat, lalu naik 1 tingkat (..) ke folder utama project
cd /d "%~dp0.."

echo ==================================================
echo MEMULAI PROSES UPDATE PENGADAAN...
echo ==================================================

:: Menghapus file lock Git jika sebelumnya sempat macet
del /f /q .git\index.lock >nul 2>&1

:: 1. Memanggil skrip untuk MENGUNDUH data (Download)
python scripts\pengadaan\download_pengadaan.py

echo.
echo ==================================================
echo MELANJUTKAN KE PROSES GENERATE DATA...
echo ==================================================

:: 2. Memanggil skrip untuk MENGOLAH data, Push Git, dan Kirim Telegram
python scripts\pengadaan\generate_pengadaan.py

echo.
echo PROSES SELESAI!
:: Menampilkan Pop-Up sukses di layar Windows selama 5 detik
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses Pengadaan Selesai!"", 5, ""Update Selesai"", 4160)(window.close)")