@echo off
:: Otomatis mendeteksi lokasi file .bat, lalu naik 1 tingkat (..) ke folder utama project
cd /d "%~dp0.."

echo ==================================================
echo MEMULAI PROSES BACKUP JSON API...
echo ==================================================

:: Memanggil Python untuk mengeksekusi skrip backup JSON
python scripts\json_backup\json_backup_download.py

echo.
echo PROSES SELESAI!
:: Menampilkan Pop-Up sukses di layar Windows selama 5 detik
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses Backup JSON Selesai!"", 5, ""Backup Selesai"", 4160)(window.close)")