@echo off
:: Otomatis mendeteksi lokasi file .bat, lalu naik 1 tingkat (..) ke folder utama project
cd /d "%~dp0.."

echo ==================================================
echo MEMULAI PROSES UPDATE RUP...
echo ==================================================

:: Memanggil Python
python scripts\rup\generate_rup.py

echo.
echo ==================================================
echo MENGIRIM DATA REKAP KE GITHUB...
echo ==================================================

:: Menghapus file lock Git jika sebelumnya sempat macet
del /f /q .git\index.lock >nul 2>&1

:: Menjalankan perintah Git
git config user.name "rizkipem-21"
git config user.email "rizki.pem@gmail.com"
git add .
git commit -m "Auto update data RUP"

:: Cukup gunakan 'git push' agar otomatis menyesuaikan reponya masing-masing
git push

echo.
echo PROSES SELESAI!
:: Menampilkan Pop-Up sukses di layar Windows selama 5 detik
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses RUP dan Push ke Server Selesai!"", 5, ""Update Selesai"", 4160)(window.close)")