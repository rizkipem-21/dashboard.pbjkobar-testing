@echo off
cd /d D:\dashboard-pbj-inaproc

echo ==================================================
echo MEMULAI PROSES UPDATE PENGADAAN...
echo ==================================================

:: (Opsional) Jika Anda masih butuh powershell untuk download, hilangkan tanda :: di baris bawah ini:
:: powershell -ExecutionPolicy Bypass -File scripts\pengadaan\download_pengadaan.ps1

:: Memanggil Python (Log, Last-Update, dan Excel dikerjakan otomatis oleh Python)
python scripts\pengadaan\generate_pengadaan.py

echo.
echo ==================================================
echo MENGIRIM DATA REKAP KE GITHUB (UNTUK CLOUDFLARE)...
echo ==================================================

:: Menghapus file lock Git jika sebelumnya sempat macet
del /f /q .git\index.lock >nul 2>&1

:: Menjalankan perintah Git
git config user.name "rizkipem-21"
git config user.email "rizki.pem@gmail.com"
git add .
git commit -m "Auto update data Pengadaan"
git push origin main

echo.
echo PROSES SELESAI!
:: Menampilkan Pop-Up sukses di layar Windows selama 5 detik
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses update data Pengadaan telah SELESAI!"", 5, ""Notifikasi Sistem"", 4160)(window.close)")