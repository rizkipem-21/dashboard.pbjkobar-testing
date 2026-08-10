@echo off
:: Otomatis mendeteksi lokasi file .bat, lalu naik 1 tingkat (..) ke folder utama project
cd /d "%~dp0.."

echo ==================================================
echo MEMULAI PROSES UPDATE DATA KBLI INAPROC...
echo ==================================================

:: Menghapus file lock Git jika sebelumnya sempat macet
del /f /q .git\index.lock >nul 2>&1

:: 1. Memanggil Skrip Python untuk Ekstraksi & Pembersihan PDF
echo [1/3] Mengekstrak PDF dan menerapkan kamus pembersihan...
python scripts\kategori_kbli\generate_kbli.py

:: 2. Proses Git untuk mempersiapkan pengunggahan data baru
echo [2/3] Memeriksa perubahan data untuk dikirim ke GitHub...
git add data_master\data_kategori_inaproc.json

:: Mengirimkan perubahan ke repositori lokal dengan catatan otomatis
git commit -m "update: sinkronisasi berkala data kategori kbli inaproc"

:: 3. Mengunggah data langsung ke server GitHub
echo [3/3] Mengunggah file ke GitHub (Git Push)...
git push

echo.
echo PROSES EKSTRAKSI DAN GIT PUSH SELESAI!
:: Menampilkan Pop-Up sukses di layar Windows selama 5 detik
mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""Proses KBLI & Git Push Selesai!"", 5, ""Update Selesai"", 4160)(window.close)")