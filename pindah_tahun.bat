@echo off
echo Memilah file JSON tahun 2025 dan 2026...

:: Buat folder jika belum ada
if not exist "2025" mkdir "2025"
if not exist "2026" mkdir "2026"

:: Pindahkan file berdasarkan nama tahun
move *2025*.json "2025\"
move *2026*.json "2026\"

echo.
echo Proses pemilahan selesai!
pause