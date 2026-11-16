import os, shutil
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
now = datetime.now(); tgl = now.strftime('%Y-%m-%d'); bulan = tgl[:7]

for tahun in (now.year, now.year - 1, now.year - 2):
    master = os.path.join(BASE_DIR, 'data', str(tahun), f'master_pengadaan_{tahun}.xlsx')
    if not os.path.exists(master): continue
    tujuan = os.path.join(BASE_DIR, 'arsip_lokal', 'pengadaan', str(tahun))
    os.makedirs(tujuan, exist_ok=True)
    nama = f'Paket Pengadaan Tahun {tahun} ({tgl}).xlsx'
    try:
        shutil.copy2(master, os.path.join(tujuan, nama))
        for f in os.listdir(tujuan):   # sisakan 1 file terbaru per bulan
            if f != nama and f'({bulan}-' in f and f.endswith('.xlsx'):
                os.remove(os.path.join(tujuan, f))
        print('Arsip:', nama)
    except OSError as e:
        print(f'Gagal arsip {nama}: {e}')