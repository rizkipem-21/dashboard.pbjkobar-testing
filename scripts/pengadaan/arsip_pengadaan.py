# ======================================================
# arsip_pengadaan.py  (LOKAL)
# Menyalin master_pengadaan_{tahun}.xlsx ke arsip_lokal/pengadaan/{tahun}/
# dengan nama bertanggal. Satu file per hari, TIDAK dipangkas.
# Dijalankan lagi di hari yang sama = file hari itu ditimpa.
# Tahun n-2 (sudah final) hanya disalin kalau isi Excel berbeda
# dari arsip terakhir (dibandingkan dengan hash isi, bukan tanggal).
# ======================================================
import os, glob, shutil, hashlib
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
now = datetime.now()
tgl = now.strftime('%Y-%m-%d')

# --- Log harian: logs/pengadaan/, simpan 30 hari. Semua print() ikut ditulis ke file ---
_LOG_DIR = os.path.join(BASE_DIR, 'logs', 'pengadaan')
os.makedirs(_LOG_DIR, exist_ok=True)
_LOG_FILE = os.path.join(_LOG_DIR, f"log_pengadaan_{datetime.now().strftime('%Y-%m-%d')}.txt")
import time as _t
for _f in os.listdir(_LOG_DIR):
    if _f.startswith('log_pengadaan_') and _f.endswith('.txt'):
        try:
            if os.path.getmtime(os.path.join(_LOG_DIR, _f)) < _t.time() - 30 * 86400:
                os.remove(os.path.join(_LOG_DIR, _f))
        except OSError: pass
_print_asli = print
def print(*args, **kwargs):
    _print_asli(*args, **kwargs)
    try:
        with open(_LOG_FILE, 'a', encoding='utf-8') as _fh:
            _fh.write(" ".join(str(a) for a in args) + '\n')
    except OSError: pass
print(f"\n=== {os.path.basename(__file__)} mulai {datetime.now().strftime('%H:%M:%S')} ===")


def hash_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for blok in iter(lambda: f.read(1024 * 1024), b''):
            h.update(blok)
    return h.hexdigest()


for tahun in (now.year, now.year - 1, now.year - 2):
    master = os.path.join(BASE_DIR, 'data', str(tahun), f'master_pengadaan_{tahun}.xlsx')
    if not os.path.exists(master): continue
    tujuan = os.path.join(BASE_DIR, 'arsip_lokal', 'pengadaan', str(tahun))
    os.makedirs(tujuan, exist_ok=True)
    nama = f'Paket Pengadaan Tahun {tahun} ({tgl}).xlsx'
    try:
        if tahun == now.year - 2:
            lama = sorted(glob.glob(os.path.join(tujuan, f'Paket Pengadaan Tahun {tahun} (*).xlsx')))
            if lama and hash_file(lama[-1]) == hash_file(master):
                print(f'Lewati {tahun}: tidak ada Excel baru sejak {os.path.basename(lama[-1])}')
                continue
        shutil.copy2(master, os.path.join(tujuan, nama))
        print('Arsip:', nama)
    except OSError as e:
        print(f'Gagal arsip {nama}: {e}')
