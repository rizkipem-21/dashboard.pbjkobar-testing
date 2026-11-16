# ======================================================
# patch_log3.py  (jalankan sekali di folder utama proyek, lalu hapus)
# Menambahkan log harian (30 hari) ke skrip yang tadinya hanya print():
#   generate_tepra.py, arsip_tepra.py  -> logs\tepra\log_tepra_TANGGAL.txt
#   arsip_pengadaan.py                 -> logs\pengadaan\log_pengadaan_TANGGAL.txt
# Semua print() tetap tampil di layar DAN ikut ditulis ke file log.
# Cadangan: file.bak3. Aman dijalankan berulang.
# ======================================================
import os
import shutil

BASE = os.path.dirname(os.path.abspath(__file__))


def baca(path):
    with open(path, 'r', encoding='utf-8', newline='') as f:
        return f.read()


def tulis(path, s):
    with open(path, 'w', encoding='utf-8', newline='') as f:
        f.write(s)


def blok_log(sub, prefix):
    return f'''
# --- Log harian: logs/{sub}/, simpan 30 hari. Semua print() ikut ditulis ke file ---
_LOG_DIR = os.path.join(BASE_DIR, 'logs', '{sub}')
os.makedirs(_LOG_DIR, exist_ok=True)
_LOG_FILE = os.path.join(_LOG_DIR, f"{prefix}_{{datetime.now().strftime('%Y-%m-%d')}}.txt")
import time as _t
for _f in os.listdir(_LOG_DIR):
    if _f.startswith('{prefix}_') and _f.endswith('.txt'):
        try:
            if os.path.getmtime(os.path.join(_LOG_DIR, _f)) < _t.time() - 30 * 86400:
                os.remove(os.path.join(_LOG_DIR, _f))
        except OSError: pass
_print_asli = print
def print(*args, **kwargs):
    _print_asli(*args, **kwargs)
    try:
        with open(_LOG_FILE, 'a', encoding='utf-8') as _fh:
            _fh.write(" ".join(str(a) for a in args) + '\\n')
    except OSError: pass
print(f"\\n=== {{os.path.basename(__file__)}} mulai {{datetime.now().strftime('%H:%M:%S')}} ===")
'''


def sisip(rel, anchor, sub, prefix, ekstra=None):
    path = os.path.join(BASE, *rel)
    nama = '\\'.join(rel)
    if not os.path.exists(path):
        print(f'[LEWATI] {nama} tidak ada')
        return
    s = baca(path)
    if '_LOG_FILE' in s:
        print(f'[SUDAH] {nama}')
        return
    nl = '\r\n' if '\r\n' in s else '\n'
    if s.count(anchor) != 1:
        print(f'[GAGAL] {nama}: baris acuan {anchor!r} ditemukan {s.count(anchor)}x (file tidak diubah)')
        return
    blok = blok_log(sub, prefix).replace('\n', nl)
    baru = s.replace(anchor, anchor + nl + blok.rstrip('\r\n'))
    if ekstra:
        lama, ganti = ekstra
        if lama in baru:
            baru = baru.replace(lama, ganti)
    shutil.copy2(path, path + '.bak3')
    tulis(path, baru)
    print(f'[OK] {nama}  (cadangan: .bak3)')


sisip(('scripts', 'pengadaan', 'generate_tepra.py'), "tahun_n  = datetime.now().year", 'tepra', 'log_tepra',
      ekstra=('"GAGAL, cek log")', '"GAGAL, cek folder logs/tepra")'))
sisip(('scripts', 'pengadaan', 'arsip_tepra.py'), "tahun = datetime.now().year", 'tepra', 'log_tepra')
sisip(('scripts', 'pengadaan', 'arsip_pengadaan.py'), "tgl = now.strftime('%Y-%m-%d')", 'pengadaan', 'log_pengadaan')

print('\nSelesai. Cek pesan [GAGAL] kalau ada. Hapus file .bak3 setelah uji berhasil.')