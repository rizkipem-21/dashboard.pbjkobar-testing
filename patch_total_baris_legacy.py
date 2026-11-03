# ============================================================================
# JUDUL  : patch_total_baris_legacy.py
# FUNGSI : Menambahkan "Total baris" pada log download LEGACY di
#          scripts\pengadaan\download_pengadaan.py dan scripts\rup\download_rup.py
# CARA   : Jalankan SEKALI dari folder utama project:  python patch_total_baris_legacy.py
#          (file asli dicadangkan sebagai .bak; aman dijalankan ulang)
# ============================================================================
import os, re, shutil

HELPER = '''def hitung_baris(data):
    """Hitung jumlah baris dari respons legacy (list, atau dict berisi data/items/results)"""
    if isinstance(data, list): return len(data)
    if isinstance(data, dict):
        for k in ('data', 'items', 'results'):
            if isinstance(data.get(k), list): return len(data[k])
        return 1
    return 0

'''

TARGET = [
    ('scripts/pengadaan/download_pengadaan.py', 'resp.json()', 'resp'),
    ('scripts/rup/download_rup.py', 'data', 'data'),
]

def patch(path, mode):
    if not os.path.exists(path):
        print(f"[LEWAT] tidak ditemukan: {path}"); return
    src = open(path, encoding='utf-8').read()
    if 'def hitung_baris' in src:
        print(f"[SUDAH] {path} sudah dipatch"); return
    if mode == 'pengadaan':
        pola = re.compile(r'(?P<a>[ \t]*)json\.dump\(resp\.json\(\), f, ensure_ascii=False, indent=2\)\n(?P<b>[ \t]*)log_print\("SUKSES"\)')
        baru = lambda m: (f'{m["a"]}data_legacy = resp.json()\n'
                          f'{m["a"]}json.dump(data_legacy, f, ensure_ascii=False, indent=2)\n'
                          f'{m["b"]}log_print(f"SUKSES (Total: {{hitung_baris(data_legacy)}} baris)")')
    else:
        pola = re.compile(r'(?P<a>[ \t]*)json\.dump\(data, f, ensure_ascii=False, indent=2\)\n(?P<b>[ \t]*)log_print\("SUKSES"\)')
        baru = lambda m: (f'{m["a"]}json.dump(data, f, ensure_ascii=False, indent=2)\n'
                          f'{m["b"]}log_print(f"SUKSES (Total: {{hitung_baris(data)}} baris)")')
    hasil, n = pola.subn(baru, src)
    if n != 1:
        print(f"[GAGAL] {path}: bagian legacy tidak ditemukan persis ({n} kecocokan), file TIDAK diubah"); return
    if 'def get_waktu_indonesia():' not in hasil:
        print(f"[GAGAL] {path}: titik sisip fungsi tidak ditemukan, file TIDAK diubah"); return
    hasil = hasil.replace('def get_waktu_indonesia():', HELPER + 'def get_waktu_indonesia():', 1)
    shutil.copy2(path, path + '.bak')
    open(path, 'w', encoding='utf-8').write(hasil)
    print(f"[OK] {path} dipatch (cadangan: {path}.bak)")

patch(TARGET[0][0], 'pengadaan')
patch(TARGET[1][0], 'rup')