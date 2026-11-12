# update_kamus_penyedia.py

import os
import sys
import json
import glob
import re
import time
import tempfile
import warnings
from datetime import datetime

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

warnings.filterwarnings('ignore')

# ======================================================
# KONFIGURASI
# ======================================================

# Lokasi file: <project>/scripts/<sesuatu>/file.py  -> sesuaikan JUMLAH_NAIK bila dipindah
JUMLAH_NAIK = 3
BASE_DIR = os.path.abspath(__file__)
for _ in range(JUMLAH_NAIK):
    BASE_DIR = os.path.dirname(BASE_DIR)

MASTER_DIR        = os.path.join(BASE_DIR, 'data_master')
KAMUS_FILE        = os.path.join(MASTER_DIR, 'kamus_penyedia_raw.json')       # mentah S6 + S7
KAMUS_GABUNG_FILE = os.path.join(MASTER_DIR, 'kamus_penyedia.json')           # hasil gabungan (dipakai dashboard)
REVIEW_FILE       = os.path.join(MASTER_DIR, 'kamus_penyedia_review_gabung.json')
GAGAL_FILE        = os.path.join(MASTER_DIR, 'kamus_penyedia_gagal.json')
SIKAP_FILE        = os.path.join(MASTER_DIR, 'kamus_sikap_raw.json')          # mentah Sikap (sudah diringkas)
GAGAL_SIKAP_FILE  = os.path.join(MASTER_DIR, 'kamus_sikap_gagal.json')

SIMPAN_SETIAP = 50     # autosave tiap N data baru
JEDA_REQUEST  = 0.2    # detik
TIMEOUT       = 15
TIMEOUT_SIKAP = 40     # respons Sikap besar

URL_SIKAP = 'https://data.inaproc.id/api/v1/sikap/detail-pelaku-usaha'

POLA_S6 = 'ekatalog_paket-e-purchasing'
POLA_S7 = 'ekatalog-archive_paket-e-purchasing'
POLA_SPPBJ_T  = 'tender_tender-ekontrak-sppbj'
POLA_SPPBJ_NT = 'tender_non-tender-ekontrak-sppbj'
KOLOM_NPWP16_SPPBJ = ('npwp_penyedia_16', 'npwp16_penyedia')

# Alasan gagal yang dianggap permanen (tidak ditembak ulang tiap run).
# Untuk mencoba ulang: hapus kamus_sikap_gagal.json
GAGAL_PERMANEN = ('404', 'data kosong', 'npwp tidak cocok')


def baca_token():
    path = os.path.join(BASE_DIR, 'token.txt')
    try:
        with open(path, 'r', encoding='utf-8') as f:
            token = f.read().strip()
    except FileNotFoundError:
        sys.exit("ERROR: token.txt tidak ditemukan di folder project!")
    if not token:
        sys.exit("ERROR: token.txt kosong! Isi dengan token terbaru.")
    return token


def buat_session(token):
    s = requests.Session()
    s.headers.update({
        'Authorization': f'Bearer {token}',
        'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                       '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'),
        'Accept': 'application/json, text/plain, */*',
    })
    retry = Retry(
        total=3, backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=['GET'],
    )
    s.mount('https://', HTTPAdapter(max_retries=retry))
    return s


tahun_n = datetime.now().year
daftar_tahun = [tahun_n - 2, tahun_n - 1, tahun_n]

# ======================================================
# UTILITAS FILE
# ======================================================

def simpan_json_atomik(path, data):
    """Tulis ke file temporer dulu, lalu ganti -> file asli tidak mungkin setengah jadi."""
    folder = os.path.dirname(path)
    fd, tmp = tempfile.mkstemp(dir=folder, suffix='.tmp')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def load_list(path):
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        return []
    try:
        with open(path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        sys.exit(f"ERROR: {path} rusak ({e}). Perbaiki/restore dulu, supaya tidak tertimpa.")
    if isinstance(data, dict):
        for k in ('data', 'items', 'results'):
            if isinstance(data.get(k), list):
                return data[k]
        sys.exit(f"ERROR: format {path} (dict) tidak dikenali.")
    return data if isinstance(data, list) else []


def load_gagal(path):
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def simpan_gagal(path, data):
    """Simpan daftar gagal; kalau kosong, hapus filenya."""
    if data:
        simpan_json_atomik(path, data)
    elif os.path.exists(path):
        os.remove(path)

def cari_file(folder, pola, tahun):
    """Cari file berdasarkan pola, apa pun awalannya (Legacy_, v1_, dll)."""
    hasil = sorted(glob.glob(os.path.join(folder, f'*{pola}_{tahun}.json')))
    if not hasil:
        print(f"[!] File tidak ditemukan: *{pola}_{tahun}.json di {folder}")
    return hasil


def get_unique_codes(file_path, kolom_id):
    """kolom_id boleh satu nama atau tuple kandidat."""
    if isinstance(kolom_id, str):
        kolom_id = (kolom_id,)

    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        return set()
    try:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
    except Exception as e:
        print(f"[!] Gagal baca {os.path.basename(file_path)}: {e}")
        return set()

    if isinstance(data, dict):
        for k in ('data', 'items', 'results'):
            if isinstance(data.get(k), list):
                data = data[k]
                break
    if not isinstance(data, list):
        return set()

    codes = set()
    for item in data:
        if not isinstance(item, dict):
            continue
        for kol in kolom_id:
            if item.get(kol):
                codes.add(str(item[kol]).strip())
                break
    return {c for c in codes if c and c.lower() != 'none'}


def angka(v):
    return re.sub(r'\D', '', str(v or ''))


# ======================================================
# PENGGABUNGAN S6 + S7 + SIKAP (V1)
# ======================================================

# Nama field sumber -> disamakan ke nama standar (dipakai di data akhir)
ALIAS_FIELD = {
    'no_telp_penyedia': 'telepon',
    'email_penyedia': 'email',
    'alamat': 'alamat_penyedia',
}
FIELD_ID = {'kode_penyedia', 'kd_penyedia', 'kd_penyedia_v1'}
FIELD_NPWP_ASAL = {'npwp_penyedia', 'npwp_16', 'npwp_15'}   # diganti jadi npwp_16 & npwp_15
KATA_BADAN = {'pt', 'cv', 'ud', 'p', 't', 'tbk', 'persero', 'perum'}
URUTAN_SUMBER = [('S6', 'v6'), ('S7', 'v5'), ('V1', 'v1')]   # urutan = prioritas nilai
NAMA_KODE = {'v6': 'epur_v6', 'v5': 'epur_v5', 'v1': 'kd_penyedia_v1'}   # nama parameter kode di data akhir


def sumber_rec(rec):
    """V1 = hasil Sikap. S7 = punya kd_penyedia (numerik). S6 = punya kode_penyedia (ULID)."""
    if 'kd_penyedia_v1' in rec:
        return 'V1'
    if rec.get('kd_penyedia') not in (None, '') and not rec.get('kode_penyedia'):
        return 'S7'
    return 'S6'


def id_dan_sumber(rec):
    s = sumber_rec(rec)
    if s == 'V1':
        kode = str(rec.get('kd_penyedia_v1') or '').strip()
    else:
        kode = str(rec.get('kd_penyedia' if s == 'S7' else 'kode_penyedia') or '').strip()
    return kode, s


def norm_nama(nama):
    teks = re.sub(r'[^a-z0-9 ]', ' ', str(nama or '').lower())
    return ' '.join(t for t in teks.split() if t not in KATA_BADAN)


def pisah_npwp(rec):
    """
    Return (npwp_16, npwp_15). Nilai disimpan APA ADANYA, tidak dipotong/ditambah.
    - 16 digit -> npwp_16
    - 15 digit -> npwp_15
    - kosong / angka palsu (0000...) / panjang lain -> diabaikan
    """
    n16 = n15 = None
    for f in ('npwp_16', 'npwp_15', 'npwp_penyedia'):
        d = angka(rec.get(f))
        if not d or len(set(d)) == 1:
            continue
        if len(d) == 16:
            n16 = n16 or d
        elif len(d) == 15:
            n15 = n15 or d
    return n16, n15


def bersih_daftar(teks):
    """Hapus item duplikat pada daftar KBLI, urutan asli dipertahankan."""
    if ';' in teks:
        sep, bagian = ';', teks.split(';')
    else:
        sep, bagian = ',', re.split(r',(?!\s)', teks)
    unik, lihat = [], set()
    for p in bagian:
        p = p.strip()
        if p and p not in lihat:
            lihat.add(p)
            unik.append(p)
    return sep.join(unik)


def seragamkan(rec):
    """Samakan nama field, pisahkan NPWP, bersihkan duplikat KBLI."""
    out = {}
    for k, v in rec.items():
        if k in FIELD_NPWP_ASAL:
            continue
        k = ALIAS_FIELD.get(k, k)
        if k.startswith('kbli') and isinstance(v, str):   # otomatis berlaku untuk kbli2025 dst.
            v = bersih_daftar(v)
        out[k] = v
    out['npwp_16'], out['npwp_15'] = pisah_npwp(rec)
    return out


def kunci_gabung(rec):
    """Kunci ini HANYA untuk mencocokkan antar sumber, tidak disimpan ke data akhir.
    NPWP 16 digit berawalan 0 dicocokkan dengan 15 digit tanpa 0 di depan."""
    n16, n15 = pisah_npwp(rec)
    if n15:
        return ('npwp', n15)
    if n16:
        return ('npwp', n16[1:]) if n16.startswith('0') else ('npwp16', n16)
    nama = norm_nama(rec.get('nama_penyedia'))
    return ('nama', nama) if nama else None


def kosong(v):
    if isinstance(v, str) and v.strip() in ('', '-'):
        return True
    return v is None or v == [] or v == {}


def sama(a, b):
    if isinstance(a, str) and isinstance(b, str):
        return ' '.join(a.lower().split()) == ' '.join(b.lower().split())
    return a == b


def gabung_bucket(b):
    hasil, recs = {}, {}
    for sumber, suf in URUTAN_SUMBER:
        if sumber in b:
            kode, rec = b[sumber]
            if kode:
                hasil[NAMA_KODE[suf]] = kode
            recs[suf] = seragamkan(rec)

    semua_field = list(dict.fromkeys(f for r in recs.values() for f in r))
    for f in semua_field:
        if f in FIELD_ID:
            continue
        ada = {suf: r[f] for suf, r in recs.items() if f in r}
        terisi = {suf: v for suf, v in ada.items() if not kosong(v)}
        if not terisi:
            hasil[f] = next(iter(ada.values()))
            continue
        unik = []
        for v in terisi.values():
            if not any(sama(v, u) for u in unik):
                unik.append(v)
        if len(unik) == 1:
            hasil[f] = unik[0]
        else:
            for suf, v in terisi.items():
                hasil[f'{f}_{suf}'] = v
    return hasil


def nama_bucket(b):
    return {norm_nama(rec.get('nama_penyedia')) for _, rec in b.values()} - {''}


def npwp_bucket(b):
    s16, s15 = set(), set()
    for _, rec in b.values():
        n16, n15 = pisah_npwp(rec)
        if n16: s16.add(n16)
        if n15: s15.add(n15)
    return s16, s15


def npwp_bentrok(b1, b2):
    a16, a15 = npwp_bucket(b1)
    c16, c15 = npwp_bucket(b2)
    return (a16 and c16 and not (a16 & c16)) or (a15 and c15 and not (a15 & c15))


def gabungkan(raw_list):
    groups = {}   # kunci -> list bucket {'S6': (kode, rec), 'S7': ..., 'V1': ...}
    for i, rec in enumerate(raw_list):
        kode, sumber = id_dan_sumber(rec)
        k = kunci_gabung(rec) or ('solo', kode or str(i))
        buckets = groups.setdefault(k, [])
        for b in buckets:
            if sumber not in b:
                b[sumber] = (kode, rec)
                break
        else:
            buckets.append({sumber: (kode, rec)})

    # ---- Pass 2: Sikap-saja dicocokkan ke e-katalog lewat nama (NPWP beda format, mis. NIK vs NPWP lama)
    semua = [b for bl in groups.values() for b in bl]
    sikap_saja = [b for b in semua if set(b) == {'V1'}]
    lainnya = [b for b in semua if 'V1' not in b]
    idx = {}
    for b in lainnya:
        for n in nama_bucket(b):
            idx.setdefault(n, []).append(b)

    dihapus, review_nama = set(), []
    for sb in sikap_saja:
        kandidat = {}
        for n in nama_bucket(sb):
            for y in idx.get(n, []):
                if 'V1' not in y and not npwp_bentrok(sb, y):
                    kandidat[id(y)] = y
        if len(kandidat) == 1:
            y = next(iter(kandidat.values()))
            y['V1'] = sb['V1']
            dihapus.add(id(sb))
            review_nama.append({
                'cocok_lewat': 'nama (npwp beda format)',
                'nama': sb['V1'][1].get('nama_penyedia'),
                'kode': {s: y[s][0] for s in y},
                'npwp_sikap': [sb['V1'][1].get('npwp_16'), sb['V1'][1].get('npwp_15')],
                'npwp_katalog': [r.get('npwp_16') or r.get('npwp_15') or r.get('npwp_penyedia')
                                 for s, (_, r) in y.items() if s != 'V1'],
            })

    hasil, review = [], list(review_nama)
    for k, buckets in groups.items():
        for b in buckets:
            if id(b) in dihapus:
                continue
            hasil.append(gabung_bucket(b))
            if k[0] == 'nama' and len(b) >= 2:
                review.append({'kunci_nama': k[1],
                               'kode': {s: b[s][0] for s in b},
                               'nama': {s: b[s][1].get('nama_penyedia') for s in b}})
    return hasil, review


def buat_kamus_gabungan(kamus_raw, sikap_raw):
    hasil, review = gabungkan(list(kamus_raw) + list(sikap_raw))
    simpan_json_atomik(KAMUS_GABUNG_FILE, hasil)
    simpan_json_atomik(REVIEW_FILE, review)
    n_dua = sum(1 for h in hasil if 'epur_v6' in h and 'epur_v5' in h)
    n_v1 = sum(1 for h in hasil if 'kd_penyedia_v1' in h)
    n_v1_saja = sum(1 for h in hasil if 'kd_penyedia_v1' in h
                    and 'epur_v6' not in h and 'epur_v5' not in h)
    print(f"[GABUNG] {len(kamus_raw)} mentah S6/S7 + {len(sikap_raw)} mentah Sikap -> {len(hasil)} penyedia "
          f"({n_dua} gabungan v6+v5, {n_v1} punya data Sikap, {n_v1_saja} hanya dari Sikap, "
          f"{len(review)} digabung via nama -> cek {os.path.basename(REVIEW_FILE)})")


# ======================================================
# FETCH API
# ======================================================

class TokenInvalid(Exception):
    pass


def fetch_data_api(session, kode, is_archive=False):
    """
    Return: (data_dict | None, alasan_gagal | None)
    Raise TokenInvalid bila 401/403.
    """
    endpoint = "ekatalog-archive" if is_archive else "ekatalog"
    url = f"https://data.inaproc.id/api/legacy/{endpoint}/penyedia-detail"

    try:
        resp = session.get(url, params={'kode_penyedia': kode}, timeout=TIMEOUT)
    except requests.RequestException as e:
        return None, f"network: {type(e).__name__}"

    if resp.status_code in (401, 403):
        raise TokenInvalid(f"HTTP {resp.status_code}")
    if resp.status_code == 404:
        return None, "404 tidak ditemukan"
    if resp.status_code != 200:
        return None, f"HTTP {resp.status_code}"

    try:
        data = resp.json()
    except ValueError:
        return None, "respon bukan JSON"

    obj = None
    if isinstance(data, list) and data and isinstance(data[0], dict):
        obj = data[0]
    elif isinstance(data, dict):
        inner = data.get('data')
        if isinstance(inner, list) and inner and isinstance(inner[0], dict):
            obj = inner[0]
        elif isinstance(inner, dict):
            obj = inner
        else:
            obj = data

    if not obj:
        return None, "data kosong"

    if not any(k in obj for k in ('nama_penyedia', 'kode_penyedia', 'kd_penyedia')):
        return None, "bentuk data tidak dikenali"

    if is_archive:
        obj.setdefault('kd_penyedia', kode)
    else:
        obj.setdefault('kode_penyedia', kode)

    return obj, None


def ratakan_izin(raw):
    """izin_usaha dari Sikap berupa string JSON. Ambil jenis, nomor, kualifikasi; buang duplikat persis."""
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except ValueError:
            return []
    if not isinstance(raw, list):
        return []
    hasil, lihat = [], set()
    for it in raw:
        if not isinstance(it, dict):
            continue
        e = {
            'jenis_izin_usaha': str(it.get('jenis_izin_usaha') or '').strip(),
            'nomor_izin_usaha': str(it.get('nomor_izin_usaha') or '').strip(),
            'kualifikasi_usaha': str(it.get('kualifikasi_usaha') or '').strip(),
        }
        k = tuple(e.values())
        if not any(k) or k in lihat:
            continue
        lihat.add(k)
        hasil.append(e)
    return hasil


def ratakan_sikap(d, npwp16):
    """Ringkas respons Sikap jadi satu baris datar. akta/pajak/pengurus/ahli/pengalaman sengaja dibuang."""
    n16 = angka(d.get('npwp16_penyedia')) or npwp16
    n15 = angka(d.get('npwp_penyedia'))
    return {
        'kd_penyedia_v1': str(d.get('kd_penyedia') or '').strip() or None,
        'npwp_16': n16 if len(n16) == 16 else None,
        'npwp_15': n15 if len(n15) == 15 else None,   # titik pemisah dibuang, digit tidak diubah
        'nama_penyedia': d.get('nama_penyedia'),
        'bentuk_usaha': d.get('bentuk_usaha'),
        'status_pelaku_usaha': d.get('status_pelaku_usaha'),
        'alamat': d.get('alamat'),
        'kabupaten': d.get('kabupaten'),
        'provinsi': d.get('provinsi'),
        'kodepos': d.get('kodepos'),
        'email': d.get('email'),
        'telepon': d.get('telepon'),
        'status_npwp': d.get('status_npwp'),
        'status_kswp': d.get('status_kswp'),
        'tgl_daftar_sikap': d.get('tgl_daftar_sikap'),
        'tgl_persetujuan_verifikasi_daftar_sikap': d.get('tgl_persetujuan_verifikasi_daftar_sikap'),
        'setuju_publikasi_data': d.get('setuju_publikasi_data'),
        'izin_usaha': ratakan_izin(d.get('izin_usaha')),
    }


def fetch_sikap(session, npwp16):
    """Return: (record_datar | None, alasan_gagal | None). Raise TokenInvalid bila 401/403."""
    try:
        resp = session.get(URL_SIKAP, params={'npwp16_penyedia': npwp16}, timeout=TIMEOUT_SIKAP)
    except requests.RequestException as e:
        return None, f"network: {type(e).__name__}"

    if resp.status_code in (401, 403):
        raise TokenInvalid(f"HTTP {resp.status_code}")
    if resp.status_code == 404:
        return None, "404 tidak ditemukan"
    if resp.status_code != 200:
        return None, f"HTTP {resp.status_code}"

    try:
        data = resp.json()
    except ValueError:
        return None, "respon bukan JSON"

    daftar = data.get('data') if isinstance(data, dict) else data
    if isinstance(daftar, dict):
        daftar = [daftar]
    if not isinstance(daftar, list) or not daftar:
        return None, "data kosong"

    cocok = [d for d in daftar if isinstance(d, dict) and angka(d.get('npwp16_penyedia')) == npwp16]
    if not cocok:
        return None, "npwp tidak cocok"
    return ratakan_sikap(cocok[0], npwp16), None


# ======================================================
# PROSES
# ======================================================

def proses_batch(session, daftar_kode, label, is_archive, kamus_list, gagal, stat):
    total = len(daftar_kode)
    for i, kode in enumerate(daftar_kode, 1):
        print(f"[{label}] {i}/{total} | ID: {kode[:10]}...", end=" ", flush=True)
        data, alasan = fetch_data_api(session, kode, is_archive)

        if data:
            kamus_list.append(data)
            gagal.pop(kode, None)
            stat['baru'] += 1
            print(f"Sukses -> {data.get('nama_penyedia', 'Tanpa Nama')}")
            if stat['baru'] % SIMPAN_SETIAP == 0:
                simpan_json_atomik(KAMUS_FILE, kamus_list)
                simpan_json_atomik(GAGAL_FILE, gagal)
                print(f"    [autosave] {len(kamus_list)} data tersimpan")
        else:
            gagal[kode] = {'sumber': label, 'alasan': alasan,
                           'waktu': datetime.now().isoformat(timespec='seconds')}
            stat['gagal'] += 1
            print(f"Gagal ({alasan})")

        time.sleep(JEDA_REQUEST)


def proses_sikap(session, daftar_npwp, sikap_list, gagal, stat):
    total = len(daftar_npwp)
    for i, n16 in enumerate(daftar_npwp, 1):
        print(f"[SIKAP] {i}/{total} | NPWP: {n16[:8]}...", end=" ", flush=True)
        rec, alasan = fetch_sikap(session, n16)

        if rec:
            sikap_list.append(rec)
            gagal.pop(n16, None)
            stat['sikap_baru'] += 1
            print(f"Sukses -> {rec.get('nama_penyedia') or 'Tanpa Nama'} "
                  f"({len(rec['izin_usaha'])} izin)")
            if stat['sikap_baru'] % SIMPAN_SETIAP == 0:
                simpan_json_atomik(SIKAP_FILE, sikap_list)
                simpan_json_atomik(GAGAL_SIKAP_FILE, gagal)
                print(f"    [autosave] {len(sikap_list)} data Sikap tersimpan")
        else:
            gagal[n16] = {'sumber': 'SIKAP', 'alasan': alasan,
                          'waktu': datetime.now().isoformat(timespec='seconds')}
            stat['sikap_gagal'] += 1
            print(f"Gagal ({alasan})")

        time.sleep(JEDA_REQUEST)


def hitung_missing_kamus(kamus_list):
    existing_ids = set()
    for item in kamus_list:
        for key in ('kode_penyedia', 'kd_penyedia'):
            if item.get(key):
                existing_ids.add(str(item[key]).strip())

    kode_s6, kode_s7 = set(), set()
    for tahun in daftar_tahun:
        d = os.path.join(BASE_DIR, 'data', str(tahun))
        for f in cari_file(d, POLA_S6, tahun):
            kode_s6 |= get_unique_codes(f, ('kode_penyedia', 'kd_penyedia'))
        for f in cari_file(d, POLA_S7, tahun):
            kode_s7 |= get_unique_codes(f, ('kd_penyedia', 'kode_penyedia'))

    missing_s6 = sorted(c for c in kode_s6 if c not in existing_ids)
    missing_s7 = sorted(c for c in kode_s7 if c not in existing_ids and c not in kode_s6)
    return missing_s6, missing_s7


def kumpulkan_npwp16():
    """NPWP 16 digit unik dari SPPBJ tender dan non tender, semua tahun."""
    semua = set()
    for tahun in daftar_tahun:
        d = os.path.join(BASE_DIR, 'data', str(tahun))
        for pola in (POLA_SPPBJ_T, POLA_SPPBJ_NT):
            for f in cari_file(d, pola, tahun):
                semua |= get_unique_codes(f, KOLOM_NPWP16_SPPBJ)

    valid = set()
    for v in semua:
        d = angka(v)
        if len(d) == 16 and len(set(d)) > 1:
            valid.add(d)
    if len(semua) != len(valid):
        print(f"[*] {len(semua) - len(valid)} NPWP SPPBJ diabaikan (bukan 16 digit / angka palsu).")
    return valid


def hitung_missing_sikap(sikap_list, gagal_sikap):
    sudah = {s.get('npwp_16') for s in sikap_list if s.get('npwp_16')}
    permanen = {k for k, v in gagal_sikap.items()
                if str(v.get('alasan', '')).startswith(GAGAL_PERMANEN)}
    return sorted(n for n in kumpulkan_npwp16() if n not in sudah and n not in permanen)


def main():
    print("=" * 55)
    print(f"  UPDATE KAMUS PENYEDIA FULL DATA (Tahun {', '.join(map(str, daftar_tahun))})")
    print("=" * 55)

    os.makedirs(MASTER_DIR, exist_ok=True)
    session = buat_session(baca_token())

    kamus_list = load_list(KAMUS_FILE)
    sikap_list = load_list(SIKAP_FILE)
    gagal = load_gagal(GAGAL_FILE)
    gagal_sikap = load_gagal(GAGAL_SIKAP_FILE)
    print(f"[*] Kamus S6/S7 termuat: {len(kamus_list)} | Sikap termuat: {len(sikap_list)}")

    missing_s6, missing_s7 = hitung_missing_kamus(kamus_list)
    missing_sikap = hitung_missing_sikap(sikap_list, gagal_sikap)
    print(f"[*] Perlu fetch: S6 {len(missing_s6)} | S7 {len(missing_s7)} | Sikap {len(missing_sikap)}")

    stat = {'baru': 0, 'gagal': 0, 'sikap_baru': 0, 'sikap_gagal': 0}

    try:
        proses_batch(session, missing_s6, 'S6', False, kamus_list, gagal, stat)
        proses_batch(session, missing_s7, 'S7', True,  kamus_list, gagal, stat)
        proses_sikap(session, missing_sikap, sikap_list, gagal_sikap, stat)
    except TokenInvalid as e:
        print(f"\n[FATAL] Token expired / IP diblokir ({e}). Perbarui token.txt.")
    except KeyboardInterrupt:
        print("\n[!] Dihentikan manual. Menyimpan data yang sudah didapat...")
    finally:
        simpan_json_atomik(KAMUS_FILE, kamus_list)
        simpan_gagal(GAGAL_FILE, gagal)                  # <-- diganti
        simpan_json_atomik(SIKAP_FILE, sikap_list)
        simpan_gagal(GAGAL_SIKAP_FILE, gagal_sikap)      # <-- diganti

    buat_kamus_gabungan(kamus_list, sikap_list)

    print(f"\n[SELESAI] S6/S7 tambah {stat['baru']} (gagal {stat['gagal']}) | "
          f"Sikap tambah {stat['sikap_baru']} (gagal {stat['sikap_gagal']})")
    if stat['gagal']:
        print(f"[*] Detail kegagalan S6/S7: {GAGAL_FILE}")
    if stat['sikap_gagal']:
        print(f"[*] Detail kegagalan Sikap: {GAGAL_SIKAP_FILE}")


if __name__ == '__main__':
    main()