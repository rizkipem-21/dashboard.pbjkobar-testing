# ======================================================
# AUTO DOWNLOAD & GENERATE PAKET PENGADAAN - MULTI TAHUN
# ======================================================

import os
import json
import re
import time
import shutil
import warnings
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ======================================================
# KONFIGURASI UTAMA
# ======================================================
BASE_DIR = r'D:\dashboard-pbj-inaproc'
TOKEN    = "inprc7642391c38774272bf57ca25ac1d4544"
HEADERS  = {
    "Authorization": f"Bearer {TOKEN}",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7"
}

tahun_n      = datetime.now().year       # Tahun berjalan
tahun_n1     = tahun_n - 1               # Tahun lalu
tahun_n2     = tahun_n - 2               # Dua tahun lalu
daftar_tahun = [tahun_n2, tahun_n1, tahun_n] # Urutan pemrosesan dari terlama ke terbaru

# 19 Daftar Endpoint API Inaproc
ENDPOINTS = [
    "rup/paket-penyedia-terumumkan",
    "rup/paket-swakelola-terumumkan",
    "tender/non-tender-ekontrak-bapbast",
    "tender/non-tender-ekontrak-kontrak",
    "tender/non-tender-ekontrak-spmkspp",
    "tender/non-tender-ekontrak-sppbj",
    "tender/non-tender-pengumuman",
    "tender/non-tender-selesai",
    "tender/pengumuman",
    "tender/tender-ekontrak-bapbast",
    "tender/tender-ekontrak-kontrak",
    "tender/tender-ekontrak-spmkspp",
    "tender/tender-ekontrak-sppbj",
    "tender/tender-selesai",
    "tender/tender-selesai-nilai",
    "tender/pencatatan-non-tender",
    "tender/pencatatan-swakelola",
    "ekatalog-archive/paket-e-purchasing",
    "ekatalog/paket-e-purchasing"
]

# ======================================================
# FUNGSI LOG & LAST-UPDATE (PENGGANTI .BAT)
# ======================================================
LOG_FILE = os.path.join(BASE_DIR, 'tools', 'log_pengadaan.txt')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

def log_print(*args, **kwargs):
    """Mencetak ke layar terminal sekaligus ke file log txt"""
    msg = " ".join(str(a) for a in args)
    print(msg, **kwargs)
    
    if 'end' in kwargs and kwargs['end'] == " ": return
    
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(msg + '\n')

def get_waktu_indonesia():
    """Mengambil waktu pasti WIB (UTC+7) dengan Bulan Indonesia"""
    tz_wib = timezone(timedelta(hours=7))
    sekarang = datetime.now(tz_wib)
    bulan_indo = {
        1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
        5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
        9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
    }
    return f"{sekarang.day} {bulan_indo[sekarang.month]} {sekarang.year} | {sekarang.strftime('%H.%M')} WIB"

# ======================================================
# HELPER FUNCTIONS
# ======================================================
def format_tgl(val):
    """Konversi format tanggal API menjadi DD-MM-YYYY"""
    if not val or (not isinstance(val, str) and pd.isna(val)):
        return ""
    try:
        tgl = str(val).strip()[:10]
        parts = tgl.split('-')
        if len(parts) == 3:
            return f"{parts[2]}-{parts[1]}-{parts[0]}"
        return ""
    except:
        return ""

def download_with_retry(url, output_path):
    """Fungsi download API Gateway dengan fitur retry hingga 5 kali"""
    max_retry = 5
    success = False
    
    for i in range(1, max_retry + 1):
        try:
            log_print(f"  Percobaan ke-{i}...", end=" ")
            response = requests.get(url, headers=HEADERS, timeout=30)
            
            if response.status_code == 200:
                data_json = response.json()
                if data_json is not None:
                    with open(output_path, 'w', encoding='utf-8') as f:
                        json.dump(data_json, f, ensure_ascii=False, indent=2)
                    log_print("-> SUKSES")
                    success = True
                    break
            else:
                log_print(f"-> Gagal HTTP {response.status_code}")
        except Exception as e:
            log_print(f"-> Gagal error: {e}")
        
        time.sleep(2)
        
    if not success:
        log_print("  -> GAGAL TOTAL -> Membuat file JSON kosong []")
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("[]")

def load_json(path):
    """Load JSON file aman ke dalam Pandas DataFrame"""
    try:
        with open(path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
            if isinstance(data, list):
                return pd.json_normalize(data)
            if isinstance(data, dict):
                for k in ['data', 'items', 'results']:
                    if k in data and isinstance(data[k], list):
                        return pd.json_normalize(data[k])
            return pd.json_normalize(data)
    except:
        return pd.DataFrame()

# ======================================================
# FUNGSI UTAMA PROSES PER TAHUN
# ======================================================
def process_tahun(tahun):
    data_dir = os.path.join(BASE_DIR, 'data', str(tahun))
    output_json = os.path.join(data_dir, f'rekap_pengadaan_{tahun}.json')
    
    # LOGIKA SKIP UTK TAHUN n-2 YANG SUDAH FINAL
    if tahun == tahun_n2 and os.path.exists(output_json):
        log_print(f"\n[SKIP] Tahun {tahun} sudah final -> Lewati download & generate")
        return None

    log_print(f'\n{"="*55}')
    log_print(f'   PROSES DATA TAHUN {tahun}')
    log_print(f'{"="*55}')

    os.makedirs(data_dir, exist_ok=True)
    is_n2 = (tahun == tahun_n2)

    # ----------------======================================
    # PHASE 1: DOWNLOAD DATA DARI API INAPROC
    # ----------------======================================
    log_print(f"Memulai proses download API untuk Tahun {tahun}...")
    for endpoint in ENDPOINTS:
        url = f"https://data.inaproc.id/api/legacy/{endpoint}?kode_klpd=D228&tahun={tahun}"
        base_name = endpoint.replace('/', '_')
        filename = f"Legacy_{base_name}_{tahun}.json"
        output_file = os.path.join(data_dir, filename)

        # Skip download per file jika tahun n-2 sudah punya file mentahnya
        if is_n2 and os.path.exists(output_file):
            log_print(f"SKIP Download (Sudah Final Lokal): {filename}")
            continue

        log_print(f"DOWNLOAD: {url}")
        download_with_retry(url, output_file)

    # ----------------======================================
    # PHASE 2: PROCESSING & COMPILING DATA
    # ------------------------------------------------======
    log_print(f"\nLoading semua file JSON hasil unduhan...")
    def p(nama): return os.path.join(data_dir, f'Legacy_{nama}_{tahun}.json')
    
    df1     = load_json(p('rup_paket-penyedia-terumumkan'))
    df1_2   = load_json(p('rup_paket-swakelola-terumumkan'))
    df2     = load_json(p('tender_non-tender-pengumuman'))
    df2_1   = load_json(p('tender_non-tender-selesai'))
    df2_2   = load_json(p('tender_non-tender-ekontrak-sppbj'))
    df2_3   = load_json(p('tender_non-tender-ekontrak-kontrak'))
    df2_4   = load_json(p('tender_non-tender-ekontrak-spmkspp'))
    df2_5   = load_json(p('tender_non-tender-ekontrak-bapbast'))
    df3     = load_json(p('tender_pencatatan-non-tender'))
    df4     = load_json(p('tender_pencatatan-swakelola'))
    df5     = load_json(p('tender_pengumuman'))
    df5_1   = load_json(p('tender_tender-selesai'))
    df5_1_1 = load_json(p('tender_tender-selesai-nilai'))
    df5_2   = load_json(p('tender_tender-ekontrak-sppbj'))
    df5_3   = load_json(p('tender_tender-ekontrak-kontrak'))
    df5_4   = load_json(p('tender_tender-ekontrak-spmkspp'))
    df5_5   = load_json(p('tender_tender-ekontrak-bapbast'))
    df6     = load_json(p('ekatalog_paket-e-purchasing'))
    df7     = load_json(p('ekatalog-archive_paket-e-purchasing'))

    # ======================================================
    # FILTER DATA JANGAN AMBIL YANG GAGAL / DIBATALKAN
    # ======================================================
    if not df2.empty and 'status_nontender' in df2.columns:
        df2 = df2[df2['status_nontender'] != 'Gagal/Batal']

    if not df5.empty and 'status_tender' in df5.columns:
        df5 = df5[df5['status_tender'] != 'Gagal/Batal']

    # FIX REKUES USER: Sumber 3 Paket Dibatalkan dibuang (tidak diambil)
    if not df3.empty and 'status_nontender_pct_ket' in df3.columns:
        df3 = df3[df3['status_nontender_pct_ket'].astype(str).str.strip() != 'Paket Dibatalkan']

    # ======================================================
    # LOGIKA RELASI DAN MAPPING DATA PBJ
    # ======================================================
    def get_set(df, col):
        if df.empty or col not in df.columns: return set()
        return set(df[col].astype(str).str.split(';').explode().str.strip())

    set_selesai   = get_set(df2_1, 'kd_nontender')
    set_sppbj     = get_set(df2_2, 'kd_nontender')
    set_kontrak   = get_set(df2_3, 'kd_nontender')
    set_spmkspp   = get_set(df2_4, 'kd_nontender')
    set_bapbast   = get_set(df2_5, 'kd_nontender')

    set_t_selesai = get_set(df5_1, 'kd_tender')
    set_t_sppbj   = get_set(df5_2, 'kd_tender')
    set_t_kontrak = get_set(df5_3, 'kd_tender')
    set_t_spmkspp = get_set(df5_4, 'kd_tender')
    set_t_bapbast = get_set(df5_5, 'kd_tender')

    # Map Nilai Kontrak & Komponen Finansial
    def build_multi_kd_map(df, kd_col, val_col):
        m = {}
        if not df.empty and val_col in df.columns:
            for _, r in df.iterrows():
                for k in str(r.get(kd_col, '')).split(';'):
                    if k.strip(): m[k.strip()] = r.get(val_col)
        return m

    map_nt_kontrak = build_multi_kd_map(df2_1, 'kd_nontender', 'nilai_kontrak')
    map_nt_pdn     = build_multi_kd_map(df2_1, 'kd_nontender', 'nilai_pdn_kontrak')
    map_nt_umk     = build_multi_kd_map(df2_1, 'kd_nontender', 'nilai_umk_kontrak')
    map_t_kontrak  = build_multi_kd_map(df5_1_1, 'kd_tender', 'nilai_kontrak')
    map_t_pdn      = build_multi_kd_map(df5_3, 'kd_tender', 'nilai_pdn_kontrak')
    map_t_umk      = build_multi_kd_map(df5_3, 'kd_tender', 'nilai_umk_kontrak')
    
    map_nt_tgl_kontrak = build_multi_kd_map(df2_3, 'kd_nontender', 'tgl_kontrak')
    map_t_tgl_kontrak  = build_multi_kd_map(df5_3, 'kd_tender', 'tgl_kontrak')
    map_nt_penyedia    = build_multi_kd_map(df2_3, 'kd_nontender', 'nama_penyedia')
    map_t_penyedia     = build_multi_kd_map(df5_3, 'kd_tender', 'nama_penyedia')

    # Load Kamus Kamus Penyedia Offline
    path_kamus = os.path.join(BASE_DIR, 'data', 'master', 'kamus_penyedia.json')
    map_offline_penyedia = {}
    if os.path.exists(path_kamus):
        try:
            with open(path_kamus, 'r', encoding='utf-8') as f:
                kamus_list = json.load(f)
                if isinstance(kamus_list, list):
                    for item in kamus_list:
                        nama = item.get('nama_penyedia', "")
                        if item.get('kode_penyedia'): map_offline_penyedia[str(item['kode_penyedia'])] = nama
                        if item.get('kd_penyedia'): map_offline_penyedia[str(item['kd_penyedia'])] = nama
        except: pass

    # Standardize RUP KD Function
    def split_kd_list(x):
        return [int(i.strip()) for i in str(x).split(';') if i.strip().isdigit()]

    def standardize_kd_rup(df, col):
        if df.empty or col not in df.columns: return df
        df[col+'_raw'] = df[col]
        df[col+'_list'] = df[col].apply(split_kd_list)
        df[col] = df[col].apply(lambda x: split_kd_list(x)[0] if len(split_kd_list(x))>0 else None)
        return df

    for d, c in [(df1,'kd_rup'), (df1_2,'kd_rup'), (df2,'kd_rup'), (df3,'kd_rup'), (df4,'kd_rup'), (df5,'kd_rup'), (df6,'rup_code'), (df7,'kd_rup')]:
        standardize_kd_rup(d, c)

    map_pagu_s1   = df1.set_index('kd_rup')['pagu'] if not df1.empty else {}
    map_pagu_s1_2 = df1_2.set_index('kd_rup')['pagu'] if not df1_2.empty else {}

    def get_pagu_multi(kd_list, tipe='s1'):
        if not isinstance(kd_list, list): return None
        total = sum(map_pagu_s1.get(k, 0) if tipe=='s1' else map_pagu_s1_2.get(k, 0) for k in kd_list)
        return total if total != 0 else None

    if not df1.empty:
        df1_clean = df1.dropna(subset=['kd_rup']).copy()
        df1_clean['kd_rup'] = df1_clean['kd_rup'].apply(lambda x: int(float(str(x).strip())))
        df1_map = df1_clean.set_index('kd_rup')
    else:
        df1_map = pd.DataFrame()

    def get_s1(kd, col):
        try:
            if pd.isna(kd) or df1_map.empty: return None
            kd_match = int(float(str(kd).strip()))
            if kd_match in df1_map.index:
                val = df1_map.loc[kd_match, col]
                return None if pd.isna(val) else val
            return None
        except: return None

    # Processing SUMBER 2 (Non Tender)
    data_s2=[]
    for _, r in df2.iterrows():
        kd = r.get('kd_rup')
        kd_list = r.get('kd_rup_list')
        kd_nt_list = [i.strip() for i in str(r.get('kd_nontender', '')).split(';')] if pd.notna(r.get('kd_nontender')) else []
        status = r.get('status_nontender')
        for k in kd_nt_list:
            if k in set_bapbast: status='BAPBAST'; break
            elif k in set_spmkspp: status='SPMKSPP'; break
            elif k in set_kontrak: status='Kontrak'; break
            elif k in set_sppbj: status='SPPBJ'; break
            elif k in set_selesai: status='Non Tender Selesai'; break

        nilai_hasil = next((map_nt_kontrak[k] for k in kd_nt_list if k in map_nt_kontrak), "N/A")
        nilai_pdn = next((map_nt_pdn[k] for k in kd_nt_list if k in map_nt_pdn), "N/A")
        nilai_umk = next((map_nt_umk[k] for k in kd_nt_list if k in map_nt_umk), "N/A")

        data_s2.append({
            'Kode RUP': r.get('kd_rup_raw'), 'Satuan Kerja': r.get('nama_satker'), 'Nama Paket': r.get('nama_paket'),
            'Metode Pengadaan': r.get('mtd_pemilihan'), 'Jenis Pengadaan': r.get('jenis_pengadaan'), 'Sumber Dana': r.get('sumber_dana'),
            'PDN': get_s1(kd, 'status_pdn'), 'UKM': get_s1(kd, 'status_ukm'), 'Nilai Pagu RUP': get_pagu_multi(kd_list, 's1'),
            'Nilai Hasil Pemilihan': nilai_hasil, 'Tanggal Kontrak': format_tgl(next((map_nt_tgl_kontrak[k] for k in kd_nt_list if k in map_nt_tgl_kontrak), "")),
            'Nama Penyedia': next((map_nt_penyedia[k] for k in kd_nt_list if k in map_nt_penyedia), ""), 'Status': status,
            'Kode Paket': r.get('kd_nontender'), 'Nilai HPS': r.get('hps'), 'Nilai PDN': nilai_pdn, 'Nilai UMK': nilai_umk,
            'Metode': 'Non Tender', 'Sumber': 'Sumber 2'
        })
    df_s2 = pd.DataFrame(data_s2)

    # Processing SUMBER 3 (Pencatatan Non Tender)
    data_s3=[]
    for _, r in df3.iterrows():
        kd = r.get('kd_rup')
        data_s3.append({
            'Kode RUP': r.get('kd_rup_raw'), 'Satuan Kerja': r.get('nama_satker'), 'Nama Paket': r.get('nama_paket'),
            'Metode Pengadaan': r.get('mtd_pemilihan'), 'Jenis Pengadaan': r.get('kategori_pengadaan'), 'Sumber Dana': r.get('sumber_dana'),
            'PDN': get_s1(kd, 'status_pdn'), 'UKM': get_s1(kd, 'status_ukm'), 'Nilai Pagu RUP': get_pagu_multi(r.get('kd_rup_list'), 's1'),
            'Nilai Hasil Pemilihan': "" if pd.isna(r.get('total_realisasi')) else r.get('total_realisasi'), 'Tanggal Kontrak': format_tgl(r.get('tgl_selesai_paket', '')),
            'Nama Penyedia': "", 'Status': r.get('status_nontender_pct_ket'), 'Kode Paket': r.get('kd_nontender_pct'), 'Nilai HPS': pd.NA,
            'Nilai PDN': r.get('nilai_pdn_pct'), 'Nilai UMK': r.get('nilai_umk_pct'), 'Metode': 'Pencatatan Non Tender', 'Sumber': 'Sumber 3'
        })
    df_s3 = pd.DataFrame(data_s3)

    # Processing SUMBER 4 (Pencatatan Swakelola)
    data_s4=[]
    swakelola_map = df1_2.set_index('kd_rup')['tipe_swakelola'] if not df1_2.empty else {}
    for _, r in df4.iterrows():
        kd = r.get('kd_rup')
        jenis = f"Swakelola {int(swakelola_map[kd])}" if kd in swakelola_map else "N/A"
        data_s4.append({
            'Kode RUP': r.get('kd_rup_raw'), 'Satuan Kerja': r.get('nama_satker'), 'Nama Paket': r.get('nama_paket'),
            'Metode Pengadaan': 'Swakelola', 'Jenis Pengadaan': jenis, 'Sumber Dana': r.get('sumber_dana'),
            'PDN': "PDN" if r.get('nilai_pdn_pct', 0)!=0 else "Tidak", 'UKM': "UKM" if r.get('nilai_umk_pct', 0)!=0 else "Tidak",
            'Nilai Pagu RUP': get_pagu_multi(r.get('kd_rup_list'), 's1_2'), 'Nilai Hasil Pemilihan': "" if pd.isna(r.get('total_realisasi')) else r.get('total_realisasi'),
            'Tanggal Kontrak': format_tgl(r.get('tgl_selesai_paket', '')), 'Nama Penyedia': "", 'Status': r.get('status_swakelola_pct_ket'),
            'Kode Paket': r.get('kd_swakelola_pct'), 'Nilai HPS': pd.NA, 'Nilai PDN': r.get('nilai_pdn_pct'), 'Nilai UMK': r.get('nilai_umk_pct'),
            'Metode': 'Pencatatan Swakelola', 'Sumber': 'Sumber 4'
        })
    df_s4 = pd.DataFrame(data_s4)

    # Tambahan Swakelola Murni (Sumber 1_2) yang belum tercatat di Sumber 4
    data_s1_2=[]
    set_all_s4 = get_set(df4, 'kd_rup_raw')
    for _, r in df1_2.iterrows():
        kd = r.get('kd_rup')
        if str(kd) not in set_all_s4:
            jenis = f"Swakelola {int(swakelola_map[kd])}" if kd in swakelola_map else "N/A"
            data_s1_2.append({
                'Kode RUP': kd, 'Satuan Kerja': r.get('nama_satker'), 'Nama Paket': r.get('nama_paket'), 'Metode Pengadaan': 'Swakelola',
                'Jenis Pengadaan': jenis, 'Sumber Dana': None, 'PDN': None, 'UKM': None, 'Nilai Pagu RUP': r.get('pagu'), 'Nilai Hasil Pemilihan': "",
                'Tanggal Kontrak': "", 'Nama Penyedia': "", 'Status': 'Pengumuman RUP', 'Kode Paket': pd.NA, 'Nilai HPS': pd.NA, 'Nilai PDN': pd.NA, 'Nilai UMK': pd.NA,
                'Metode': 'Swakelola', 'Sumber': 'Sumber 1_2'
            })
    df_s1_2 = pd.DataFrame(data_s1_2)

    # Processing SUMBER 5 (Tender / Seleksi)
    data_s5=[]
    for _, r in df5.iterrows():
        kd = r.get('kd_rup')
        kd_t_list = [i.strip() for i in str(r.get('kd_tender', '')).split(';')] if pd.notna(r.get('kd_tender')) else []
        status = r.get('status_tender')
        for k in kd_t_list:
            if k in set_t_bapbast: status='BAPBAST'; break
            elif k in set_t_spmkspp: status='SPMKSPP'; break
            elif k in set_t_kontrak: status='Kontrak'; break
            elif k in set_t_sppbj: status='SPPBJ'; break
            elif k in set_t_selesai: status='Tender Selesai'; break

        nilai_hasil = next((map_t_kontrak[k] for k in kd_t_list if k in map_t_kontrak), "N/A")
        nilai_pdn = next((map_t_pdn[k] for k in kd_t_list if k in map_t_pdn), "N/A")
        nilai_umk = next((map_t_umk[k] for k in kd_t_list if k in map_t_umk), "N/A")

        mtd = str(r.get('mtd_pemilihan', '')).strip().lower()
        kat_metode = 'Seleksi' if mtd == 'seleksi' else ('Tender Cepat' if mtd == 'tender cepat' else 'Tender')

        data_s5.append({
            'Kode RUP': r.get('kd_rup_raw'), 'Satuan Kerja': r.get('nama_satker'), 'Nama Paket': r.get('nama_paket'),
            'Metode Pengadaan': r.get('mtd_pemilihan'), 'Jenis Pengadaan': r.get('jenis_pengadaan'), 'Sumber Dana': r.get('sumber_dana'),
            'PDN': get_s1(kd, 'status_pdn'), 'UKM': get_s1(kd, 'status_ukm'), 'Nilai Pagu RUP': get_pagu_multi(r.get('kd_rup_list'), 's1'),
            'Nilai Hasil Pemilihan': nilai_hasil, 'Tanggal Kontrak': format_tgl(next((map_t_tgl_kontrak[k] for k in kd_t_list if k in map_t_tgl_kontrak), "")),
            'Nama Penyedia': next((map_t_penyedia[k] for k in kd_t_list if k in map_t_penyedia), ""), 'Status': status,
            'Kode Paket': r.get('kd_tender'), 'Nilai HPS': r.get('hps'), 'Nilai PDN': nilai_pdn, 'Nilai UMK': nilai_umk,
            'Metode': kat_metode, 'Sumber': 'Sumber 5'
        })
    df_s5 = pd.DataFrame(data_s5)

    # Processing SUMBER 6 (E-Purchasing V6)
    data_s6=[]
    for _, r in df6.iterrows():
        kd = r.get('rup_code')
        nilai_hasil = r.get('total')
        if pd.isna(nilai_hasil): nilai_hasil=""
        try:
            kd_match = int(float(str(kd).strip()))
            is_match = not df1_map.empty and kd_match in df1_map.index
        except: is_match = False

        if is_match:
            nilai_pdn_val = nilai_hasil if str(df1_map.loc[kd_match, 'status_pdn']).strip().upper() == 'PDN' else 0
            nilai_umk_val = nilai_hasil if str(df1_map.loc[kd_match, 'status_ukm']).strip().upper() == 'UKM' else 0
        else: nilai_pdn_val, nilai_umk_val = "N/A", "N/A"

        kode_p = str(r.get('kode_penyedia', ""))
        nama_p = map_offline_penyedia.get(kode_p, kode_p) if kode_p and kode_p != "None" else ""

        data_s6.append({
            'Kode RUP': r.get('rup_code_raw'), 'Satuan Kerja': r.get('nama_satker'), 'Nama Paket': r.get('rup_name'),
            'Metode Pengadaan': 'E-Purchasing', 'Jenis Pengadaan': get_s1(kd, 'jenis_pengadaan'), 'Sumber Dana': r.get('funding_source'),
            'PDN': get_s1(kd, 'status_pdn'), 'UKM': get_s1(kd, 'status_ukm'), 'Nilai Pagu RUP': get_pagu_multi(r.get('rup_code_list'), 's1'),
            'Nilai Hasil Pemilihan': nilai_hasil, 'Tanggal Kontrak': "", 'Nama Penyedia': nama_p, 'Status': r.get('status'),
            'Kode Paket': r.get('order_id'), 'Nilai HPS': pd.NA, 'Nilai PDN': nilai_pdn_val, 'Nilai UMK': nilai_umk_val,
            'Metode': 'E-Purchasing V6', 'Sumber': 'Sumber 6'
        })
    df_s6 = pd.DataFrame(data_s6)

    # Processing SUMBER 7 (E-Purchasing V5 Archive)
    data_s7=[]
    for _, r in df7.iterrows():
        kd = r.get('kd_rup')
        nilai_hasil = r.get('total_harga')
        if pd.isna(nilai_hasil): nilai_hasil=""
        try:
            kd_match = int(float(str(kd).strip()))
            is_match = not df1_map.empty and kd_match in df1_map.index
        except: is_match = False

        if is_match:
            nilai_pdn_val = nilai_hasil if str(df1_map.loc[kd_match, 'status_pdn']).strip().upper() == 'PDN' else 0
            nilai_umk_val = nilai_hasil if str(df1_map.loc[kd_match, 'status_ukm']).strip().upper() == 'UKM' else 0
        else: nilai_pdn_val, nilai_umk_val = "N/A", "N/A"

        kode_p = str(r.get('kd_penyedia', ""))
        nama_p = map_offline_penyedia.get(kode_p, kode_p) if kode_p and kode_p != "None" else ""

        data_s7.append({
            'Kode RUP': r.get('kd_rup_raw'), 'Satuan Kerja': r.get('nama_satker') if pd.notna(r.get('nama_satker')) else get_s1(kd, 'nama_satker'),
            'Nama Paket': r.get('nama_paket'), 'Metode Pengadaan': 'E-Purchasing', 'Jenis Pengadaan': get_s1(kd, 'jenis_pengadaan'), 'Sumber Dana': r.get('nama_sumber_dana'),
            'PDN': get_s1(kd, 'status_pdn'), 'UKM': get_s1(kd, 'status_ukm'), 'Nilai Pagu RUP': get_pagu_multi(r.get('kd_rup_list'), 's1'),
            'Nilai Hasil Pemilihan': nilai_hasil, 'Tanggal Kontrak': "", 'Nama Penyedia': nama_p, 'Status': r.get('paket_status_str'),
            'Kode Paket': r.get('kd_paket'), 'Nilai HPS': pd.NA, 'Nilai PDN': nilai_pdn_val, 'Nilai UMK': nilai_umk_val,
            'Metode': 'E-Purchasing V5', 'Sumber': 'Sumber 7'
        })
    df_s7 = pd.DataFrame(data_s7)

    # Swakelola/Penyedia Murni (Sumber 1) yang sama sekali belum dieksekusi proses pengadaannya
    set_all_executed = get_set(df2, 'kd_rup_raw').union(get_set(df3, 'kd_rup_raw')).union(get_set(df4, 'kd_rup_raw'))\
                      .union(get_set(df5, 'kd_rup_raw')).union(get_set(df6, 'rup_code_raw')).union(get_set(df7, 'kd_rup_raw'))
    
    data_s1=[]
    for _, r in df1.iterrows():
        kd = r.get('kd_rup')
        if str(kd) not in set_all_executed:
            data_s1.append({
                'Kode RUP': kd, 'Satuan Kerja': r.get('nama_satker'), 'Nama Paket': r.get('nama_paket'), 'Metode Pengadaan': r.get('metode_pengadaan'),
                'Jenis Pengadaan': r.get('jenis_pengadaan'), 'Sumber Dana': None, 'PDN': 'PDN' if r.get('status_pdn')=='PDN' else 'Non-PDN',
                'UKM': 'UKM' if r.get('status_ukm')=='UKM' else 'Non-UKM', 'Nilai Pagu RUP': r.get('pagu'), 'Nilai Hasil Pemilihan': "",
                'Tanggal Kontrak': "", 'Nama Penyedia': "", 'Status': 'Belum Proses', 'Kode Paket': pd.NA, 'Nilai HPS': pd.NA, 'Nilai PDN': pd.NA, 'Nilai UMK': pd.NA,
                'Metode': r.get('metode_pengadaan'), 'Sumber': 'Sumber 1'
            })
    df_s1 = pd.DataFrame(data_s1)

    # Gabungkan semua sumber data secara aman tanpa duplikasi kronis
    final_df = pd.concat([df_s2, df_s3, df_s4, df_s1_2, df_s5, df_s6, df_s7, df_s1], ignore_index=True)
    final_df = final_df.map(lambda x: re.sub(r'[\x00-\x1F]', '', str(x)) if isinstance(x, str) else x)

    cols = [
        'Kode RUP', 'Satuan Kerja', 'Nama Paket', 'Metode Pengadaan', 'Jenis Pengadaan',
        'Sumber Dana', 'PDN', 'UKM', 'Nilai Pagu RUP', 'Nilai Hasil Pemilihan',
        'Tanggal Kontrak', 'Nama Penyedia', 'Status', 'Kode Paket', 'Nilai HPS',
        'Nilai PDN', 'Nilai UMK', 'Metode', 'Sumber'
    ]
    final_df = final_df[cols]
    final_df['PDN'] = final_df['PDN'].fillna("N/A")
    final_df['UKM'] = final_df['UKM'].fillna("N/A")
    final_df = final_df.fillna("")

    # SIMPAN REKAP FORMAT JSON
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(final_df.to_dict(orient='records'), f, ensure_ascii=False, indent=2)
    log_print(f"JSON Rekap sukses dibuat: {output_json}")

    # ----------------======================================
    # PHASE 3: GENERATE EXCEL MAESTRO VIA OPENPYXL
    # ----------------======================================
    tahun_label = str(df1['tahun_anggaran'].iloc[0]) if not df1.empty and 'tahun_anggaran' in df1.columns else str(tahun)
    tgl_now = datetime.now().strftime('%Y-%m-%d')
    nama_file_excel = f'Paket Pengadaan Tahun {tahun_label} ({tgl_now}) (Api Gateway Legacy).xlsx'
    
    output_dir_excel = os.path.join(BASE_DIR, 'output', 'pengadaan', str(tahun))
    os.makedirs(output_dir_excel, exist_ok=True)
    output_excel_path = os.path.join(output_dir_excel, nama_file_excel)

    kolom_angka = ['Nilai Pagu RUP', 'Nilai Hasil Pemilihan', 'Nilai HPS', 'Nilai PDN', 'Nilai UMK']
    excel_df = final_df.copy()

    for col in ['Nilai Pagu RUP', 'Nilai HPS']:
        if col in excel_df.columns: excel_df[col] = pd.to_numeric(excel_df[col], errors='coerce')

    def safe_numeric(val):
        if val in ("N/A", ""): return val
        try: return float(val)
        except: return val

    for col in ['Nilai Hasil Pemilihan', 'Nilai PDN', 'Nilai UMK']:
        if col in excel_df.columns: excel_df[col] = excel_df[col].apply(safe_numeric)

    excel_df.to_excel(output_excel_path, index=False, sheet_name='Pengadaan')

    # Styling Excel Menggunakan Aturan Identik LPSE
    wb = load_workbook(output_excel_path)
    ws = wb['Pengadaan']

    header_fill = PatternFill('solid', start_color='1F4E79')
    header_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    data_font = Font(name='Arial', size=10)
    
    fill_putih = PatternFill('solid', start_color='FFFFFF')
    fill_biru_muda = PatternFill('solid', start_color='DCE6F1')
    border_thin = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                         top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

    lebar_kolom = {
        'Kode RUP': 18, 'Satuan Kerja': 38, 'Nama Paket': 50, 'Metode Pengadaan': 22, 'Jenis Pengadaan': 32,
        'Sumber Dana': 14, 'PDN': 10, 'UKM': 10, 'Nilai Pagu RUP': 20, 'Nilai Hasil Pemilihan': 20,
        'Tanggal Kontrak': 18, 'Nama Penyedia': 35, 'Status': 22, 'Kode Paket': 20, 'Nilai HPS': 20,
        'Nilai PDN': 18, 'Nilai UMK': 18, 'Metode': 22, 'Sumber': 12
    }

    for i, col in enumerate(final_df.columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = lebar_kolom.get(col, 15)

    for cell in ws[1]:
        cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, header_align, border_thin
    ws.row_dimensions[1].height = 32

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
        fill = fill_putih if row_idx % 2 == 0 else fill_biru_muda
        for cell in row:
            col_name = final_df.columns[cell.column - 1]
            cell.font, cell.fill, cell.border = data_font, fill, border_thin
            if col_name in kolom_angka:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col_name in ('PDN', 'UKM', 'Sumber Dana', 'Metode', 'Sumber', 'Status'):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(vertical='center', wrap_text=False)

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    wb.save(output_excel_path)
    log_print(f"Excel Berwarna Berhasil Terformat: {output_excel_path}")

    # Kopi Hasil ke folder web terintegrasi
    master_excel = os.path.join(data_dir, f'master_pengadaan_{tahun}.xlsx')
    shutil.copy2(output_excel_path, master_excel)
    
    log_print(f'SELESAI GENERATE TAHUN {tahun} | Total paket data: {len(final_df)}')
    return len(final_df)

# ======================================================
# PROGRAM BERJALAN UTAMA (__main__)
# ======================================================
if __name__ == '__main__':
    log_print("\n" + "="*55)
    log_print(f"START {get_waktu_indonesia()}")
    log_print(f"AUTOMATION SYSTEM PENGADAAN BARANG DAN JASA")
    log_print(f"Target Sinkronisasi: {daftar_tahun}")
    log_print("="*55)

    total_seluruh_paket = 0
    for t in daftar_tahun:
        hasil_proses = process_tahun(t)
        if hasil_proses:
            total_seluruh_paket += hasil_proses

    # ----------------------------------------------------
    # UPDATE FILE LAST-UPDATE UNTUK WEBSITE
    # ----------------------------------------------------
    with open(os.path.join(BASE_DIR, 'data', 'last-update-pengadaan.txt'), 'w', encoding='utf-8') as f:
        f.write(get_waktu_indonesia())

    log_print("\n" + "="*55)
    log_print(f"SINKRONISASI PBJ ALL-TAHUN SELESAI PADA {get_waktu_indonesia()} | Total: {total_seluruh_paket} Paket")
    log_print("="*55)