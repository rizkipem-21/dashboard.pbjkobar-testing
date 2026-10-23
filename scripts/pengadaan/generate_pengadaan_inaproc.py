# ======================================================
# file: generate_pengadaan_inaproc.py
# ======================================================

import os
import json
import re
import shutil
import warnings
import pandas as pd
import sys
import subprocess
import time
import requests
import urllib3
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Mengabaikan peringatan SSL jika memakai verify=False
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
import config_rahasia

warnings.filterwarnings('ignore')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

tahun_n      = datetime.now().year       
tahun_n1     = tahun_n - 1                
tahun_n2     = tahun_n - 2                
daftar_tahun = [tahun_n, tahun_n1, tahun_n2] 

LOG_FILE = os.path.join(BASE_DIR, 'tools', 'log_pengadaan.txt')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

def log_print(*args, **kwargs):
    msg = " ".join(str(a) for a in args)
    print(msg, **kwargs)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(msg + '\n')

def get_waktu_indonesia():
    tz_wib = timezone(timedelta(hours=7))
    sekarang = datetime.now(tz_wib)
    bulan_indo = {1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April', 5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'}
    return f"{sekarang.day} {bulan_indo[sekarang.month]} {sekarang.year} | {sekarang.strftime('%H.%M')} WIB"

def sync_to_github():
    log_print("\n==================================================")
    log_print("MENGIRIM DATA PENGADAAN KE GITHUB DARI PYTHON...")
    log_print("==================================================")

    waktu_sekarang = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        subprocess.run(["git", "config", "user.name", "rizkipem-21"], cwd=BASE_DIR)
        subprocess.run(["git", "config", "user.email", "rizki.pem@gmail.com"], cwd=BASE_DIR)
        subprocess.run(["git", "add", "."], capture_output=True, text=True, cwd=BASE_DIR)
        
        commit_msg = f"Auto update Pengadaan Inaproc {waktu_sekarang}"
        subprocess.run(["git", "commit", "-m", commit_msg], capture_output=True, text=True, cwd=BASE_DIR)
        
        res_push = subprocess.run(["git", "push"], capture_output=True, text=True, cwd=BASE_DIR)
        if res_push.returncode == 0:
            log_print("✅ Push ke GitHub BERHASIL")
            return True, "✅ Push ke GitHub BERHASIL"
        else:
            error_git = res_push.stderr.strip()
            log_print(f"❌ Push ke GitHub GAGAL: {error_git}")
            return False, f"❌ Push ke GitHub GAGAL:\n`{error_git}`"
    except Exception as e:
        log_print(f"❌ Terjadi kesalahan pada eksekusi Git: {str(e)}")
        return False, f"❌ Terjadi kesalahan pada eksekusi Git:\n`{str(e)}`"

def format_tgl(val):
    if not val or (not isinstance(val, str) and pd.isna(val)): return ""
    try:
        parts = str(val).strip()[:10].split('-')
        if len(parts) == 3: return f"{parts[2]}-{parts[1]}-{parts[0]}"
        return ""
    except: return ""

def get_file_path(data_dir, base_name, tahun):
    v1_path = os.path.join(data_dir, f"v1_{base_name}_{tahun}.json")
    if os.path.exists(v1_path): return v1_path
    return os.path.join(data_dir, f"Legacy_{base_name}_{tahun}.json")

def load_json(path):
    try:
        with open(path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
            if isinstance(data, list): 
                return pd.json_normalize(data) if data else pd.DataFrame()
            if isinstance(data, dict):
                for k in ['data', 'items', 'results']:
                    if k in data and isinstance(data[k], list): 
                        return pd.json_normalize(data[k])
                return pd.DataFrame() 
    except Exception as e:
        return pd.DataFrame()

def kelola_arsip_bulanan(folder_path, tahun):
    if not os.path.exists(folder_path): return
    folder_arsip_lokal = os.path.join(BASE_DIR, 'arsip_lokal', 'pengadaan', str(tahun))
    os.makedirs(folder_arsip_lokal, exist_ok=True)
    
    file_excel = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('~$')]
    arsip_bulanan = {}
    
    for f in file_excel:
        match = re.search(r"\((\d{4}-\d{2}-\d{2})\)", f)
        if match:
            bulan = match.group(1)[:7]
            if bulan not in arsip_bulanan: arsip_bulanan[bulan] = []
            arsip_bulanan[bulan].append((match.group(1), f))
            
    for bulan, list_file in arsip_bulanan.items():
        list_file.sort(key=lambda x: x[0])
        for tgl, nama_file in list_file[:-1]:
            path_sumber = os.path.join(folder_path, nama_file)
            path_tujuan = os.path.join(folder_arsip_lokal, nama_file)
            
            berhasil_pindah = False
            while not berhasil_pindah:
                try: 
                    shutil.move(path_sumber, path_tujuan)
                    log_print(f"📦 Arsip harian diamankan ke lokal: {nama_file}")
                    berhasil_pindah = True
                except PermissionError:
                    log_print(f"⚠️ Akses Ditolak: File lama {nama_file} sedang terbuka.")
                    tutup_sukses = False
                    try:
                        import win32com.client
                        excel_app = win32com.client.GetActiveObject("Excel.Application")
                        for wb_app in excel_app.Workbooks:
                            if wb_app.Name == nama_file:
                                wb_app.Close(SaveChanges=False)
                                tutup_sukses = True
                                break
                    except Exception: pass
                    
                    if not tutup_sukses:
                        os.system(f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""File lama {nama_file} sedang terbuka. Tutup file di Excel agar bisa diarsipkan!"", 5, ""Peringatan Excel"", 48)(window.close)")')
                    time.sleep(3)
                except Exception as e: 
                    berhasil_pindah = True

def update_daftar_arsip_json(folder_path):
    if not os.path.exists(folder_path): return
    file_excel = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    file_excel.sort(reverse=True) 
    arsip_list = [{"nama_file": f} for f in file_excel]
    path_json = os.path.join(folder_path, 'daftar_arsip.json')
    try:
        with open(path_json, 'w', encoding='utf-8') as f:
            json.dump(arsip_list, f, indent=4)
    except: pass

def get_kategori_status(sumber, s):
    sumber, s = str(sumber).lower(), str(s).lower()
    if "sumber 1" in sumber: return "Belum Proses"
    if "sumber 2" in sumber or "sumber 5" in sumber: return "Sudah Selesai" if "bapbast" in s else "Sedang Berjalan"
    if "sumber 3" in sumber or "sumber 4" in sumber: return "Sudah Selesai" if "paket selesai" in s or "selesai" in s else "Sedang Berjalan"
    if "sumber 6" in sumber: return "Sudah Selesai" if "payment_outside" in s or "completed" in s else "Sedang Berjalan"
    if "sumber 7" in sumber: return "Sudah Selesai" if "paket selesai" in s or "selesai" in s else "Sedang Berjalan"
    if "bapbast" in s or "payment" in s or "completed" in s: return "Sudah Selesai"
    if "pengumuman rup" in s or s == "" or s == "-": return "Belum Proses"
    return "Sedang Berjalan"

def fetch_inaproc_data(tahun):
    url = f"https://data.inaproc.id/api/v1/dashboard/realisasi/table?tahun={tahun}&jenis=4&instansi=D228&limit=1000"
    try:
        res = requests.get(url, timeout=45, verify=False)
        if res.status_code == 200:
            return res.json().get('data', [])
        else:
            log_print(f"❌ API Inaproc Error {res.status_code}: {res.text}")
            return []
    except Exception as e:
        log_print(f"❌ Exception API Inaproc: {e}")
        return []

def process_tahun(tahun):
    data_dir = os.path.join(BASE_DIR, 'data', str(tahun))
    output_dir_excel = os.path.join(BASE_DIR, 'output', 'pengadaan', str(tahun))
    output_json = os.path.join(data_dir, f'rekap_pengadaan_{tahun}.json')
        
    if tahun == tahun_n2 and os.path.exists(output_dir_excel) and os.path.exists(output_json):
        if any(f.startswith('Paket Pengadaan Tahun') and f.endswith('.xlsx') for f in os.listdir(output_dir_excel)):
            log_print(f"\n[SKIP] Data Tahun {tahun} sudah lengkap & final -> Lewati generate")
            return None

    log_print(f'\n{"="*55}\n   GENERATE DATA TAHUN {tahun} (BASE INAPROC)\n{"="*55}')
    
    # 1. TARIK DATA DASAR API INAPROC
    log_print("🔄 Mengambil data Inaproc API...")
    data_inaproc = fetch_inaproc_data(tahun)
    if not data_inaproc:
        log_print(f"⚠️ Data API Inaproc kosong/gagal untuk tahun {tahun}. Skip tahun ini.")
        return 0

    log_print("🔄 Memuat file JSON pendukung...")
    def p(nama): return get_file_path(data_dir, nama, tahun)
    
    df1     = load_json(p('rup_paket-penyedia-terumumkan'))
    df1_2   = load_json(p('rup_paket-swakelola-terumumkan'))
    df1_3   = load_json(p('rup_paket-penyedia'))
    df1_4   = load_json(p('rup_paket-swakelola'))
    df1_5   = load_json(p('rup_history-kaji-ulang'))
    df_ang_p = load_json(p('rup_paket-anggaran-penyedia'))
    df_ang_s = load_json(p('rup_paket-anggaran-swakelola'))
    df_prog  = load_json(p('rup_program-master'))
    df_keg   = load_json(p('rup_kegiatan-master'))
    df_sub   = load_json(p('rup_sub-kegiatan-master'))
     
    df2     = load_json(p('tender_non-tender-pengumuman'))
    df2_1   = load_json(p('tender_non-tender-selesai'))
    df2_2   = load_json(p('tender_non-tender-ekontrak-sppbj'))
    df2_3   = load_json(p('tender_non-tender-ekontrak-kontrak'))
    df2_4   = load_json(p('tender_non-tender-ekontrak-spmkspp'))
    df2_5   = load_json(p('tender_non-tender-ekontrak-bapbast'))
    df3     = load_json(p('tender_pencatatan-non-tender'))
    df3_1   = load_json(p('tender_pencatatan-non-tender-realisasi'))
    df4     = load_json(p('tender_pencatatan-swakelola'))
    df4_1   = load_json(p('tender_pencatatan-swakelola-realisasi'))
    df5     = load_json(p('tender_pengumuman'))
    df5_1   = load_json(p('tender_tender-selesai'))
    df5_1_1 = load_json(p('tender_tender-selesai-nilai'))
    df5_2   = load_json(p('tender_tender-ekontrak-sppbj'))
    df5_3   = load_json(p('tender_tender-ekontrak-kontrak'))
    df5_4   = load_json(p('tender_tender-ekontrak-spmkspp'))
    df5_5   = load_json(p('tender_tender-ekontrak-bapbast'))
    df_sikap_nt = load_json(p('sikap_penilaian-kinerja-penyedia-non-tender'))
    df_sikap_t  = load_json(p('sikap_penilaian-kinerja-penyedia-tender'))
    df6     = load_json(p('ekatalog_paket-e-purchasing'))
    df7     = load_json(p('ekatalog-archive_paket-e-purchasing'))
    df7_1   = load_json(p('ekatalog-archive_instansi-satker'))

    if not df2.empty and 'status_nontender' in df2.columns: df2 = df2[df2['status_nontender'] != 'Gagal/Batal']
    if not df5.empty and 'status_tender' in df5.columns: df5 = df5[df5['status_tender'] != 'Gagal/Batal']
    if not df3.empty and 'status_nontender_pct_ket' in df3.columns: df3 = df3[df3['status_nontender_pct_ket'].astype(str).str.strip() != 'Paket Dibatalkan']
    if not df6.empty and 'status' in df6.columns: df6 = df6[~df6['status'].isin(['CANCELLED_ON_NEGOTIATION', 'CANCELLED_ON_REVIEW', 'CANCELLED'])]

    def get_set(df, col): return set(df[col].astype(str).str.split(';').explode().str.strip()) if not df.empty and col in df.columns else set()
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

    def parse_sikap(df, key_col):
        res = {}
        if df.empty or key_col not in df.columns: return res
        for _, r in df.iterrows():
            for k in str(r.get(key_col, '')).split(';'):
                kd = k.strip()
                if not kd or kd == 'nan': continue
                if kd not in res:
                    res[kd] = {'kualitas': 0.0, 'biaya': 0.0, 'waktu': 0.0, 'layanan': 0.0, 'has_data': False}
                indikator = str(r.get('indikator_penilaian', '')).lower()
                try: nilai = float(r.get('nilai_indikator', 0))
                except: nilai = 0.0
                if "kualitas dan kuantitas" in indikator: res[kd]['kualitas'] = nilai; res[kd]['has_data'] = True
                elif "biaya" in indikator: res[kd]['biaya'] = nilai; res[kd]['has_data'] = True
                elif "waktu" in indikator: res[kd]['waktu'] = nilai; res[kd]['has_data'] = True
                elif "layanan" in indikator: res[kd]['layanan'] = nilai; res[kd]['has_data'] = True
        return res

    map_sikap_nt = parse_sikap(df_sikap_nt, 'kd_nontender')
    map_sikap_t  = parse_sikap(df_sikap_t, 'kd_tender')
    set_kinerja_nt = {k for k, v in map_sikap_nt.items() if v['has_data']}
    set_kinerja_t  = {k for k, v in map_sikap_t.items() if v['has_data']}

    def get_kinerja_vals(kd_list, map_sikap):
        if not kd_list: return "", "", "", "", "", ""
        for k in kd_list:
            if k in map_sikap and map_sikap[k]['has_data']:
                d = map_sikap[k]
                skor_angka = (d['kualitas'] * 0.30) + (d['biaya'] * 0.20) + (d['waktu'] * 0.30) + (d['layanan'] * 0.20)
                if skor_angka >= 3: kata = "Sangat Baik"
                elif skor_angka >= 2: kata = "Baik"
                elif skor_angka >= 1: kata = "Cukup"
                else: kata = "Buruk"
                return d['kualitas'], d['biaya'], d['waktu'], d['layanan'], round(skor_angka, 2), kata
        return "", "", "", "", "", ""

    def build_multi_kd_map(df, kd_col, val_col):
        m = {}
        if not df.empty and val_col in df.columns:
            for _, r in df.iterrows():
                for k in str(r.get(kd_col, '')).split(';'):
                    if k.strip(): m[k.strip()] = r.get(val_col)
        return m

    map_nt_kontrak = build_multi_kd_map(df2_1, 'kd_nontender', 'nilai_negosiasi')
    map_nt_pdn     = build_multi_kd_map(df2_1, 'kd_nontender', 'nilai_pdn_kontrak')
    map_nt_umk     = build_multi_kd_map(df2_1, 'kd_nontender', 'nilai_umk_kontrak')
    map_t_kontrak  = build_multi_kd_map(df5_1_1, 'kd_tender', 'nilai_negosiasi')
    map_t_pdn      = build_multi_kd_map(df5_3, 'kd_tender', 'nilai_pdn_kontrak')
    map_t_umk      = build_multi_kd_map(df5_3, 'kd_tender', 'nilai_umk_kontrak')
    
    map_nt_tgl_kontrak = build_multi_kd_map(df2_3, 'kd_nontender', 'tgl_kontrak')
    map_t_tgl_kontrak  = build_multi_kd_map(df5_3, 'kd_tender', 'tgl_kontrak')
    map_nt_penyedia    = build_multi_kd_map(df2_1, 'kd_nontender', 'nama_penyedia')
    map_nt_npwp        = build_multi_kd_map(df2_1, 'kd_nontender', 'npwp_penyedia')
    map_nt_alamat      = build_multi_kd_map(df2_1, 'kd_nontender', 'alamat_penyedia')
    map_t_penyedia     = build_multi_kd_map(df5_1_1, 'kd_tender', 'nama_penyedia')
    map_t_npwp         = build_multi_kd_map(df5_1_1, 'kd_tender', 'npwp_penyedia')
    map_t_alamat       = build_multi_kd_map(df5_1_1, 'kd_tender', 'alamat_penyedia')
    map_t_mak          = build_multi_kd_map(df5_1, 'kd_tender', 'mak')

    def get_mapped_val(kd_list, map_dict, is_date=False, is_numeric=False):
        if not kd_list: return ""
        for k in kd_list:
            if k in map_dict:
                val = map_dict[k]
                if pd.isna(val) or val is None or str(val).strip() == "" or str(val).lower() in ["nan", "none"]:
                    return "tidak ada data (null)"
                if is_date:
                    t = format_tgl(val)
                    return t if t else "tidak ada data (null)"
                if is_numeric:
                    try: return float(val)
                    except: return "tidak ada data (null)"
                return str(val).strip()
        return ""

    map_nt_nil_neg = build_multi_kd_map(df2_1, 'kd_nontender', 'nilai_negosiasi')
    map_nt_npwp16 = build_multi_kd_map(df2_1, 'kd_nontender', 'npwp16_penyedia')
    map_nt_kualifikasi = build_multi_kd_map(df2_1, 'kd_nontender', 'kualifikasi_paket')
    map_nt_no_sppbj = build_multi_kd_map(df2_2, 'kd_nontender', 'no_sppbj')
    map_nt_tgl_sppbj = build_multi_kd_map(df2_2, 'kd_nontender', 'tgl_sppbj')
    map_nt_no_kontrak = build_multi_kd_map(df2_3, 'kd_nontender', 'no_kontrak')
    map_nt_nil_kontrak = build_multi_kd_map(df2_3, 'kd_nontender', 'nilai_kontrak')
    map_nt_jenis_kontrak = build_multi_kd_map(df2_3, 'kd_nontender', 'jenis_kontrak')
    map_nt_status_kontrak = build_multi_kd_map(df2_3, 'kd_nontender', 'status_kontrak')
    map_nt_no_spmkspp = build_multi_kd_map(df2_4, 'kd_nontender', 'no_spmk_spp')
    map_nt_tgl_spmkspp = build_multi_kd_map(df2_4, 'kd_nontender', 'tgl_spmk_spp')
    map_nt_no_bast = build_multi_kd_map(df2_5, 'kd_nontender', 'no_bast')
    map_nt_tgl_bast = build_multi_kd_map(df2_5, 'kd_nontender', 'tgl_bast')
    map_nt_no_bap = build_multi_kd_map(df2_5, 'kd_nontender', 'no_bap')
    map_nt_tgl_bap = build_multi_kd_map(df2_5, 'kd_nontender', 'tgl_bap')

    map_t_nil_neg = build_multi_kd_map(df5_1_1, 'kd_tender', 'nilai_negosiasi')
    map_t_npwp16 = build_multi_kd_map(df5_1_1, 'kd_tender', 'npwp16_penyedia')
    map_t_kualifikasi = build_multi_kd_map(df5_1, 'kd_tender', 'kualifikasi_paket')
    map_t_no_sppbj = build_multi_kd_map(df5_2, 'kd_tender', 'no_sppbj')
    map_t_tgl_sppbj = build_multi_kd_map(df5_2, 'kd_tender', 'tgl_sppbj')
    map_t_no_kontrak = build_multi_kd_map(df5_3, 'kd_tender', 'no_kontrak')
    map_t_nil_kontrak = build_multi_kd_map(df5_3, 'kd_tender', 'nilai_kontrak')
    map_t_jenis_kontrak = build_multi_kd_map(df5_3, 'kd_tender', 'jenis_kontrak')
    map_t_status_kontrak = build_multi_kd_map(df5_3, 'kd_tender', 'status_kontrak')
    map_t_no_spmkspp = build_multi_kd_map(df5_4, 'kd_tender', 'no_spmk_spp')
    map_t_tgl_spmkspp = build_multi_kd_map(df5_4, 'kd_tender', 'tgl_spmk_spp')
    map_t_no_bast = build_multi_kd_map(df5_5, 'kd_tender', 'no_bast')
    map_t_tgl_bast = build_multi_kd_map(df5_5, 'kd_tender', 'tgl_bast')
    map_t_no_bap = build_multi_kd_map(df5_5, 'kd_tender', 'no_bap')
    map_t_tgl_bap = build_multi_kd_map(df5_5, 'kd_tender', 'tgl_bap')

    path_kamus = os.path.join(BASE_DIR, 'data_master', 'kamus_penyedia.json')
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

    # ================== PEMETAAN KAJI ULANG ==================
    status_terakhir = {}
    kaji_ulang_dict = {}
    if not df1_5.empty and 'kd_rup_lama' in df1_5.columns and 'kd_rup_baru' in df1_5.columns:
        for _, r in df1_5.iterrows():
            try:
                lama = int(float(str(r['kd_rup_lama']).strip()))
                baru = int(float(str(r['kd_rup_baru']).strip()))
                alasan = str(r.get('alasan_kajiulang', '')).lower()
                jenis_rev = str(r.get('jenis_revisi', '')).lower()
                tgl = str(r.get('tgl_kaji_ulang', ''))
                
                kematian = ['batal', 'pembatalan', 'hapus']
                is_batal = any(k in alasan for k in kematian) or any(k in jenis_rev for k in kematian)
                is_aktif = ('aktif' in alasan or 'pengaktifan' in jenis_rev) and not is_batal
                
                def catat_sejarah(k_id, batal_flag, tgl_str):
                    if k_id not in status_terakhir:
                        status_terakhir[k_id] = {'tgl': tgl_str, 'is_batal': batal_flag}
                    else:
                        if tgl_str >= status_terakhir[k_id]['tgl']:
                            status_terakhir[k_id] = {'tgl': tgl_str, 'is_batal': batal_flag}

                if is_batal:
                    catat_sejarah(lama, True, tgl)
                    catat_sejarah(baru, True, tgl)
                elif is_aktif:
                    catat_sejarah(lama, False, tgl)
                    catat_sejarah(baru, False, tgl)
                
                if lama != baru and not is_batal:
                    kaji_ulang_dict[lama] = baru
            except: pass

    blacklist_rup = {k for k, v in status_terakhir.items() if v['is_batal']}

    path_manual_konsol = os.path.join(BASE_DIR, 'data_master', 'manual_konsolidasi.json')
    if os.path.exists(path_manual_konsol):
        try:
            with open(path_manual_konsol, 'r', encoding='utf-8') as f:
                manual_data = json.load(f)
                if isinstance(manual_data, list):
                    for item in manual_data:
                        if item.get("status_aktif", False) and str(item.get("tahun_anggaran")) == str(tahun):
                            try:
                                lama = int(float(item["kode_konsol_lama"]))
                                baru = int(float(item["kode_konsol_baru"]))
                                kaji_ulang_dict[lama] = baru
                            except: pass
        except: pass

    def is_related(p1, p2):
        try:
            a, b = int(p1), int(p2)
            if a == b: return True
            curr, visited = a, set()
            while curr in kaji_ulang_dict and curr not in visited:
                visited.add(curr)
                curr = kaji_ulang_dict[curr]
                if curr == b: return True
            curr, visited = b, set()
            while curr in kaji_ulang_dict and curr not in visited:
                visited.add(curr)
                curr = kaji_ulang_dict[curr]
                if curr == a: return True
            return False
        except: return False

    def bersihkan_dan_ambil_kanan(raw_rup_str):
        if pd.isna(raw_rup_str) or not str(raw_rup_str).strip(): return ""
        parts = [p.strip() for p in str(raw_rup_str).split(';') if p.strip()]
        if len(parts) <= 1: return ";".join(parts)
            
        families = []
        for p in parts:
            found_family = False
            for fam in families:
                if any(is_related(p, member) for member in fam):
                    fam.append(p)
                    found_family = True
                    break
            if not found_family: families.append([p])
                
        if len(families) == 1: return parts[-1]
            
        if len(parts) % 2 == 0:
            mid = len(parts) // 2
            kiri, kanan = parts[:mid], parts[mid:]
            is_pasangan_identik = True
            kanan_terpakai = set()
            for k in kiri:
                punya_pasangan = False
                for idx_kn, kn in enumerate(kanan):
                    if idx_kn not in kanan_terpakai and is_related(k, kn):
                        punya_pasangan = True
                        kanan_terpakai.add(idx_kn)
                        break
                if not punya_pasangan:
                    is_pasangan_identik = False
                    break
            if is_pasangan_identik: return ";".join(kanan)
                
        rightmost_set = set(fam[-1] for fam in families)
        final_parts = [p for p in parts if p in rightmost_set and not rightmost_set.remove(p)]
        return ";".join(final_parts)

    def cari_rup_baru(kode_rup_bersih):
        if not kode_rup_bersih: return ""
        parts = kode_rup_bersih.split(';')
        results = []
        for p in parts:
            if p.isdigit():
                curr = int(p)
                if curr in blacklist_rup: continue
                awal = curr
                visited = set()
                while curr in kaji_ulang_dict and curr not in visited:
                    visited.add(curr)
                    next_rup = kaji_ulang_dict[curr]
                    if next_rup in blacklist_rup:
                        curr = None
                        break
                    curr = next_rup
                if curr is not None and curr != awal: results.append(str(curr))
        return ";".join(list(dict.fromkeys(results)))

    def split_kd_list(x): return [int(i.strip()) for i in str(x).split(';') if i.strip().isdigit()]

    def standardize_kd_rup(df, col):
        if df.empty or col not in df.columns: return df
        df[col+'_raw'] = df[col]
        df[col+'_list'] = df[col].apply(split_kd_list)
        df[col] = df[col].apply(lambda x: split_kd_list(x)[0] if len(split_kd_list(x))>0 else None)
        return df

    for d, c in [(df1,'kd_rup'), (df1_2,'kd_rup'), (df1_3,'kd_rup'), (df1_4,'kd_rup'), (df2,'kd_rup'), (df3,'kd_rup'), (df4,'kd_rup'), (df5,'kd_rup'), (df6,'rup_code'), (df7,'kd_rup')]:
        standardize_kd_rup(d, c)

    map_pagu_s1   = df1.drop_duplicates(subset=['kd_rup']).set_index('kd_rup')['pagu'] if not df1.empty and 'kd_rup' in df1.columns else {}
    map_pagu_s1_2 = df1_2.drop_duplicates(subset=['kd_rup']).set_index('kd_rup')['pagu'] if not df1_2.empty and 'kd_rup' in df1_2.columns else {}
    map_pagu_s1_3 = df1_3.drop_duplicates(subset=['kd_rup']).set_index('kd_rup')['pagu'] if not df1_3.empty and 'kd_rup' in df1_3.columns else {}
    map_pagu_s1_4 = df1_4.drop_duplicates(subset=['kd_rup']).set_index('kd_rup')['pagu'] if not df1_4.empty and 'kd_rup' in df1_4.columns else {}

    def get_pagu_multi(kd_list, tipe='s1'):
        if not kd_list: return None
        total, added = 0, set() 
        for k in kd_list:
            try: k_int = int(float(str(k).strip()))
            except: continue
            if k_int in added: continue
            added.add(k_int)
            if tipe == 's1': val = map_pagu_s1.get(k_int, map_pagu_s1_3.get(k_int, 0))
            else: val = map_pagu_s1_2.get(k_int, map_pagu_s1_4.get(k_int, 0))
                
            if (pd.isna(val) or float(val) == 0) and k_int in kaji_ulang_dict:
                k_baru = kaji_ulang_dict[k_int]
                if tipe == 's1': val = map_pagu_s1.get(k_baru, map_pagu_s1_3.get(k_baru, 0))
                else: val = map_pagu_s1_2.get(k_baru, map_pagu_s1_4.get(k_baru, 0))
            total += float(val) if not pd.isna(val) else 0
        return total if total != 0 else None

    def build_df_map(df_source):
        if not df_source.empty:
            df_clean = df_source.dropna(subset=['kd_rup']).drop_duplicates(subset=['kd_rup']).copy()
            df_clean['kd_rup'] = df_clean['kd_rup'].apply(lambda x: int(float(str(x).strip())))
            return df_clean.set_index('kd_rup')
        return pd.DataFrame()

    df1_map, df1_2_map, df1_3_map, df1_4_map = build_df_map(df1), build_df_map(df1_2), build_df_map(df1_3), build_df_map(df1_4)
    map_status_konsol = {}
    if not df1_3.empty and 'kd_rup' in df1_3.columns and 'status_konsolidasi' in df1_3.columns:
        map_status_konsol = {int(float(str(k).strip())): v for k, v in zip(df1_3['kd_rup'], df1_3['status_konsolidasi']) if pd.notna(k) and str(k).strip().isdigit()}

    def clean_kd(val):
        if pd.isna(val) or str(val).strip() in ["", "nan", "None"]: return ""
        try: return str(int(float(val)))
        except: return str(val).strip()

    def build_anggaran_map(df_ang):
        if df_ang.empty: return {}
        ang_map = {}
        for _, row in df_ang.iterrows():
            kd = row.get('kd_rup')
            if pd.isna(kd): continue
            try: kd_int = int(float(str(kd).strip()))
            except: continue
            
            if kd_int not in ang_map: ang_map[kd_int] = {'sd': [], 'mak': [], 'kd_keg': [], 'kd_sub': []}
            sd = str(row.get('sumber_dana', '')).strip()
            mak = str(row.get('mak', '')).strip()
            keg = clean_kd(row.get('kd_kegiatan'))
            sub = clean_kd(row.get('kd_subkegiatan'))
            
            if sd and sd != 'nan': ang_map[kd_int]['sd'].append(sd)
            if mak and mak != 'nan': ang_map[kd_int]['mak'].append(mak)
            if keg: ang_map[kd_int]['kd_keg'].append(keg)
            if sub: ang_map[kd_int]['kd_sub'].append(sub)
            
        for k in ang_map:
            ang_map[k]['sd'] = "; ".join(list(dict.fromkeys(ang_map[k]['sd'])))
            ang_map[k]['mak'] = "; ".join(list(dict.fromkeys(ang_map[k]['mak'])))
            ang_map[k]['kd_keg'] = list(dict.fromkeys(ang_map[k]['kd_keg']))
            ang_map[k]['kd_sub'] = list(dict.fromkeys(ang_map[k]['kd_sub']))
        return ang_map

    map_ang_p = build_anggaran_map(df_ang_p)
    map_ang_s = build_anggaran_map(df_ang_s)

    map_prog = dict(zip(df_prog['kd_program_str'].astype(str).str.strip(), df_prog['nama_program'])) if not df_prog.empty and 'kd_program_str' in df_prog.columns else {}
    map_keg = {clean_kd(k): str(v).strip() for k, v in zip(df_keg['kd_kegiatan'], df_keg['nama_kegiatan']) if clean_kd(k)} if not df_keg.empty and 'kd_kegiatan' in df_keg.columns else {}
    map_sub = {clean_kd(k): str(v).strip() for k, v in zip(df_sub['kd_subkegiatan'], df_sub['nama_subkegiatan']) if clean_kd(k)} if not df_sub.empty and 'kd_subkegiatan' in df_sub.columns else {}

    def get_nama_program(mak_str):
        if not mak_str or pd.isna(mak_str): return ""
        maks = [m.strip() for m in str(mak_str).split(';')]
        progs = []
        for m in maks:
            parts = m.split('.')
            if len(parts) >= 3:
                kd_prog = f"{parts[0]}.{parts[1]}.{parts[2]}"
                if kd_prog in map_prog: progs.append(map_prog[kd_prog])
        return "; ".join(list(dict.fromkeys(progs)))

    def get_kegiatan_sub(kd_list, is_swakelola, is_sub=False):
        if not kd_list: return ""
        map_ang = map_ang_s if is_swakelola else map_ang_p
        map_master = map_sub if is_sub else map_keg
        key_kd = 'kd_sub' if is_sub else 'kd_keg'
        hasil = []
        for k in kd_list:
            try: k_int = int(float(str(k).strip()))
            except: continue
            kds = map_ang.get(k_int, {}).get(key_kd, [])
            for kd in kds:
                if str(kd) in map_master: hasil.append(map_master[str(kd)])
        if not hasil: return ""
        return "; ".join(list(dict.fromkeys(hasil)))

    def get_anggaran_multi(kd_list, map_ang, key):
        if not kd_list: return ""
        res = [map_ang[k][key] for k in kd_list if k in map_ang and map_ang[k][key]]
        return "; ".join(list(dict.fromkeys(res)))

    def get_status_konsolidasi(rup_val):
        if pd.isna(rup_val) or not str(rup_val).strip(): return ""
        valid_rups = [k for k in str(rup_val).split(';') if k.strip().isdigit()]
        if len(valid_rups) > 1: return "Konsolidasi"
        elif len(valid_rups) == 1: return map_status_konsol.get(int(valid_rups[0]), "")
        return ""

    # ================= GENERATE DATA FINAL LAMA UNTUK LOOKUP =================
    log_print("🔄 Mengolah kamus detail paket untuk pengayaan (enrichment)...")
    
    # 1. Bangun dataframe lama HANYA untuk di-lookup, tidak langsung dijadikan output (Sesuai original script)
    # df_s2 -> df_s7 mirip kode lama tapi disederhanakan sebagai Mega Dictionary
    old_by_paket = {}
    old_by_rup = {}
    
    # helper memasukkan data tender ke dict lookup
    def add_to_lookup(k_paket, k_rup, dict_info):
        k_paket_str = str(k_paket).strip()
        k_rup_str = str(k_rup).strip()
        if k_paket_str and k_paket_str.lower() not in ['nan', 'none', '<na>', '']:
            for k in k_paket_str.split(';'):
                if k.strip(): old_by_paket[k.strip()] = dict_info
        if k_rup_str and k_rup_str.lower() not in ['nan', 'none', '<na>', '']:
            for r in k_rup_str.split(';'):
                if r.strip(): old_by_rup[r.strip()] = dict_info

    # 1. Tender (df5)
    for _, r in df5.iterrows():
        kd_t_list = [i.strip() for i in str(r.get('kd_tender', '')).split(';')] if pd.notna(r.get('kd_tender')) else []
        raw_rup = r.get('kd_rup_raw')
        kode_rup_asli_kotor = bersihkan_dan_ambil_kanan(raw_rup)
        cleaned_list = [int(i.strip()) for i in kode_rup_asli_kotor.split(';') if i.strip().isdigit() and int(i.strip()) not in blacklist_rup]
        if not cleaned_list: continue
        kode_rup_asli = ";".join(map(str, cleaned_list))
        
        status_awal = r.get('status_tender')
        if pd.notna(status_awal) and 'berlangsung' in str(status_awal).lower():
            found_in_selesai = any(k in s for k in kd_t_list for s in [set_kinerja_t, set_t_bapbast, set_t_spmkspp, set_t_kontrak, set_t_sppbj, set_t_selesai])
            status = 'Berlangsung' if found_in_selesai else 'Pemilihan Penyedia'
        else:
            status = status_awal
            for k in kd_t_list:
                if k in set_kinerja_t: status='Kinerja Dinilai'; break
                elif k in set_t_bapbast: status='BAPBAST'; break
                elif k in set_t_spmkspp: status='SPMKSPP'; break
                elif k in set_t_kontrak: status='Kontrak'; break
                elif k in set_t_sppbj: status='SPPBJ'; break
                elif k in set_t_selesai: status='Tender Selesai'; break
                
        kual, bia, wak, lay, s_ang, s_kat = get_kinerja_vals(kd_t_list, map_sikap_t)
        
        dict_info = {
            'Kualifikasi Paket': get_mapped_val(kd_t_list, map_t_kualifikasi),
            'Nilai Negosiasi': get_mapped_val(kd_t_list, map_t_nil_neg, is_numeric=True),
            'No SPPBJ': get_mapped_val(kd_t_list, map_t_no_sppbj), 'Tanggal SPPBJ': get_mapped_val(kd_t_list, map_t_tgl_sppbj, is_date=True),
            'No Kontrak': get_mapped_val(kd_t_list, map_t_no_kontrak), 'Tanggal Kontrak': get_mapped_val(kd_t_list, map_t_tgl_kontrak, is_date=True),
            'Jenis Kontrak': get_mapped_val(kd_t_list, map_t_jenis_kontrak), 'Nilai Kontrak': get_mapped_val(kd_t_list, map_t_nil_kontrak, is_numeric=True),
            'Status Kontrak': get_mapped_val(kd_t_list, map_t_status_kontrak), 'No SPMK SPP': get_mapped_val(kd_t_list, map_t_no_spmkspp),
            'Tanggal SPMK SPP': get_mapped_val(kd_t_list, map_t_tgl_spmkspp, is_date=True), 'No BAST': get_mapped_val(kd_t_list, map_t_no_bast),
            'Tanggal BAST': get_mapped_val(kd_t_list, map_t_tgl_bast, is_date=True), 'No BAP': get_mapped_val(kd_t_list, map_t_no_bap),
            'Tanggal BAP': get_mapped_val(kd_t_list, map_t_tgl_bap, is_date=True),
            'NPWP 15': next((map_t_npwp[k] for k in kd_t_list if k in map_t_npwp), ""), 'NPWP 16': get_mapped_val(kd_t_list, map_t_npwp16),
            'Alamat': next((map_t_alamat[k] for k in kd_t_list if k in map_t_alamat), ""),
            'Status': status, 'Nilai HPS': r.get('hps'), 'Nilai PDN': next((map_t_pdn[k] for k in kd_t_list if k in map_t_pdn), "N/A"), 
            'Nilai UMK': next((map_t_umk[k] for k in kd_t_list if k in map_t_umk), "N/A"),
            'Penilaian Kualitas dan kuantitas dengan indikator kesesuaian': kual, 'Penilaian Biaya dengan indikator kemampuan pengendalian biaya': bia,
            'Penilaian Waktu dengan indikator ketepatan': wak, 'Penilaian Layanan dengan indikator komunikasi dan tingkat respon': lay,
            'Kinerja Penyedia dalam angka': s_ang, 'Kinerja Penyedia dalam kata': s_kat,
        }
        add_to_lookup(r.get('kd_tender'), kode_rup_asli, dict_info)

    # 2. Non-Tender (df2)
    for _, r in df2.iterrows():
        kd_nt_list = [i.strip() for i in str(r.get('kd_nontender', '')).split(';')] if pd.notna(r.get('kd_nontender')) else []
        raw_rup = r.get('kd_rup_raw')
        kode_rup_asli_kotor = bersihkan_dan_ambil_kanan(raw_rup)
        cleaned_list = [int(i.strip()) for i in kode_rup_asli_kotor.split(';') if i.strip().isdigit() and int(i.strip()) not in blacklist_rup]
        if not cleaned_list: continue
        kode_rup_asli = ";".join(map(str, cleaned_list))
        
        status_awal = r.get('status_nontender')
        if pd.notna(status_awal) and 'berlangsung' in str(status_awal).lower():
            found_in_selesai = any(k in s for k in kd_nt_list for s in [set_kinerja_nt, set_bapbast, set_spmkspp, set_kontrak, set_sppbj, set_selesai])
            status = 'Berlangsung' if found_in_selesai else 'Pemilihan Penyedia'
        else:
            status = status_awal
            for k in kd_nt_list:
                if k in set_kinerja_nt: status='Kinerja Dinilai'; break
                elif k in set_bapbast: status='BAPBAST'; break
                elif k in set_spmkspp: status='SPMKSPP'; break
                elif k in set_kontrak: status='Kontrak'; break
                elif k in set_sppbj: status='SPPBJ'; break
                elif k in set_selesai: status='Non Tender Selesai'; break
                
        kual, bia, wak, lay, s_ang, s_kat = get_kinerja_vals(kd_nt_list, map_sikap_nt)
        
        dict_info = {
            'Kualifikasi Paket': get_mapped_val(kd_nt_list, map_nt_kualifikasi),
            'Nilai Negosiasi': get_mapped_val(kd_nt_list, map_nt_nil_neg, is_numeric=True),
            'No SPPBJ': get_mapped_val(kd_nt_list, map_nt_no_sppbj), 'Tanggal SPPBJ': get_mapped_val(kd_nt_list, map_nt_tgl_sppbj, is_date=True),
            'No Kontrak': get_mapped_val(kd_nt_list, map_nt_no_kontrak), 'Tanggal Kontrak': get_mapped_val(kd_nt_list, map_nt_tgl_kontrak, is_date=True),
            'Jenis Kontrak': get_mapped_val(kd_nt_list, map_nt_jenis_kontrak), 'Nilai Kontrak': get_mapped_val(kd_nt_list, map_nt_nil_kontrak, is_numeric=True),
            'Status Kontrak': get_mapped_val(kd_nt_list, map_nt_status_kontrak), 'No SPMK SPP': get_mapped_val(kd_nt_list, map_nt_no_spmkspp),
            'Tanggal SPMK SPP': get_mapped_val(kd_nt_list, map_nt_tgl_spmkspp, is_date=True), 'No BAST': get_mapped_val(kd_nt_list, map_nt_no_bast),
            'Tanggal BAST': get_mapped_val(kd_nt_list, map_nt_tgl_bast, is_date=True), 'No BAP': get_mapped_val(kd_nt_list, map_nt_no_bap),
            'Tanggal BAP': get_mapped_val(kd_nt_list, map_nt_tgl_bap, is_date=True),
            'NPWP 15': next((map_nt_npwp[k] for k in kd_nt_list if k in map_nt_npwp), ""), 'NPWP 16': get_mapped_val(kd_nt_list, map_nt_npwp16),
            'Alamat': next((map_nt_alamat[k] for k in kd_nt_list if k in map_nt_alamat), ""),
            'Status': status, 'Nilai HPS': r.get('hps'), 'Nilai PDN': next((map_nt_pdn[k] for k in kd_nt_list if k in map_nt_pdn), "N/A"), 
            'Nilai UMK': next((map_nt_umk[k] for k in kd_nt_list if k in map_nt_umk), "N/A"),
            'Penilaian Kualitas dan kuantitas dengan indikator kesesuaian': kual, 'Penilaian Biaya dengan indikator kemampuan pengendalian biaya': bia,
            'Penilaian Waktu dengan indikator ketepatan': wak, 'Penilaian Layanan dengan indikator komunikasi dan tingkat respon': lay,
            'Kinerja Penyedia dalam angka': s_ang, 'Kinerja Penyedia dalam kata': s_kat,
        }
        add_to_lookup(r.get('kd_nontender'), kode_rup_asli, dict_info)

    # 3. EPurchasing V6 (df6)
    for _, r in df6.iterrows():
        raw_rup = r.get('rup_code_raw')
        kode_rup_asli_kotor = bersihkan_dan_ambil_kanan(raw_rup)
        cleaned_list = [int(i.strip()) for i in kode_rup_asli_kotor.split(';') if i.strip().isdigit() and int(i.strip()) not in blacklist_rup]
        if not cleaned_list: continue
        kode_rup_asli = ";".join(map(str, cleaned_list))
        
        dict_info = {
            'No Kontrak': f"EP-{r.get('order_id')}" if pd.notna(r.get('order_id')) and str(r.get('order_id')).strip() else "", 
            'NPWP 15': str(r.get('npwp_penyedia', r.get('npwp', ''))), 'Alamat': str(r.get('alamat_penyedia', r.get('alamat', ''))),
            'Status': r.get('status')
        }
        add_to_lookup(r.get('order_id'), kode_rup_asli, dict_info)

    # 4. EPurchasing V5 (df7)
    for _, r in df7.iterrows():
        raw_rup = r.get('kd_rup_raw')
        kode_rup_asli_kotor = bersihkan_dan_ambil_kanan(raw_rup)
        cleaned_list = [int(i.strip()) for i in kode_rup_asli_kotor.split(';') if i.strip().isdigit() and int(i.strip()) not in blacklist_rup]
        if not cleaned_list: continue
        kode_rup_asli = ";".join(map(str, cleaned_list))
        
        paket_status_str = r.get('paket_status_str')
        status_s7 = r.get('status_paket') if paket_status_str == "Paket Proses" else (r.get('paket_status_str') if paket_status_str == "Paket Selesai" else paket_status_str)
        
        dict_info = {
            'NPWP 15': str(r.get('npwp_penyedia', r.get('npwp', ''))), 'Alamat': str(r.get('alamat_penyedia', r.get('alamat', ''))),
            'Status': status_s7
        }
        add_to_lookup(r.get('no_paket'), kode_rup_asli, dict_info)

    cols = [
        'Kode Paket', 'Kode RUP', 'Kode RUP Baru', 'Nama Instansi', 'Satuan Kerja', 'Nama Paket', 
        'Metode Pemilihan', 'Jenis Pengadaan', 'Tahun Anggaran', 'Nama Program', 'Nama Kegiatan', 
        'Nama Sub Kegiatan', 'Sumber Dana', 'MAK', 'PDN', 'UKM', 'Status Konsolidasi', 
        'Kualifikasi Paket', 'Nilai Pagu RUP', 'Nilai Hasil Pemilihan', 'Nilai Negosiasi', 
        'No SPPBJ', 'Tanggal SPPBJ', 'No Kontrak', 'Tanggal Kontrak', 'Jenis Kontrak', 
        'Nilai Kontrak', 'Status Kontrak', 'No SPMK SPP', 'Tanggal SPMK SPP', 'No BAST', 
        'Tanggal BAST', 'No BAP', 'Tanggal BAP', 'Nama Penyedia', 'NPWP 15', 'NPWP 16', 
        'Alamat', 'Status', 'Nilai HPS', 'Nilai PDN', 'Nilai UMK', 
        'Penilaian Kualitas dan kuantitas dengan indikator kesesuaian', 'Penilaian Biaya dengan indikator kemampuan pengendalian biaya',
        'Penilaian Waktu dengan indikator ketepatan', 'Penilaian Layanan dengan indikator komunikasi dan tingkat respon',
        'Kinerja Penyedia dalam angka', 'Kinerja Penyedia dalam kata',
        'Cara Pengadaan', 'Sumber'
    ]

    log_print("🔄 Melakukan pemetaan dan pengayaan data API Inaproc...")
    enriched_data = []
    
    for item in data_inaproc:
        # DATA DASAR INAPROC
        base_jenis = item.get('jenis_pengadaan', '')
        base_paket = str(item.get('kode_paket', '')).strip()
        raw_rup = str(item.get('kode_rup', '')).replace(',', ';')
        base_metode = item.get('metode_pengadaan', '')
        base_instansi = item.get('nama_instansi', '')
        base_nama_paket = item.get('nama_paket', '')
        base_penyedia = item.get('nama_penyedia', '')
        base_satker = item.get('nama_satuan_kerja', '')
        base_nilai = item.get('total_nilai', '')
        
        # PEMBERSIHAN RUP
        kode_rup_asli_kotor = bersihkan_dan_ambil_kanan(raw_rup)
        cleaned_list = [int(i.strip()) for i in kode_rup_asli_kotor.split(';') if i.strip().isdigit() and int(i.strip()) not in blacklist_rup]
        kode_rup_asli = ";".join(map(str, cleaned_list))
        kode_rup_baru = cari_rup_baru(kode_rup_asli)

        new_row = {c: "" for c in cols}
        new_row['Kode Paket'] = base_paket
        new_row['Kode RUP'] = kode_rup_asli
        new_row['Kode RUP Baru'] = kode_rup_baru
        new_row['Nama Instansi'] = base_instansi
        new_row['Satuan Kerja'] = base_satker
        new_row['Nama Paket'] = base_nama_paket
        new_row['Metode Pemilihan'] = base_metode
        new_row['Jenis Pengadaan'] = base_jenis
        new_row['Nama Penyedia'] = base_penyedia
        new_row['Nilai Hasil Pemilihan'] = base_nilai
        new_row['Tahun Anggaran'] = tahun
        new_row['Sumber'] = 'API Inaproc'

        # PENCARIAN DETAIL DARI LOKAL
        matched_row = None
        if base_paket and base_paket in old_by_paket:
            matched_row = old_by_paket[base_paket]
        else:
            for r in cleaned_list:
                if str(r) in old_by_rup:
                    matched_row = old_by_rup[str(r)]
                    break

        if matched_row is not None:
            for c in cols:
                # Kolom Data Dasar tidak boleh ditimpa
                if c not in ['Kode Paket', 'Kode RUP', 'Kode RUP Baru', 'Nama Instansi', 'Satuan Kerja', 
                             'Nama Paket', 'Metode Pemilihan', 'Jenis Pengadaan', 'Nama Penyedia', 
                             'Nilai Hasil Pemilihan', 'Tahun Anggaran', 'Sumber']:
                    new_row[c] = matched_row.get(c, "")
                    
        # PERHITUNGAN ULANG UNTUK PENGANGGARAN (JIka tidak dapat dari kamus lokal)
        is_swakelola = 'swakelola' in str(base_metode).lower()
        map_ang_target = map_ang_s if is_swakelola else map_ang_p
        
        if not new_row.get('Sumber Dana'):
            sd = get_anggaran_multi(cleaned_list, map_ang_target, 'sd')
            if not sd and kode_rup_baru:
                list_baru = [int(i.strip()) for i in kode_rup_baru.split(';') if i.strip().isdigit()]
                sd = get_anggaran_multi(list_baru, map_ang_target, 'sd')
            new_row['Sumber Dana'] = sd
            
        if not new_row.get('MAK'):
            mak = get_anggaran_multi(cleaned_list, map_ang_target, 'mak')
            if not mak and kode_rup_baru:
                list_baru = [int(i.strip()) for i in kode_rup_baru.split(';') if i.strip().isdigit()]
                mak = get_anggaran_multi(list_baru, map_ang_target, 'mak')
            new_row['MAK'] = mak
            new_row['Nama Program'] = get_nama_program(new_row['MAK'])
            
        if not new_row.get('Nama Kegiatan'):
            keg = get_kegiatan_sub(cleaned_list, is_swakelola, is_sub=False)
            if not keg and kode_rup_baru:
                list_baru = [int(i.strip()) for i in kode_rup_baru.split(';') if i.strip().isdigit()]
                keg = get_kegiatan_sub(list_baru, is_swakelola, is_sub=False)
            new_row['Nama Kegiatan'] = keg
            
        if not new_row.get('Nama Sub Kegiatan'):
            sub = get_kegiatan_sub(cleaned_list, is_swakelola, is_sub=True)
            if not sub and kode_rup_baru:
                list_baru = [int(i.strip()) for i in kode_rup_baru.split(';') if i.strip().isdigit()]
                sub = get_kegiatan_sub(list_baru, is_swakelola, is_sub=True)
            new_row['Nama Sub Kegiatan'] = sub
            
        if new_row.get('Nilai Pagu RUP') in ["", None, pd.NA]:
            pagu = get_pagu_multi(cleaned_list, 's1_2' if is_swakelola else 's1')
            new_row['Nilai Pagu RUP'] = pagu if pagu else ""
            
        if not new_row.get('Status Konsolidasi'):
            new_row['Status Konsolidasi'] = get_status_konsolidasi(kode_rup_asli)

        enriched_data.append(new_row)

    df_detail_inaproc = pd.DataFrame(enriched_data)
    df_detail_inaproc = df_detail_inaproc.map(lambda x: re.sub(r'[\x00-\x1F]', '', str(x)) if isinstance(x, str) else x)
    
    df_detail_inaproc['PDN'] = df_detail_inaproc['PDN'].replace("", "N/A")
    df_detail_inaproc['UKM'] = df_detail_inaproc['UKM'].replace("", "N/A")
    df_detail_inaproc['Status'] = df_detail_inaproc['Status'].apply(lambda x: str(x).replace('_', ' ') if pd.notna(x) else x)

    def safe_numeric(val):
        if val in ("N/A", "", None) or pd.isna(val): return 0.0
        try: return float(val)
        except: return 0.0

    def aggregate_text(series):
        valid = [str(x).strip() for x in series if pd.notna(x) and str(x).strip() not in ("N/A", "")]
        return "; ".join(list(dict.fromkeys(valid)))

    def aggregate_raw_status(group, pagu_rup, sum_hasil):
        def get_score(s_text):
            s_low = str(s_text).strip().lower()
            if 'batal' in s_low or 'gagal' in s_low or 'cancel' in s_low: return -1
            elif 'kinerja dinilai' in s_low: return 110
            elif 'bapbast' in s_low or 'payment' in s_low or 'completed' in s_low or 'paket selesai' in s_low: return 100
            elif 'spmkspp' in s_low: return 60
            elif 'kontrak' in s_low: return 50
            elif 'sppbj' in s_low: return 40
            elif 'selesai' in s_low: return 30
            elif 'pengumuman' in s_low and 'rup' not in s_low: return 20
            elif 'pengumuman rup' in s_low: return 10
            elif s_low == '' or s_low == '-': return 0
            else: return 5

        scored_statuses = []
        all_batal, has_proses_berjalan = True, False
        for _, row in group.iterrows():
            raw_s = str(row['Status']).strip()
            score = get_score(raw_s)
            if score != -1: all_batal = False
            if 0 < score < 100 and score != 10: has_proses_berjalan = True
            if raw_s: scored_statuses.append((score, raw_s))
            
        if all_batal or not scored_statuses: return "; ".join(list(dict.fromkeys([s for score, s in scored_statuses])))
        
        valid_scored = [item for item in scored_statuses if item[0] != -1]
        if not valid_scored: return "; ".join(list(dict.fromkeys([s for score, s in scored_statuses])))

        if has_proses_berjalan:
            filtered_berjalan = [item for item in valid_scored if item[0] < 100]
            if filtered_berjalan:
                max_score = max([score for score, s in filtered_berjalan])
                return "; ".join(list(dict.fromkeys([s for score, s in filtered_berjalan if score == max_score])))

        max_score = max([score for score, s in valid_scored])
        hasil_status = "; ".join(list(dict.fromkeys([s for score, s in valid_scored if score == max_score])))
        
        persentase_terpakai = (sum_hasil / pagu_rup) * 100 if pagu_rup > 0 else 100
        if len(group) > 1 and max_score == 100 and persentase_terpakai < 80: return "Sedang Berjalan (Sisa Pagu)"
        return hasil_status

    df_agregasi = df_detail_inaproc.copy()
    for c in ['Nilai HPS', 'Nilai Hasil Pemilihan', 'Nilai PDN', 'Nilai UMK', 'Nilai Pagu RUP', 'Nilai Negosiasi', 'Nilai Kontrak']:
        df_agregasi[c] = df_agregasi[c].apply(safe_numeric)

    df_agregasi['Kode RUP_Group'] = df_agregasi['Kode RUP'].apply(lambda x: str(x).split(';')[0].strip() if x else "UNKNOWN")
    grouped = df_agregasi.groupby('Kode RUP_Group')

    rekap_data = []
    for kode_rup, group in grouped:
        first_row = group.iloc[0]
        sum_hps, sum_hasil = group['Nilai HPS'].sum(), group['Nilai Hasil Pemilihan'].sum()
        sum_pdn, sum_umk = group['Nilai PDN'].sum(), group['Nilai UMK'].sum()
        sum_neg = group['Nilai Negosiasi'].sum()
        sum_kon = group['Nilai Kontrak'].sum()
        pagu_rup = first_row['Nilai Pagu RUP']

        rekap_data.append({
            'Kode Paket': aggregate_text(group['Kode Paket']), 'Kode RUP': first_row['Kode RUP'], 'Kode RUP Baru': first_row['Kode RUP Baru'], 
            'Nama Instansi': first_row['Nama Instansi'], 'Satuan Kerja': first_row['Satuan Kerja'], 'Nama Paket': first_row['Nama Paket'], 'Metode Pemilihan': first_row['Metode Pemilihan'],
            'Jenis Pengadaan': first_row['Jenis Pengadaan'], 'Tahun Anggaran': first_row['Tahun Anggaran'], 
            'Nama Program': first_row['Nama Program'], 'Nama Kegiatan': first_row['Nama Kegiatan'], 'Nama Sub Kegiatan': first_row['Nama Sub Kegiatan'],
            'Sumber Dana': first_row['Sumber Dana'], 'MAK': first_row['MAK'], 'PDN': first_row['PDN'], 'UKM': first_row['UKM'], 'Status Konsolidasi': first_row['Status Konsolidasi'],
            'Kualifikasi Paket': aggregate_text(group['Kualifikasi Paket']), 'Nilai Pagu RUP': pagu_rup if pagu_rup != 0 else "", 
            'Nilai Hasil Pemilihan': sum_hasil if sum_hasil != 0 else "", 'Nilai Negosiasi': sum_neg if sum_neg != 0 else "",
            'No SPPBJ': aggregate_text(group['No SPPBJ']), 'Tanggal SPPBJ': aggregate_text(group['Tanggal SPPBJ']),
            'No Kontrak': aggregate_text(group['No Kontrak']), 'Tanggal Kontrak': aggregate_text(group['Tanggal Kontrak']), 
            'Jenis Kontrak': aggregate_text(group['Jenis Kontrak']), 'Nilai Kontrak': sum_kon if sum_kon != 0 else "",
            'Status Kontrak': aggregate_text(group['Status Kontrak']), 'No SPMK SPP': aggregate_text(group['No SPMK SPP']),
            'Tanggal SPMK SPP': aggregate_text(group['Tanggal SPMK SPP']), 'No BAST': aggregate_text(group['No BAST']),
            'Tanggal BAST': aggregate_text(group['Tanggal BAST']), 'No BAP': aggregate_text(group['No BAP']),
            'Tanggal BAP': aggregate_text(group['Tanggal BAP']),
            'Nama Penyedia': aggregate_text(group['Nama Penyedia']), 'NPWP 15': aggregate_text(group['NPWP 15']), 'NPWP 16': aggregate_text(group['NPWP 16']), 
            'Alamat': aggregate_text(group['Alamat']),
            'Status': aggregate_raw_status(group, pagu_rup, sum_hasil), 'Nilai HPS': sum_hps if sum_hps != 0 else "",
            'Nilai PDN': sum_pdn if sum_pdn != 0 else "", 'Nilai UMK': sum_umk if sum_umk != 0 else "",
            'Penilaian Kualitas dan kuantitas dengan indikator kesesuaian': aggregate_text(group['Penilaian Kualitas dan kuantitas dengan indikator kesesuaian']),
            'Penilaian Biaya dengan indikator kemampuan pengendalian biaya': aggregate_text(group['Penilaian Biaya dengan indikator kemampuan pengendalian biaya']),
            'Penilaian Waktu dengan indikator ketepatan': aggregate_text(group['Penilaian Waktu dengan indikator ketepatan']),
            'Penilaian Layanan dengan indikator komunikasi dan tingkat respon': aggregate_text(group['Penilaian Layanan dengan indikator komunikasi dan tingkat respon']),
            'Kinerja Penyedia dalam angka': aggregate_text(group['Kinerja Penyedia dalam angka']),
            'Kinerja Penyedia dalam kata': aggregate_text(group['Kinerja Penyedia dalam kata']),
            'Cara Pengadaan': aggregate_text(group['Cara Pengadaan']), 'Sumber': aggregate_text(group['Sumber'])
        })
    df_rekap = pd.DataFrame(rekap_data)

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(df_detail_inaproc.to_dict(orient='records'), f, ensure_ascii=False, separators=(',', ':'))
    log_print(f"JSON Rekap Inaproc sukses dibuat: {output_json}")

    kolom_angka_baku = ['Nilai Pagu RUP', 'Nilai Hasil Pemilihan', 'Nilai Negosiasi', 'Nilai Kontrak', 'Nilai HPS', 'Nilai PDN', 'Nilai UMK']
    excel_df_detail = df_detail_inaproc.copy()
    for col in kolom_angka_baku:
        if col in excel_df_detail.columns:
            excel_df_detail[col] = excel_df_detail[col].apply(lambda x: safe_numeric(x) if safe_numeric(x) != 0 else "")

    tahun_label = str(df_detail_inaproc['Tahun Anggaran'].iloc[0]) if not df_detail_inaproc.empty else str(tahun)
    nama_file_excel = f'Paket Pengadaan Tahun {tahun_label} ({datetime.now().strftime("%Y-%m-%d")}).xlsx'
    os.makedirs(output_dir_excel, exist_ok=True)
    output_excel_path = os.path.join(output_dir_excel, nama_file_excel)

    berhasil_simpan = False
    while not berhasil_simpan:
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                df_rekap.to_excel(writer, index=False, sheet_name='1. Rekap per RUP')
                excel_df_detail.to_excel(writer, index=False, sheet_name='2. Detail per Paket')

            wb = load_workbook(output_excel_path)
            header_fill = PatternFill('solid', start_color='1F4E79')
            header_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
            header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            data_font = Font(name='Arial', size=10)
            fill_putih = PatternFill('solid', start_color='FFFFFF')
            fill_biru_muda = PatternFill('solid', start_color='DCE6F1')
            border_thin = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

            def style_sheet(ws_name, df_ref):
                ws = wb[ws_name]
                for i, _ in enumerate(df_ref.columns, start=1): ws.column_dimensions[get_column_letter(i)].width = 15
                for cell in ws[1]: cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, header_align, border_thin
                ws.row_dimensions[1].height = 32
                for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
                    fill = fill_putih if row_idx % 2 == 0 else fill_biru_muda
                    for cell in row:
                        col_name = df_ref.columns[cell.column - 1]
                        cell.font, cell.fill, cell.border = data_font, fill, border_thin
                        if col_name in kolom_angka_baku and str(cell.value).replace('.', '', 1).isdigit():
                            cell.number_format = '#,##0'
                            cell.alignment = Alignment(horizontal='right', vertical='center')
                        else: cell.alignment = Alignment(vertical='center', wrap_text=False)
                ws.freeze_panes = 'A2'
                ws.auto_filter.ref = ws.dimensions

            style_sheet('1. Rekap per RUP', df_rekap)
            style_sheet('2. Detail per Paket', excel_df_detail)
            wb.save(output_excel_path)
            berhasil_simpan = True
            log_print(f"✅ EXCEL Pengadaan Inaproc sukses dibuat: {output_excel_path}")
            
        except PermissionError:
            import time
            log_print(f"⚠️ Akses Ditolak: File {nama_file_excel} sedang terbuka.")
            tutup_sukses = False
            try:
                import win32com.client
                excel_app = win32com.client.GetActiveObject("Excel.Application")
                for wb_app in excel_app.Workbooks:
                    if wb_app.Name == nama_file_excel:
                        wb_app.Close(SaveChanges=False)
                        tutup_sukses = True
                        break
            except Exception: pass
            
            if not tutup_sukses:
                os.system(f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""File {nama_file_excel} sedang terbuka. Tutup file di Excel agar proses berlanjut!"", 5, ""Peringatan Excel"", 48)(window.close)")')
            time.sleep(3)
    
    master_path = os.path.join(data_dir, f'master_pengadaan_{tahun}.xlsx')
    nama_file_master = f'master_pengadaan_{tahun}.xlsx'
    berhasil_salin = False
    
    while not berhasil_salin:
        try:
            shutil.copy2(output_excel_path, master_path)
            log_print(f"✅ EXCEL Master sukses disalin: {master_path}")
            berhasil_salin = True
        except PermissionError:
            import time
            log_print(f"⚠️ Akses Ditolak: File {nama_file_master} sedang terbuka.")
            tutup_sukses = False
            try:
                import win32com.client
                excel_app = win32com.client.GetActiveObject("Excel.Application")
                for wb_app in excel_app.Workbooks:
                    if wb_app.Name == nama_file_master:
                        wb_app.Close(SaveChanges=False)
                        tutup_sukses = True
                        break
            except Exception: pass
            
            if not tutup_sukses:
                os.system(f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""File {nama_file_master} sedang terbuka. Tutup file di Excel agar proses berlanjut!"", 5, ""Peringatan Excel"", 48)(window.close)")')
            time.sleep(3)
        except Exception as e:
            break
            
    kelola_arsip_bulanan(output_dir_excel, tahun)
    update_daftar_arsip_json(output_dir_excel)
    
    log_print(f'SELESAI GENERATE INAPROC TAHUN {tahun} | Total data: {len(df_detail_inaproc)}')
    return len(df_detail_inaproc)


def kirim_telegram_aman(pesan):
    if len(pesan) > 4000: pesan = pesan[:4000] + "\n...[TERPOTONG]"
    import requests
    import time
    url = f"https://api.telegram.org/bot{config_rahasia.BOT_TOKEN}/sendMessage"
    
    for percobaan in range(1, 4):
        try:
            res = requests.post(url, data={"chat_id": config_rahasia.CHAT_ID, "text": pesan}, timeout=30)
            if res.status_code != 200: 
                log_print(f"🚨 GAGAL KIRIM TELEGRAM (STATUS {res.status_code}): {res.text}")
            break
        except Exception as e:
            if percobaan < 3:
                time.sleep(5)
            else:
                log_print(f"🚨 GAGAL KONEKSI KE TELEGRAM: {str(e)}")

if __name__ == '__main__':
    log_print("\n" + "="*55)
    log_print(f"START GENERATE PENGADAAN INAPROC {get_waktu_indonesia()}")
    log_print("="*55)

    daftar_error_api = []
    path_error = os.path.join(BASE_DIR, 'scripts', 'pengadaan', 'error_api_pengadaan.json')
    if os.path.exists(path_error):
        try:
            with open(path_error, 'r', encoding='utf-8') as f: daftar_error_api = json.load(f)
            os.remove(path_error)
        except: pass

    for t in daftar_tahun: process_tahun(t)

    with open(os.path.join(BASE_DIR, "data", "last-update-pengadaan.txt"), "w", encoding='utf-8') as f:
        f.write(get_waktu_indonesia())

    # --- EKSEKUSI TEPRA ---
    log_print("\n" + "="*50)
    log_print("START GENERATE TEPRA")
    status_tepra = "✅ Data TEPRA berhasil dibuat."
    try:
        path_tepra = os.path.join(BASE_DIR, 'scripts', 'pengadaan', 'generate_tepra.py')
        res_tepra = subprocess.run([sys.executable, path_tepra], capture_output=True, text=True)
        if res_tepra.returncode == 0:
            log_print("PROSES TEPRA SUKSES\n" + res_tepra.stdout.strip())
        else:
            log_print(f"GAGAL PROSES TEPRA:\n{res_tepra.stderr}")
            status_tepra = "⚠️ Gagal membuat data TEPRA (Cek Log)."
    except Exception as e:
        log_print(f"ERROR TEPRA: {str(e)}")
        status_tepra = "⚠️ Error sistem saat eksekusi TEPRA."

    git_sukses, pesan_git = sync_to_github()

    durasi_str = "Tidak diketahui"
    file_start = os.path.join(BASE_DIR, 'tools', 'start_time_pengadaan.txt')
    if os.path.exists(file_start):
        try:
            with open(file_start, 'r') as f:
                waktu_mulai = float(f.read().strip())
            durasi_detik = int(time.time() - waktu_mulai)
            durasi_str = str(timedelta(seconds=durasi_detik))
            os.remove(file_start) 
        except: pass

    if len(daftar_error_api) > 0 or not git_sukses:
        pesan_ringkasan = "🚨 LAPORAN UPDATE SISTEM (INAPROC) 🚨\n\n"
        if len(daftar_error_api) > 0:
            teks_error = "⚠️ GAGAL DOWNLOAD API:\n" + "".join([f"{err.replace('_', ' ')}\n" for err in daftar_error_api])
            if len(teks_error) > 3500: teks_error = teks_error[:3500] + "\n... [DAFTAR ERROR DIPOTONG] ...\n"
            pesan_ringkasan += teks_error + "\n"
        pesan_ringkasan += f"📊 STATUS TEPRA: {status_tepra}\n\n🌐 STATUS GITHUB:\n{pesan_git}\n\n⏱ Durasi Total: {durasi_str}\n📅 Waktu: {get_waktu_indonesia()}"
        kirim_telegram_aman(pesan_ringkasan)
    else:
        pesan_sukses = f"✅ UPDATE PENGADAAN INAPROC BERHASIL ✅\n\nSeluruh data berhasil diolah.\n📊 STATUS TEPRA: {status_tepra}\n\n🌐 STATUS GITHUB:\n{pesan_git}\n\n⏱ Durasi Total: {durasi_str}\n📅 Waktu: {get_waktu_indonesia()}"
        kirim_telegram_aman(pesan_sukses)
        
    log_print(f"\nPROSES SELESAI SELURUHNYA PADA {get_waktu_indonesia()}")