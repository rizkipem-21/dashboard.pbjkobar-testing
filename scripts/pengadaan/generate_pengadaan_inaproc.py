# ======================================================
# ini file generate_pengadaan_inaproc.py
# DATA DASAR = Inaproc (dashboard/realisasi/table)
# Kolom lain dilengkapi dengan logika lama (lewat Kode Paket / Kode RUP)
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
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
import config_rahasia

warnings.filterwarnings('ignore')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

tahun_n      = datetime.now().year
tahun_n1     = tahun_n - 1
tahun_n2     = tahun_n - 2
daftar_tahun = [tahun_n, tahun_n1, tahun_n2]

# MENGGUNAKAN LOG TUNGGAL (Sama dengan script download)
LOG_FILE = os.path.join(BASE_DIR, 'tools', 'log_pengadaan.txt')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

TEKS_TIDAK_COCOK = "Tidak Cocok"

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

        commit_msg = f"Auto update Pengadaan {waktu_sekarang}"
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

            # Jika data adalah list langsung
            if isinstance(data, list):
                return pd.json_normalize(data) if data else pd.DataFrame()

            # Jika data adalah dictionary (respons API standar)
            if isinstance(data, dict):
                for k in ['data', 'items', 'results']:
                    if k in data and isinstance(data[k], list):
                        return pd.json_normalize(data[k])

                return pd.DataFrame()

    except Exception as e:
        return pd.DataFrame()

# ======================================================
# HELPER BARU: NORMALISASI KUNCI PENCOCOKAN
# (semua kunci dijadikan teks tanpa ".0" agar angka vs teks tidak gagal cocok)
# ======================================================
def norm_kd(v):
    if v is None: return ""
    try:
        if pd.isna(v): return ""
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    if s.lower() in ("nan", "none", "<na>", "nat"): return ""
    if re.fullmatch(r"-?\d+\.0+", s): s = s.split('.')[0]
    return s

def val(r, k, default=""):
    """Ambil nilai dari baris (Series/dict); NaN/None -> default"""
    v = r.get(k, default) if r is not None else default
    try:
        if v is None or pd.isna(v): return default
    except (TypeError, ValueError):
        pass
    return v

def to_float(v):
    try:
        if v in ("N/A", "", None) or pd.isna(v): return 0.0
        return float(str(v).replace(',', ''))
    except: return 0.0

def build_row_index(df, col):
    """Kamus {kode_paket_teks: baris pertama} dari sebuah tabel sumber"""
    idx = {}
    if df.empty or col not in df.columns: return idx
    for _, r in df.iterrows():
        for k in str(r[col]).split(';'):
            k = norm_kd(k)
            if k and k not in idx: idx[k] = r
    return idx

def resolve_tipe_sumber(sumber_transaksi):
    """Menentukan tabel pencari berdasarkan kolom sumber_transaksi Inaproc"""
    s = str(sumber_transaksi or '').strip().lower()
    if not s: return None
    if 'swakelola' in s: return 'df4'
    if 'pencatatan' in s: return 'df3'
    if 'non tender' in s or 'non-tender' in s or 'nontender' in s: return 'df2'
    if 'tender' in s: return 'df5'
    if 'katalog' in s or 'purchasing' in s:
        return 'df7' if ('5' in s) else 'df6'
    return None

def hitung_status(status_awal, kd_list, rantai):
    """Logika status lama. rantai = [(set_kode, label), ...] urut prioritas"""
    if pd.notna(status_awal) and 'berlangsung' in str(status_awal).lower():
        ada = any(k in s for k in kd_list for s, _ in rantai)
        return 'Berlangsung' if ada else 'Pemilihan Penyedia'
    for k in kd_list:
        for s, label in rantai:
            if k in s: return label
    return status_awal

def kelola_arsip_bulanan(folder_path, tahun):
    if not os.path.exists(folder_path): return

    folder_arsip_lokal = os.path.join(BASE_DIR, 'arsip_lokal', 'pengadaan', str(tahun))
    os.makedirs(folder_arsip_lokal, exist_ok=True)

    # Abaikan file temporary Excel yang berawalan ~$
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
                    import time
                    log_print(f"⚠️ Akses Ditolak: File lama {nama_file} sedang terbuka.")
                    tutup_sukses = False
                    try:
                        import win32com.client
                        excel_app = win32com.client.GetActiveObject("Excel.Application")
                        for wb_app in excel_app.Workbooks:
                            if wb_app.Name == nama_file:
                                log_print(f"Menutup otomatis file: {nama_file}...")
                                wb_app.Close(SaveChanges=False)
                                tutup_sukses = True
                                break
                    except Exception: pass

                    if not tutup_sukses:
                        log_print("Memunculkan pop-up peringatan arsip...")
                        os.system(f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""File lama {nama_file} sedang terbuka. Tutup file di Excel agar bisa diarsipkan!"", 5, ""Peringatan Excel"", 48)(window.close)")')
                    time.sleep(3)
                except Exception as e:
                    log_print(f"⚠️ Gagal memindah arsip {nama_file}: {str(e)}")
                    berhasil_pindah = True

def update_daftar_arsip_json(folder_path):
    """Membaca sisa file Excel di folder output dan memperbarui daftar_arsip.json"""
    if not os.path.exists(folder_path): return

    # Ambil semua file excel, urutkan dari yang terbaru (Z-A)
    file_excel = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    file_excel.sort(reverse=True)

    # Buat format JSON (sesuaikan dengan format yang dibaca website Anda)
    arsip_list = [{"nama_file": f} for f in file_excel]

    # Simpan ke daftar_arsip.json
    path_json = os.path.join(folder_path, 'daftar_arsip.json')
    try:
        with open(path_json, 'w', encoding='utf-8') as f:
            json.dump(arsip_list, f, indent=4)
        log_print(f"📝 Daftar arsip JSON diperbarui: {path_json}")
    except Exception as e:
        log_print(f"⚠️ Gagal memperbarui daftar_arsip.json: {str(e)}")

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

# Kolom pelengkap (di luar 9 kolom data dasar Inaproc) yang diisi per baris
KOLOM_DETAIL = [
    'Tahun Anggaran', 'Sumber Dana', 'MAK', 'PDN', 'UKM', 'Kualifikasi Paket', 'Nilai Pagu RUP',
    'Nilai Negosiasi', 'No SPPBJ', 'Tanggal SPPBJ', 'No Kontrak', 'Tanggal Kontrak', 'Jenis Kontrak',
    'Nilai Kontrak', 'Status Kontrak', 'No SPMK SPP', 'Tanggal SPMK SPP', 'No BAST', 'Tanggal BAST',
    'No BAP', 'Tanggal BAP', 'NPWP 15', 'NPWP 16', 'Alamat', 'Status', 'Nilai HPS', 'Nilai PDN',
    'Nilai UMK', 'Cara Pengadaan', 'Sumber'
]

def process_tahun(tahun):
    # Deklarasikan semua path di awal agar dikenali oleh seluruh fungsi
    data_dir = os.path.join(BASE_DIR, 'data', str(tahun))
    output_dir_excel = os.path.join(BASE_DIR, 'output', 'pengadaan', str(tahun))
    output_json = os.path.join(data_dir, f'rekap_pengadaan_{tahun}.json')

    # ---------------------------------------------------------
    # LOGIKA SKIP: HANYA melewati tahun n-2 (sudah final)
    # ---------------------------------------------------------
    if tahun == tahun_n2 and os.path.exists(output_dir_excel) and os.path.exists(output_json):
        file_sudah_ada = any(f.startswith('Paket Pengadaan Tahun') and f.endswith('.xlsx') for f in os.listdir(output_dir_excel))
        if file_sudah_ada:
            log_print(f"\n[SKIP] Data Tahun {tahun} sudah lengkap & final -> Lewati generate")
            return None
    # ---------------------------------------------------------

    log_print(f'\n{"="*55}\n   GENERATE DATA TAHUN {tahun}\n{"="*55}')
    def p(nama): return get_file_path(data_dir, nama, tahun)

    # ===== SUMBER DATA DASAR BARU (INAPROC) =====
    df_ina  = load_json(p('dashboard_realisasi_table'))

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

    # ===== INDEKS PENCOCOKAN KODE PAKET (dibuat SEBELUM filter batal, agar paket Inaproc tetap ketemu) =====
    idx_map = {
        'df2': build_row_index(df2, 'kd_nontender'),
        'df3': build_row_index(df3, 'kd_nontender_pct'),
        'df4': build_row_index(df4, 'kd_swakelola_pct'),
        'df5': build_row_index(df5, 'kd_tender'),
        'df6': build_row_index(df6, 'order_id'),
        'df7': build_row_index(df7, 'kd_paket'),   # V5: Kode Paket Inaproc = kd_paket (BUKAN no_paket)
    }

    # Filter lama (hanya dipakai untuk menentukan executed_rups / Pengumuman RUP)
    if not df2.empty and 'status_nontender' in df2.columns: df2 = df2[df2['status_nontender'] != 'Gagal/Batal']
    if not df5.empty and 'status_tender' in df5.columns: df5 = df5[df5['status_tender'] != 'Gagal/Batal']
    if not df3.empty and 'status_nontender_pct_ket' in df3.columns: df3 = df3[df3['status_nontender_pct_ket'].astype(str).str.strip() != 'Paket Dibatalkan']
    if not df6.empty and 'status' in df6.columns: df6 = df6[~df6['status'].isin(['CANCELLED_ON_NEGOTIATION', 'CANCELLED_ON_REVIEW', 'CANCELLED'])]

    def get_set(df, col):
        if df.empty or col not in df.columns: return set()
        return {norm_kd(x) for x in df[col].astype(str).str.split(';').explode() if norm_kd(x)}

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

    # ================== EKSTRAKSI NILAI SIKAP ==================
    def parse_sikap(df, key_col):
        res = {}
        if df.empty or key_col not in df.columns: return res
        for _, r in df.iterrows():
            for k in str(r.get(key_col, '')).split(';'):
                kd = norm_kd(k)
                if not kd: continue
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
    # ===========================================================

    def build_multi_kd_map(df, kd_col, val_col):
        m = {}
        if not df.empty and val_col in df.columns:
            for _, r in df.iterrows():
                for k in str(r.get(kd_col, '')).split(';'):
                    k = norm_kd(k)
                    if k: m[k] = r.get(val_col)
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

    # === FUNGSI HELPER NULL & MAP ===
    def get_mapped_val(kd_list, map_dict, is_date=False, is_numeric=False):
        if not kd_list: return "" # Jika tidak ada kode paket sama sekali, biarkan kosong

        for k in kd_list:
            if k in map_dict:
                # KONDISI 2: Paket SUDAH SAMPAI tahap ini (kodenya ada di dictionary tahap ini)
                v = map_dict[k]

                # Jika sudah sampai tahapnya tapi data benar-benar kosong dari LKPP
                if pd.isna(v) or v is None or str(v).strip() == "" or str(v).lower() in ["nan", "none"]:
                    return "tidak ada data (null)"

                if is_date:
                    t = format_tgl(v)
                    return t if t else "tidak ada data (null)"
                if is_numeric:
                    try: return float(v)
                    except: return "tidak ada data (null)"

                return str(v).strip()

        # KONDISI 1: paket belum sampai ke tahap ini
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
    # ==================================

    # =========================================================================
    # BLOK: PEMBERSIHAN DATA & PEMETAAN KAJI ULANG (TIDAK DIUBAH)
    # =========================================================================
    status_terakhir = {} # Menyimpan riwayat terakhir: {rup_id: {'tgl': '...', 'is_batal': True/False}}
    kaji_ulang_dict = {}

    if not df1_5.empty and 'kd_rup_lama' in df1_5.columns and 'kd_rup_baru' in df1_5.columns:
        for _, r in df1_5.iterrows():
            try:
                lama = int(float(str(r['kd_rup_lama']).strip()))
                baru = int(float(str(r['kd_rup_baru']).strip()))
                alasan = str(r.get('alasan_kajiulang', '')).lower()
                jenis_rev = str(r.get('jenis_revisi', '')).lower()
                tgl = str(r.get('tgl_kaji_ulang', ''))

                # Cek Indikator Batal dan Aktif (Saringan Kata Kematian & Hakim Waktu)
                kematian = ['batal', 'pembatalan', 'hapus']
                is_batal = any(k in alasan for k in kematian) or any(k in jenis_rev for k in kematian)

                # Pengaktifan HANYA SAH jika tidak ada unsur kematian di atas
                is_aktif = ('aktif' in alasan or 'pengaktifan' in jenis_rev) and not is_batal

                # Fungsi internal pencatat sejarah
                def catat_sejarah(k_id, batal_flag, tgl_str):
                    if k_id not in status_terakhir:
                        status_terakhir[k_id] = {'tgl': tgl_str, 'is_batal': batal_flag}
                    else:
                        # Update status HANYA JIKA tanggal kaji ulangnya lebih baru/sama
                        if tgl_str >= status_terakhir[k_id]['tgl']:
                            status_terakhir[k_id] = {'tgl': tgl_str, 'is_batal': batal_flag}

                # Tentukan nasib paket berdasarkan waktu
                if is_batal:
                    catat_sejarah(lama, True, tgl)
                    catat_sejarah(baru, True, tgl)
                elif is_aktif:
                    catat_sejarah(lama, False, tgl)
                    catat_sejarah(baru, False, tgl)

                # Masukkan ke kamus jika ini revisi normal (bukan batal)
                if lama != baru and not is_batal:
                    kaji_ulang_dict[lama] = baru
            except: pass

    # Blacklist HANYA untuk paket yang nasib TERAKHIRNYA adalah batal
    # (dipakai untuk baris Sumber 1 / 1_2 dan pelacakan Kode RUP Baru; TIDAK dipakai membuang baris Inaproc)
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
        """Sensor: Mengecek relasi maju-mundur antara dua kode RUP"""
        try:
            a, b = int(p1), int(p2)
            if a == b: return True
            # Cek a -> b
            curr, visited = a, set()
            while curr in kaji_ulang_dict and curr not in visited:
                visited.add(curr)
                curr = kaji_ulang_dict[curr]
                if curr == b: return True
            # Cek b -> a
            curr, visited = b, set()
            while curr in kaji_ulang_dict and curr not in visited:
                visited.add(curr)
                curr = kaji_ulang_dict[curr]
                if curr == a: return True
            return False
        except: return False

    def bersihkan_dan_ambil_kanan(raw_rup_str):
        if raw_rup_str is None or pd.isna(raw_rup_str) or not str(raw_rup_str).strip(): return ""
        parts = [norm_kd(p) for p in str(raw_rup_str).split(';')]
        parts = [p for p in parts if p]

        if len(parts) <= 1:
            return ";".join(parts)

        # TAHAP 1: PEMETAAN SILSILAH (Melihat Gambaran Besar)
        families = []
        for pp in parts:
            found_family = False
            for fam in families:
                if any(is_related(pp, member) for member in fam):
                    fam.append(pp)
                    found_family = True
                    break
            if not found_family:
                families.append([pp]) # Buat keluarga baru jika mandiri

        # TAHAP 2: KLASIFIKASI & EKSEKUSI
        # SKENARIO 1: 1 HISTORY PANJANG (Baik Genap/Ganjil)
        if len(families) == 1:
            return parts[-1] # Langsung comot paling kanan

        # SKENARIO 2: PASANGAN IDENTIK (Acak / Cermin / Berurutan)
        if len(parts) % 2 == 0:
            mid = len(parts) // 2
            kiri = parts[:mid]
            kanan = parts[mid:]

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

            if is_pasangan_identik:
                return ";".join(kanan) # Belah 2, Ambil mutlak Kanan

        # SKENARIO 3: LONCATAN / JOMBLO / CAMPURAN (A B C)
        rightmost_set = set(fam[-1] for fam in families)
        final_parts = []

        for pp in parts:
            if pp in rightmost_set:
                final_parts.append(pp)
                rightmost_set.remove(pp) # Kunci agar tidak duplikat

        return ";".join(final_parts)

    def cari_rup_baru(kode_rup_bersih):
        """Melacak RUP sampai ke ujung sejarahnya (Looping)"""
        if not kode_rup_bersih: return ""
        parts = kode_rup_bersih.split(';')
        results = []
        for pp in parts:
            if pp.isdigit():
                curr = int(pp)
                if curr in blacklist_rup: continue

                awal = curr
                visited = set()

                # Telusuri terus ke depan sampai ujung rantai
                while curr in kaji_ulang_dict and curr not in visited:
                    visited.add(curr)
                    next_rup = kaji_ulang_dict[curr]
                    # Berhenti jika ujungnya ternyata paket mati
                    if next_rup in blacklist_rup:
                        curr = None
                        break
                    curr = next_rup

                if curr is not None and curr != awal:
                    results.append(str(curr))
        return ";".join(list(dict.fromkeys(results)))
    # =========================================================================

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

            if tipe == 's1':
                v = map_pagu_s1.get(k_int, map_pagu_s1_3.get(k_int, 0))
            else:
                v = map_pagu_s1_2.get(k_int, map_pagu_s1_4.get(k_int, 0))

            if (pd.isna(v) or float(v) == 0) and k_int in kaji_ulang_dict:
                k_baru = kaji_ulang_dict[k_int]
                if tipe == 's1':
                    v = map_pagu_s1.get(k_baru, map_pagu_s1_3.get(k_baru, 0))
                else:
                    v = map_pagu_s1_2.get(k_baru, map_pagu_s1_4.get(k_baru, 0))

            total += float(v) if not pd.isna(v) else 0
        return total if total != 0 else None

    def build_df_map(df_source):
        if not df_source.empty and 'kd_rup' in df_source.columns:
            df_clean = df_source.dropna(subset=['kd_rup']).drop_duplicates(subset=['kd_rup']).copy()
            df_clean['kd_rup'] = df_clean['kd_rup'].apply(lambda x: int(float(str(x).strip())))
            return df_clean.set_index('kd_rup')
        return pd.DataFrame()

    df1_map, df1_2_map, df1_3_map, df1_4_map = build_df_map(df1), build_df_map(df1_2), build_df_map(df1_3), build_df_map(df1_4)
    map_status_konsol = {}
    if not df1_3.empty and 'kd_rup' in df1_3.columns and 'status_konsolidasi' in df1_3.columns:
        # Kunci diubah ke int lewat int(float(..)) agar tetap terbaca walau kolom kd_rup berubah jadi float (mis. "65400376.0")
        for k, v in zip(df1_3['kd_rup'], df1_3['status_konsolidasi']):
            try: k_int = int(float(str(k).strip()))
            except: continue
            if pd.notna(v) and str(v).strip(): map_status_konsol[k_int] = v

    def get_s1(kd, col, tipe='s1'):
        try:
            if kd is None or pd.isna(kd): return None
            kd_match = int(float(str(kd).strip()))
            if tipe == 's1':
                if not df1_map.empty and kd_match in df1_map.index and col in df1_map.columns and not pd.isna(df1_map.loc[kd_match, col]): return df1_map.loc[kd_match, col]
                if not df1_3_map.empty and kd_match in df1_3_map.index and col in df1_3_map.columns and not pd.isna(df1_3_map.loc[kd_match, col]): return df1_3_map.loc[kd_match, col]
            elif tipe == 's1_2':
                if not df1_2_map.empty and kd_match in df1_2_map.index and col in df1_2_map.columns and not pd.isna(df1_2_map.loc[kd_match, col]): return df1_2_map.loc[kd_match, col]
                if not df1_4_map.empty and kd_match in df1_4_map.index and col in df1_4_map.columns and not pd.isna(df1_4_map.loc[kd_match, col]): return df1_4_map.loc[kd_match, col]
            return None
        except: return None

    def clean_kd(v):
        if pd.isna(v) or str(v).strip() in ["", "nan", "None"]: return ""
        try: return str(int(float(v)))
        except: return str(v).strip()

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

    # Membangun kamus master
    map_prog = dict(zip(df_prog['kd_program_str'].astype(str).str.strip(), df_prog['nama_program'])) if not df_prog.empty and 'kd_program_str' in df_prog.columns else {}

    map_keg = {}
    if not df_keg.empty and 'kd_kegiatan' in df_keg.columns:
        for k, v in zip(df_keg['kd_kegiatan'], df_keg['nama_kegiatan']):
            k_cln = clean_kd(k)
            if k_cln: map_keg[k_cln] = str(v).strip()

    map_sub = {}
    if not df_sub.empty and 'kd_subkegiatan' in df_sub.columns:
        for k, v in zip(df_sub['kd_subkegiatan'], df_sub['nama_subkegiatan']):
            k_cln = clean_kd(k)
            if k_cln: map_sub[k_cln] = str(v).strip()

    def get_nama_program(mak_str):
        if not mak_str or pd.isna(mak_str): return ""
        maks = [m.strip() for m in str(mak_str).split(';')]
        progs = []
        for m in maks:
            parts = m.split('.')
            if len(parts) >= 3:
                kd_prog = f"{parts[0]}.{parts[1]}.{parts[2]}" # Mengambil 3 blok pertama (misal: 2.11.01)
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
                if str(kd) in map_master:
                    hasil.append(map_master[str(kd)])

        if not hasil: return ""
        unik = list(dict.fromkeys(hasil))
        return "; ".join(unik)

    def get_anggaran_multi(kd_list, map_ang, key):
        if not kd_list: return ""
        res = [map_ang[k][key] for k in kd_list if k in map_ang and map_ang[k][key]]
        return "; ".join(list(dict.fromkeys(res)))

    # ---------------------------------------------------------------
    # executed_rups: RUP yang SUDAH terealisasi (tidak boleh muncul lagi sebagai "Pengumuman RUP")
    # = logika lama (tabel transaksi) + seluruh Kode RUP dari Inaproc
    # ---------------------------------------------------------------
    executed_rups = set()

    def tambah_executed(raw_k):
        bersih = bersihkan_dan_ambil_kanan(raw_k)
        for k_str in bersih.split(';'):
            if k_str.strip().isdigit():
                curr = int(k_str.strip())
                executed_rups.add(curr)
                # Lacak sampai ujung riwayat agar RUP Baru tidak lolos ke Pengumuman RUP
                visited = set()
                while curr in kaji_ulang_dict and curr not in visited:
                    visited.add(curr)
                    curr = kaji_ulang_dict[curr]
                    executed_rups.add(curr)

    for df_trans, col_name in [(df2,'kd_rup_raw'), (df3,'kd_rup_raw'), (df4,'kd_rup_raw'), (df5,'kd_rup_raw'), (df6,'rup_code_raw'), (df7,'kd_rup_raw')]:
        if not df_trans.empty and col_name in df_trans.columns:
            for raw_k in df_trans[col_name]:
                tambah_executed(raw_k)

    if not df_ina.empty and 'kode_rup' in df_ina.columns:
        for raw_k in df_ina['kode_rup']:
            tambah_executed(raw_k)

    def format_satker(nama, kd):
        n = str(nama).strip() if pd.notna(nama) and str(nama).lower() not in ['nan', 'none', ''] else ""
        k = str(kd).strip() if pd.notna(kd) and str(kd).lower() not in ['nan', 'none', ''] else ""
        try: k = str(int(float(k)))
        except: pass
        if n and k: return f"{n} - {k}"
        return n if n else ""

    # =========================================================================
    # PETA PELENGKAP PER SUMBER (pencatatan non tender & swakelola)
    # =========================================================================
    map_s3_npwp, map_s3_alamat = {}, {}
    if not df3_1.empty and 'nama_penyedia' in df3_1.columns:
        kd_kolom_3_1 = next((c for c in df3_1.columns if c.lower().strip() in ['kd_nontender_pct', 'kd_nontender', 'kode_paket']), None)
        if kd_kolom_3_1:
            for pkt_id, group in df3_1.groupby(kd_kolom_3_1):
                npwp_set, alamat_set = [], []
                for _, row in group.iterrows():
                    npwp = str(row.get('npwp_penyedia', row.get('npwp', ''))).strip()
                    alamat = str(row.get('alamat_penyedia', row.get('alamat', ''))).strip()
                    if npwp and npwp.lower() not in ['', 'nan', 'none', '-'] and npwp not in npwp_set: npwp_set.append(npwp)
                    if alamat and alamat.lower() not in ['', 'nan', 'none', '-'] and alamat not in alamat_set: alamat_set.append(alamat)
                if npwp_set: map_s3_npwp[norm_kd(pkt_id)] = "; ".join(npwp_set)
                if alamat_set: map_s3_alamat[norm_kd(pkt_id)] = "; ".join(alamat_set)

    map_s4_npwp = {}
    if not df4_1.empty and 'nama_pelaksana' in df4_1.columns:
        kd_kolom_4_1 = next((c for c in df4_1.columns if c.lower().strip() in ['kd_swakelola_pct', 'kd_swakelola', 'kode_paket']), None)
        if kd_kolom_4_1:
            for pkt_id, group in df4_1.groupby(kd_kolom_4_1):
                npwp_set = []
                for _, row in group.iterrows():
                    npwp = str(row.get('npwp_pelaksana', row.get('npwp', ''))).strip()
                    if npwp and npwp.lower() not in ['', 'nan', 'none', '-'] and npwp not in npwp_set: npwp_set.append(npwp)
                if npwp_set: map_s4_npwp[norm_kd(pkt_id)] = "; ".join(npwp_set)

    # Cadangan Kode RUP: pencatatan non tender realisasi (kolom kd_rup_paket)
    map_s3_rup_realisasi = {}
    if not df3_1.empty and 'kd_rup_paket' in df3_1.columns:
        kd_kolom_rup = next((c for c in df3_1.columns if c.lower().strip() in ['kd_nontender_pct', 'kd_nontender', 'kode_paket']), None)
        if kd_kolom_rup:
            for pkt_id, group in df3_1.groupby(kd_kolom_rup):
                rups = []
                for x in group['kd_rup_paket']:
                    x = norm_kd(x)
                    if x and x not in rups: rups.append(x)
                if rups: map_s3_rup_realisasi[norm_kd(pkt_id)] = ";".join(rups)

    # =========================================================================
    # BAGIAN BARU: BANGUN BARIS REALISASI DARI INAPROC (DATA DASAR)
    # 9 kolom data dasar diisi dari Inaproc & TIDAK diubah lagi setelah ini
    # =========================================================================
    rantai_nt = [(set_kinerja_nt, 'Kinerja Dinilai'), (set_bapbast, 'BAPBAST'), (set_spmkspp, 'SPMKSPP'),
                 (set_kontrak, 'Kontrak'), (set_sppbj, 'SPPBJ'), (set_selesai, 'Non Tender Selesai')]
    rantai_t  = [(set_kinerja_t, 'Kinerja Dinilai'), (set_t_bapbast, 'BAPBAST'), (set_t_spmkspp, 'SPMKSPP'),
                 (set_t_kontrak, 'Kontrak'), (set_t_sppbj, 'SPPBJ'), (set_t_selesai, 'Tender Selesai')]

    def _peta_swakelola(df):
        if df.empty or 'kd_rup' not in df.columns or 'tipe_swakelola' not in df.columns: return {}
        return df.dropna(subset=['kd_rup']).drop_duplicates(subset=['kd_rup']).set_index('kd_rup')['tipe_swakelola'].to_dict()
    swakelola_map = _peta_swakelola(df1_2)
    swakelola_map_fallback = _peta_swakelola(df1_4)
    jumlah_metode_cadangan = 0
    jumlah_jenis_cadangan = 0

    data_inaproc = []
    jumlah_tidak_cocok = 0
    jumlah_rup_cadangan = {}
    perlu_dicek = []   # baris yang perlu dicek manual (Tidak Cocok / Kode RUP kosong)
    nama_tabel = {'df2': 'tender_non-tender-pengumuman (kd_nontender)', 'df3': 'tender_pencatatan-non-tender (kd_nontender_pct)',
                  'df4': 'tender_pencatatan-swakelola (kd_swakelola_pct)', 'df5': 'tender_pengumuman (kd_tender)',
                  'df6': 'ekatalog_paket-e-purchasing (order_id)', 'df7': 'ekatalog-archive_paket-e-purchasing (kd_paket)'}

    for _, ina in df_ina.iterrows():
        kode_paket = norm_kd(ina.get('kode_paket'))
        kd_list = [kode_paket] if kode_paket else []

        # --- Cari baris sumber berdasarkan Kode Paket ---
        tipe = resolve_tipe_sumber(ina.get('sumber_transaksi'))
        kandidat = [tipe] if tipe else ['df5', 'df2', 'df3', 'df4', 'df6', 'df7']
        r, tipe_found = None, None
        for t in kandidat:
            hit = idx_map[t].get(kode_paket)
            if hit is not None:
                r, tipe_found = hit, t
                break

        # --- Kode RUP: tetap lewat kode pembersihan lama ---
        # Urutan sumber: 1) Inaproc  2) tabel sumber (kd_rup / rup_code)  3) pencatatan non tender realisasi (kd_rup_paket)
        kode_rup_bersih = bersihkan_dan_ambil_kanan(ina.get('kode_rup'))
        asal_rup = 'Inaproc'
        if not kode_rup_bersih and r is not None:
            kode_rup_bersih = bersihkan_dan_ambil_kanan(r.get('rup_code' if tipe_found == 'df6' else 'kd_rup'))
            if kode_rup_bersih: asal_rup = 'Tabel sumber'
        if not kode_rup_bersih and kode_paket in map_s3_rup_realisasi:
            kode_rup_bersih = bersihkan_dan_ambil_kanan(map_s3_rup_realisasi[kode_paket])
            if kode_rup_bersih: asal_rup = 'Realisasi (kd_rup_paket)'
        cleaned_list = [int(i.strip()) for i in kode_rup_bersih.split(';') if i.strip().isdigit()]
        kode_rup_asli = ";".join(map(str, cleaned_list))
        kode_rup_baru = cari_rup_baru(kode_rup_asli)
        kd_lookup = cleaned_list[0] if cleaned_list else None
        if asal_rup != 'Inaproc':
            jumlah_rup_cadangan[asal_rup] = jumlah_rup_cadangan.get(asal_rup, 0) + 1
            tambah_executed(kode_rup_asli)   # RUP hasil cadangan juga dianggap sudah terealisasi

        total_nilai = val(ina, 'total_nilai', "")

        # --- Metode Pemilihan: jika kosong / "-" -> ambil seperti kode lama (dari tabel sumber, lalu RUP) ---
        metode = val(ina, 'metode_pengadaan')
        if norm_kd(metode) in ('', '-'):
            m_cad = ""
            if r is not None:
                if tipe_found == 'df4': m_cad = 'Swakelola'
                elif tipe_found in ('df6', 'df7'): m_cad = 'E-Purchasing'
                else: m_cad = val(r, 'mtd_pemilihan')
            if norm_kd(m_cad) in ('', '-'):
                m_cad = get_s1(kd_lookup, 'metode_pengadaan', 's1') or ""
            if norm_kd(m_cad) not in ('', '-'):
                metode = m_cad
                jumlah_metode_cadangan += 1

        # --- Jenis Pengadaan: jika kosong / "-" -> ambil seperti kode lama (dari tabel sumber, lalu RUP) ---
        jenis = val(ina, 'jenis_pengadaan')
        if norm_kd(jenis) in ('', '-'):
            j_cad = ""
            if r is not None:
                if tipe_found == 'df3': j_cad = val(r, 'kategori_pengadaan')
                elif tipe_found == 'df4':
                    tsw = val(r, 'tipe_swakelola', None)
                    if tsw is None: tsw = swakelola_map.get(kd_lookup, swakelola_map_fallback.get(kd_lookup))
                    try: j_cad = f"Swakelola {int(float(tsw))}"
                    except: j_cad = ""
                elif tipe_found in ('df6', 'df7'): j_cad = get_s1(kd_lookup, 'jenis_pengadaan', 's1') or ""
                else: j_cad = val(r, 'jenis_pengadaan')
            if norm_kd(j_cad) in ('', '-'):
                j_cad = get_s1(kd_lookup, 'jenis_pengadaan', 's1') or ""
            if norm_kd(j_cad) not in ('', '-'):
                jenis = j_cad
                jumlah_jenis_cadangan += 1

        # --- 9 KOLOM DATA DASAR (dari Inaproc) ---
        base = {
            'Kode Paket': kode_paket,
            'Kode RUP': kode_rup_asli,
            'Kode RUP Baru': kode_rup_baru,
            'Nama Instansi': val(ina, 'nama_instansi'),
            'Satuan Kerja': val(ina, 'nama_satuan_kerja'),
            'Nama Paket': val(ina, 'nama_paket'),
            'Metode Pemilihan': metode,
            'Jenis Pengadaan': jenis,
            'Nama Penyedia': val(ina, 'nama_penyedia'),
            'Nilai Hasil Pemilihan': total_nilai,
        }

        det = {c: "" for c in KOLOM_DETAIL}

        # Kolom berbasis Kode RUP (berlaku untuk semua, termasuk "Tidak Cocok")
        is_swa = (tipe_found == 'df4') or ('swakelola' in str(metode).lower())
        det['PDN'] = get_s1(kd_lookup, 'status_pdn', 's1')
        det['UKM'] = get_s1(kd_lookup, 'status_ukm', 's1')
        det['Nilai Pagu RUP'] = get_pagu_multi(cleaned_list, 's1_2' if is_swa else 's1')
        det['Sumber Dana'] = val(ina, 'sumber_dana')
        det['Tahun Anggaran'] = val(ina, 'tahun_anggaran')

        if r is None:
            # Kode Paket tidak ditemukan di tabel pencarinya
            det['Status'] = TEKS_TIDAK_COCOK
            jumlah_tidak_cocok += 1

        elif tipe_found == 'df2':
            status = hitung_status(r.get('status_nontender'), kd_list, rantai_nt)
            det.update({
                'Tahun Anggaran': val(r, 'tahun_anggaran', det['Tahun Anggaran']),
                'Sumber Dana': val(r, 'sumber_dana', det['Sumber Dana']), 'MAK': val(r, 'mak'),
                'Kualifikasi Paket': get_mapped_val(kd_list, map_nt_kualifikasi),
                'Nilai Negosiasi': get_mapped_val(kd_list, map_nt_nil_neg, is_numeric=True),
                'No SPPBJ': get_mapped_val(kd_list, map_nt_no_sppbj), 'Tanggal SPPBJ': get_mapped_val(kd_list, map_nt_tgl_sppbj, is_date=True),
                'No Kontrak': get_mapped_val(kd_list, map_nt_no_kontrak), 'Tanggal Kontrak': get_mapped_val(kd_list, map_nt_tgl_kontrak, is_date=True),
                'Jenis Kontrak': get_mapped_val(kd_list, map_nt_jenis_kontrak), 'Nilai Kontrak': get_mapped_val(kd_list, map_nt_nil_kontrak, is_numeric=True),
                'Status Kontrak': get_mapped_val(kd_list, map_nt_status_kontrak), 'No SPMK SPP': get_mapped_val(kd_list, map_nt_no_spmkspp),
                'Tanggal SPMK SPP': get_mapped_val(kd_list, map_nt_tgl_spmkspp, is_date=True), 'No BAST': get_mapped_val(kd_list, map_nt_no_bast),
                'Tanggal BAST': get_mapped_val(kd_list, map_nt_tgl_bast, is_date=True), 'No BAP': get_mapped_val(kd_list, map_nt_no_bap),
                'Tanggal BAP': get_mapped_val(kd_list, map_nt_tgl_bap, is_date=True),
                'NPWP 15': next((map_nt_npwp[k] for k in kd_list if k in map_nt_npwp), ""), 'NPWP 16': get_mapped_val(kd_list, map_nt_npwp16),
                'Alamat': next((map_nt_alamat[k] for k in kd_list if k in map_nt_alamat), ""),
                'Status': status, 'Nilai HPS': val(r, 'hps'),
                'Nilai PDN': next((map_nt_pdn[k] for k in kd_list if k in map_nt_pdn), "N/A"),
                'Nilai UMK': next((map_nt_umk[k] for k in kd_list if k in map_nt_umk), "N/A"),
                'Cara Pengadaan': 'Non Tender', 'Sumber': 'Sumber 2'
            })

        elif tipe_found == 'df3':
            det.update({
                'Tahun Anggaran': val(r, 'tahun_anggaran', det['Tahun Anggaran']),
                'Sumber Dana': val(r, 'sumber_dana', det['Sumber Dana']),
                'MAK': get_anggaran_multi(cleaned_list, map_ang_p, 'mak'),
                'NPWP 15': map_s3_npwp.get(kode_paket, ""), 'Alamat': map_s3_alamat.get(kode_paket, ""),
                'Status': val(r, 'status_nontender_pct_ket'),
                'Nilai PDN': val(r, 'nilai_pdn_pct'), 'Nilai UMK': val(r, 'nilai_umk_pct'),
                'Cara Pengadaan': 'Pencatatan Non Tender', 'Sumber': 'Sumber 3'
            })

        elif tipe_found == 'df4':
            det.update({
                'Tahun Anggaran': val(r, 'tahun_anggaran', det['Tahun Anggaran']),
                'Sumber Dana': val(r, 'sumber_dana', det['Sumber Dana']),
                'MAK': get_anggaran_multi(cleaned_list, map_ang_s, 'mak'),
                'PDN': "PDN" if to_float(r.get('nilai_pdn_pct', 0)) != 0 else "Tidak",
                'UKM': "UKM" if to_float(r.get('nilai_umk_pct', 0)) != 0 else "Tidak",
                'NPWP 15': map_s4_npwp.get(kode_paket, ""),
                'Status': val(r, 'status_swakelola_pct_ket'),
                'Nilai PDN': val(r, 'nilai_pdn_pct'), 'Nilai UMK': val(r, 'nilai_umk_pct'),
                'Cara Pengadaan': 'Pencatatan Swakelola', 'Sumber': 'Sumber 4'
            })

        elif tipe_found == 'df5':
            status = hitung_status(r.get('status_tender'), kd_list, rantai_t)
            mtd = str(r.get('mtd_pemilihan', '')).strip().lower()
            kat_metode = 'Seleksi' if mtd == 'seleksi' else ('Tender Cepat' if mtd == 'tender cepat' else 'Tender')
            det.update({
                'Tahun Anggaran': val(r, 'tahun_anggaran', det['Tahun Anggaran']),
                'Sumber Dana': val(r, 'sumber_dana', det['Sumber Dana']),
                'MAK': next((map_t_mak[k] for k in kd_list if k in map_t_mak), ""),
                'Kualifikasi Paket': get_mapped_val(kd_list, map_t_kualifikasi),
                'Nilai Negosiasi': get_mapped_val(kd_list, map_t_nil_neg, is_numeric=True),
                'No SPPBJ': get_mapped_val(kd_list, map_t_no_sppbj), 'Tanggal SPPBJ': get_mapped_val(kd_list, map_t_tgl_sppbj, is_date=True),
                'No Kontrak': get_mapped_val(kd_list, map_t_no_kontrak), 'Tanggal Kontrak': get_mapped_val(kd_list, map_t_tgl_kontrak, is_date=True),
                'Jenis Kontrak': get_mapped_val(kd_list, map_t_jenis_kontrak), 'Nilai Kontrak': get_mapped_val(kd_list, map_t_nil_kontrak, is_numeric=True),
                'Status Kontrak': get_mapped_val(kd_list, map_t_status_kontrak), 'No SPMK SPP': get_mapped_val(kd_list, map_t_no_spmkspp),
                'Tanggal SPMK SPP': get_mapped_val(kd_list, map_t_tgl_spmkspp, is_date=True), 'No BAST': get_mapped_val(kd_list, map_t_no_bast),
                'Tanggal BAST': get_mapped_val(kd_list, map_t_tgl_bast, is_date=True), 'No BAP': get_mapped_val(kd_list, map_t_no_bap),
                'Tanggal BAP': get_mapped_val(kd_list, map_t_tgl_bap, is_date=True),
                'NPWP 15': next((map_t_npwp[k] for k in kd_list if k in map_t_npwp), ""), 'NPWP 16': get_mapped_val(kd_list, map_t_npwp16),
                'Alamat': next((map_t_alamat[k] for k in kd_list if k in map_t_alamat), ""),
                'Status': status, 'Nilai HPS': val(r, 'hps'),
                'Nilai PDN': next((map_t_pdn[k] for k in kd_list if k in map_t_pdn), "N/A"),
                'Nilai UMK': next((map_t_umk[k] for k in kd_list if k in map_t_umk), "N/A"),
                'Cara Pengadaan': kat_metode, 'Sumber': 'Sumber 5'
            })

        elif tipe_found == 'df6':
            nilai_hasil = total_nilai
            status_pdn_katalog, status_ukm_katalog = det['PDN'], det['UKM']
            nilai_pdn_val = nilai_hasil if str(status_pdn_katalog).strip().upper() == 'PDN' else 0
            nilai_umk_val = nilai_hasil if str(status_ukm_katalog).strip().upper() == 'UKM' else 0
            det.update({
                'Tahun Anggaran': val(r, 'fiscal_year', det['Tahun Anggaran']),
                'Sumber Dana': val(r, 'funding_source', det['Sumber Dana']), 'MAK': val(r, 'mak'),
                'No Kontrak': f"EP-{kode_paket}" if kode_paket else "",
                'NPWP 15': str(val(r, 'npwp_penyedia', val(r, 'npwp', ''))),
                'Alamat': str(val(r, 'alamat_penyedia', val(r, 'alamat', ''))),
                'Status': val(r, 'status'),
                'Nilai PDN': nilai_pdn_val, 'Nilai UMK': nilai_umk_val,
                'Cara Pengadaan': 'E-Purchasing V6', 'Sumber': 'Sumber 6'
            })

        elif tipe_found == 'df7':
            nilai_hasil = total_nilai
            status_pdn_katalog, status_ukm_katalog = det['PDN'], det['UKM']
            nilai_pdn_val = nilai_hasil if str(status_pdn_katalog).strip().upper() == 'PDN' else 0
            nilai_umk_val = nilai_hasil if str(status_ukm_katalog).strip().upper() == 'UKM' else 0
            paket_status_str = r.get('paket_status_str')
            status_s7 = r.get('status_paket') if paket_status_str == "Paket Proses" else paket_status_str
            det.update({
                'Tahun Anggaran': val(r, 'tahun_anggaran', det['Tahun Anggaran']),
                'Sumber Dana': val(r, 'nama_sumber_dana', det['Sumber Dana']),
                'MAK': get_anggaran_multi(cleaned_list, map_ang_p, 'mak'),
                'NPWP 15': str(val(r, 'npwp_penyedia', val(r, 'npwp', ''))),
                'Alamat': str(val(r, 'alamat_penyedia', val(r, 'alamat', ''))),
                'Status': status_s7,
                'Nilai PDN': nilai_pdn_val, 'Nilai UMK': nilai_umk_val,
                'Cara Pengadaan': 'E-Purchasing V5', 'Sumber': 'Sumber 7'
            })

        row_out = {}
        row_out.update(det)
        row_out.update(base)            # data dasar Inaproc selalu menang
        row_out['Status Realisasi'] = 'Realisasi'
        data_inaproc.append(row_out)

        masalah = []
        if r is None:
            masalah.append("Kode Paket tidak ditemukan di " + (nama_tabel[tipe] if tipe else "semua tabel sumber"))
        if not kode_rup_asli:
            masalah.append("Kode RUP kosong (Inaproc, tabel sumber, dan kd_rup_paket)")
        if masalah:
            perlu_dicek.append({
                'Masalah': " | ".join(masalah), 'Kode Paket': kode_paket, 'Kode RUP': kode_rup_asli,
                'Sumber Transaksi (Inaproc)': val(ina, 'sumber_transaksi'), 'Tahapan (Inaproc)': val(ina, 'tahapan_pengadaan'),
                'Status Paket (Inaproc)': val(ina, 'status_paket'), 'Satuan Kerja': base['Satuan Kerja'],
                'Nama Paket': base['Nama Paket'], 'Nilai Hasil Pemilihan': base['Nilai Hasil Pemilihan']
            })

    df_inaproc_rows = pd.DataFrame(data_inaproc)
    if jumlah_metode_cadangan:
        log_print(f"ℹ️ Tahun {tahun}: {jumlah_metode_cadangan} baris Metode Pemilihan diisi dari cadangan (Inaproc kosong / '-')")
    if jumlah_jenis_cadangan:
        log_print(f"ℹ️ Tahun {tahun}: {jumlah_jenis_cadangan} baris Jenis Pengadaan diisi dari cadangan (Inaproc kosong)")
    for asal, jml in jumlah_rup_cadangan.items():
        log_print(f"ℹ️ Tahun {tahun}: {jml} baris Kode RUP-nya diisi dari cadangan: {asal}")
    df_perlu_dicek = pd.DataFrame(perlu_dicek)
    if jumlah_tidak_cocok:
        log_print(f"ℹ️ Tahun {tahun}: {jumlah_tidak_cocok} dari {len(data_inaproc)} baris Inaproc berstatus '{TEKS_TIDAK_COCOK}' (Kode Paket tidak ditemukan di tabel pencarinya)")
    if perlu_dicek:
        log_print(f"🔎 Tahun {tahun}: {len(perlu_dicek)} baris PERLU DICEK MANUAL -> lihat sheet '3. Perlu Dicek' di Excel.")

    # =========================================================================
    # BARIS "PENGUMUMAN RUP" (BELUM REALISASI) - DARI SUMBER 1 & 1_2 (LOGIKA LAMA)
    # =========================================================================
    swakelola_map = df1_2.set_index('kd_rup')['tipe_swakelola'] if (not df1_2.empty and 'tipe_swakelola' in df1_2.columns) else {}

    data_s1_2 = []
    for _, r in df1_2.iterrows():
        kd = r.get('kd_rup')
        kd_int = None
        try:
            kd_int = int(float(str(kd).strip()))
            if kd_int in blacklist_rup: continue
        except: pass

        if kd not in executed_rups:
            jenis = f"Swakelola {int(swakelola_map[kd])}" if kd in swakelola_map else "N/A"
            kode_rup_baru_s1 = cari_rup_baru(str(kd)) if pd.notna(kd) else ""
            data_s1_2.append({
                'Kode Paket': pd.NA, 'Kode RUP': kd, 'Kode RUP Baru': kode_rup_baru_s1,
                'Nama Instansi': 'KAB. KOTAWARINGIN BARAT', 'Satuan Kerja': format_satker(r.get('nama_satker'), r.get('kd_satker_str')), 'Nama Paket': r.get('nama_paket'), 'Metode Pemilihan': 'Swakelola',
                'Jenis Pengadaan': jenis, 'Tahun Anggaran': r.get('tahun_anggaran'), 'Sumber Dana': map_ang_s.get(kd_int, {}).get('sd', ''), 'MAK': map_ang_s.get(kd_int, {}).get('mak', ''), 'PDN': None, 'UKM': None,
                'Kualifikasi Paket': "", 'Nilai Pagu RUP': r.get('pagu'), 'Nilai Hasil Pemilihan': "",
                'Nilai Negosiasi': "", 'No SPPBJ': "", 'Tanggal SPPBJ': "", 'No Kontrak': "", 'Tanggal Kontrak': "", 'Jenis Kontrak': "", 'Nilai Kontrak': "",
                'Status Kontrak': "", 'No SPMK SPP': "", 'Tanggal SPMK SPP': "", 'No BAST': "", 'Tanggal BAST': "", 'No BAP': "", 'Tanggal BAP': "",
                'Nama Penyedia': "", 'NPWP 15': "", 'NPWP 16': "", 'Alamat': "", 'Status': 'Pengumuman RUP', 'Nilai HPS': pd.NA, 'Nilai PDN': pd.NA, 'Nilai UMK': pd.NA,
                'Cara Pengadaan': 'Swakelola', 'Sumber': 'Sumber 1_2', 'Status Realisasi': 'Belum Realisasi'
            })
    df_s1_2 = pd.DataFrame(data_s1_2).drop_duplicates(ignore_index=True) if data_s1_2 else pd.DataFrame()

    data_s1 = []
    for _, r in df1.iterrows():
        kd = r.get('kd_rup')
        kd_int = None
        try:
            kd_int = int(float(str(kd).strip()))
            if kd_int in blacklist_rup: continue
        except: pass

        if kd not in executed_rups:
            kode_rup_baru_s1 = cari_rup_baru(str(kd)) if pd.notna(kd) else ""
            data_s1.append({
                'Kode Paket': pd.NA, 'Kode RUP': kd, 'Kode RUP Baru': kode_rup_baru_s1,
                'Nama Instansi': 'KAB. KOTAWARINGIN BARAT', 'Satuan Kerja': format_satker(r.get('nama_satker'), r.get('kd_satker_str')), 'Nama Paket': r.get('nama_paket'), 'Metode Pemilihan': r.get('metode_pengadaan'),
                'Jenis Pengadaan': r.get('jenis_pengadaan'), 'Tahun Anggaran': r.get('tahun_anggaran'), 'Sumber Dana': map_ang_p.get(kd_int, {}).get('sd', ''), 'MAK': map_ang_p.get(kd_int, {}).get('mak', ''), 'PDN': 'PDN' if r.get('status_pdn')=='PDN' else 'Non-PDN',
                'UKM': 'UKM' if r.get('status_ukm')=='UKM' else 'Non-UKM',
                'Kualifikasi Paket': "", 'Nilai Pagu RUP': r.get('pagu'), 'Nilai Hasil Pemilihan': "",
                'Nilai Negosiasi': "", 'No SPPBJ': "", 'Tanggal SPPBJ': "", 'No Kontrak': "", 'Tanggal Kontrak': "", 'Jenis Kontrak': "", 'Nilai Kontrak': "",
                'Status Kontrak': "", 'No SPMK SPP': "", 'Tanggal SPMK SPP': "", 'No BAST': "", 'Tanggal BAST': "", 'No BAP': "", 'Tanggal BAP': "",
                'Nama Penyedia': "", 'NPWP 15': "", 'NPWP 16': "", 'Alamat': "", 'Status': 'Pengumuman RUP', 'Nilai HPS': pd.NA, 'Nilai PDN': pd.NA, 'Nilai UMK': pd.NA,
                'Cara Pengadaan': r.get('metode_pengadaan'), 'Sumber': 'Sumber 1', 'Status Realisasi': 'Belum Realisasi'
            })
    df_s1 = pd.DataFrame(data_s1).drop_duplicates(ignore_index=True) if data_s1 else pd.DataFrame()

    # Baris Inaproc TIDAK di-drop_duplicates (jumlah baris harus sesuai Inaproc)
    bagian = [d for d in [df_inaproc_rows, df_s1_2, df_s1] if not d.empty]
    final_df = pd.concat(bagian, ignore_index=True) if bagian else pd.DataFrame()
    final_df = final_df.map(lambda x: re.sub(r'[\x00-\x1F]', '', str(x)) if isinstance(x, str) else x)

    def get_status_konsolidasi(rup_val):
        if pd.isna(rup_val) or not str(rup_val).strip(): return ""
        valid_rups = [k for k in str(rup_val).split(';') if k.strip().isdigit()]
        if len(valid_rups) > 1: return "Konsolidasi"
        elif len(valid_rups) == 1: return map_status_konsol.get(int(valid_rups[0]), "")
        return ""
    final_df['Status Konsolidasi'] = final_df['Kode RUP'].apply(get_status_konsolidasi)

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
        'Cara Pengadaan', 'Sumber', 'Status Realisasi'
    ]
    final_df = final_df.reindex(columns=cols).fillna("")

    # Isi Sumber Dana, MAK, Program, Kegiatan, Sub Kegiatan, & Kinerja berdasarkan Kode RUP
    # (9 kolom data dasar Inaproc TIDAK disentuh di loop ini)
    for idx, row in final_df.iterrows():
        kd_rup_str = str(row['Kode RUP'])
        kd_baru_str = str(row['Kode RUP Baru'])

        cleaned_list = [int(i.strip()) for i in kd_rup_str.split(';') if i.strip().isdigit()]
        list_baru = [int(i.strip()) for i in kd_baru_str.split(';') if i.strip().isdigit()]

        is_swakelola = ('Swakelola' in str(row['Sumber']) or 'Swakelola' in str(row['Jenis Pengadaan'])
                        or 'Sumber 4' in str(row['Sumber']) or 'Sumber 1_2' in str(row['Sumber'])
                        or str(row['Metode Pemilihan']).strip().lower() == 'swakelola')
        map_ang_target = map_ang_s if is_swakelola else map_ang_p

        # 1. Tarik Sumber Dana terpusat dari Anggaran
        sd = get_anggaran_multi(cleaned_list, map_ang_target, 'sd')
        if not sd and list_baru: sd = get_anggaran_multi(list_baru, map_ang_target, 'sd')
        if sd: final_df.at[idx, 'Sumber Dana'] = sd

        # 2. Tarik MAK terpusat dari Anggaran
        mak = get_anggaran_multi(cleaned_list, map_ang_target, 'mak')
        if not mak and list_baru: mak = get_anggaran_multi(list_baru, map_ang_target, 'mak')
        if mak: final_df.at[idx, 'MAK'] = mak

        # 3. Tarik Program menggunakan MAK yang baru ditarik
        final_df.at[idx, 'Nama Program'] = get_nama_program(final_df.at[idx, 'MAK'])

        # 4. Tarik Kegiatan
        keg = get_kegiatan_sub(cleaned_list, is_swakelola, is_sub=False)
        if not keg and list_baru: keg = get_kegiatan_sub(list_baru, is_swakelola, is_sub=False)
        final_df.at[idx, 'Nama Kegiatan'] = keg

        # 5. Tarik Sub Kegiatan
        sub = get_kegiatan_sub(cleaned_list, is_swakelola, is_sub=True)
        if not sub and list_baru: sub = get_kegiatan_sub(list_baru, is_swakelola, is_sub=True)
        final_df.at[idx, 'Nama Sub Kegiatan'] = sub

        # 6. Tarik Penilaian Kinerja (Sumber 2 & 5)
        sumber = str(row['Sumber'])
        kual, bia, wak, lay, s_ang, s_kat = "", "", "", "", "", ""
        if "Sumber 2" in sumber:
            kd_nt_list = [norm_kd(i) for i in str(row['Kode Paket']).split(';')] if str(row['Kode Paket']).strip() else []
            kual, bia, wak, lay, s_ang, s_kat = get_kinerja_vals(kd_nt_list, map_sikap_nt)
        elif "Sumber 5" in sumber:
            kd_t_list = [norm_kd(i) for i in str(row['Kode Paket']).split(';')] if str(row['Kode Paket']).strip() else []
            kual, bia, wak, lay, s_ang, s_kat = get_kinerja_vals(kd_t_list, map_sikap_t)

        if s_ang != "":
            final_df.at[idx, 'Penilaian Kualitas dan kuantitas dengan indikator kesesuaian'] = kual
            final_df.at[idx, 'Penilaian Biaya dengan indikator kemampuan pengendalian biaya'] = bia
            final_df.at[idx, 'Penilaian Waktu dengan indikator ketepatan'] = wak
            final_df.at[idx, 'Penilaian Layanan dengan indikator komunikasi dan tingkat respon'] = lay
            final_df.at[idx, 'Kinerja Penyedia dalam angka'] = s_ang
            final_df.at[idx, 'Kinerja Penyedia dalam kata'] = s_kat

    final_df['PDN'] = final_df['PDN'].replace("", "N/A")
    final_df['UKM'] = final_df['UKM'].replace("", "N/A")
    final_df['Status'] = final_df['Status'].apply(lambda x: str(x).replace('_', ' ') if pd.notna(x) else x)

    def safe_numeric(v):
        if v in ("N/A", "", None) or pd.isna(v): return 0.0
        try: return float(v)
        except: return 0.0

    def aggregate_text(series):
        valid = [str(x).strip() for x in series if pd.notna(x) and str(x).strip() not in ("N/A", "")]
        return "; ".join(list(dict.fromkeys(valid)))

    def aggregate_raw_status(group, pagu_rup, sum_hasil):
        def get_score(s_text):
            s_low = str(s_text).strip().lower()
            if 'tidak cocok' in s_low: return 0
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

    df_agregasi = final_df.copy()
    for c in ['Nilai HPS', 'Nilai Hasil Pemilihan', 'Nilai PDN', 'Nilai UMK', 'Nilai Pagu RUP', 'Nilai Negosiasi', 'Nilai Kontrak']:
        df_agregasi[c] = df_agregasi[c].apply(safe_numeric)

    # Baris tanpa Kode RUP tidak digabung jadi satu grup "UNKNOWN", tapi per Kode Paket
    df_agregasi['Kode RUP_Group'] = [
        (str(rup).split(';')[0].strip() if str(rup).strip() else f"TANPA_RUP_{kp if str(kp).strip() else i}")
        for i, (rup, kp) in enumerate(zip(df_agregasi['Kode RUP'], df_agregasi['Kode Paket']))
    ]
    grouped = df_agregasi.groupby('Kode RUP_Group')

    rekap_data = []
    for kode_rup, group in grouped:
        first_row = group.iloc[0]
        sum_hps, sum_hasil = group['Nilai HPS'].sum(), group['Nilai Hasil Pemilihan'].sum()
        sum_pdn, sum_umk = group['Nilai PDN'].sum(), group['Nilai UMK'].sum()
        sum_neg = group['Nilai Negosiasi'].sum()
        sum_kon = group['Nilai Kontrak'].sum()
        pagu_rup = first_row['Nilai Pagu RUP']
        status_realisasi = 'Realisasi' if (group['Status Realisasi'] == 'Realisasi').any() else 'Belum Realisasi'

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
            'Cara Pengadaan': aggregate_text(group['Cara Pengadaan']), 'Sumber': aggregate_text(group['Sumber']),
            'Status Realisasi': status_realisasi
        })
    df_rekap = pd.DataFrame(rekap_data)

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(final_df.to_dict(orient='records'), f, ensure_ascii=False, separators=(',', ':'))
    log_print(f"JSON Rekap sukses dibuat: {output_json}")

    kolom_angka_baku = ['Nilai Pagu RUP', 'Nilai Hasil Pemilihan', 'Nilai Negosiasi', 'Nilai Kontrak', 'Nilai HPS', 'Nilai PDN', 'Nilai UMK']
    excel_df_detail = final_df.copy()
    for col in kolom_angka_baku:
        if col in excel_df_detail.columns:
            excel_df_detail[col] = excel_df_detail[col].apply(lambda x: safe_numeric(x) if safe_numeric(x) != 0 else "")

    tahun_label = str(df1['tahun_anggaran'].iloc[0]) if not df1.empty and 'tahun_anggaran' in df1.columns else str(tahun)
    nama_file_excel = f'Paket Pengadaan Tahun {tahun_label} ({datetime.now().strftime("%Y-%m-%d")}).xlsx'
    output_dir_excel = os.path.join(BASE_DIR, 'output', 'pengadaan', str(tahun))
    os.makedirs(output_dir_excel, exist_ok=True)
    output_excel_path = os.path.join(output_dir_excel, nama_file_excel)

    # === SIMPAN & STYLING EXCEL DENGAN PENGAMAN AUTO-CLOSE ===
    berhasil_simpan = False
    while not berhasil_simpan:
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                df_rekap.to_excel(writer, index=False, sheet_name='1. Rekap per RUP')
                excel_df_detail.to_excel(writer, index=False, sheet_name='2. Detail per Paket')
                if not df_perlu_dicek.empty:
                    df_perlu_dicek.to_excel(writer, index=False, sheet_name='3. Perlu Dicek')

            wb = load_workbook(output_excel_path)
            header_fill = PatternFill('solid', start_color='1F4E79')
            header_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
            header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            data_font = Font(name='Arial', size=10)
            fill_putih = PatternFill('solid', start_color='FFFFFF')
            fill_biru_muda = PatternFill('solid', start_color='DCE6F1')
            border_thin = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

            def style_sheet(ws_name, df_ref, dict_lebar, list_kolom_angka):
                ws = wb[ws_name]
                for i, col in enumerate(df_ref.columns, start=1): ws.column_dimensions[get_column_letter(i)].width = 15 # Paksa semua ukuran menjadi 15
                for cell in ws[1]: cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, header_align, border_thin
                ws.row_dimensions[1].height = 32
                for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
                    fill = fill_putih if row_idx % 2 == 0 else fill_biru_muda
                    for cell in row:
                        col_name = df_ref.columns[cell.column - 1]
                        cell.font, cell.fill, cell.border = data_font, fill, border_thin
                        if col_name in list_kolom_angka and str(cell.value).replace('.', '', 1).isdigit():
                            cell.number_format = '#,##0'
                            cell.alignment = Alignment(horizontal='right', vertical='center')
                        else: cell.alignment = Alignment(vertical='center', wrap_text=False)
                ws.freeze_panes = 'A2'
                ws.auto_filter.ref = ws.dimensions

            lebar_baku = {} # Dikosongkan karena tidak lagi diperlukan
            style_sheet('1. Rekap per RUP', df_rekap, lebar_baku, kolom_angka_baku)
            style_sheet('2. Detail per Paket', excel_df_detail, lebar_baku, kolom_angka_baku)
            if not df_perlu_dicek.empty:
                style_sheet('3. Perlu Dicek', df_perlu_dicek, lebar_baku, ['Nilai Hasil Pemilihan'])
            wb.save(output_excel_path)
            berhasil_simpan = True
            log_print(f"✅ EXCEL Pengadaan sukses dibuat: {output_excel_path}")

        except PermissionError:
            import time
            log_print(f"⚠️ Akses Ditolak: File {nama_file_excel} sedang terbuka.")
            tutup_sukses = False
            try:
                import win32com.client
                excel_app = win32com.client.GetActiveObject("Excel.Application")
                for wb_app in excel_app.Workbooks:
                    if wb_app.Name == nama_file_excel:
                        log_print(f"Menutup otomatis file: {nama_file_excel}...")
                        wb_app.Close(SaveChanges=False)
                        tutup_sukses = True
                        break
            except Exception: pass

            if not tutup_sukses:
                log_print("Memunculkan pop-up peringatan...")
                os.system(f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""File {nama_file_excel} sedang terbuka. Tutup file di Excel agar proses berlanjut!"", 5, ""Peringatan Excel"", 48)(window.close)")')
            time.sleep(3)
    # ========================================================

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
                        log_print(f"Menutup otomatis file: {nama_file_master}...")
                        wb_app.Close(SaveChanges=False)
                        tutup_sukses = True
                        break
            except Exception:
                pass

            if not tutup_sukses:
                log_print("Memunculkan pop-up peringatan...")
                os.system(f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup(""File {nama_file_master} sedang terbuka. Tutup file di Excel agar proses berlanjut!"", 5, ""Peringatan Excel"", 48)(window.close)")')
            time.sleep(3)
        except Exception as e:
            log_print(f"⚠️ Gagal menyalin EXCEL Master (Error Lain): {str(e)}")
            break # Keluar dari loop jika errornya bukan karena file terbuka

    kelola_arsip_bulanan(output_dir_excel, tahun)
    update_daftar_arsip_json(output_dir_excel)

    log_print(f'SELESAI GENERATE TAHUN {tahun} | Total data: {len(final_df)}')
    return len(final_df)


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
                log_print(f"⚠️ Koneksi Telegram lambat. Mencoba ulang ({percobaan}/3)...")
                time.sleep(5)
            else:
                log_print(f"🚨 GAGAL KONEKSI KE TELEGRAM (JARINGAN PUTUS): {str(e)}")


if __name__ == '__main__':
    log_print("\n" + "="*55)
    log_print(f"START GENERATE PENGADAAN (INAPROC) {get_waktu_indonesia()}")
    log_print("="*55)

    # 1. BACA ERROR DI AWAL
    daftar_error_api = []
    path_error = os.path.join(BASE_DIR, 'scripts', 'pengadaan', 'error_api_pengadaan.json')
    if os.path.exists(path_error):
        try:
            with open(path_error, 'r', encoding='utf-8') as f: daftar_error_api = json.load(f)
            os.remove(path_error)
        except: pass

    # 2. CEK TOTAL URL (Disesuaikan dengan tahun yang tidak di-skip)
    total_target = 0
    path_url = os.path.join(BASE_DIR, 'scripts', 'pengadaan', 'url_pengadaan.txt')
    if os.path.exists(path_url):
        with open(path_url, 'r', encoding='utf-8') as f:
            jumlah_url = len([line for line in f if line.strip()])

        tahun_diproses = 0
        for t in daftar_tahun:
            if t == tahun_n2 and os.path.exists(os.path.join(BASE_DIR, 'data', str(t), f'rekap_pengadaan_{t}.json')):
                continue
            tahun_diproses += 1

        total_target = jumlah_url * tahun_diproses

    # 3. LOGIKA BERHENTI JIKA GAGAL TOTAL
    if total_target > 0 and len(daftar_error_api) >= total_target:
        pesan_gagal = f"🚨 LAPORAN UPDATE (PENGADAAN) 🚨\n\n⚠️ GAGAL TOTAL DOWNLOAD API!\nTidak ada data baru. Proses dihentikan.\n\n📊 STATUS EKSEKUSI:\n🔹 Generate Pengadaan: Batal (API Error)\n🔹 Generate TEPRA: Batal (API Error)\n\n📅 Waktu: {get_waktu_indonesia()}"
        kirim_telegram_aman(pesan_gagal)
        log_print("GAGAL TOTAL. Skrip berhenti.")
        sys.exit(0)

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
    # ----------------------

    # 1. PUSH KE GITHUB
    git_sukses, pesan_git = sync_to_github()

    # --- MENGHITUNG DURASI TOTAL & AUTO DELETE ---
    durasi_str = "Tidak diketahui"
    file_start = os.path.join(BASE_DIR, 'tools', 'start_time_pengadaan.txt')
    if os.path.exists(file_start):
        try:
            with open(file_start, 'r') as f:
                waktu_mulai = float(f.read().strip())
            durasi_detik = int(time.time() - waktu_mulai)
            durasi_str = str(timedelta(seconds=durasi_detik))
            os.remove(file_start) # Auto-delete file sementara
        except: pass
    # ---------------------------------------------

    # 2. KIRIM TELEGRAM
    if len(daftar_error_api) > 0 or not git_sukses:
        pesan_ringkasan = "🚨 LAPORAN UPDATE SISTEM (PENGADAAN) 🚨\n\n"
        if len(daftar_error_api) > 0:
            teks_error = "⚠️ GAGAL DOWNLOAD API:\n" + "".join([f"{err.replace('_', ' ')}\n" for err in daftar_error_api])
            if len(teks_error) > 3500: teks_error = teks_error[:3500] + "\n... [DAFTAR ERROR DIPOTONG] ...\n"
            pesan_ringkasan += teks_error + "\n"
        pesan_ringkasan += f"📊 STATUS TEPRA: {status_tepra}\n\n🌐 STATUS GITHUB:\n{pesan_git}\n\n⏱ Durasi Total: {durasi_str}\n📅 Waktu: {get_waktu_indonesia()}"
        kirim_telegram_aman(pesan_ringkasan)
    else:
        pesan_sukses = f"✅ UPDATE PENGADAAN BERHASIL ✅\n\nSeluruh data berhasil diolah.\n📊 STATUS TEPRA: {status_tepra}\n\n🌐 STATUS GITHUB:\n{pesan_git}\n\n⏱ Durasi Total: {durasi_str}\n📅 Waktu: {get_waktu_indonesia()}"
        kirim_telegram_aman(pesan_sukses)

    log_print(f"\nPROSES SELESAI SELURUHNYA PADA {get_waktu_indonesia()}")