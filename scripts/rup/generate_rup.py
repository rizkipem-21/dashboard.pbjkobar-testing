# ======================================================
# GENERATE RUP MULTI TAHUN (DOWNLOAD RETRY + JSON + EXCEL)
# ======================================================

import pandas as pd
import json
import os
import re
import requests
import time
from datetime import datetime, timedelta, timezone
import warnings
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

warnings.filterwarnings("ignore")

# ======================================================
# KONFIGURASI TAHUN DINAMIS (AUTOPILOT)
# ======================================================

# Otomatis mengambil jalur folder utama (naik 1 tingkat dari folder 'scripts')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

tahun_n  = datetime.now().year       # Tahun berjalan
tahun_n1 = tahun_n - 1               # Tahun lalu
tahun_n2 = tahun_n - 2               # Dua tahun lalu
daftar_tahun = [tahun_n, tahun_n1, tahun_n2]

# Konfigurasi API
# Membaca token murni HANYA dari file token.txt
try:
    with open(os.path.join(BASE_DIR, 'token.txt'), 'r', encoding='utf-8') as f:
        TOKEN = f.read().strip()
        if not TOKEN:
            print("ERROR: File token.txt kosong! Silakan isi dengan token terbaru.")
            sys.exit(1) # Menghentikan skrip
except FileNotFoundError:
    print("ERROR: File token.txt tidak ditemukan di folder project! Buat filenya dan isi dengan token terbaru.")
    sys.exit(1) # Menghentikan skrip
    
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}
ENDPOINTS = [
    "rup/paket-penyedia-terumumkan",
    "rup/paket-swakelola-terumumkan",
    "rup/master-satker",
    "rup/program-master",
    "rup/struktur-anggaran-pd"
]

# ======================================================
# FUNGSI LOG & LAST-UPDATE (PENGGANTI .BAT)
# ======================================================
LOG_FILE = os.path.join(BASE_DIR, 'tools', 'log_rup.txt')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

def log_print(*args, **kwargs):
    """Mencetak ke layar terminal sekaligus ke file log txt"""
    msg = " ".join(str(a) for a in args)
    print(msg, **kwargs) # Cetak ke layar
    
    # Jangan tulis enter bawaan print yang ada di 'end' ke file log untuk merapikan
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
# FUNGSI 1: DOWNLOAD DATA API DENGAN RETRY
# ======================================================
def download_data_api_with_retry(tahun):
    log_print(f"\n--- MENGUNDUH DATA TAHUN {tahun} ---")
    data_dir = os.path.join(BASE_DIR, 'data', str(tahun))
    os.makedirs(data_dir, exist_ok=True)

    for ep in ENDPOINTS:
        url = f"https://data.inaproc.id/api/legacy/{ep}?kode_klpd=D228&tahun={tahun}"
        base_name = ep.replace('/', '_')
        filename = f"Legacy_{base_name}_{tahun}.json"
        output_path = os.path.join(data_dir, filename)

        max_retry = 5
        success = False

        log_print(f"\nDOWNLOAD: {url}")
        
        for i in range(1, max_retry + 1):
            try:
                log_print(f"  Percobaan ke-{i}...", end=" ")
                response = requests.get(url, headers=HEADERS, timeout=60)
                
                if response.status_code == 200:
                    data = response.json()
                    with open(output_path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                    log_print("SUKSES")
                    success = True
                    break
                else:
                    log_print(f"GAGAL (Status {response.status_code})")
            except Exception as e:
                log_print(f"ERROR: {e}")
            
            if i < max_retry:
                time.sleep(2) # Jeda 2 detik sebelum coba lagi

        if not success:
            log_print(f"  GAGAL TOTAL -> buat file kosong")
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write("[]")

# ======================================================
# FUNGSI BANTUAN LOAD DATA & PEMBERSIHAN
# ======================================================
def load_json_local(path):
    try:
        if not os.path.exists(path): return []
        with open(path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
            if isinstance(data, list): return data
            if isinstance(data, dict): return data.get('data', [])
            return []
    except:
        return []

def bersihkan_arsip_bulanan(folder_path):
    """Fitur untuk menghapus riwayat Excel bulan lalu, sisakan hanya 1 file akhir bulan"""
    if not os.path.exists(folder_path): return
    
    bulan_ini = datetime.now().strftime("%Y-%m")
    file_excel = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    arsip_bulanan = {}
    for f in file_excel:
        match = re.search(r"\((\d{4}-\d{2}-\d{2})\)", f)
        if match:
            tanggal = match.group(1)
            bulan = tanggal[:7]
            if bulan not in arsip_bulanan:
                arsip_bulanan[bulan] = []
            arsip_bulanan[bulan].append((tanggal, f))
            
    for bulan, list_file in arsip_bulanan.items():
        if bulan != bulan_ini:
            # Urutkan dari tanggal awal ke akhir, hapus semua kecuali hari terakhir
            list_file.sort(key=lambda x: x[0])
            file_yang_dihapus = list_file[:-1]
            for tgl, nama_file in file_yang_dihapus:
                try: os.remove(os.path.join(folder_path, nama_file))
                except Exception: pass

# ======================================================
# FUNGSI 2: PROSES DATA & GENERATE EXCEL
# ======================================================
def process_tahun(tahun):
    log_print(f"\n--- MEMPROSES DATA TAHUN {tahun} ---")
    data_dir = os.path.join(BASE_DIR, 'data', str(tahun))
    
    # Path Sumber
    s_master    = os.path.join(data_dir, f"Legacy_rup_master-satker_{tahun}.json")
    s_penyedia  = os.path.join(data_dir, f"Legacy_rup_paket-penyedia-terumumkan_{tahun}.json")
    s_swakelola = os.path.join(data_dir, f"Legacy_rup_paket-swakelola-terumumkan_{tahun}.json")
    s_program   = os.path.join(data_dir, f"Legacy_rup_program-master_{tahun}.json")
    s_struktur  = os.path.join(data_dir, f"Legacy_rup_struktur-anggaran-pd_{tahun}.json")

    # Load ke DataFrame
    df_master    = pd.DataFrame(load_json_local(s_master))
    df_penyedia  = pd.DataFrame(load_json_local(s_penyedia))
    df_swakelola = pd.DataFrame(load_json_local(s_swakelola))
    df_program   = pd.DataFrame(load_json_local(s_program))
    df_struktur  = pd.DataFrame(load_json_local(s_struktur))

    if df_master.empty: 
        log_print(f"Data Master kosong. Melewati tahun {tahun}.")
        return 0

    # 1. Master Satker
    if 'kd_satker' in df_master.columns:
        df_master['kd_satker'] = pd.to_numeric(df_master['kd_satker'], errors='coerce')
    
    df_master = df_master[df_master['tahun_aktif'].astype(str).str.contains(str(tahun), na=False)]
    master_satker = df_master[['kd_satker', 'nama_satker']].drop_duplicates().dropna(subset=['kd_satker'])
    master_satker['kd_satker'] = master_satker['kd_satker'].astype(int)
    master_satker.rename(columns={'nama_satker': 'Satuan Kerja'}, inplace=True)

    # 2. Aggregasi Pagu
    for d in [df_penyedia, df_swakelola, df_program, df_struktur]:
        if not d.empty and 'kd_satker' in d.columns:
            d['kd_satker'] = pd.to_numeric(d['kd_satker'], errors='coerce')

    rup_penyedia = df_penyedia.groupby('kd_satker', as_index=False)['pagu'].sum().rename(columns={'pagu': 'RUP Penyedia'}) if not df_penyedia.empty else pd.DataFrame(columns=['kd_satker', 'RUP Penyedia'])
    rup_swakelola = df_swakelola.groupby('kd_satker', as_index=False)['pagu'].sum().rename(columns={'pagu': 'RUP Swakelola'}) if not df_swakelola.empty else pd.DataFrame(columns=['kd_satker', 'RUP Swakelola'])
    
    # 3. Pagu Program
    if not df_program.empty:
        kolom_ada = [c for c in ['kd_satker', 'nama_program', 'kd_program'] if c in df_program.columns]
        df_program = df_program.drop_duplicates(subset=kolom_ada)
        df_program = df_program[~df_program['nama_program'].astype(str).str.contains(r'( M$|\(M\)$)', regex=True)]
        pagu_program = df_program.groupby('kd_satker', as_index=False)['pagu_program'].sum().rename(columns={'pagu_program': 'Pagu Program'})
    else:
        pagu_program = pd.DataFrame(columns=['kd_satker', 'Pagu Program'])

    # 4. Struktur Anggaran
    struktur = df_struktur.groupby('kd_satker', as_index=False)['belanja_pengadaan'].sum().rename(columns={'belanja_pengadaan': 'Pagu Pengadaan'}) if not df_struktur.empty else pd.DataFrame(columns=['kd_satker', 'Pagu Pengadaan'])

    # 5. Merge Semua
    df = master_satker.merge(pagu_program, on='kd_satker', how='left')
    df = df.merge(rup_penyedia, on='kd_satker', how='left')
    df = df.merge(rup_swakelola, on='kd_satker', how='left')
    df = df.merge(struktur, on='kd_satker', how='left')
    df.fillna(0, inplace=True)

    # 6. Kalkulasi (Aman dari ZeroDivisionError)
    df['RUP Penyedia'] = pd.to_numeric(df['RUP Penyedia'], errors='coerce').fillna(0)
    df['RUP Swakelola'] = pd.to_numeric(df['RUP Swakelola'], errors='coerce').fillna(0)
    df['Pagu Pengadaan'] = pd.to_numeric(df['Pagu Pengadaan'], errors='coerce').fillna(0)

    df['Total RUP Terumumkan'] = df['RUP Penyedia'] + df['RUP Swakelola']
    df['Selisih RUP Terumumkan'] = df['Total RUP Terumumkan'] - df['Pagu Pengadaan']
    
    df['Persentase'] = (
        df['Total RUP Terumumkan'].astype(float) / 
        df['Pagu Pengadaan'].astype(float).replace(0, float('nan'))
    ).fillna(0) * 100

    # 7. Finalisasi Data
    df_final = df[['Satuan Kerja', 'Pagu Program', 'Pagu Pengadaan', 'RUP Penyedia', 
                  'RUP Swakelola', 'Total RUP Terumumkan', 'Selisih RUP Terumumkan', 'Persentase']]
    df_final = df_final.sort_values('Satuan Kerja').reset_index(drop=True)

    # ======================================================
    # SIMPAN JSON
    # ======================================================
    output_json = os.path.join(data_dir, f"rekap_rup_{tahun}.json")
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(df_final.to_dict(orient='records'), f, ensure_ascii=False, indent=2)

    # ======================================================
    # SIMPAN EXCEL (HISTORY SAJA)
    # ======================================================
    tgl = datetime.now().strftime('%Y-%m-%d')
    nama_file_history = f"Rekap RUP Tahun {tahun} ({tgl}) Legacy.xlsx"
    
    output_history_dir = os.path.join(BASE_DIR, "output", "rup", str(tahun))
    os.makedirs(output_history_dir, exist_ok=True)
    
    path_history = os.path.join(output_history_dir, nama_file_history)

    df_final.to_excel(path_history, index=False, sheet_name='Rekap RUP')

    # Formatting
    wb = load_workbook(path_history)
    ws = wb['Rekap RUP']
    h_fill = PatternFill('solid', start_color='1F4E79')
    h_font = Font(name='Arial', bold=True, color='FFFFFF')
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment = h_fill, h_font, Alignment(horizontal='center')
    
    ws.column_dimensions['A'].width = 45
    for col in ['B','C','D','E','F','G']: ws.column_dimensions[col].width = 22
    ws.column_dimensions['H'].width = 12

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            if cell.column in [2,3,4,5,6,7]: cell.number_format = '#,##0'
            if cell.column == 8: cell.number_format = '0.00"%"'

    wb.save(path_history)

    # ----------------------------------------------------
    # FITUR PEMBERSIHAN ARSIP BULANAN & BUKU DAFTAR (RUP)
    # ----------------------------------------------------
    bersihkan_arsip_bulanan(output_history_dir)
    if os.path.exists(output_history_dir):
        file_tersisa = [f for f in os.listdir(output_history_dir) if f.endswith('.xlsx')]
        with open(os.path.join(output_history_dir, 'daftar_arsip.json'), 'w') as f:
            json.dump(file_tersisa, f)

    log_print(f"DONE -> JSON: rekap_rup_{tahun}.json")
    log_print(f"DONE -> EXCEL: {path_history}")
    return len(df_final)

# ======================================================
# MAIN EXECUTION
# ======================================================
if __name__ == "__main__":
    log_print("\n==================================================")
    log_print(f"START {get_waktu_indonesia()}")
    log_print("AUTO GENERATE RUP (DOWNLOAD RETRY & MULTI TAHUN)")
    log_print("==================================================")

    total_all = 0
    for t in daftar_tahun:
        output_json_cek = os.path.join(BASE_DIR, 'data', str(t), f'rekap_rup_{t}.json')
        
        # LOGIKA SKIP UNTUK TAHUN LALU (n-1 dan n-2)
        if t != tahun_n and os.path.exists(output_json_cek):
            log_print(f"\n[SKIP] Tahun {t} sudah ada (final).")
            continue
        
        download_data_api_with_retry(t)
        total_all += process_tahun(t)
    
    # ----------------------------------------------------
    # UPDATE FILE LAST-UPDATE UNTUK WEBSITE
    # ----------------------------------------------------
    with open(os.path.join(BASE_DIR, "data", "last-update-rup.txt"), "w", encoding='utf-8') as f:
        f.write(get_waktu_indonesia())
    
    log_print("\n" + "="*50)
    log_print(f"PROSES SELESAI SELURUHNYA PADA {get_waktu_indonesia()}")
    log_print("==================================================")