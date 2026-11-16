### ini file generate_rup.py ###

# ======================================================
# GENERATE RUP MULTI TAHUN (JSON + EXCEL + GIT + TELEGRAM)
# ======================================================
import pandas as pd
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
import warnings
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
import config_rahasia

warnings.filterwarnings("ignore")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

tahun_n  = datetime.now().year       
tahun_n1 = tahun_n - 1               
tahun_n2 = tahun_n - 2               
daftar_tahun = [tahun_n, tahun_n1, tahun_n2]

LOG_DIR = os.path.join(BASE_DIR, 'logs', 'rup')
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, f"log_rup_{datetime.now().strftime('%Y-%m-%d')}.txt")
# Simpan log 30 hari terakhir saja, yang lebih tua dihapus
import time as _t
for _f in os.listdir(LOG_DIR):
    if _f.startswith('log_rup_') and _f.endswith('.txt'):
        try:
            if os.path.getmtime(os.path.join(LOG_DIR, _f)) < _t.time() - 30 * 86400:
                os.remove(os.path.join(LOG_DIR, _f))
        except OSError: pass

_baris_tertunda = ""

def log_print(*args, **kwargs):
    global _baris_tertunda
    msg = " ".join(str(a) for a in args)
    print(msg, **kwargs)
    if kwargs.get('end') == " ":
        _baris_tertunda += msg + " "
        return
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(_baris_tertunda + msg + '\n')
    except OSError:
        pass
    _baris_tertunda = ""

def get_waktu_indonesia():
    tz_wib = timezone(timedelta(hours=7))
    sekarang = datetime.now(tz_wib)
    bulan_indo = {1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April', 5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'}
    return f"{sekarang.day} {bulan_indo[sekarang.month]} {sekarang.year} | {sekarang.strftime('%H.%M')} WIB"


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

def get_file_path(data_dir, base_name, tahun):
    v1_path = os.path.join(data_dir, f"v1_{base_name}_{tahun}.json")
    legacy_path = os.path.join(data_dir, f"Legacy_{base_name}_{tahun}.json")
    if os.path.exists(v1_path): return v1_path
    if os.path.exists(legacy_path): return legacy_path
    return legacy_path 

def process_tahun(tahun):
    log_print(f"\n--- MEMPROSES DATA TAHUN {tahun} ---")
    data_dir = os.path.join(BASE_DIR, 'data', str(tahun))
    
    s_master    = get_file_path(data_dir, "rup_master-satker", tahun)
    s_penyedia  = get_file_path(data_dir, "rup_paket-penyedia-terumumkan", tahun)
    s_swakelola = get_file_path(data_dir, "rup_paket-swakelola-terumumkan", tahun)
    s_program   = get_file_path(data_dir, "rup_program-master", tahun)
    s_struktur  = get_file_path(data_dir, "rup_struktur-anggaran-pd", tahun)

    df_master    = pd.DataFrame(load_json_local(s_master))
    df_penyedia  = pd.DataFrame(load_json_local(s_penyedia))
    df_swakelola = pd.DataFrame(load_json_local(s_swakelola))
    df_program   = pd.DataFrame(load_json_local(s_program))
    df_struktur  = pd.DataFrame(load_json_local(s_struktur))

    if df_master.empty: 
        log_print(f"Data Master kosong. Melewati tahun {tahun}.")
        return 0

    if 'kd_satker' in df_master.columns:
        df_master['kd_satker'] = pd.to_numeric(df_master['kd_satker'], errors='coerce')
    
    df_master = df_master[df_master['tahun_aktif'].astype(str).str.contains(str(tahun), na=False)]
    
    # Ambil kd_satker_str jika tersedia di data master
    cols_to_extract = ['kd_satker', 'nama_satker']
    if 'kd_satker_str' in df_master.columns: cols_to_extract.append('kd_satker_str')
        
    master_satker = df_master[cols_to_extract].drop_duplicates().dropna(subset=['kd_satker'])
    master_satker['kd_satker'] = master_satker['kd_satker'].astype(int)
    
    # Format Satuan Kerja menggunakan kd_satker_str
    def format_satker_master(row):
        nama = str(row['nama_satker']).strip()
        # Ambil kd_satker_str, jika kebetulan kosong, fallback ke kd_satker
        kd = str(row.get('kd_satker_str', row['kd_satker'])).strip()
        
        if pd.notna(row['nama_satker']) and nama.lower() not in ['nan', 'none', '']:
            return f"{nama} - {kd}"
        return kd

    master_satker['Satuan Kerja'] = master_satker.apply(format_satker_master, axis=1)
    
    # Hapus kolom teks agar tidak mengganggu proses merge selanjutnya
    cols_to_drop = ['nama_satker']
    if 'kd_satker_str' in master_satker.columns: cols_to_drop.append('kd_satker_str')
    master_satker.drop(columns=cols_to_drop, inplace=True)

    for d in [df_penyedia, df_swakelola, df_program, df_struktur]:
        if not d.empty and 'kd_satker' in d.columns:
            d['kd_satker'] = pd.to_numeric(d['kd_satker'], errors='coerce')

    rup_penyedia = df_penyedia.groupby('kd_satker', as_index=False).agg(**{'RUP Penyedia': ('pagu', 'sum'), 'Paket Penyedia': ('pagu', 'count')}) if not df_penyedia.empty else pd.DataFrame(columns=['kd_satker', 'RUP Penyedia', 'Paket Penyedia'])
    rup_swakelola = df_swakelola.groupby('kd_satker', as_index=False).agg(**{'RUP Swakelola': ('pagu', 'sum'), 'Paket Swakelola': ('pagu', 'count')}) if not df_swakelola.empty else pd.DataFrame(columns=['kd_satker', 'RUP Swakelola', 'Paket Swakelola'])

    if not df_program.empty:
        kolom_ada = [c for c in ['kd_satker', 'nama_program', 'kd_program'] if c in df_program.columns]
        df_program = df_program.drop_duplicates(subset=kolom_ada)
        df_program = df_program[~df_program['nama_program'].astype(str).str.contains(r'( M$|\(M\)$)', regex=True)]
        pagu_program = df_program.groupby('kd_satker', as_index=False)['pagu_program'].sum().rename(columns={'pagu_program': 'Pagu Program'})
    else:
        pagu_program = pd.DataFrame(columns=['kd_satker', 'Pagu Program'])

    struktur = df_struktur.groupby('kd_satker', as_index=False)['belanja_pengadaan'].sum().rename(columns={'belanja_pengadaan': 'Pagu Pengadaan'}) if not df_struktur.empty else pd.DataFrame(columns=['kd_satker', 'Pagu Pengadaan'])

    df = master_satker.merge(pagu_program, on='kd_satker', how='left')
    df = df.merge(rup_penyedia, on='kd_satker', how='left')
    df = df.merge(rup_swakelola, on='kd_satker', how='left')
    df = df.merge(struktur, on='kd_satker', how='left')
    df.fillna(0, inplace=True)

    for col in ['RUP Penyedia', 'Paket Penyedia', 'RUP Swakelola', 'Paket Swakelola', 'Pagu Pengadaan']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df['Total RUP Terumumkan'] = df['RUP Penyedia'] + df['RUP Swakelola']
    df['Total Paket Terumumkan'] = df['Paket Penyedia'] + df['Paket Swakelola']
    df['Selisih RUP Terumumkan'] = df['Total RUP Terumumkan'] - df['Pagu Pengadaan']
    
    df['Persentase'] = (
        df['Total RUP Terumumkan'].astype(float) / 
        df['Pagu Pengadaan'].astype(float).replace(0, float('nan'))
    ).fillna(0) * 100

    # Menyisipkan kolom paket di sebelah kolom pagu
    df_final = df[['Satuan Kerja', 'Pagu Program', 'Pagu Pengadaan', 'Paket Penyedia', 'RUP Penyedia', 
                    'Paket Swakelola', 'RUP Swakelola', 'Total Paket Terumumkan', 'Total RUP Terumumkan', 
                    'Selisih RUP Terumumkan', 'Persentase']]
    df_final = df_final.sort_values('Satuan Kerja').reset_index(drop=True)

    output_json = os.path.join(data_dir, f"rekap_rup_{tahun}.json")
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(df_final.to_dict(orient='records'), f, ensure_ascii=False, indent=2)

    tgl = datetime.now().strftime('%Y-%m-%d')
    nama_file_history = f"Rekap RUP Tahun {tahun} ({tgl}).xlsx"
    output_history_dir = os.path.join(BASE_DIR, "arsip_lokal", "rup", str(tahun))
    os.makedirs(output_history_dir, exist_ok=True)
    path_history = os.path.join(output_history_dir, nama_file_history)

    df_final.to_excel(path_history, index=False, sheet_name='Rekap RUP')

    wb = load_workbook(path_history)
    ws = wb['Rekap RUP']
    h_fill = PatternFill('solid', start_color='1F4E79')
    h_font = Font(name='Arial', bold=True, color='FFFFFF')
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment = h_fill, h_font, Alignment(horizontal='center')
    
    ws.column_dimensions['A'].width = 45
    for col in ['B','C','E','G','I','J']: ws.column_dimensions[col].width = 22 # Kolom Nilai Uang
    for col in ['D','F','H']: ws.column_dimensions[col].width = 18             # Kolom Jumlah Paket
    ws.column_dimensions['K'].width = 12                                       # Kolom Persentase

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            if cell.column in range(2, 11): # Kolom B sampai J (Uang & Paket)
                cell.number_format = '#,##0'
            if cell.column == 11: # Kolom K (Persentase)
                cell.number_format = '0.00"%"'

    wb.save(path_history)

    log_print(f"DONE -> JSON: rekap_rup_{tahun}.json")
    log_print(f"DONE -> EXCEL: {path_history}")
    return len(df_final)

if __name__ == "__main__":
    waktu_mulai = float(os.environ.get('RUP_START', time.time()))

    log_print("\n==================================================")
    log_print(f"START GENERATE RUP {get_waktu_indonesia()}")
    log_print("==================================================")

    # Proses tiap tahun (tahun final n-1 & n-2 dilewati bila rekap sudah ada)
    peringatan = []
    for t in daftar_tahun:
        output_json_cek = os.path.join(BASE_DIR, 'data', str(t), f'rekap_rup_{t}.json')
        if t != tahun_n and os.path.exists(output_json_cek):
            log_print(f"\n[SKIP] Tahun {t} sudah ada (final).")
            continue
        hasil = process_tahun(t)
        if hasil == 0:
            peringatan.append(f"• Tahun {t}: data master kosong, tahun dilewati (hasil lama dipertahankan)")

    with open(os.path.join(BASE_DIR, "data", "last-update-rup.txt"), "w", encoding='utf-8') as f:
        f.write(get_waktu_indonesia())

    durasi_str = str(timedelta(seconds=int(time.time() - waktu_mulai)))

    log_print("\n" + "="*50)
    log_print(f"PROSES SELESAI SELURUHNYA PADA {get_waktu_indonesia()}")
    log_print("==================================================")

    # Catatan dari skrip sebelumnya (detail_rup.py & konsolidasi_rup.py)
    temp_tg = os.path.join(BASE_DIR, 'logs', 'rup', 'temp_tg_rup.txt')
    info_tambahan = ""
    if os.path.exists(temp_tg):
        try:
            with open(temp_tg, 'r', encoding='utf-8') as f:
                info_tambahan = f.read()
        except: pass

    blok_peringatan = ("⚠️ PERINGATAN:\n" + "\n".join(peringatan) + "\n\n") if peringatan else ""

    # Laporan error API ada di laporan download_json.py, bukan di sini
    if peringatan:
        pesan = "🚨 LAPORAN UPDATE SISTEM (RUP) 🚨\n\n" + blok_peringatan
        pesan += f"📊 RINGKASAN EKSEKUSI:\n{info_tambahan}🔹 Generate RUP: Selesai\n\n"
    else:
        pesan = f"✅ UPDATE RUP HARIAN SELESAI ✅\n\nSeluruh proses ekstraksi dan generate data berhasil.\n\n📊 RINGKASAN EKSEKUSI:\n{info_tambahan}🔹 Generate RUP: Selesai\n\n"
    pesan += f"⏱ Durasi Total: {durasi_str}\n📅 Waktu: {get_waktu_indonesia()}"
    kirim_telegram_aman(pesan)

    # Bersihkan file sementara setelah laporan terkirim
    if os.path.exists(temp_tg):
        try: os.remove(temp_tg)
        except: pass