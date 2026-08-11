# ======================================================
# FASE 1: EXTRACT (DOWNLOAD DATA PENGADAAN)
# ======================================================

import os
import json
import time
import requests
import sys
import re
from datetime import datetime, timedelta, timezone
import warnings

warnings.filterwarnings('ignore')

# ======================================================
# KONFIGURASI UTAMA
# ======================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    with open(os.path.join(BASE_DIR, 'token.txt'), 'r', encoding='utf-8') as f:
        TOKEN = f.read().strip()
        if not TOKEN:
            print("ERROR: File token.txt kosong! Silakan isi dengan token terbaru.")
            sys.exit(1)
except FileNotFoundError:
    print("ERROR: File token.txt tidak ditemukan di folder project!")
    sys.exit(1)
        
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept": "application/json, text/plain, */*"
}

tahun_n      = datetime.now().year       
tahun_n1     = tahun_n - 1               
tahun_n2     = tahun_n - 2               
daftar_tahun = [tahun_n, tahun_n1, tahun_n2] 

daftar_error_api = [] # Variabel penampung daftar gagal

# MENGGUNAKAN LOG TUNGGAL
LOG_FILE = os.path.join(BASE_DIR, 'tools', 'log_pengadaan.txt')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

def log_print(*args, **kwargs):
    msg = " ".join(str(a) for a in args)
    print(msg, **kwargs)
    if 'end' in kwargs and kwargs['end'] == " ": return
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(msg + '\n')

def get_waktu_indonesia():
    tz_wib = timezone(timedelta(hours=7))
    sekarang = datetime.now(tz_wib)
    bulan_indo = {1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April', 5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'}
    return f"{sekarang.day} {bulan_indo[sekarang.month]} {sekarang.year} | {sekarang.strftime('%H.%M')} WIB"

def download_data_pengadaan(tahun, is_n2, data_dir):
    txt_path = os.path.join(BASE_DIR, 'scripts', 'pengadaan', 'url_pengadaan.txt')
    if not os.path.exists(txt_path): return
    with open(txt_path, 'r', encoding='utf-8') as f:
        urls = [line.strip() for line in f if line.strip() and not line.strip().startswith(('#', '='))]

    for raw_url in urls:
        target_url = raw_url.replace('{tahun}', str(tahun))
        is_v1 = '/v1/' in target_url
        tipe = "v1" if is_v1 else "Legacy"

        match = re.search(r'api/(?:v1|legacy)/(.*?)\?', target_url)
        if not match: continue
        base_name = match.group(1).replace('/', '_')
        
        filename = f"{tipe}_{base_name}_{tahun}.json"
        output_path = os.path.join(data_dir, filename)
        log_print("") # Tambahan enter otomatis

        if is_n2 and os.path.exists(output_path): continue

        # -- TAMBAHAN LOGIKA BACKUP HARIAN (COPY) --
        if os.path.exists(output_path):
            tgl_file = datetime.fromtimestamp(os.path.getmtime(output_path)).date()
            if tgl_file < datetime.now().date():
                import shutil
                
                # Deteksi 4 digit tahun (contoh: 2024, 2025, 2026) dari nama file
                match_tahun = re.search(r'(20\d{2})', filename)
                
                if match_tahun:
                    tahun_file = match_tahun.group(1)
                    f_arsip = os.path.join(BASE_DIR, 'arsip_json', str(tgl_file.year), f"{tgl_file.month:02d}", f"{tgl_file.day:02d}", tahun_file)
                    folder_log = f"{tgl_file}/{tahun_file}"
                else:
                    f_arsip = os.path.join(BASE_DIR, 'arsip_json', str(tgl_file.year), f"{tgl_file.month:02d}", f"{tgl_file.day:02d}")
                    folder_log = f"{tgl_file}"
                
                os.makedirs(f_arsip, exist_ok=True)
                shutil.copy2(output_path, os.path.join(f_arsip, filename))
                log_print(f"[BACKUP] {filename} di-copy ke arsip {folder_log}")
        # -----------------------------------

        log_print(f"DOWNLOAD [{tipe.upper()}]: {target_url}")

        if is_v1:
            all_data = []
            cursor = None
            req_count = 1
            first_response = None
            last_error = "Unknown Error" # Penampung error v1
            
            while True:
                url_cursor = target_url
                if cursor:
                    sep = "&" if "?" in target_url else "?"
                    url_cursor = f"{target_url}{sep}cursor={cursor}"

                success = False
                resp_data = None
                
                for i in range(1, 6):
                    try:
                        log_print(f"  Request ke-{req_count} (Coba {i})...", end=" ")
                        resp = requests.get(url_cursor, headers=HEADERS, timeout=150)
                        if resp.status_code == 200:
                            resp_data = resp.json()
                            log_print("SUKSES")
                            success = True
                            break
                        else:
                            last_error = f"Status {resp.status_code}"
                            log_print(f"GAGAL ({last_error})")
                    except Exception as e:
                        last_error = f"Koneksi Terputus / {str(e)}"
                        log_print(f"ERROR: {last_error}")
                    
                    if i < 5 and not success:
                        waktu_tunggu = i * 5
                        log_print(f"    -> [!] Menunggu {waktu_tunggu} detik agar server Inaproc stabil...")
                        time.sleep(waktu_tunggu)
                
                if not success: break
                if req_count == 1: first_response = resp_data
                
                # JARING PENGAMAN: Mencegah error jika server membalas {"data": null}
                if resp_data and 'data' in resp_data:
                    isi_data = resp_data['data']
                    if isi_data is not None:
                        all_data.extend(isi_data)
                elif resp_data and isinstance(resp_data, list):
                    all_data.extend(resp_data)
                
                if resp_data and 'meta' in resp_data and resp_data['meta'].get('has_more'):
                    cursor = resp_data['meta'].get('cursor')
                    req_count += 1
                    time.sleep(1)
                else: break
            
            if not success:
                if not os.path.exists(output_path):
                    with open(output_path, 'w', encoding='utf-8') as f: f.write("[]")
                else:
                    log_print(f"  -> [Aman] Proses terputus, mempertahankan file {filename} lama yang utuh.")
                daftar_error_api.append(f"❌ PENGADAAN V1 ({tahun}) - {base_name} ({last_error})")
                
            else:
                # LOGIKA PENGAMAN: Mencegah file utuh tertimpa balasan API kosong (null)
                if len(all_data) == 0:
                    if os.path.exists(output_path):
                        log_print(f"  -> [Aman] Data API kosong (null), mempertahankan file {filename} lama yang utuh.")
                    else:
                        with open(output_path, 'w', encoding='utf-8') as f: 
                            f.write("[]")
                        log_print(f"  -> File baru dibuat dengan daftar kosong [] (Total: 0 baris)")
                else:
                    with open(output_path, 'w', encoding='utf-8') as f:
                        json.dump(all_data, f, ensure_ascii=False, indent=2)
                    log_print(f"  -> Disimpan ke {filename} (Total: {len(all_data)} baris)")

        else:
            success = False
            last_error = "Unknown Error" # Penampung error legacy
            
            for i in range(1, 6):
                try:
                    log_print(f"  Percobaan ke-{i}...", end=" ")
                    resp = requests.get(target_url, headers=HEADERS, timeout=150)
                    if resp.status_code == 200:
                        with open(output_path, 'w', encoding='utf-8') as f:
                            json.dump(resp.json(), f, ensure_ascii=False, indent=2)
                        log_print("SUKSES")
                        success = True
                        break
                    else:
                        last_error = f"Status {resp.status_code}"
                        log_print(f"GAGAL ({last_error})")
                except Exception as e:
                    last_error = f"Koneksi Terputus / {str(e)}"
                    log_print(f"ERROR: {last_error}")
                
                if i < 5 and not success:
                    waktu_tunggu = i * 5
                    log_print(f"    -> [!] Menunggu {waktu_tunggu} detik agar server Inaproc stabil...")
                    time.sleep(waktu_tunggu)
            
            if not success:
                if not os.path.exists(output_path):
                    with open(output_path, 'w', encoding='utf-8') as f: f.write("[]")
                else:
                    log_print(f"  -> [Aman] Server gagal diakses, mempertahankan file {filename}.")
                daftar_error_api.append(f"❌ PENGADAAN Legacy ({tahun}) - {base_name} ({last_error})")

if __name__ == '__main__':
    # Catat waktu mulai untuk menghitung durasi total
    waktu_mulai = time.time()
    with open(os.path.join(BASE_DIR, 'tools', 'start_time_pengadaan.txt'), 'w') as f:
        f.write(str(waktu_mulai))

    log_print("\n" + "="*55)
    log_print(f"START DOWNLOAD PENGADAAN {get_waktu_indonesia()}")
    log_print("="*55)

    for t in daftar_tahun:
        data_dir = os.path.join(BASE_DIR, 'data', str(t))
        os.makedirs(data_dir, exist_ok=True)
        log_print(f'\n{"="*55}\n   DOWNLOAD DATA TAHUN {t}\n{"="*55}')
        
        if t == tahun_n2 and os.path.exists(os.path.join(data_dir, f'rekap_pengadaan_{t}.json')):
            log_print(f"\n[SKIP] Tahun {t} sudah final -> Lewati download")
            continue
            
        download_data_pengadaan(t, (t == tahun_n2), data_dir)

    # SIMPAN DAFTAR ERROR KE FOLDER SCRIPTS/PENGADAAN
    path_error = os.path.join(BASE_DIR, 'scripts', 'pengadaan', 'error_api_pengadaan.json')
    with open(path_error, 'w', encoding='utf-8') as f:
        json.dump(daftar_error_api, f)

    log_print("\nPROSES DOWNLOAD SELESAI!")