# ======================================================
# 1. DOWNLOAD RUP MULTI TAHUN
# ======================================================
import json
import os
import re
import requests
import time
import sys
from datetime import datetime, timedelta, timezone

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

tahun_n  = datetime.now().year       
tahun_n1 = tahun_n - 1               
tahun_n2 = tahun_n - 2               
daftar_tahun = [tahun_n, tahun_n1, tahun_n2]

try:
    with open(os.path.join(BASE_DIR, 'token.txt'), 'r', encoding='utf-8') as f:
        TOKEN = f.read().strip()
        if not TOKEN:
            print("ERROR: File token.txt kosong!")
            sys.exit(1)
except FileNotFoundError:
    print("ERROR: File token.txt tidak ditemukan di folder project!")
    sys.exit(1)
    
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}

LOG_FILE = os.path.join(BASE_DIR, 'tools', 'log_rup.txt')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
daftar_error_api = []

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

def download_data_api_with_retry(tahun):
    global daftar_error_api
    log_print(f"\n--- MENGUNDUH DATA TAHUN {tahun} ---")
    data_dir = os.path.join(BASE_DIR, 'data', str(tahun))
    os.makedirs(data_dir, exist_ok=True)

    txt_path = os.path.join(BASE_DIR, 'scripts', 'rup', 'url_rup.txt')
    if not os.path.exists(txt_path):
        log_print(f"ERROR: File sumber URL tidak ditemukan di {txt_path}")
        return

    with open(txt_path, 'r', encoding='utf-8') as f:
        urls = [line.strip() for line in f if line.strip()]

    for raw_url in urls:
        target_url = raw_url.replace('{tahun}', str(tahun))
        is_v1 = '/v1/' in target_url
        tipe = "v1" if is_v1 else "Legacy"

        match = re.search(r'api/(?:v1|legacy)/(.*?)\?', target_url)
        base_name = match.group(1).replace('/', '_') if match else "unknown"
        filename = f"{tipe}_{base_name}_{tahun}.json"
        output_path = os.path.join(data_dir, filename)

        log_print(f"\nDOWNLOAD [{tipe.upper()}]: {target_url}")

        if is_v1:
            all_data = []
            cursor = None
            req_count = 1
            first_response = None
            success = False 
            last_error = "Unknown Error" # Penampung error v1

            while True:
                url_cursor = target_url
                if cursor:
                    sep = "&" if "?" in target_url else "?"
                    url_cursor = f"{target_url}{sep}cursor={cursor}"

                max_retry = 5
                page_success = False
                resp_data = None

                for i in range(1, max_retry + 1):
                    try:
                        log_print(f"  Request ke-{req_count} (Percobaan {i})...", end=" ")
                        response = requests.get(url_cursor, headers=HEADERS, timeout=150)
                        if response.status_code == 200:
                            resp_data = response.json()
                            log_print("SUKSES")
                            page_success = True
                            success = True
                            break
                        else:
                            last_error = f"Status {response.status_code}"
                            log_print(f"GAGAL ({last_error})")
                    except Exception as e:
                        last_error = f"Koneksi Terputus / {str(e)}"
                        log_print(f"ERROR: {last_error}")
                    
                    if i < max_retry and not page_success:
                        waktu_tunggu = i * 5
                        log_print(f"    -> [!] Menunggu {waktu_tunggu} detik agar server Inaproc stabil...")
                        time.sleep(waktu_tunggu)

                if not page_success:
                    log_print("  GAGAL TOTAL pada request ini. Berhenti ditarik.")
                    success = False 
                    break

                if req_count == 1: first_response = resp_data

                if resp_data and 'data' in resp_data:
                    all_data.extend(resp_data['data'])
                elif resp_data and isinstance(resp_data, list):
                    all_data.extend(resp_data)

                if resp_data and 'meta' in resp_data and resp_data['meta'].get('has_more'):
                    cursor = resp_data['meta'].get('cursor')
                    req_count += 1
                    time.sleep(1) 
                else:
                    break

            if not success:
                if not os.path.exists(output_path):
                    with open(output_path, 'w', encoding='utf-8') as f: f.write("[]")
                else:
                    log_print(f"  -> [Aman] Proses terputus, mempertahankan file {filename} lama yang utuh.")
                daftar_error_api.append(f"❌ RUP V1 ({tahun}) - {base_name} ({last_error})")    
            else:
                with open(output_path, 'w', encoding='utf-8') as f:
                    if len(all_data) == 0 and first_response:
                        json.dump(first_response, f, ensure_ascii=False, indent=2)
                    else:
                        json.dump(all_data, f, ensure_ascii=False, indent=2)
                log_print(f"  -> Disimpan ke {filename} (Total: {len(all_data)} baris)")

        else:
            max_retry = 5
            success = False
            last_error = "Unknown Error" # Penampung error legacy

            for i in range(1, max_retry + 1):
                try:
                    log_print(f"  Percobaan ke-{i}...", end=" ")
                    response = requests.get(target_url, headers=HEADERS, timeout=150)
                    if response.status_code == 200:
                        data = response.json()
                        with open(output_path, 'w', encoding='utf-8') as f:
                            json.dump(data, f, ensure_ascii=False, indent=2)
                        log_print("SUKSES")
                        success = True
                        break
                    else:
                        last_error = f"Status {response.status_code}"
                        log_print(f"GAGAL ({last_error})")
                except Exception as e:
                    last_error = f"Koneksi Terputus / {str(e)}"
                    log_print(f"ERROR: {last_error}")

                if i < max_retry and not success:
                    waktu_tunggu = i * 5
                    log_print(f"    -> [!] Menunggu {waktu_tunggu} detik agar server Inaproc stabil...")
                    time.sleep(waktu_tunggu)

            if not success:
                if not os.path.exists(output_path):
                    log_print(f"  GAGAL TOTAL -> buat file kosong")
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write("[]")
                else:
                    log_print(f"  -> [Aman] Server gagal diakses, mempertahankan file {filename} yang sudah ada.")
                daftar_error_api.append(f"❌ RUP Legacy ({tahun}) - {base_name} ({last_error})") 

if __name__ == "__main__":
    log_print("\n==================================================")
    log_print(f"START DOWNLOAD RUP {get_waktu_indonesia()}")
    log_print("==================================================")

    for t in daftar_tahun:
        output_json_cek = os.path.join(BASE_DIR, 'data', str(t), f'rekap_rup_{t}.json')
        if t != tahun_n and os.path.exists(output_json_cek):
            log_print(f"\n[SKIP] Tahun {t} sudah ada (final).")
            continue
        download_data_api_with_retry(t)

    path_error = os.path.join(BASE_DIR, 'scripts', 'rup', 'error_api_rup.json')
    os.makedirs(os.path.dirname(path_error), exist_ok=True)
    with open(path_error, 'w', encoding='utf-8') as f:
        json.dump(daftar_error_api, f)
        
    log_print(f"\nPROSES DOWNLOAD SELESAI PADA {get_waktu_indonesia()}")