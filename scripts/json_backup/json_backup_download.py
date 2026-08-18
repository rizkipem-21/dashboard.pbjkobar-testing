# ======================================================
# DOWNLOAD JSON UNTUK BACKUP (HANYA TAHUN BERJALAN)
# ======================================================
import json
import os
import re
import requests
import time
import sys
import shutil
from datetime import datetime, timedelta, timezone
import warnings

warnings.filterwarnings("ignore")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Memastikan bisa import config_rahasia dari root folder
sys.path.append(os.path.abspath(os.path.join(BASE_DIR)))
try:
    import config_rahasia
except ImportError:
    print("WARNING: config_rahasia.py tidak ditemukan. Notifikasi Telegram mungkin gagal.")

# Konfigurasi Waktu
tz_wib = timezone(timedelta(hours=7))
waktu_mulai = datetime.now(tz_wib)
tahun_n = waktu_mulai.year
bulan_str = f"{waktu_mulai.month:02d}"
hari_str = f"{waktu_mulai.day:02d}"

ARSIP_DIR = os.path.join(BASE_DIR, 'arsip_json', str(tahun_n), bulan_str, hari_str)
os.makedirs(ARSIP_DIR, exist_ok=True)

try:
    with open(os.path.join(BASE_DIR, 'token.txt'), 'r', encoding='utf-8') as f:
        TOKEN = f.read().strip()
        if not TOKEN: sys.exit(1)
except FileNotFoundError:
    sys.exit(1)
    
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}

LOG_FILE = os.path.join(BASE_DIR, 'tools', 'log_json_backup.txt')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

def log_print(*args, **kwargs):
    msg = " ".join(str(a) for a in args)
    print(msg, **kwargs) 
    if 'end' in kwargs and kwargs['end'] == " ": return
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(msg + '\n')

def kirim_telegram_aman(pesan):
    if len(pesan) > 4000: pesan = pesan[:4000] + "\n...[TERPOTONG]"
    try:
        url = f"https://api.telegram.org/bot{config_rahasia.BOT_TOKEN}/sendMessage"
        requests.post(url, data={"chat_id": config_rahasia.CHAT_ID, "text": pesan, "parse_mode": "HTML"}, timeout=10)
    except: pass

if __name__ == "__main__":
    log_print("\n" + "="*50)
    log_print(f"START BACKUP JSON API PADA {hari_str}/{bulan_str}/{tahun_n} | {waktu_mulai.strftime('%H.%M')} WIB")
    log_print("="*50)

    txt_path = os.path.join(BASE_DIR, 'scripts', 'json_backup', 'url_backup.txt')
    if not os.path.exists(txt_path):
        log_print(f"ERROR: File sumber URL tidak ditemukan di {txt_path}")
        sys.exit(1)

    with open(txt_path, 'r', encoding='utf-8') as f:
        urls = [line.strip() for line in f if line.strip() and not line.strip().startswith(('#', '='))]

    stat_copy = 0
    stat_dl = 0
    stat_gagal = 0

    # Gabungkan target Tahun N dan Tahun N-1 menjadi antrean tugas
    daftar_tahun = [tahun_n, tahun_n - 1]
    tasks = [(t, u) for t in daftar_tahun for u in urls]

    tahun_aktif = None
    for tahun, raw_url in tasks:
        if tahun != tahun_aktif:
            log_print("\n" + "="*50)
            log_print(f"MEMPROSES TAHUN {tahun}")
            log_print("="*50)
            tahun_aktif = tahun

        target_url = raw_url.replace('{tahun}', str(tahun))
        is_v1 = '/v1/' in target_url
        tipe = "v1" if is_v1 else "Legacy"

        match = re.search(r'api/(?:v1|legacy)/(.*?)\?', target_url)
        base_name = match.group(1).replace('/', '_') if match else "unknown"
        filename = f"{tipe}_{base_name}_{tahun}.json"
        
        # --- LOGIKA SUB-FOLDER TAHUN (Adaptasi dari diskusi kemarin) ---
        match_tahun = re.search(r'(20\d{2})', filename)
        if match_tahun:
            tahun_file = match_tahun.group(1)
            arsip_target = os.path.join(ARSIP_DIR, tahun_file)
        else:
            arsip_target = ARSIP_DIR
        
        os.makedirs(arsip_target, exist_ok=True)
        output_path = os.path.join(arsip_target, filename)
        # ---------------------------------------------------------------

        # FITUR SKIP 1: Cek apakah sistem utama sudah mendownloadnya hari ini di folder data
        path_data_utama = os.path.join(BASE_DIR, 'data', str(tahun), filename)
        if os.path.exists(path_data_utama):
            tgl_file = datetime.fromtimestamp(os.path.getmtime(path_data_utama)).date()
            if tgl_file == waktu_mulai.date():
                shutil.copy2(path_data_utama, output_path)
                log_print(f"\n[SKIP] {filename} sudah ditarik sistem utama hari ini. -> Di-copy ke arsip.")
                stat_copy += 1
                continue

        # FITUR SKIP 2: Cek apakah file sudah terlanjur didownload langsung ke arsip hari ini
        if os.path.exists(output_path):
            log_print(f"\n[SKIP] {filename} sudah ada di arsip hari ini. -> Kuota API dihemat.")
            stat_copy += 1
            continue

        log_print(f"\nDOWNLOAD [{tipe.upper()}]: {target_url}")

        if is_v1:
            all_data = []
            cursor = None
            req_count = 1
            first_response = None
            success = False 

            while True:
                url_cursor = target_url
                if cursor:
                    sep = "&" if "?" in target_url else "?"
                    url_cursor = f"{target_url}{sep}cursor={cursor}"

                page_success = False
                resp_data = None

                for i in range(1, 6):
                    try:
                        log_print(f"  Request ke-{req_count} (Coba {i})...", end=" ")
                        response = requests.get(url_cursor, headers=HEADERS, timeout=150)
                        if response.status_code == 200:
                            resp_data = response.json()
                            log_print("SUKSES")
                            page_success = True
                            success = True
                            break
                        else:
                            log_print(f"GAGAL (Status {response.status_code})")
                    except Exception as e:
                        log_print(f"ERROR: Terputus / {str(e)}")
                    
                    if i < 5 and not page_success: time.sleep(i * 5)

                if not page_success:
                    log_print("  GAGAL TOTAL pada request ini. Berhenti ditarik.")
                    success = False 
                    break

                if req_count == 1: first_response = resp_data

                if resp_data and 'data' in resp_data:
                    isi_data = resp_data['data']
                    if isi_data is not None: all_data.extend(isi_data)
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
                stat_gagal += 1
            else:
                if len(all_data) == 0:
                    with open(output_path, 'w', encoding='utf-8') as f: f.write("[]")
                    log_print(f"  -> File dibuat dengan daftar kosong [] (Total: 0 baris)")
                else:
                    with open(output_path, 'w', encoding='utf-8') as f:
                        json.dump(all_data, f, ensure_ascii=False, indent=2)
                    log_print(f"  -> Disimpan ke {filename} (Total: {len(all_data)} baris)")
                stat_dl += 1

        else: # LOGIKA LEGACY
            success = False
            for i in range(1, 6):
                try:
                    log_print(f"  Percobaan ke-{i}...", end=" ")
                    response = requests.get(target_url, headers=HEADERS, timeout=150)
                    if response.status_code == 200:
                        with open(output_path, 'w', encoding='utf-8') as f:
                            json.dump(response.json(), f, ensure_ascii=False, indent=2)
                        log_print("SUKSES")
                        success = True
                        break
                    else:
                        log_print(f"GAGAL (Status {response.status_code})")
                except Exception as e:
                    log_print(f"ERROR: Terputus / {str(e)}")

                if i < 5 and not success: time.sleep(i * 5)

            if not success:
                if not os.path.exists(output_path):
                    log_print(f"  GAGAL TOTAL -> buat file kosong")
                    with open(output_path, 'w', encoding='utf-8') as f: f.write("[]")
                stat_gagal += 1
            else:
                stat_dl += 1

    durasi = str(datetime.now(tz_wib) - waktu_mulai).split('.')[0]
    log_print("\n" + "="*50)
    log_print(f"PROSES BACKUP JSON SELESAI | Durasi: {durasi}")
    log_print("="*50)

    # Kirim Laporan ke Telegram
    pesan_tg = (
        f"✅ <b>BACKUP JSON API SELESAI</b> ✅\n\n"
        f"📅 <b>Tanggal:</b> {waktu_mulai.strftime('%d/%m/%Y')}\n"
        f"⏱ <b>Durasi:</b> {durasi}\n\n"
        f"<b>Statistik:</b>\n"
        f"🔄 Copy Lokal (Skip API): {stat_copy} file\n"
        f"⬇️ Sukses Download: {stat_dl} file\n"
        f"❌ Gagal/Error: {stat_gagal} file"
    )
    kirim_telegram_aman(pesan_tg)