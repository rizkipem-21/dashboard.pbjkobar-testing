import os
import json
import time
import requests
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

# ======================================================
# KONFIGURASI PATH DAN TAHUN OTOMATIS
# ======================================================
BASE_DIR = r'D:\rup-2026-inaproc'
MASTER_DIR = os.path.join(BASE_DIR, 'data', 'master')
KAMUS_FILE = os.path.join(MASTER_DIR, 'kamus_penyedia.json')

TOKEN = 'inprc7642391c38774272bf57ca25ac1d4544'
HEADERS = {
    'Authorization': f'Bearer {TOKEN}',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*'
}

tahun_n  = datetime.now().year
tahun_n1 = tahun_n - 1
tahun_n2 = tahun_n - 2
daftar_tahun = [tahun_n2, tahun_n1, tahun_n]

def get_unique_codes(file_path, id_column):
    """Mengekstrak ID unik dari file JSON."""
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        return set()
    try:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
            if isinstance(data, dict):
                for k in ['data', 'items', 'results']:
                    if k in data and isinstance(data[k], list):
                        data = data[k]
                        break
            if not isinstance(data, list): return set()
            codes = set(str(item.get(id_column)) for item in data if item.get(id_column))
            return {c.strip() for c in codes if c.lower() != "none" and c.strip() != ""}
    except Exception:
        return set()

def fetch_data_api(kode, is_archive=False):
    """Tembak API dan ambil parameter/data penyedia sesuai aslinya"""
    endpoint = "ekatalog-archive" if is_archive else "ekatalog"
    url = f"https://data.inaproc.id/api/legacy/{endpoint}/penyedia-detail?kode_penyedia={kode}"
    
    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            
            # API Inaproc mengembalikan data berupa List/Array -> [ { ... } ]
            if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
                objek_penyedia = data[0]
                
                # Memastikan ID penyedia (S6/S7) selalu tersimpan sebagai pengaman
                if is_archive and 'kd_penyedia' not in objek_penyedia:
                    objek_penyedia['kd_penyedia'] = kode
                elif not is_archive and 'kode_penyedia' not in objek_penyedia:
                    objek_penyedia['kode_penyedia'] = kode
                    
                return objek_penyedia
                
            elif isinstance(data, dict):
                if isinstance(data.get('data'), list) and len(data['data']) > 0:
                    return data['data'][0]
                elif isinstance(data.get('data'), dict):
                    return data['data']
                else:
                    return data
                    
            return None
                
        elif response.status_code == 401:
            print("\n[FATAL] Token Expired / IP Diblokir!")
            return "UNAUTHORIZED"
        else:
            return None
            
    except Exception as e:
        return None

def main():
    print("="*55)
    print(f"  UPDATE KAMUS PENYEDIA FULL DATA (Tahun {tahun_n2}, {tahun_n1}, {tahun_n})")
    print("="*55)

    os.makedirs(MASTER_DIR, exist_ok=True)

    # 1. Load Kamus (Format sekarang: Array List [ {...}, {...} ])
    kamus_list = []
    if os.path.exists(KAMUS_FILE) and os.path.getsize(KAMUS_FILE) > 0:
        with open(KAMUS_FILE, 'r', encoding='utf-8') as f:
            kamus_list = json.load(f)
        print(f"[*] Kamus termuat. Isi: {len(kamus_list)} data penyedia.")

    # 2. Ambil semua ID yang SUDAH ADA di dalam kamus agar tidak terduplikasi
    existing_ids = set()
    for item in kamus_list:
        if item.get('kode_penyedia'):
            existing_ids.add(str(item['kode_penyedia']))
        if item.get('kd_penyedia'):
            existing_ids.add(str(item['kd_penyedia']))

    # 3. Kumpulkan Kode Unik dari file sumber json lokal
    semua_kode_s6 = set()
    semua_kode_s7 = set()

    for tahun in daftar_tahun:
        data_dir = os.path.join(BASE_DIR, 'data', str(tahun))
        file_s6 = os.path.join(data_dir, f'Legacy_ekatalog_paket-e-purchasing_{tahun}.json')
        file_s7 = os.path.join(data_dir, f'Legacy_ekatalog-archive_paket-e-purchasing_{tahun}.json')
        semua_kode_s6.update(get_unique_codes(file_s6, 'kode_penyedia'))
        semua_kode_s7.update(get_unique_codes(file_s7, 'kd_penyedia'))

    # 4. Cari kode yang belum ada di dalam existing_ids
    missing_s6 = [c for c in semua_kode_s6 if c not in existing_ids]
    missing_s7 = [c for c in semua_kode_s7 if c not in existing_ids and c not in missing_s6]

    total_missing = len(missing_s6) + len(missing_s7)
    if total_missing == 0:
        print("\n[SELESAI] Semua kode sudah ada di kamus.")
        return

    print(f"\n[*] Mulai fetch {total_missing} data baru dari API Inaproc...")
    count_new = 0

    try:
        # Pengecekan Sumber 6
        for i, kode in enumerate(missing_s6, 1):
            print(f"[S6] {i}/{len(missing_s6)} | ID: {kode[:10]}...", end=" ")
            data_penyedia = fetch_data_api(kode, is_archive=False)
            
            if data_penyedia == "UNAUTHORIZED": break
            if data_penyedia:
                kamus_list.append(data_penyedia) # Tambahkan ke dalam array
                count_new += 1
                print(f"Sukses -> {data_penyedia.get('nama_penyedia', 'Tanpa Nama')}")
            else:
                print("Gagal / Kosong")
            time.sleep(0.2)

        # Pengecekan Sumber 7
        for i, kode in enumerate(missing_s7, 1):
            print(f"[S7] {i}/{len(missing_s7)} | ID: {kode[:10]}...", end=" ")
            data_penyedia = fetch_data_api(kode, is_archive=True)
            
            if data_penyedia == "UNAUTHORIZED": break
            if data_penyedia:
                kamus_list.append(data_penyedia) # Tambahkan ke dalam array
                count_new += 1
                print(f"Sukses -> {data_penyedia.get('nama_penyedia', 'Tanpa Nama')}")
            else:
                print("Gagal / Kosong")
            time.sleep(0.2)
            
    except KeyboardInterrupt:
        print("\n[!] Dihentikan manual oleh user. Menyimpan data yang sudah didapat...")

    # 5. Simpan ke JSON dengan format Array List
    with open(KAMUS_FILE, 'w', encoding='utf-8') as f:
        json.dump(kamus_list, f, ensure_ascii=False, indent=2)

    print(f"\n[SELESAI] Tambah {count_new} penyedia. Total di kamus: {len(kamus_list)}")

if __name__ == '__main__':
    main()