import pandas as pd
import json
import os

# Sesuaikan dengan tahun yang error di log Anda (misal: 2024)
tahun = 2024
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) # Sesuaikan jika posisi file ini berbeda

# Kita cek untuk Penyedia (S1) dan Swakelola (S1_2)
file_penyedia = os.path.join(BASE_DIR, 'data', str(tahun), f'v1_rup_paket-penyedia-terumumkan_{tahun}.json')
file_swakelola = os.path.join(BASE_DIR, 'data', str(tahun), f'v1_rup_paket-swakelola-terumumkan_{tahun}.json')

def cek_duplikat(path_file, nama_sumber):
    if not os.path.exists(path_file):
        print(f"File {nama_sumber} tidak ditemukan di: {path_file}")
        return

    with open(path_file, 'r', encoding='utf-8-sig') as f:
        data = json.load(f)
    
    # Normalisasi data sesuai struktur Inaproc V1
    if isinstance(data, dict):
        if 'data' in data: data = data['data']
        elif 'items' in data: data = data['items']

    df = pd.json_normalize(data)
    
    if 'kd_rup' not in df.columns:
        print(f"Kolom 'kd_rup' tidak ada di {nama_sumber}.")
        return

    # Cari yang duplikat
    df_duplikat = df[df.duplicated(subset=['kd_rup'], keep=False)]
    
    print(f"\n=== HASIL CEK {nama_sumber.upper()} TAHUN {tahun} ===")
    if df_duplikat.empty:
        print("-> AMAN: Tidak ada Kode RUP yang duplikat.")
    else:
        print(f"-> DITEMUKAN: Ada {len(df_duplikat)} baris yang saling duplikat!")
        # Menampilkan Kode RUP dan Nama Paketnya agar mudah dicek manual
        kolom_tampil = ['kd_rup', 'nama_paket'] if 'nama_paket' in df.columns else ['kd_rup']
        print(df_duplikat[kolom_tampil].sort_values('kd_rup').to_string(index=False))

# Eksekusi pengecekan
cek_duplikat(file_penyedia, "Penyedia")
cek_duplikat(file_swakelola, "Swakelola")