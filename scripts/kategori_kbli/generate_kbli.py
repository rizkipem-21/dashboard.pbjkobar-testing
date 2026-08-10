import os
import io
import json
import requests
import pdfplumber
import pandas as pd
import re 

from datetime import datetime, timedelta, timezone

def get_waktu_indonesia():
    tz_wib = timezone(timedelta(hours=7))
    sekarang = datetime.now(tz_wib)
    hari_indo = {0: 'Senin', 1: 'Selasa', 2: 'Rabu', 3: 'Kamis', 4: 'Jumat', 5: 'Sabtu', 6: 'Minggu'}
    bulan_indo = {1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April', 5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'}
    
    nama_hari = hari_indo[sekarang.weekday()]
    nama_bulan = bulan_indo[sekarang.month]
    return f"{nama_hari}, {sekarang.day} {nama_bulan} {sekarang.year} | {sekarang.strftime('%H:%M:%S')} WIB"

# ==========================================
# KONFIGURASI FOLDER DINAMIS
# ==========================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))

# --- SISTEM LOGGING ---
LOG_FILE = os.path.join(BASE_DIR, 'tools', 'log_kbli.txt')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

_original_print = print

def log_print(*args, **kwargs):
    msg = " ".join(str(a) for a in args)
    _original_print(msg, **kwargs)
    if 'end' in kwargs and kwargs['end'] == " ": return
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(msg + '\n')

print = log_print
# --------------------------------------------

URL_FILE = os.path.join(SCRIPT_DIR, 'url_kategori.txt')
OUTPUT_DIR = os.path.join(BASE_DIR, 'data_master')
OUTPUT_JSON = os.path.join(OUTPUT_DIR, 'data_kategori_inaproc.json')

def main():
    # Menangkap waktu mulai untuk kalkulasi durasi
    tz_wib = timezone(timedelta(hours=7))
    waktu_mulai_obj = datetime.now(tz_wib)
    
    print("="*50)
    print("MEMULAI PROSES EKSTRAKSI DATA KBLI INAPROC")
    print(f"Waktu Mulai: {get_waktu_indonesia()}")
    print("="*50)

    if not os.path.exists(URL_FILE):
        print(f"❌ File tidak ditemukan: {URL_FILE}")
        return
    
    urls = []
    with open(URL_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            if line.strip():
                parts = line.split('->')
                url = parts[0].strip()
                judul = parts[1].split('(')[0].strip() if len(parts) > 1 else "PDF"
                urls.append((url, judul))
        
    if not urls:
        print("❌ File url_kategori.txt kosong! Silakan isi dengan link PDF.")
        return

    semua_baris = []

    for indeks, (url, judul) in enumerate(urls, start=1):
        print(f"\n⏳ [{indeks}/{len(urls)}] Memproses {judul} dari: {url}")
        try:
            jumlah_baris_awal = len(semua_baris)
            response = requests.get(url, timeout=60)
            response.raise_for_status()
            pdf_bytes = io.BytesIO(response.content)
            
            with pdfplumber.open(pdf_bytes) as pdf:
                global_kbli_kode = ""
                global_kbli_deskripsi = ""
                sedang_baca_kbli = False
                
                for page_num, page in enumerate(pdf.pages, 1):
                    teks = page.extract_text() or ""
                    
                    pola_kbli_teks = re.findall(r'([A-Za-z\s\-\,]+?)\s*\((\d{5})\)', teks)
                    if pola_kbli_teks:
                        global_kbli_deskripsi = " | ".join([m[0].strip().replace('\n', ' ') for m in pola_kbli_teks])
                        global_kbli_kode = " | ".join([m[1] for m in pola_kbli_teks])
                    
                    tabel_halaman = page.extract_tables()
                    for tabel in tabel_halaman:
                        if not tabel or len(tabel[0]) < 3:
                            continue
                            
                        # 1. PENGONTROL SAKLAR (Prioritas Dibalik + Regex Anti-Bocor)
                        header_teks = str(tabel[:2]).replace('\\n', ' ').replace('\n', ' ').lower()
                        kolom1_header = str(tabel[0][0]).replace('\n', ' ').strip().lower() if tabel[0][0] else ""
                        
                        # PRIORITAS 1: SENSOR MATI (Cek Keras Tabel Sampah Dulu)
                        if "keterangan" in header_teks or "profil" in header_teks or "informasi pokok" in header_teks or "informasi utama" in header_teks or "koleksi" in header_teks:
                            sedang_baca_kbli = False
                        elif kolom1_header in ['no', 'no.', 'nomor', 'atribut', 'syarat', 'kriteria', 'harga', 'spesifikasi', 'uraian', 'informasi']:
                            sedang_baca_kbli = False
                            
                        # PRIORITAS 2: SENSOR NYALA (Pendeteksi Tingkat 1 atau I yang Akurat)
                        elif re.search(r'\btingkat\s*(i|1)\b', header_teks) or re.search(r'\bkategori tingkat\s*(i|1)\b', header_teks):
                            sedang_baca_kbli = True
                                
                        if not sedang_baca_kbli:
                            continue
                            
                        # 2. PROSES BARIS
                        for baris in tabel:
                            baris_bersih = [str(sel).replace('\n', ' ').strip() if sel else None for sel in baris]
                            
                            # Abaikan baris judul (Menggunakan Regex agar akurat untuk Romawi & Angka Arab)
                            gabungan_teks = str(baris).replace('\\n', ' ').replace('\n', ' ').lower()
                            if re.search(r'\btingkat\s*(i|1)\b', gabungan_teks) and re.search(r'\btingkat\s*(ii|2)\b', gabungan_teks):
                                continue
                                
                            if all(sel is None or str(sel).strip() in ['', 'none'] for sel in baris_bersih):
                                continue
                                
                            # 3. JARING PENGAMAN GANDA (Hanya untuk sampah absolut)
                            kolom1 = str(baris_bersih[0]).strip() if baris_bersih and baris_bersih[0] else ""
                            kolom1_lower = kolom1.lower()
                            
                            if kolom1_lower in ['no', 'no.', 'nomor', 'atribut', 'syarat', 'kriteria', 'harga', 'spesifikasi', 'uraian', 'informasi', 'kategori tingkat i', 'kategori tingkat 1', 'kategori produk']:
                                continue
                            
                            if re.match(r'^[\d\W]+$', kolom1):
                                continue
                                
                            # 4. PENYESUAIAN KOLOM
                            panjang_kolom = len(baris_bersih)
                            if 3 <= panjang_kolom <= 4:
                                baris_bersih = baris_bersih[:3]
                                baris_bersih.extend([global_kbli_kode, global_kbli_deskripsi, None, None])
                            elif 5 <= panjang_kolom <= 6:
                                baris_bersih = baris_bersih[:5]
                                baris_bersih.extend([None, None])
                            elif panjang_kolom >= 7:
                                baris_bersih = baris_bersih[:7]
                            else:
                                continue
                                
                            semua_baris.append(baris_bersih)
            
            baris_baru = len(semua_baris) - jumlah_baris_awal
            
            if baris_baru > 0:
                print(f"✅ Selesai ekstrak: {url} (+{baris_baru} baris)")
            else:
                print(f"⚠️ PDF terbaca, namun tidak ada format tabel KBLI yang valid: {url}")
                
        except Exception as e:
            print(f"❌ Gagal memproses {url}: {e}")

    if not semua_baris:
        print("⚠️ Tidak ada data tabel yang berhasil diekstrak. Periksa kembali struktur PDF.")
        return

    # Transformasi Pandas
    print("🧹 Membersihkan data sel yang digabung (Forward Fill & Explode)...")
    kolom = [
        "Kategori_Tingkat_I", "Kategori_Tingkat_II", "Kategori_Tingkat_III",
        "KBLI_2020_Kode", "KBLI_2020_Deskripsi", "KBLI_2025_Kode", "KBLI_2025_Deskripsi"
    ]
    
    df = pd.DataFrame(semua_baris, columns=kolom)
    
    # 1. Standarisasi nilai kosong
    df.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
    df.replace('None', pd.NA, inplace=True)
    df.replace(r'^\s*-\s*$', pd.NA, regex=True, inplace=True) # Hancurkan strip "-" menjadi NaN
    
    # 2. Forward Fill KHUSUS Kategori Terlebih Dahulu
    kolom_kategori = ["Kategori_Tingkat_I", "Kategori_Tingkat_II", "Kategori_Tingkat_III"]
    df[kolom_kategori] = df[kolom_kategori].ffill()
    
    # 3. Forward Fill KBLI Terisolasi (Mencegah KBLI tumpah ke kategori/tabel lain)
    kolom_kbli = ['KBLI_2020_Kode', 'KBLI_2020_Deskripsi', 'KBLI_2025_Kode', 'KBLI_2025_Deskripsi']
    df[kolom_kbli] = df.groupby(kolom_kategori)[kolom_kbli].ffill()
    
    # 4. Sapu Bersih Tabel Phantom (Hapus baris yang tidak memiliki KBLI sama sekali)
    df.dropna(subset=['KBLI_2020_Kode', 'KBLI_2025_Kode'], how='all', inplace=True)
    df.dropna(how='all', inplace=True)

    if 'Kategori_Tingkat_I' in df.columns:
        df = df[~df['Kategori_Tingkat_I'].astype(str).str.contains('Kategori Tingkat', na=False, case=False)]

    # 5. Explode Data (Memecah KBLI gabungan '|' menjadi baris-baris terpisah)
    df['KBLI_2020_Kode'] = df['KBLI_2020_Kode'].astype(str).str.split(r'\s*\|\s*')
    df['KBLI_2020_Deskripsi'] = df['KBLI_2020_Deskripsi'].astype(str).str.split(r'\s*\|\s*')
    
    df = df.explode(['KBLI_2020_Kode', 'KBLI_2020_Deskripsi'], ignore_index=True)
    df.replace(['nan', '<NA>', 'NaN'], pd.NA, inplace=True)
    
    # 6. Pemisahan Otomatis Kata Menempel (HurufKecilBesar -> HurufKecil Besar)
    kolom_teks = ["Kategori_Tingkat_I", "Kategori_Tingkat_II", "Kategori_Tingkat_III", "KBLI_2020_Deskripsi", "KBLI_2025_Deskripsi"]
    for col in kolom_teks:
        if col in df.columns:
            mask = df[col].notna()
            if mask.any():
                df.loc[mask, col] = df.loc[mask, col].astype(str).str.replace(r'([a-z])([A-Z])', r'\1 \2', regex=True)

    csv_kamus = os.path.join('data_master', 'kamus_pembersihan.csv')

    print(f"\n[INFO] Sedang mencari kamus di: {csv_kamus}")
    
    if os.path.exists(csv_kamus):
        df_kamus = pd.read_csv(csv_kamus, sep=None, engine='python')
        print(f"[INFO] Kamus ditemukan! Memproses {len(df_kamus)} kata perbaikan...")
        
        for _, row in df_kamus.iterrows():
            if 'kata_salah' in row and 'kata_benar' in row:
                kata_salah_mentah = str(row['kata_salah']).strip()
                pola_regex = kata_salah_mentah.replace(' ', r'\s+') 
                kata_benar = str(row['kata_benar']).strip()
                
                for col in df.columns:
                    mask = df[col].notna()
                    if mask.any():
                        df.loc[mask, col] = df.loc[mask, col].astype(str).str.replace(pola_regex, kata_benar, regex=True).str.strip()
                        
        print("[INFO] Pembersihan teks berhasil diterapkan ke semua kolom.\n")
    else:
        print("\n[PERINGATAN] File kamus_pembersihan.csv TIDAK DITEMUKAN!\n")

    # =========================================================
    # 5. SAPU BERSIH DUPLIKAT (Mengatasi tabel yang diulang di PDF)
    # =========================================================
    jumlah_sebelum = len(df)
    df = df.drop_duplicates(ignore_index=True)
    jumlah_setelah = len(df)
    if jumlah_sebelum != jumlah_setelah:
        print(f"🗑️ Berhasil membuang {jumlah_sebelum - jumlah_setelah} baris data duplikat (Tabel berulang).")

    # Mengubah nilai NaN dari Pandas menjadi None agar menjadi 'null' di format JSON
    df = df.astype(object).where(pd.notnull(df), None)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    print("💾 Menyimpan ke format JSON...")
    hasil_json = df.to_dict(orient='records')
    
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(hasil_json, f, ensure_ascii=False, indent=4)
        
    print(f"✅ Selesai! {len(df)} baris data berhasil disimpan ke:\n📂 {OUTPUT_JSON}")
    
    # --- MENGHITUNG DURASI WAKTU ---
    waktu_selesai_obj = datetime.now(tz_wib)
    durasi = waktu_selesai_obj - waktu_mulai_obj
    durasi_str = str(durasi).split('.')[0] # Menghilangkan format microsecond
    
    print("="*50)
    print(f"PROSES SELESAI PADA: {get_waktu_indonesia()}")
    print(f"⏳ TOTAL DURASI PROSES: {durasi_str}")
    print("="*50)

if __name__ == "__main__":
    main()