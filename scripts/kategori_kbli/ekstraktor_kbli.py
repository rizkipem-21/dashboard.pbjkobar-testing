import os
import io
import json
import requests
import pdfplumber
import pandas as pd

# ==========================================
# KONFIGURASI FOLDER DINAMIS
# ==========================================
# Menentukan lokasi direktori saat ini (folder Kategori_KBLI)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Mundur 2 langkah untuk mencapai folder utama (dashboard.pbjkobar-testing)
BASE_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))

# Path untuk file input dan output
URL_FILE = os.path.join(SCRIPT_DIR, 'url_kategori.txt')
OUTPUT_DIR = os.path.join(BASE_DIR, 'data_master')
OUTPUT_JSON = os.path.join(OUTPUT_DIR, 'data_kategori_inaproc.json')

def main():
    print("="*50)
    print("MAMULAI PROSES EKSTRAKSI DATA KBLI INAPROC")
    print("="*50)

    # 1. Baca URL dari txt
    if not os.path.exists(URL_FILE):
        print(f"❌ File tidak ditemukan: {URL_FILE}")
        return
    
    # Membaca semua baris yang ada isi teksnya
    with open(URL_FILE, 'r', encoding='utf-8') as f:
        urls = [line.split('->')[0].strip() for line in f if line.strip()]
        
    if not urls:
        print("❌ File url_kategori.txt kosong! Silakan isi dengan link PDF.")
        return

    semua_baris = [] # Wadah utama ditaruh di luar loop agar menampung semua data

    # 2 & 3. Unduh dan Ekstrak Tabel secara berurutan
    for indeks, url in enumerate(urls, start=1):
        print(f"\n⏳ [{indeks}/{len(urls)}] Memproses PDF dari: {url}")
        try:
            response = requests.get(url, timeout=60)
            response.raise_for_status()
            pdf_bytes = io.BytesIO(response.content)
            
            with pdfplumber.open(pdf_bytes) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    teks = page.extract_text() or ""
                    
                    if "Kategori Produk" in teks or "KBLI" in teks or "Bidang Usaha" in teks:
                        tabel_halaman = page.extract_tables()
                        for tabel in tabel_halaman:
                            if tabel and len(tabel[0]) >= 7:
                                for baris in tabel:
                                    gabungan_teks = str(baris).lower()
                                    if "tingkat i" not in gabungan_teks and "kode" not in gabungan_teks:
                                        baris_bersih = [str(sel).replace('\n', ' ').strip() if sel else None for sel in baris[:7]]
                                        semua_baris.append(baris_bersih)
            print(f"✅ Selesai ekstrak: {url}")
        except Exception as e:
            print(f"❌ Gagal memproses link ke-{indeks} ({url}): {e}")

    if not semua_baris:
        print("⚠️ Tidak ada data tabel yang berhasil diekstrak. Periksa kembali struktur PDF.")
        return

    # 4. Transformasi & Pembersihan (Forward Fill) dengan Pandas
    print("🧹 Membersihkan data sel yang digabung (Forward Fill)...")
    kolom = [
        "Kategori_Tingkat_I", "Kategori_Tingkat_II", "Kategori_Tingkat_III",
        "KBLI_2020_Kode", "KBLI_2020_Deskripsi", "KBLI_2025_Kode", "KBLI_2025_Deskripsi"
    ]
    
    df = pd.DataFrame(semua_baris, columns=kolom)
    
    # Standardisasi format kosong menjadi format NaN milik Pandas
    df.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
    df.replace('None', pd.NA, inplace=True)
    
    # Eksekusi Forward Fill pada 3 kolom pertama
    kolom_ffill = ["Kategori_Tingkat_I", "Kategori_Tingkat_II", "Kategori_Tingkat_III"]
    df[kolom_ffill] = df[kolom_ffill].ffill()
    
    # Buang baris yang benar-benar kosong di semua kolom
    df.dropna(how='all', inplace=True)

    # === BLOK OTOMATISASI PEMBERSIHAN DATA (FINAL) ===
    # 1. Hapus baris sampah yang berisi judul/header tabel
    if 'Kategori_Tingkat_I' in df.columns:
        df = df[~df['Kategori_Tingkat_I'].astype(str).str.contains('Kategori Tingkat', na=False, case=False)]

    # 2. Terapkan Kamus Perbaikan CSV
    csv_kamus = os.path.join('data_master', 'kamus_pembersihan.csv')
    print(f"\n[INFO] Sedang mencari kamus di: {csv_kamus}")
    
    if os.path.exists(csv_kamus):
        df_kamus = pd.read_csv(csv_kamus, sep=None, engine='python')
        print(f"[INFO] Kamus ditemukan! Memproses {len(df_kamus)} kata perbaikan...")
        
        for _, row in df_kamus.iterrows():
            if 'kata_salah' in row and 'kata_benar' in row:
                # Ubah spasi menjadi \s+ agar mendeteksi spasi biasa maupun enter/karakter gaib PDF
                kata_salah_mentah = str(row['kata_salah']).strip()
                pola_regex = kata_salah_mentah.replace(' ', r'\s+') 
                kata_benar = str(row['kata_benar']).strip()
                
                for col in df.columns:
                    # Ambil baris yang ada isinya saja (menghindari error NaN)
                    mask = df[col].notna()
                    if mask.any():
                        # Paksa menjadi teks dan terapkan perbaikan dari kamus
                        df.loc[mask, col] = df.loc[mask, col].astype(str).str.replace(pola_regex, kata_benar, regex=True).str.strip()
                        
        print("[INFO] Pembersihan teks berhasil diterapkan ke semua kolom.\n")
    else:
        print("\n[PERINGATAN] File kamus_pembersihan.csv TIDAK DITEMUKAN!\n")
    # =========================================

    # BARIS TAMBAHAN: Ubah semua NaN menjadi None agar menjadi 'null' di JSON
    df = df.astype(object).where(pd.notnull(df), None)

    # 5. Simpan Hasil ke JSON
    # Pastikan folder data_master ada
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    print("💾 Menyimpan ke format JSON...")
    hasil_json = df.to_dict(orient='records')
    
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(hasil_json, f, ensure_ascii=False, indent=4)
        
    print(f"✅ Selesai! {len(df)} baris data berhasil disimpan ke:\n📂 {OUTPUT_JSON}")

if __name__ == "__main__":
    main()