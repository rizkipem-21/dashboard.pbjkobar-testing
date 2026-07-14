import os
import io
import requests
import pdfplumber

# Set path file
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_INPUT = os.path.join(SCRIPT_DIR, 'url_kategori.txt')
FILE_OUTPUT = os.path.join(SCRIPT_DIR, 'url_dengan_nama.txt')

def ekstrak_judul(url):
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        pdf_bytes = io.BytesIO(response.content)
        
        with pdfplumber.open(pdf_bytes) as pdf:
            # Cek 5 halaman pertama saja agar cepat dan efisien
            for page in pdf.pages[:5]:
                tabel_halaman = page.extract_tables()
                
                for tabel in tabel_halaman:
                    for baris in tabel:
                        # Gabungkan teks dalam baris untuk pengecekan cepat (huruf kecil semua)
                        gabungan_teks = " ".join([str(sel).lower() for sel in baris if sel])
                        
                        # Jika menemukan baris "Nama Kategori Produk"
                        if "nama kategori produk" in gabungan_teks:
                            # Bersihkan sel dari kosong atau sekadar titik dua (:)
                            sel_berisi = [str(sel).strip().replace('\n', ' ') for sel in baris if sel and str(sel).strip() not in ['', ':']]
                            
                            # Logikanya: ['1.', 'Nama Kategori Produk', 'Peralatan Pendidikan']
                            # Kita ambil elemen paling akhir dari list tersebut
                            if len(sel_berisi) >= 2:
                                return sel_berisi[-1]
            
            return "KATEGORI TIDAK DITEMUKAN DI TABEL BAB 1"
                
    except Exception as e:
        return f"GAGAL UNDUH/BACA: {str(e)[:30]}"

def main():
    if not os.path.exists(FILE_INPUT):
        print(f"❌ File input tidak ditemukan: {FILE_INPUT}")
        return

    # Baca semua link dari file
    with open(FILE_INPUT, 'r', encoding='utf-8') as f:
        urls = [line.strip() for line in f if line.strip()]

    print(f"Ditemukan {len(urls)} link. Memulai pencarian di Bab 1...\n")

    hasil = []
    for i, url in enumerate(urls, 1):
        print(f"[{i}/{len(urls)}] Memeriksa: {url}")
        judul = ekstrak_judul(url)
        print(f"  -> {judul}")
        hasil.append(f"{url} -> {judul}")

    # Simpan hasil ke file teks baru
    with open(FILE_OUTPUT, 'w', encoding='utf-8') as f:
        f.write("\n".join(hasil))
        
    print(f"\n✅ Selesai! Hasil pemetaan disimpan di: {FILE_OUTPUT}")

if __name__ == "__main__":
    main()