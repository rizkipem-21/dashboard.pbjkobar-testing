import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime

# ==========================================
# SETUP PATH FILE JSON
# ==========================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MASTER_DIR = os.path.join(BASE_DIR, 'data_master')
DATA_FILE = os.path.join(MASTER_DIR, 'manual_konsolidasi.json')

os.makedirs(MASTER_DIR, exist_ok=True)

# ==========================================
# DAFTAR SATUAN KERJA
# ==========================================
SATKER_LIST = [
  "Badan Kepegawaian dan Pengembangan Sumber Daya Manusia",
  "Badan Kesatuan Bangsa dan Politik",
  "Badan Keuangan dan Aset Daerah",
  "Badan Penanggulangan Bencana Daerah",
  "Badan Pendapatan Daerah",
  "Badan Perencanaan Pembangunan, Riset dan Inovasi Daerah",
  "Dinas Kepemudaan dan Olahraga",
  "Dinas Kependudukan dan Pencatatan Sipil",
  "Dinas Kesehatan",
  "Dinas Komunikasi, Informatika, Statistik dan Persandian",
  "Dinas Lingkungan Hidup",
  "Dinas Pariwisata",
  "Dinas Pekerjaan Umum dan Penataan Ruang",
  "Dinas Pemadam Kebakaran dan Penyelamatan",
  "Dinas Pemberdayaan Masyarakat dan Desa",
  "Dinas Pemberdayaan Perempuan dan Perlindungan Anak, Pengendalian Penduduk dan Keluarga Berencana",
  "Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu",
  "Dinas Pendidikan dan Kebudayaan",
  "Dinas Perhubungan",
  "Dinas Perikanan dan Ketahanan Pangan",
  "Dinas Perindustrian, Perdagangan, Koperasi, Usaha Kecil dan Menengah",
  "Dinas Perpustakaan dan Kearsipan",
  "Dinas Pertanian",
  "Dinas Perumahan Rakyat, Kawasan Pemukiman dan Pertanahan",
  "Dinas Sosial",
  "Dinas Tenaga Kerja dan Transmigrasi",
  "Inspektorat Daerah",
  "Instalasi Farmasi",
  "Kecamatan Arut Selatan",
  "Kecamatan Arut Utara",
  "Kecamatan Kotawaringin Lama",
  "Kecamatan Kumai",
  "Kecamatan Pangkalan Banteng",
  "Kecamatan Pangkalan Lada",
  "Laboratorium Kesehatan Daerah",
  "Puskesmas Arut Selatan",
  "Puskesmas Arut Utara",
  "Puskesmas Ipuh Bangun Jaya",
  "Puskesmas Karang Mulya",
  "Puskesmas Kotawaringin Lama",
  "Puskesmas Kumai",
  "Puskesmas Kumpai Batu Atas",
  "Puskesmas Madurejo",
  "Puskesmas Mendawai",
  "Puskesmas Natai Pelingkau",
  "Puskesmas Pandu Sanjaya",
  "Puskesmas Pangkalan Lada",
  "Puskesmas Riam Durian",
  "Puskesmas Runtu",
  "Puskesmas Sambi",
  "Puskesmas Semanggang",
  "Puskesmas Sungai Rangit",
  "Puskesmas Teluk Bogam",
  "Rumah Sakit Kutaringin",
  "Rumah Sakit Umum Daerah Sultan Imanuddin",
  "Satuan Polisi Pamong Praja",
  "Sekretariat DPRD",
  "Sekretariat Daerah"
]

# ==========================================
# FUNGSI DATA
# ==========================================
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return []
    return []

def save_data(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# ==========================================
# KELAS APLIKASI GUI
# ==========================================
class AppKonsolidasi(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Panel Admin - Riwayat Revisi Konsolidasi")
        
        # UBAH BAGIAN INI: Ukuran awal lebih ramah layar kecil, lalu otomatis Fullscreen
        self.geometry("1100x600")
        try:
            self.state('zoomed') # Perintah agar otomatis Maximize di Windows
        except:
            pass
            
        self.configure(padx=20, pady=20)
        
        self.db_data = load_data()
        self.edit_index = None
        
        self.vcmd_angka = (self.register(self.hanya_angka), '%P')
        
        # Inisialisasi Style untuk Treeview Row Height
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)
        
        self.create_input_frame()
        self.create_filter_frame()
        self.create_table_frame()
        self.refresh_table()

    def hanya_angka(self, nilai_input):
        if nilai_input == "":
            return True
        return nilai_input.isdigit()

    def cek_panjang_rup(self, event, widget_entry):
        nilai = widget_entry.get().strip()
        if nilai and len(nilai) < 8:
            messagebox.showwarning("Digit Kurang!", "Kode RUP harus terdiri dari minimal 8 digit angka.\nSilakan lengkapi kembali.", parent=self)
            self.after(10, lambda: widget_entry.focus_set())

    def filter_satker(self, event):
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Tab'):
            return
            
        ketikan = self.ent_satker.get()
        if ketikan == "":
            self.ent_satker['values'] = SATKER_LIST
        else:
            hasil_filter = [item for item in SATKER_LIST if ketikan.lower() in item.lower()]
            self.ent_satker['values'] = hasil_filter
            
        self.ent_satker.event_generate('<Down>')

    def klik_satker(self, event):
        self.ent_satker.config(state="normal")
        self.ent_satker.focus_set()
        self.ent_satker.event_generate('<Down>')
        return "break"

    def focusout_satker(self, event):
        isi = self.ent_satker.get().strip()
        if isi and isi not in SATKER_LIST:
            messagebox.showwarning("Satker Tidak Valid", "Silakan pilih Satuan Kerja yang valid dari daftar dropdown.", parent=self)
            self.ent_satker.set("")
            self.after(10, lambda: self.ent_satker.focus_set())
            self.ent_satker.config(state="normal")
            return
            
        self.ent_satker.config(state="readonly")
        self.ent_satker['values'] = SATKER_LIST

    def create_input_frame(self):
        frame_input = ttk.LabelFrame(self, text="  Input Riwayat Konsolidasi Baru  ", padding=(15, 15))
        frame_input.pack(fill="x", pady=(0, 15))

        # Baris 1
        ttk.Label(frame_input, text="Kode RUP Lama:").grid(row=0, column=0, sticky="w", pady=5)
        self.ent_lama = ttk.Entry(frame_input, width=30, validate='key', validatecommand=self.vcmd_angka)
        self.ent_lama.grid(row=0, column=1, padx=(10, 30), pady=5)
        self.ent_lama.bind("<FocusOut>", lambda e: self.cek_panjang_rup(e, self.ent_lama))

        ttk.Label(frame_input, text="Kode RUP Baru:").grid(row=0, column=2, sticky="w", pady=5)
        self.ent_baru = ttk.Entry(frame_input, width=30, validate='key', validatecommand=self.vcmd_angka)
        self.ent_baru.grid(row=0, column=3, padx=(10, 0), pady=5)
        self.ent_baru.bind("<FocusOut>", lambda e: self.cek_panjang_rup(e, self.ent_baru))

        # Baris 2
        ttk.Label(frame_input, text="Nama Paket/Ket:").grid(row=1, column=0, sticky="w", pady=5)
        self.ent_nama = ttk.Entry(frame_input, width=30)
        self.ent_nama.grid(row=1, column=1, padx=(10, 30), pady=5)

        ttk.Label(frame_input, text="Satuan Kerja:").grid(row=1, column=2, sticky="w", pady=5)
        self.ent_satker = ttk.Combobox(frame_input, width=27, values=SATKER_LIST, height=10, state="readonly")
        self.ent_satker.grid(row=1, column=3, padx=(10, 0), pady=5)
        
        self.ent_satker.bind('<Button-1>', self.klik_satker)
        self.ent_satker.bind('<KeyRelease>', self.filter_satker)
        self.ent_satker.bind('<FocusOut>', self.focusout_satker)

        # Baris 3
        ttk.Label(frame_input, text="Alasan Revisi:").grid(row=2, column=0, sticky="w", pady=5)
        self.ent_alasan = ttk.Entry(frame_input, width=30)
        self.ent_alasan.grid(row=2, column=1, padx=(10, 30), pady=5)

        ttk.Label(frame_input, text="Tahun Anggaran:").grid(row=2, column=2, sticky="w", pady=5)
        self.ent_tahun = ttk.Combobox(frame_input, width=27, values=["2024", "2025", "2026"], state="readonly")
        self.ent_tahun.grid(row=2, column=3, padx=(10, 0), pady=5)

        # Tombol Simpan
        self.btn_simpan = ttk.Button(frame_input, text="Simpan Data", command=self.simpan_baru)
        self.btn_simpan.grid(row=3, column=0, columnspan=4, pady=(15, 0))

    def create_filter_frame(self):
        frame_filter = ttk.LabelFrame(self, text="  Panel Filter Pencarian  ", padding=(10, 10))
        frame_filter.pack(fill="x", pady=(0, 15))

        ttk.Label(frame_filter, text="Box Pencarian Universal (Cari RUP Lama / Baru / Alasan):").grid(row=0, column=0, sticky="w", padx=5)
        self.ent_cari = ttk.Entry(frame_filter, width=45)
        self.ent_cari.grid(row=0, column=1, padx=10, sticky="w")
        self.ent_cari.bind("<KeyRelease>", lambda e: self.refresh_table())

        ttk.Label(frame_filter, text="Filter Status:").grid(row=0, column=2, padx=(30, 10), sticky="w")
        self.var_aktif = tk.BooleanVar(value=True)
        self.var_nonaktif = tk.BooleanVar(value=True)

        self.chk_aktif = ttk.Checkbutton(frame_filter, text="AKTIF", variable=self.var_aktif, command=self.refresh_table)
        self.chk_aktif.grid(row=0, column=3, padx=10)

        self.chk_nonaktif = ttk.Checkbutton(frame_filter, text="NON-AKTIF", variable=self.var_nonaktif, command=self.refresh_table)
        self.chk_nonaktif.grid(row=0, column=4, padx=10)

    def create_table_frame(self):
        frame_tabel = ttk.LabelFrame(self, text="  Database Riwayat  ", padding=(10, 10))
        frame_tabel.pack(fill="both", expand=True)

        kolom = ("no", "status", "tahun", "lama", "baru", "nama", "satker", "alasan", "tanggal", "diubah")
        self.tree = ttk.Treeview(frame_tabel, columns=kolom, show="headings", selectmode="browse")
        
        self.tree.heading("no", text="No.")
        self.tree.heading("status", text="Status")
        self.tree.heading("tahun", text="Tahun")
        self.tree.heading("lama", text="RUP Lama")
        self.tree.heading("baru", text="RUP Baru")
        self.tree.heading("nama", text="Keterangan")
        self.tree.heading("satker", text="Satuan Kerja")
        self.tree.heading("alasan", text="Alasan")
        self.tree.heading("tanggal", text="Tgl Input")
        self.tree.heading("diubah", text="Terakhir Diubah")

        self.tree.column("no", width=45, anchor="center")
        self.tree.column("status", width=90, anchor="center")
        self.tree.column("tahun", width=65, anchor="center")
        self.tree.column("lama", width=95, anchor="center")
        self.tree.column("baru", width=95, anchor="center")
        self.tree.column("nama", width=220)
        self.tree.column("satker", width=200)
        self.tree.column("alasan", width=200)
        self.tree.column("tanggal", width=95, anchor="center")
        self.tree.column("diubah", width=140, anchor="center")

        scrollbar = ttk.Scrollbar(frame_tabel, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        frame_aksi = tk.Frame(self)
        frame_aksi.pack(pady=(15, 0))

        btn_toggle = ttk.Button(frame_aksi, text="Ubah Status Aktif / Non-Aktif", command=self.toggle_status)
        btn_toggle.grid(row=0, column=0, padx=10)

        btn_edit = ttk.Button(frame_aksi, text="Edit Data Selected", command=self.muat_data_edit)
        btn_edit.grid(row=0, column=1, padx=10)

        btn_hapus = ttk.Button(frame_aksi, text="Hapus Data (Permanen)", command=self.hapus_data)
        btn_hapus.grid(row=0, column=2, padx=10)

        self.tree.tag_configure('nonaktif', foreground='gray', font=('Arial', 9, 'overstrike'))
        self.tree.tag_configure('aktif', foreground='black')

    def is_teks_valid(self, teks):
        return len(teks) > 1 and any(c.isalnum() for c in teks)

    def muat_data_edit(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Pilih Data", "Silakan klik/pilih salah satu data di tabel untuk diedit.")
            return
        
        index = int(self.tree.item(selected_item[0], "tags")[1])
        self.edit_index = index
        row = self.db_data[index]

        self.ent_lama.delete(0, tk.END)
        self.ent_baru.delete(0, tk.END)
        self.ent_nama.delete(0, tk.END)
        self.ent_satker.config(state="normal")
        self.ent_satker.set("")
        self.ent_alasan.delete(0, tk.END)
        self.ent_tahun.config(state="normal")
        self.ent_tahun.set("")

        self.ent_lama.insert(0, row.get("kode_konsol_lama", ""))
        self.ent_baru.insert(0, row.get("kode_konsol_baru", ""))
        self.ent_nama.insert(0, row.get("nama_paket_keterangan", ""))
        self.ent_satker.set(row.get("nama_satker", ""))
        self.ent_satker.config(state="readonly")
        self.ent_alasan.insert(0, row.get("alasan_revisi", ""))
        self.ent_tahun.set(row.get("tahun_anggaran", ""))
        self.ent_tahun.config(state="readonly")

        self.btn_simpan.config(text="Perbarui Data")
        messagebox.showinfo("Mode Edit", "Data telah ditarik ke atas. Silakan ubah lalu klik 'Perbarui Data'.")

    def simpan_baru(self):
        lama = self.ent_lama.get().strip()
        baru = self.ent_baru.get().strip()
        nama = self.ent_nama.get().strip()
        satker = self.ent_satker.get().strip()
        alasan = self.ent_alasan.get().strip()
        tahun = self.ent_tahun.get().strip()

        if not lama or not baru or not nama or not satker or not alasan or not tahun:
            messagebox.showwarning("Peringatan", "Semua kolom data (RUP Lama, RUP Baru, Nama/Ket, Satuan Kerja, Alasan, Tahun Anggaran) WAJIB diisi semua!")
            return

        if len(lama) < 8 or len(baru) < 8:
            messagebox.showwarning("Peringatan", "Kode RUP Lama and Kode RUP Baru harus terdiri dari minimal 8 digit angka!")
            return

        if not self.is_teks_valid(nama) or not self.is_teks_valid(satker) or not self.is_teks_valid(alasan):
            messagebox.showwarning("Peringatan", "Input teks tidak valid!\n\nTidak boleh hanya berisi 1 karakter atau murni simbol tanpa huruf/angka.")
            return

        if self.edit_index is not None:
            self.db_data[self.edit_index]["kode_konsol_lama"] = lama
            self.db_data[self.edit_index]["kode_konsol_baru"] = baru
            self.db_data[self.edit_index]["nama_paket_keterangan"] = nama
            self.db_data[self.edit_index]["nama_satker"] = satker
            self.db_data[self.edit_index]["alasan_revisi"] = alasan
            self.db_data[self.edit_index]["tahun_anggaran"] = int(tahun)
            self.db_data[self.edit_index]["terakhir_diubah"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            self.edit_index = None
            self.btn_simpan.config(text="Simpan Data")
            msg_sukses = "Data berhasil diperbarui secara aman!"
        else:
            now = datetime.now()
            data_baru = {
                "kode_konsol_lama": lama,
                "kode_konsol_baru": baru,
                "nama_paket_keterangan": nama,
                "nama_satker": satker,
                "alasan_revisi": alasan,
                "tahun_anggaran": int(tahun),
                "tanggal_input": now.strftime("%Y-%m-%d"),
                "waktu_input": now.strftime("%H:%M:%S"),
                "status_aktif": True,
                "terakhir_diubah": "-"
            }
            self.db_data.append(data_baru)
            msg_sukses = "Data berhasil ditambahkan!"

        save_data(self.db_data)
        self.refresh_table()

        self.ent_lama.delete(0, tk.END)
        self.ent_baru.delete(0, tk.END)
        self.ent_nama.delete(0, tk.END)
        self.ent_satker.set("")  
        self.ent_alasan.delete(0, tk.END)
        self.ent_tahun.set("")
        
        self.ent_satker.config(state="readonly")
        self.ent_tahun.config(state="readonly")
        messagebox.showinfo("Sukses", msg_sukses)

    def toggle_status(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Pilih Data", "Silakan klik/pilih salah satu data di tabel terlebih dahulu.")
            return
        
        index = int(self.tree.item(selected_item[0], "tags")[1])
        status_sekarang = self.db_data[index]["status_aktif"]
        
        self.db_data[index]["status_aktif"] = not status_sekarang
        save_data(self.db_data)
        self.refresh_table()

    def hapus_data(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Pilih Data", "Silakan klik/pilih salah satu data di tabel untuk dihapus.")
            return
        
        jawaban = simpledialog.askstring(
            "Konfirmasi Hapus", 
            'PENTING: Data yang dihapus tidak dapat dikembalikan.\n\nKetik kata "HAPUS" (huruf besar) untuk melanjutkan:',
            parent=self
        )
        
        if jawaban == "HAPUS":
            index = int(self.tree.item(selected_item[0], "tags")[1])
            del self.db_data[index]
            save_data(self.db_data)
            self.refresh_table()
            messagebox.showinfo("Sukses", "Data berhasil dihapus secara permanen.")
        elif jawaban is not None:
            messagebox.showerror("Gagal", 'Kata kunci salah! Anda harus mengetik "HAPUS".\nProses penghapusan dibatalkan.')

    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
            
        kata_kunci = self.ent_cari.get().strip().lower()
        show_aktif = self.var_aktif.get()
        show_nonaktif = self.var_nonaktif.get()
        
        no_urut = 1
        for i, row in enumerate(self.db_data):
            status_aktif = row.get("status_aktif", True)
            
            # 1. Filter Status Checklist
            if status_aktif and not show_aktif:
                continue
            if not status_aktif and not show_nonaktif:
                continue
                
            # 2. Filter Box Pencarian Universal (Lama, Baru, Alasan)
            lama_str = str(row.get("kode_konsol_lama", "")).lower()
            baru_str = str(row.get("kode_konsol_baru", "")).lower()
            alasan_str = str(row.get("alasan_revisi", "")).lower()
            
            if kata_kunci and (kata_kunci not in lama_str and kata_kunci not in baru_str and kata_kunci not in alasan_str):
                continue
                
            status_teks = "AKTIF" if status_aktif else "NON-AKTIF"
            tag = "aktif" if status_aktif else "nonaktif"
            
            self.tree.insert("", "end", values=(
                no_urut,
                status_teks,
                row.get("tahun_anggaran", "-"),
                row.get("kode_konsol_lama", ""),
                row.get("kode_konsol_baru", ""),
                row.get("nama_paket_keterangan", ""),
                row.get("nama_satker", ""),
                row.get("alasan_revisi", ""),
                row.get("tanggal_input", ""),
                row.get("terakhir_diubah", "-")
            ), tags=(tag, str(i)))
            no_urut += 1

if __name__ == "__main__":
    app = AppKonsolidasi()
    app.mainloop()