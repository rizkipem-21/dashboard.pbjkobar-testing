import os
import json
import threading
import tkinter as tk
from tkinter import ttk, messagebox

# ==========================================
# SETUP PATH UTAMA
# ==========================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data')

class AppPelacakJSON(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Detective - Pelacak & Perbandingan Kode Paket JSON")
        self.geometry("1100x700")
        self.minsize(950, 600)
        self.configure(padx=10, pady=10)
        
        try:
            self.state('zoomed')
        except:
            pass

        self.hasil_pencarian = {} 
        self.mode_perbandingan = False 
        self.dict_file_viewer = {} # Menyimpan path lengkap file untuk Tab 2
        
        # === MEMBUAT TAB MENU (NOTEBOOK) ===
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)
        
        self.tab1 = ttk.Frame(self.notebook, padding=10)
        self.tab2 = ttk.Frame(self.notebook, padding=10)
        
        self.notebook.add(self.tab1, text=" 🔍 Pencari & Pembanding ")
        self.notebook.add(self.tab2, text=" 📂 Viewer Data JSON ")
        
        # Setup Konten Tab 1
        self.create_top_frame_tab1()
        self.create_main_layout_tab1()
        
        # Setup Konten Tab 2
        self.create_layout_tab2()

    # ==========================================
    # KONTEN TAB 1: PENCARI & PEMBANDING
    # ==========================================
    def create_top_frame_tab1(self):
        self.frame_top = ttk.LabelFrame(self.tab1, text="  Mesin Pencari & Pembanding  ", padding=(15, 15))
        self.frame_top.pack(side="top", fill="x", pady=(0, 15))

        ttk.Label(self.frame_top, text="Kata Kunci 1:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.ent_cari = ttk.Entry(self.frame_top, width=30, font=('Arial', 11))
        self.ent_cari.grid(row=0, column=1, padx=(0, 15), sticky="w")
        self.ent_cari.bind("<Return>", lambda e: self.mulai_pencarian())

        self.btn_mode_banding = ttk.Button(self.frame_top, text="🆚 Aktifkan Perbandingan", command=self.toggle_mode_perbandingan)
        self.btn_mode_banding.grid(row=0, column=2, padx=(0, 15), sticky="w")

        self.btn_cari = ttk.Button(self.frame_top, text="🔍 Lacak Data", command=self.mulai_pencarian)
        self.btn_cari.grid(row=0, column=3, padx=(0, 10), sticky="w")

        self.lbl_status = ttk.Label(self.frame_top, text="Siap digunakan.", foreground="gray")
        self.lbl_status.grid(row=0, column=4, sticky="e", padx=(20, 0))
        self.frame_top.columnconfigure(4, weight=1) 

        self.lbl_cari_2 = ttk.Label(self.frame_top, text="Kata Kunci 2:", font=('Arial', 10, 'bold'))
        self.ent_cari_2 = ttk.Entry(self.frame_top, width=30, font=('Arial', 11))
        self.ent_cari_2.bind("<Return>", lambda e: self.mulai_pencarian())

    def toggle_mode_perbandingan(self):
        self.mode_perbandingan = not self.mode_perbandingan
        if self.mode_perbandingan:
            self.btn_mode_banding.config(text="✖️ Matikan Perbandingan")
            self.lbl_cari_2.grid(row=1, column=0, sticky="w", padx=(0, 5), pady=(10, 0))
            self.ent_cari_2.grid(row=1, column=1, padx=(0, 15), sticky="w", pady=(10, 0))
            self.lbl_status.grid(row=1, column=4, sticky="e") 
            self.frame_p2.pack(side="left", fill="both", expand=True, padx=(5, 0))
            self.lbl_status.config(text="Mode Perbandingan Aktif. Masukkan Kode ke-2.", foreground="#0288d1")
        else:
            self.btn_mode_banding.config(text="🆚 Aktifkan Perbandingan")
            self.lbl_cari_2.grid_forget()
            self.ent_cari_2.grid_forget()
            self.ent_cari_2.delete(0, tk.END)
            self.lbl_status.grid(row=0, column=4, sticky="e")
            self.frame_p2.pack_forget()
            self.txt_preview_2.delete(1.0, tk.END)
            self.lbl_status.config(text="Kembali ke Mode Tunggal.", foreground="gray")

    def create_main_layout_tab1(self):
        paned_window = ttk.PanedWindow(self.tab1, orient=tk.HORIZONTAL)
        paned_window.pack(fill="both", expand=True)

        frame_kiri = ttk.LabelFrame(paned_window, text="  Ditemukan di File Berikut:  ", padding=5)
        paned_window.add(frame_kiri, weight=1)

        kolom = ("tahun", "file", "jumlah")
        self.tree = ttk.Treeview(frame_kiri, columns=kolom, show="headings", selectmode="browse")
        self.tree.heading("tahun", text="Tahun")
        self.tree.heading("file", text="Nama File JSON")
        self.tree.heading("jumlah", text="Match")
        self.tree.column("tahun", width=60, anchor="center")
        self.tree.column("file", width=250, anchor="w")
        self.tree.column("jumlah", width=80, anchor="center")

        scrollbar_tree = ttk.Scrollbar(frame_kiri, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar_tree.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_tree.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.tampilkan_isi_data)

        self.frame_kanan_container = ttk.Frame(paned_window)
        paned_window.add(self.frame_kanan_container, weight=3)

        self.frame_p1 = ttk.LabelFrame(self.frame_kanan_container, text="  Hasil Kata Kunci 1  ", padding=5)
        self.frame_p1.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.txt_preview_1 = tk.Text(self.frame_p1, wrap="none", font=("Consolas", 10), bg="#f8fafc", fg="#0f172a")
        scroll_y1 = ttk.Scrollbar(self.frame_p1, orient="vertical", command=self.txt_preview_1.yview)
        scroll_x1 = ttk.Scrollbar(self.frame_p1, orient="horizontal", command=self.txt_preview_1.xview)
        self.txt_preview_1.configure(yscrollcommand=scroll_y1.set, xscrollcommand=scroll_x1.set)
        scroll_x1.pack(side="bottom", fill="x")
        self.txt_preview_1.pack(side="left", fill="both", expand=True)
        scroll_y1.pack(side="right", fill="y")

        self.frame_p2 = ttk.LabelFrame(self.frame_kanan_container, text="  Hasil Kata Kunci 2  ", padding=5)
        self.txt_preview_2 = tk.Text(self.frame_p2, wrap="none", font=("Consolas", 10), bg="#f8fafc", fg="#1e293b")
        scroll_y2 = ttk.Scrollbar(self.frame_p2, orient="vertical", command=self.txt_preview_2.yview)
        scroll_x2 = ttk.Scrollbar(self.frame_p2, orient="horizontal", command=self.txt_preview_2.xview)
        self.txt_preview_2.configure(yscrollcommand=scroll_y2.set, xscrollcommand=scroll_x2.set)
        scroll_x2.pack(side="bottom", fill="x")
        self.txt_preview_2.pack(side="left", fill="both", expand=True)
        scroll_y2.pack(side="right", fill="y")

    # ==========================================
    # LOGIKA PROSES TAB 1
    # ==========================================
    def mulai_pencarian(self):
        kw1 = self.ent_cari.get().strip()
        kw2 = self.ent_cari_2.get().strip() if self.mode_perbandingan else ""

        if len(kw1) < 3:
            messagebox.showwarning("Peringatan", "Kata kunci 1 terlalu pendek! Masukkan minimal 3 karakter.")
            return
        if self.mode_perbandingan and len(kw2) < 3:
            messagebox.showwarning("Peringatan", "Mode Perbandingan Aktif! Masukkan Kata Kunci 2 minimal 3 karakter.")
            return

        self.btn_cari.config(state="disabled")
        self.btn_mode_banding.config(state="disabled")
        self.lbl_status.config(text="Sedang menyisir data JSON...", foreground="blue")
        
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.txt_preview_1.delete(1.0, tk.END)
        self.txt_preview_2.delete(1.0, tk.END)
        self.hasil_pencarian.clear()

        thread = threading.Thread(target=self.proses_pencarian, args=(kw1, kw2))
        thread.daemon = True
        thread.start()

    def proses_pencarian(self, kw1, kw2):
        if not os.path.exists(DATA_DIR):
            self.update_ui_selesai("Folder 'data' tidak ditemukan!")
            return

        total_file_diperiksa = 0
        total_ditemukan = 0

        for root, dirs, files in os.walk(DATA_DIR):
            nama_folder = os.path.basename(root)
            tahun = nama_folder if nama_folder.isdigit() else "-"

            for file in files:
                if file.endswith(".json"):
                    total_file_diperiksa += 1
                    filepath = os.path.join(root, file)
                    m1, m2 = self.ekstrak_data_json(filepath, kw1, kw2)
                    
                    if m1 or m2:
                        total_ditemukan += 1
                        key_unik = f"{tahun}|{file}"
                        self.hasil_pencarian[key_unik] = {"m1": m1, "m2": m2}
                        info_jumlah = f"KW1: {len(m1)}"
                        if self.mode_perbandingan:
                            info_jumlah += f" | KW2: {len(m2)}"
                        self.after(0, self.tambah_baris_tabel, tahun, file, info_jumlah, key_unik)

        pesan_akhir = f"Selesai! Diperiksa: {total_file_diperiksa} file | Ditemukan: {total_ditemukan} file."
        self.after(0, self.update_ui_selesai, pesan_akhir)

    def ekstrak_data_json(self, filepath, kw1, kw2=""):
        m1, m2 = [], []
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                content = f.read()
            has_kw1 = kw1 in content
            has_kw2 = kw2 in content if kw2 else False
            if not has_kw1 and not has_kw2:
                return m1, m2

            data = json.loads(content)
            items = []
            if isinstance(data, list):
                items = data
            elif isinstance(data, dict):
                if 'data' in data and isinstance(data['data'], list):
                    items = data['data']
                else:
                    items = [data]

            for item in items:
                item_str = json.dumps(item)
                if has_kw1 and kw1 in item_str:
                    m1.append(item)
                if has_kw2 and kw2 in item_str:
                    m2.append(item)
        except:
            pass
        return m1, m2

    def tambah_baris_tabel(self, tahun, file, jumlah, key_unik):
        self.tree.insert("", "end", values=(tahun, file, jumlah), tags=(key_unik,))

    def update_ui_selesai(self, pesan):
        self.lbl_status.config(text=pesan, foreground="green")
        self.btn_cari.config(state="normal")
        self.btn_mode_banding.config(state="normal")
        if not self.hasil_pencarian:
            self.txt_preview_1.insert(tk.END, ">>> KODE TIDAK DITEMUKAN DI FILE MANAPUN <<<")
            if self.mode_perbandingan:
                self.txt_preview_2.insert(tk.END, ">>> KODE TIDAK DITEMUKAN DI FILE MANAPUN <<<")

    def tampilkan_isi_data(self, event):
        selected_item = self.tree.selection()
        if not selected_item: return

        key_unik = self.tree.item(selected_item[0], "tags")[0]
        data_packet = self.hasil_pencarian.get(key_unik, {"m1": [], "m2": []})

        self.txt_preview_1.delete(1.0, tk.END)
        self.txt_preview_2.delete(1.0, tk.END)
        
        if data_packet["m1"]:
            for i, data in enumerate(data_packet["m1"], 1):
                self.txt_preview_1.insert(tk.END, f"=== MATCH KW1 KE-{i} ===\n")
                self.txt_preview_1.insert(tk.END, json.dumps(data, indent=4, ensure_ascii=False, sort_keys=True) + "\n\n")
        else:
            self.txt_preview_1.insert(tk.END, "--- Tidak ada match untuk Kata Kunci 1 di file ini ---")

        if self.mode_perbandingan:
            if data_packet["m2"]:
                for i, data in enumerate(data_packet["m2"], 1):
                    self.txt_preview_2.insert(tk.END, f"=== MATCH KW2 KE-{i} ===\n")
                    self.txt_preview_2.insert(tk.END, json.dumps(data, indent=4, ensure_ascii=False, sort_keys=True) + "\n\n")
            else:
                self.txt_preview_2.insert(tk.END, "--- Tidak ada match untuk Kata Kunci 2 di file ini ---")


    # ==========================================
    # KONTEN TAB 2: VIEWER DATA JSON & FILTER
    # ==========================================
    def create_layout_tab2(self):
        # Frame Top Control Tab 2
        frame_top_t2 = ttk.LabelFrame(self.tab2, text="  Kontrol Penjelajah Folder  ", padding=(15, 15))
        frame_top_t2.pack(side="top", fill="x", pady=(0, 15))
        
        btn_scan_all = ttk.Button(frame_top_t2, text="🔍 Scan JSON Folder", command=self.scan_seluruh_folder_json)
        btn_scan_all.grid(row=0, column=0, padx=(0, 20), sticky="w")
        
        ttk.Label(frame_top_t2, text="Cari di Teks Terbuka:", font=('Arial', 10, 'bold')).grid(row=0, column=1, sticky="w", padx=(0, 5))
        self.ent_filter_internal = ttk.Entry(frame_top_t2, width=35, font=('Arial', 11))
        self.ent_filter_internal.grid(row=0, column=2, padx=(0, 10), sticky="w")
        self.ent_filter_internal.bind("<Return>", lambda e: self.eksekusi_cari_internal_text())
        
        btn_cari_internal = ttk.Button(frame_top_t2, text="⚡ Cari Kata", command=self.eksekusi_cari_internal_text)
        btn_cari_internal.grid(row=0, column=3, sticky="w")
        
        # Main Split Frame Tab 2
        paned_t2 = ttk.PanedWindow(self.tab2, orient=tk.HORIZONTAL)
        paned_t2.pack(fill="both", expand=True)
        
        # Sisi Kiri: Daftar File
        frame_kiri_t2 = ttk.LabelFrame(paned_t2, text="  Daftar Semua File JSON:  ", padding=5)
        paned_t2.add(frame_kiri_t2, weight=1)
        
        self.tree_viewer = ttk.Treeview(frame_kiri_t2, columns=("tahun", "file"), show="headings", selectmode="browse")
        self.tree_viewer.heading("tahun", text="Tahun")
        self.tree_viewer.heading("file", text="Nama File")
        self.tree_viewer.column("tahun", width=60, anchor="center")
        self.tree_viewer.column("file", width=250, anchor="w")
        
        scroll_tree_t2 = ttk.Scrollbar(frame_kiri_t2, orient="vertical", command=self.tree_viewer.yview)
        self.tree_viewer.configure(yscroll=scroll_tree_t2.set)
        self.tree_viewer.pack(side="left", fill="both", expand=True)
        scroll_tree_t2.pack(side="right", fill="y")
        self.tree_viewer.bind("<<TreeviewSelect>>", self.baca_dan_tampilkan_file_viewer)
        
        # Sisi Kanan: Pratinjau Teks
        frame_kanan_t2 = ttk.LabelFrame(paned_t2, text="  Isi Konten File (Urut A-Z):  ", padding=5)
        paned_t2.add(frame_kanan_t2, weight=3)
        
        self.txt_viewer = tk.Text(frame_kanan_t2, wrap="none", font=("Consolas", 10), bg="#ffffff", fg="#020617")
        scroll_y_t2 = ttk.Scrollbar(frame_kanan_t2, orient="vertical", command=self.txt_viewer.yview)
        scroll_x_t2 = ttk.Scrollbar(frame_kanan_t2, orient="horizontal", command=self.txt_viewer.xview)
        self.txt_viewer.configure(yscrollcommand=scroll_y_t2.set, xscrollcommand=scroll_x_t2.set)
        
        scroll_x_t2.pack(side="bottom", fill="x")
        self.txt_viewer.pack(side="left", fill="both", expand=True)
        scroll_y_t2.pack(side="right", fill="y")
        
        # Daftarkan konfigurasi style highlight pencarian internal
        self.txt_viewer.tag_config('highlight_search', background='#eab308', foreground='#000000')

    # ==========================================
    # LOGIKA PROSES TAB 2
    # ==========================================
    def scan_seluruh_folder_json(self):
        for i in self.tree_viewer.get_children():
            self.tree_viewer.delete(i)
        self.dict_file_viewer.clear()
        self.txt_viewer.delete(1.0, tk.END)
        
        if not os.path.exists(DATA_DIR):
            messagebox.showerror("Error", f"Folder '{DATA_DIR}' tidak ditemukan!")
            return
            
        # --- Kumpulkan data dulu ke dalam list ---
        temp_list = []
        for root, dirs, files in os.walk(DATA_DIR):
            nama_folder = os.path.basename(root)
            tahun = nama_folder if nama_folder.isdigit() else "-"
            for file in files:
                if file.endswith(".json"):
                    full_path = os.path.join(root, file)
                    temp_list.append((tahun, file, full_path))
        
        # --- Urutkan: Tahun (Terbaru/Descending), lalu File (A-Z/Ascending) ---
        temp_list.sort(key=lambda x: (-int(x[0]) if x[0].isdigit() else 0, x[1].lower()))
        
        # --- Masukkan ke dalam tabel ---
        counter = 0
        for tahun, file, full_path in temp_list:
            counter += 1
            key_id = f"v_file_{counter}"
            self.dict_file_viewer[key_id] = full_path
            self.tree_viewer.insert("", "end", values=(tahun, file), tags=(key_id,))
            
        messagebox.showinfo("Scan Selesai", f"Berhasil memuat {counter} file JSON ke dalam daftar list.")

    def baca_dan_tampilkan_file_viewer(self, event):
        selected = self.tree_viewer.selection()
        if not selected: return
        
        key_id = self.tree_viewer.item(selected[0], "tags")[0]
        full_path = self.dict_file_viewer.get(key_id)
        
        self.txt_viewer.delete(1.0, tk.END)
        self.ent_filter_internal.delete(0, tk.END)
        
        if not full_path or not os.path.exists(full_path):
            self.txt_viewer.insert(tk.END, "❌ Gagal memuat file, path tidak valid.")
            return
            
        try:
            with open(full_path, 'r', encoding='utf-8-sig') as f:
                raw_data = json.load(f)
                
            # Formatting rapi urut alfabetis A-Z sesuai request
            pretty_json = json.dumps(raw_data, indent=4, ensure_ascii=False, sort_keys=True)
            self.txt_viewer.insert(tk.END, pretty_json)
        except Exception as err:
            self.txt_viewer.insert(tk.END, f"❌ Terjadi kesalahan membaca JSON:\n{str(err)}")

    def eksekusi_cari_internal_text(self):
        # Bersihkan highlight pencarian lama terlebih dahulu
        self.txt_viewer.tag_remove('highlight_search', '1.0', tk.END)
        
        keyword = self.ent_filter_internal.get().strip()
        if not keyword: 
            return
            
        posisi_index = '1.0'
        jumlah_cocok = 0
        
        while True:
            # Gunakan fungsi bawaan widget text untuk pencarian koordinat string karakter
            posisi_index = self.txt_viewer.search(keyword, posisi_index, nocase=True, stopindex=tk.END)
            if not posisi_index: 
                break
                
            jumlah_cocok += 1
            akhir_posisi = f"{posisi_index}+{len(keyword)}c"
            self.txt_viewer.tag_add('highlight_search', posisi_index, akhir_posisi)
            
            if jumlah_cocok == 1:
                # Otomatis geser scroll viewport ke temuan kata pertama
                self.txt_viewer.see(posisi_index)
                
            posisi_index = akhir_posisi
            
        if jumlah_cocok == 0:
            messagebox.showinfo("Pencarian", f"Kata '{keyword}' tidak ditemukan pada teks JSON ini.")


if __name__ == "__main__":
    app = AppPelacakJSON()
    app.mainloop()