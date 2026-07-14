import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import re
import traceback

class RekonsiliasiApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplikasi Rekonsiliasi Data Pengadaan (Auto-Center)")
        
        window_width = 900
        window_height = 700
        self.center_window(self.root, window_width, window_height)
        self.root.configure(bg="#f5f5f5")
        
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TButton", font=("Arial", 10), padding=5)
        self.style.configure("Action.TButton", font=("Arial", 10, "bold"), background="#0288d1", foreground="white")
        self.style.configure("Scan.TButton", font=("Arial", 9), background="#e0e0e0")
        
        # Variabel Tab 1 (Realisasi)
        self.file_master_path = ""
        self.file_realisasi_path = ""
        self.master_sheet = None
        self.realisasi_sheet = None
        self.master_status_col = None
        self.realisasi_status_col = None
        self.master_status_vars = {}
        self.realisasi_status_vars = {}
        self.var_all_master = tk.BooleanVar(value=False)
        self.var_all_realisasi = tk.BooleanVar(value=False)

        # Variabel Tab 2 (RUP)
        self.file_master_rup_path = ""
        self.file_csv_rup_path = ""
        self.master_rup_sheet = None
        self.csv_rup_sheet = None
        
        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === NOTEBOOK (TABS) ===
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.tab_realisasi = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.tab_realisasi, text=" 1. Rekonsiliasi REALISASI ")
        
        self.tab_rup = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.tab_rup, text=" 2. Rekonsiliasi RUP ")
        
        # --- TAMBAHAN TAB 3 ---
        self.tab_master = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.tab_master, text=" 3. Master Prod vs Test ")
        
        self.file_master_prod_path = ""
        self.file_master_test_path = ""
        self.master_prod_sheet = None
        self.master_test_sheet = None
        # ----------------------
        
        self.setup_tab_realisasi()
        self.setup_tab_rup()
        self.setup_tab_master() # Panggil setup tab 3
        
        # === SHARED LOG AREA ===
        ttk.Label(main_frame, text="Status & Log Proses:", font=("Arial", 9, "bold")).pack(anchor=tk.W, pady=(0,2))
        self.txt_log = tk.Text(main_frame, height=8, font=("Consolas", 10), bg="#1e1e1e", fg="#d4d4d4", wrap=tk.WORD)
        self.txt_log.pack(fill=tk.BOTH, expand=False)
        self.log("Aplikasi siap. Silakan pilih tab menu di atas sesuai kebutuhan Anda.")

    def center_window(self, window, width, height):
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")

    def log(self, pesan):
        self.txt_log.insert(tk.END, pesan + "\n")
        self.txt_log.see(tk.END)

    # ==========================================
    # HELPER UMUM (FILE & KOLOM)
    # ==========================================
    def cek_dan_pilih_sheet(self, file_path):
        if file_path.endswith(('.xlsx', '.xls')):
            try:
                xl = pd.ExcelFile(file_path)
                sheets = xl.sheet_names
                if len(sheets) > 1:
                    return self.buka_dialog_pilih_sheet(sheets, os.path.basename(file_path))
                elif len(sheets) == 1:
                    return sheets[0]
            except Exception as e:
                self.log(f"[ERROR] Gagal mendeteksi sheet: {e}")
        return None

    def buka_dialog_pilih_sheet(self, sheets, nama_file):
        popup = tk.Toplevel(self.root)
        popup.title("Pilih Sheet")
        popup.resizable(False, False)
        popup.transient(self.root)
        popup.grab_set()
        
        popup_width = 380
        popup_height = 160
        self.center_window(popup, popup_width, popup_height)
        
        ttk.Label(popup, text=f"File '{nama_file}' memiliki lebih dari 1 sheet.\nSilakan pilih sheet yang akan digunakan:", font=("Arial", 10), justify=tk.LEFT).pack(pady=(15,5), padx=20, anchor=tk.W)
        
        pilihan = tk.StringVar()
        combobox = ttk.Combobox(popup, textvariable=pilihan, values=sheets, state="readonly", width=35)
        combobox.pack(pady=10, padx=20)
        combobox.current(0)
        
        def oke(): popup.destroy()
        ttk.Button(popup, text="Gunakan Sheet Ini", command=oke).pack(pady=10)
        self.root.wait_window(popup)
        return pilihan.get()

    def load_file(self, path, sheet_name=None):
        if path.endswith('.csv'):
            try: return pd.read_csv(path, sep=None, engine='python', encoding='utf-8-sig')
            except: return pd.read_csv(path, sep=None, engine='python', encoding='latin1')
        return pd.read_excel(path, sheet_name=sheet_name)

    def ekstrak_dan_bersihkan_kode(self, val):
        """Memecah titik koma dan membersihkan format menjadi murni deretan angka"""
        if pd.isna(val): return []
        raw_str = str(val).strip()
        if not raw_str or raw_str.lower() in ['nan', 'none', '-']: return []
        
        hasil = []
        parts = raw_str.split(';')
        for p in parts:
            p_clean = str(p).strip()
            if p_clean.endswith('.0'): p_clean = p_clean[:-2]
            # Hapus semua karakter selain angka (sangat aman)
            p_angka_saja = re.sub(r'\D', '', p_clean)
            if p_angka_saja:
                hasil.append(p_angka_saja)
        return list(set(hasil)) # Hapus duplikat internal

    # ==========================================
    # SETUP & FUNGSI TAB 1: REALISASI
    # ==========================================
    def setup_tab_realisasi(self):
        frame_input = ttk.Frame(self.tab_realisasi)
        frame_input.pack(fill=tk.X, pady=(0, 10))
        
        frame_left_input = ttk.Frame(frame_input)
        frame_left_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 15))
        ttk.Label(frame_left_input, text="1. File Master Pengadaan:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0,4))
        frame_master = ttk.Frame(frame_left_input)
        frame_master.pack(fill=tk.X)
        ttk.Button(frame_master, text="Pilih Master", command=self.pilih_master_realisasi).pack(side=tk.LEFT, padx=(0,8))
        self.lbl_master = ttk.Label(frame_master, text="Belum ada file...", font=("Arial", 9, "italic"), foreground="gray")
        self.lbl_master.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        frame_right_input = ttk.Frame(frame_input)
        frame_right_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(15, 0))
        ttk.Label(frame_right_input, text="2. File Realisasi:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0,4))
        frame_realisasi = ttk.Frame(frame_right_input)
        frame_realisasi.pack(fill=tk.X)
        ttk.Button(frame_realisasi, text="Pilih Realisasi", command=self.pilih_realisasi).pack(side=tk.LEFT, padx=(0,8))
        self.lbl_realisasi = ttk.Label(frame_realisasi, text="Belum ada file...", font=("Arial", 9, "italic"), foreground="gray")
        self.lbl_realisasi.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(self.tab_realisasi, text="3. Filter Status (Beri Ceklis Pada Status yang Diaktifkan):", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(10,5))
        
        frame_status = ttk.Frame(self.tab_realisasi)
        frame_status.pack(fill=tk.BOTH, expand=True, pady=(0,15))
        
        # MASTER STATUS UI
        frame_left = ttk.LabelFrame(frame_status, text=" Status MASTER ", padding=10)
        frame_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5))
        ttk.Button(frame_left, text="🔍 Scan Status", style="Scan.TButton", command=self.scan_status_master).pack(anchor=tk.W, pady=(0,5))
        self.cb_all_m = ttk.Checkbutton(frame_left, text="Pilih Semua", variable=self.var_all_master, command=lambda: self.toggle_all(self.var_all_master, self.master_status_vars))
        self.cb_all_m.pack(anchor=tk.W, pady=(2,5))
        
        self.canvas_m, self.scroll_frame_m = self.create_scrollable_frame(frame_left)
        
        # REALISASI STATUS UI
        frame_right = ttk.LabelFrame(frame_status, text=" Status REALISASI ", padding=10)
        frame_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5,0))
        ttk.Button(frame_right, text="🔍 Scan Status", style="Scan.TButton", command=self.scan_status_realisasi).pack(anchor=tk.W, pady=(0,5))
        self.cb_all_r = ttk.Checkbutton(frame_right, text="Pilih Semua", variable=self.var_all_realisasi, command=lambda: self.toggle_all(self.var_all_realisasi, self.realisasi_status_vars))
        self.cb_all_r.pack(anchor=tk.W, pady=(2,5))
        
        self.canvas_r, self.scroll_frame_r = self.create_scrollable_frame(frame_right)
        
        ttk.Button(self.tab_realisasi, text="🚀 PROSES DATA REALISASI (3 SHEET)", style="Action.TButton", command=self.proses_data_realisasi).pack(fill=tk.X, pady=(5,0))

    def create_scrollable_frame(self, parent):
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(container, highlightthickness=1, highlightbackground="#cccccc", bg="white")
        sb = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg="white")
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        return canvas, scroll_frame

    def pilih_master_realisasi(self):
        p = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if p:
            self.file_master_path = p
            self.master_sheet = self.cek_dan_pilih_sheet(p)
            info = f" ({self.master_sheet})" if self.master_sheet else ""
            self.lbl_master.config(text=f"{os.path.basename(p)}{info}", font=("Arial", 9, "bold"), foreground="#2e7d32")
            self.log(f"[REALISASI] Master: {os.path.basename(p)}{info}")

    def pilih_realisasi(self):
        p = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if p:
            self.file_realisasi_path = p
            self.realisasi_sheet = self.cek_dan_pilih_sheet(p)
            info = f" ({self.realisasi_sheet})" if self.realisasi_sheet else ""
            self.lbl_realisasi.config(text=f"{os.path.basename(p)}{info}", font=("Arial", 9, "bold"), foreground="#2e7d32")
            self.log(f"[REALISASI] Data: {os.path.basename(p)}{info}")

    def toggle_all(self, var_all, var_dict):
        for var in var_dict.values(): var.set(var_all.get())

    def find_status_column(self, df):
        if "Status" in df.columns: return "Status"
        for col in df.columns:
            if "status" in col.lower(): return col
        return None

    def scan_status(self, df, scroll_frame, var_dict, var_all):
        col_status = self.find_status_column(df)
        if not col_status: return None
        status_list = sorted([str(x).strip() for x in df[col_status].dropna().unique()])
        var_all.set(False)
        for widget in scroll_frame.winfo_children(): widget.destroy()
        var_dict.clear()
        for item in status_list:
            var = tk.BooleanVar(value=False)
            tk.Checkbutton(scroll_frame, text=item, variable=var, bg="white", anchor="w", font=("Arial", 9)).pack(fill=tk.X, padx=5, pady=2)
            var_dict[item] = var
        return col_status, len(status_list)

    def scan_status_master(self):
        if not self.file_master_path: return messagebox.showerror("Error", "Pilih File Master dulu!")
        try:
            df = self.load_file(self.file_master_path, self.master_sheet)
            res = self.scan_status(df, self.scroll_frame_m, self.master_status_vars, self.var_all_master)
            if res:
                self.master_status_col = res[0]
                self.log(f"-> Master: Memuat {res[1]} status.")
            else: messagebox.showerror("Error", "Kolom Status tidak ditemukan!")
        except Exception as e: messagebox.showerror("Error", str(e))

    def scan_status_realisasi(self):
        if not self.file_realisasi_path: return messagebox.showerror("Error", "Pilih File Realisasi dulu!")
        try:
            df = self.load_file(self.file_realisasi_path, self.realisasi_sheet)
            res = self.scan_status(df, self.scroll_frame_r, self.realisasi_status_vars, self.var_all_realisasi)
            if res:
                self.realisasi_status_col = res[0]
                self.log(f"-> Realisasi: Memuat {res[1]} status.")
            else: messagebox.showerror("Error", "Kolom Status tidak ditemukan!")
        except Exception as e: messagebox.showerror("Error", str(e))

    def cari_kolom_paket(self, df_m, df_r):
        variasi = ["Kode Paket", "Kode_Paket", "kode_paket", "No Paket", "No_Paket", "ID Paket", "Kd_Paket", "kd_paket"]
        m_col = next((c for c in df_m.columns if c.lower().strip() in [v.lower() for v in variasi]), None)
        r_col = next((c for c in df_r.columns if c.lower().strip() in [v.lower() for v in variasi]), None)
        return m_col, r_col

    def proses_data_realisasi(self):
        if not self.file_master_path or not self.file_realisasi_path:
            return messagebox.showerror("Error", "File Master dan Realisasi harus diisi!")
            
        m_sel = [s for s, v in self.master_status_vars.items() if v.get()]
        r_sel = [s for s, v in self.realisasi_status_vars.items() if v.get()]
        
        if not m_sel and self.master_status_vars: return messagebox.showwarning("Peringatan", "Centang minimal 1 status Master!")
        if not r_sel and self.realisasi_status_vars: return messagebox.showwarning("Peringatan", "Centang minimal 1 status Realisasi!")
            
        self.log("\n--- Mulai Rekonsiliasi REALISASI ---")
        try:
            df_m = self.load_file(self.file_master_path, self.master_sheet)
            df_r = self.load_file(self.file_realisasi_path, self.realisasi_sheet)
            
            m_key, r_key = self.cari_kolom_paket(df_m, df_r)
            if not m_key or not r_key:
                return messagebox.showerror("Error", "Kolom 'Kode Paket' tidak ditemukan di salah satu file!")
            self.log(f"✓ Kolom Master: '{m_key}', Kolom Realisasi: '{r_key}'")
            
            if self.master_status_col and m_sel: df_m = df_m[df_m[self.master_status_col].astype(str).str.strip().isin(m_sel)].copy()
            if self.realisasi_status_col and r_sel: df_r = df_r[df_r[self.realisasi_status_col].astype(str).str.strip().isin(r_sel)].copy()
                
            self.log("✓ Menormalisasi kode paket...")
            # Menggunakan fungsi ekstrak yang aman
            df_m['Kode_Normal'] = df_m[m_key].apply(lambda x: self.ekstrak_dan_bersihkan_kode(x)[0] if self.ekstrak_dan_bersihkan_kode(x) else None)
            df_r['Kode_Normal'] = df_r[r_key].apply(lambda x: self.ekstrak_dan_bersihkan_kode(x)[0] if self.ekstrak_dan_bersihkan_kode(x) else None)
            
            set_m = set(df_m['Kode_Normal'].dropna().unique())
            set_r = set(df_r['Kode_Normal'].dropna().unique())
            
            df_hanya_r = df_r[~df_r['Kode_Normal'].isin(set_m)].copy().drop(columns=['Kode_Normal'])
            df_hanya_m = df_m[~df_m['Kode_Normal'].isin(set_r)].copy().drop(columns=['Kode_Normal'])
            df_kedua = df_m[df_m['Kode_Normal'].isin(set_r)].copy().drop(columns=['Kode_Normal'])
            
            self.log(f"-> Hanya di Realisasi: {len(df_hanya_r)} | Hanya di Master: {len(df_hanya_m)} | Match: {len(df_kedua)}")
            
            pesan = f"Data Realisasi berhasil diproses!\n\nRingkasan:\n- Hanya di Realisasi: {len(df_hanya_r)}\n- Hanya di Master: {len(df_hanya_m)}\n- Cocok di Keduanya: {len(df_kedua)}\n\nApakah Anda ingin menyimpan hasilnya ke Excel?"
            if messagebox.askyesno("Konfirmasi Simpan", pesan):
                save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="Rekonsiliasi_Realisasi.xlsx")
                if save_path:
                    with pd.ExcelWriter(save_path, engine='openpyxl') as w:
                        df_hanya_r.to_excel(w, sheet_name='Hanya di Realisasi', index=False)
                        df_hanya_m.to_excel(w, sheet_name='Hanya di Master', index=False)
                        df_kedua.to_excel(w, sheet_name='Ada di Kedua Sheet', index=False)
                    self.log(f"[SUKSES] Disimpan di: {save_path}")
                    messagebox.showinfo("Sukses", "File Excel berhasil disimpan!")
            else:
                self.log("[INFO] Proses selesai. File tidak disimpan.")
                
        except Exception as e:
            self.log(f"[ERROR] {traceback.format_exc()}")
            messagebox.showerror("Error", f"Terjadi kesalahan:\n{e}")

    # ==========================================
    # SETUP & FUNGSI TAB 2: RUP
    # ==========================================
    def setup_tab_rup(self):
        ttk.Label(self.tab_rup, text="Rekonsiliasi berdasarkan KODE RUP (Tanpa Filter Status)", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0,15))
        
        # Frame Input Master RUP
        f_m = ttk.Frame(self.tab_rup)
        f_m.pack(fill=tk.X, pady=5)
        ttk.Label(f_m, text="1. File Master Pengadaan (Sheet 1):", font=("Arial", 10, "bold"), width=30).pack(side=tk.LEFT)
        ttk.Button(f_m, text="Pilih Master", command=self.pilih_master_rup).pack(side=tk.LEFT, padx=10)
        self.lbl_master_rup = ttk.Label(f_m, text="Belum ada file...", font=("Arial", 9, "italic"), foreground="gray")
        self.lbl_master_rup.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Frame Input CSV RUP
        f_c = ttk.Frame(self.tab_rup)
        f_c.pack(fill=tk.X, pady=15)
        ttk.Label(f_c, text="2. File Data RUP (CSV/Excel):", font=("Arial", 10, "bold"), width=30).pack(side=tk.LEFT)
        ttk.Button(f_c, text="Pilih Data RUP", command=self.pilih_csv_rup).pack(side=tk.LEFT, padx=10)
        self.lbl_csv_rup = ttk.Label(f_c, text="Belum ada file...", font=("Arial", 9, "italic"), foreground="gray")
        self.lbl_csv_rup.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(self.tab_rup, text="Catatan Sistem:\n- Sistem memprioritaskan kolom 'Kode RUP Baru' jika tersedia.\n- Sistem otomatis memecah Kode RUP yang digabung dengan titik koma (;).", justify=tk.LEFT).pack(anchor=tk.W, pady=(10,20))

        ttk.Button(self.tab_rup, text="🚀 PROSES DATA RUP (3 SHEET)", style="Action.TButton", command=self.proses_data_rup).pack(fill=tk.X, pady=5)

    def pilih_master_rup(self):
        p = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if p:
            self.file_master_rup_path = p
            self.master_rup_sheet = self.cek_dan_pilih_sheet(p)
            info = f" ({self.master_rup_sheet})" if self.master_rup_sheet else ""
            self.lbl_master_rup.config(text=f"{os.path.basename(p)}{info}", font=("Arial", 9, "bold"), foreground="#2e7d32")
            self.log(f"[RUP] Master: {os.path.basename(p)}{info}")

    def pilih_csv_rup(self):
        p = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if p:
            self.file_csv_rup_path = p
            self.csv_rup_sheet = self.cek_dan_pilih_sheet(p)
            info = f" ({self.csv_rup_sheet})" if self.csv_rup_sheet else ""
            self.lbl_csv_rup.config(text=f"{os.path.basename(p)}{info}", font=("Arial", 9, "bold"), foreground="#2e7d32")
            self.log(f"[RUP] Data CSV/Excel RUP: {os.path.basename(p)}{info}")

    def cari_kolom_rup(self, df):
        variasi = ["Kode RUP", "kode_rup", "Kode_RUP", "ID RUP"]
        return next((c for c in df.columns if c.lower().strip() in [v.lower() for v in variasi]), None)

    def cari_kolom_rup_baru(self, df):
        variasi = ["Kode RUP Baru", "kode_rup_baru", "Kode_RUP_Baru"]
        return next((c for c in df.columns if c.lower().strip() in [v.lower() for v in variasi]), None)

    def proses_data_rup(self):
        if not self.file_master_rup_path or not self.file_csv_rup_path:
            return messagebox.showerror("Error", "File Master dan Data RUP harus diisi!")
            
        self.log("\n--- Mulai Rekonsiliasi RUP (Titik Koma Dipecah) ---")
        try:
            df_m = self.load_file(self.file_master_rup_path, self.master_rup_sheet)
            df_r = self.load_file(self.file_csv_rup_path, self.csv_rup_sheet)
            
            # Cari Kolom
            m_rup_col = self.cari_kolom_rup(df_m)
            m_rup_baru_col = self.cari_kolom_rup_baru(df_m)
            r_rup_col = self.cari_kolom_rup(df_r)
            
            if not m_rup_col: return messagebox.showerror("Error", "Kolom 'Kode RUP' tidak ditemukan di File Master!")
            if not r_rup_col: return messagebox.showerror("Error", "Kolom 'Kode RUP' tidak ditemukan di File RUP!")
            
            self.log(f"✓ Kolom Master: '{m_rup_col}', Kolom RUP Baru: '{m_rup_baru_col or 'TIDAK ADA'}'")
            self.log(f"✓ Kolom CSV RUP: '{r_rup_col}'")
            
            # --- FUNGSI KHUSUS LOGIKA MASTER RUP ---
            def get_final_rup_list(row):
                """Mengekstrak RUP Baru jika ada, jika tidak pakai RUP Lama, lalu pisah titik koma"""
                k_baru = str(row[m_rup_baru_col]).strip() if m_rup_baru_col and pd.notna(row[m_rup_baru_col]) else ""
                k_lama = str(row[m_rup_col]).strip() if pd.notna(row[m_rup_col]) else ""
                
                # Prioritas: Jika k_baru ada isinya (dan bukan sekadar strip/kosong), gunakan k_baru
                if k_baru and k_baru.lower() not in ["", "-", "nan", "none"]:
                    target_str = k_baru
                else:
                    target_str = k_lama
                    
                return self.ekstrak_dan_bersihkan_kode(target_str)

            self.log("✓ Menganalisis prioritas & memecah titik koma di Data Master...")
            df_m['List_RUP_Final'] = df_m.apply(get_final_rup_list, axis=1)
            
            # EXPLODE: Gandakan baris jika list memiliki lebih dari 1 RUP
            df_m_explode = df_m.explode('List_RUP_Final')
            df_m_explode = df_m_explode[df_m_explode['List_RUP_Final'].notna()] # Buang yang kosong
            
            self.log("✓ Memecah titik koma di Data CSV RUP...")
            df_r['List_RUP_Final'] = df_r[r_rup_col].apply(self.ekstrak_dan_bersihkan_kode)
            df_r_explode = df_r.explode('List_RUP_Final')
            df_r_explode = df_r_explode[df_r_explode['List_RUP_Final'].notna()]
            
            # MENGAMBIL SET KODE
            set_m = set(df_m_explode['List_RUP_Final'].unique())
            set_r = set(df_r_explode['List_RUP_Final'].unique())
            
            self.log("Membandingkan data...")
            df_hanya_r = df_r_explode[~df_r_explode['List_RUP_Final'].isin(set_m)].copy().drop(columns=['List_RUP_Final'])
            df_hanya_m = df_m_explode[~df_m_explode['List_RUP_Final'].isin(set_r)].copy().drop(columns=['List_RUP_Final'])
            df_kedua = df_m_explode[df_m_explode['List_RUP_Final'].isin(set_r)].copy().drop(columns=['List_RUP_Final'])
            
            # Hapus duplikat pasca-explode agar tampilan Excel tetap rapi jika ada dobel tak sengaja
            df_hanya_r = df_hanya_r.drop_duplicates()
            df_hanya_m = df_hanya_m.drop_duplicates()
            df_kedua = df_kedua.drop_duplicates()

            self.log(f"-> RUP Belum Ditarik (Hanya di CSV): {len(df_hanya_r)} | RUP Tak Dikenali (Hanya di Master): {len(df_hanya_m)} | Match: {len(df_kedua)}")
            
            jumlah_hanya_r = len(set_r - set_m)
            jumlah_hanya_m = len(set_m - set_r)
            jumlah_cocok = len(set_m & set_r)

            pesan = f"Data RUP berhasil diproses!\n\nRingkasan (Total Kode RUP Unik):\n- Hanya di CSV RUP: {jumlah_hanya_r}\n- Hanya di Master: {jumlah_hanya_m}\n- Cocok di Keduanya: {jumlah_cocok}\n\nApakah Anda ingin menyimpan hasilnya ke Excel?"
            if messagebox.askyesno("Konfirmasi Simpan", pesan):
                save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="Rekonsiliasi_RUP.xlsx")
                if save_path:
                    with pd.ExcelWriter(save_path, engine='openpyxl') as w:
                        df_hanya_r.to_excel(w, sheet_name='Hanya di CSV RUP', index=False)
                        df_hanya_m.to_excel(w, sheet_name='Hanya di Master', index=False)
                        df_kedua.to_excel(w, sheet_name='Cocok di Keduanya', index=False)
                    self.log(f"[SUKSES] Disimpan di: {save_path}")
                    messagebox.showinfo("Sukses", "File Excel berhasil disimpan!")
            else:
                self.log("[INFO] Proses selesai. File tidak disimpan.")
                
        except Exception as e:
            self.log(f"[ERROR] {traceback.format_exc()}")
            messagebox.showerror("Error", f"Terjadi kesalahan:\n{e}")

# ==========================================
    # SETUP & FUNGSI TAB 3: MASTER VS MASTER
    # ==========================================
    def setup_tab_master(self):
        ttk.Label(self.tab_master, text="Rekonsiliasi Master Produksi vs Master Testing", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=(0,15))
        
        # Frame Input Prod
        f_p = ttk.Frame(self.tab_master)
        f_p.pack(fill=tk.X, pady=5)
        ttk.Label(f_p, text="1. File Master Produksi:", font=("Arial", 10, "bold"), width=25).pack(side=tk.LEFT)
        ttk.Button(f_p, text="Pilih Produksi", command=self.pilih_master_prod).pack(side=tk.LEFT, padx=10)
        self.lbl_master_prod = ttk.Label(f_p, text="Belum ada file...", font=("Arial", 9, "italic"), foreground="gray")
        self.lbl_master_prod.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Frame Input Test
        f_t = ttk.Frame(self.tab_master)
        f_t.pack(fill=tk.X, pady=15)
        ttk.Label(f_t, text="2. File Master Testing:", font=("Arial", 10, "bold"), width=25).pack(side=tk.LEFT)
        ttk.Button(f_t, text="Pilih Testing", command=self.pilih_master_test).pack(side=tk.LEFT, padx=10)
        self.lbl_master_test = ttk.Label(f_t, text="Belum ada file...", font=("Arial", 9, "italic"), foreground="gray")
        self.lbl_master_test.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(self.tab_master, text="Catatan Sistem:\n- Sistem menggunakan Logika Kunci Ganda (Hybrid Key).\n- Jika Kode Paket kosong, sistem otomatis mencocokkan menggunakan Kode RUP.", justify=tk.LEFT).pack(anchor=tk.W, pady=(10,20))

        ttk.Button(self.tab_master, text="🚀 PROSES KOMPARASI MASTER", style="Action.TButton", command=self.proses_data_master).pack(fill=tk.X, pady=5)

    def pilih_master_prod(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.file_master_prod_path = p
            self.master_prod_sheet = self.cek_dan_pilih_sheet(p)
            info = f" ({self.master_prod_sheet})" if self.master_prod_sheet else ""
            self.lbl_master_prod.config(text=f"{os.path.basename(p)}{info}", font=("Arial", 9, "bold"), foreground="#2e7d32")
            self.log(f"[MASTER] File Produksi: {os.path.basename(p)}{info}")

    def pilih_master_test(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.file_master_test_path = p
            self.master_test_sheet = self.cek_dan_pilih_sheet(p)
            info = f" ({self.master_test_sheet})" if self.master_test_sheet else ""
            self.lbl_master_test.config(text=f"{os.path.basename(p)}{info}", font=("Arial", 9, "bold"), foreground="#2e7d32")
            self.log(f"[MASTER] File Testing: {os.path.basename(p)}{info}")

    def proses_data_master(self):
        if not self.file_master_prod_path or not self.file_master_test_path:
            return messagebox.showerror("Error", "File Produksi dan Testing harus diisi!")
            
        self.log("\n--- Mulai Komparasi MASTER vs MASTER ---")
        try:
            df_p = self.load_file(self.file_master_prod_path, self.master_prod_sheet)
            df_t = self.load_file(self.file_master_test_path, self.master_test_sheet)
            
            p_paket, p_rup = self.cari_kolom_paket(df_p, df_p)[0], self.cari_kolom_rup(df_p)
            t_paket, t_rup = self.cari_kolom_paket(df_t, df_t)[0], self.cari_kolom_rup(df_t)
            
            if not p_paket or not p_rup or not t_paket or not t_rup:
                return messagebox.showerror("Error", "Kolom 'Kode Paket' atau 'Kode RUP' tidak ditemukan di salah satu file!")

            # Logika Hybrid Key
            def buat_hybrid_key(row, col_pkt, col_rup):
                val_pkt = self.ekstrak_dan_bersihkan_kode(row.get(col_pkt, ""))
                val_rup = self.ekstrak_dan_bersihkan_kode(row.get(col_rup, ""))
                
                pkt = val_pkt[0] if val_pkt else ""
                rup = val_rup[0] if val_rup else ""
                
                # Tambahkan prefix agar angka RUP tidak keliru terbaca sebagai Kode Paket
                if pkt and pkt != "": return f"PKT_{pkt}"
                if rup and rup != "": return f"RUP_{rup}"
                return None

            self.log("✓ Membuat Hybrid Key (Memprioritaskan Kode Paket, fallback Kode RUP)...")
            df_p['Hybrid_Key'] = df_p.apply(lambda r: buat_hybrid_key(r, p_paket, p_rup), axis=1)
            df_t['Hybrid_Key'] = df_t.apply(lambda r: buat_hybrid_key(r, t_paket, t_rup), axis=1)
            
            # Hapus baris yang benar-benar tidak punya Kode Paket maupun Kode RUP
            df_p = df_p.dropna(subset=['Hybrid_Key'])
            df_t = df_t.dropna(subset=['Hybrid_Key'])
            
            set_p = set(df_p['Hybrid_Key'].unique())
            set_t = set(df_t['Hybrid_Key'].unique())
            
            df_hanya_p = df_p[~df_p['Hybrid_Key'].isin(set_t)].copy().drop(columns=['Hybrid_Key'])
            df_hanya_t = df_t[~df_t['Hybrid_Key'].isin(set_p)].copy().drop(columns=['Hybrid_Key'])
            df_kedua = df_p[df_p['Hybrid_Key'].isin(set_t)].copy().drop(columns=['Hybrid_Key'])
            
            self.log(f"-> Hanya di Produksi: {len(df_hanya_p)} | Hanya di Testing: {len(df_hanya_t)} | Match: {len(df_kedua)}")
            
            pesan = f"Komparasi Master berhasil diproses!\n\nRingkasan:\n- Hanya di Produksi: {len(df_hanya_p)}\n- Hanya di Testing: {len(df_hanya_t)}\n- Cocok di Keduanya: {len(df_kedua)}\n\nApakah Anda ingin menyimpan hasilnya ke Excel?"
            if messagebox.askyesno("Konfirmasi Simpan", pesan):
                save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="Rekonsiliasi_Master_Prod_vs_Test.xlsx")
                if save_path:
                    with pd.ExcelWriter(save_path, engine='openpyxl') as w:
                        df_hanya_p.to_excel(w, sheet_name='Hanya di Produksi', index=False)
                        df_hanya_t.to_excel(w, sheet_name='Hanya di Testing', index=False)
                        df_kedua.to_excel(w, sheet_name='Cocok di Keduanya', index=False)
                    self.log(f"[SUKSES] Disimpan di: {save_path}")
                    messagebox.showinfo("Sukses", "File Excel berhasil disimpan!")
            else:
                self.log("[INFO] Proses selesai. File tidak disimpan.")
                
        except Exception as e:
            self.log(f"[ERROR] {traceback.format_exc()}")
            messagebox.showerror("Error", f"Terjadi kesalahan:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RekonsiliasiApp(root)
    root.mainloop()