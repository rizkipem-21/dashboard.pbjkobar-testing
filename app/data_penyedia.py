import json
import os
import tkinter as tk
from tkinter import ttk, messagebox

# ==========================================
# KONFIGURASI LOKASI FILE JSON
# ==========================================
NAMA_FOLDER_DATA = "data_master"
NAMA_FILE_JSON = "kamus_penyedia.json"


def ambil_path_file_json():
    """
    Mencari file JSON dengan dua prioritas:
    1. app/data_master/kamus_penyedia.json
    2. data_master/kamus_penyedia.json (satu folder di atas app)
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(base_dir)

    daftar_path = [
        os.path.join(base_dir, NAMA_FOLDER_DATA, NAMA_FILE_JSON),
        os.path.join(parent_dir, NAMA_FOLDER_DATA, NAMA_FILE_JSON),
    ]

    for path in daftar_path:
        if os.path.isfile(path):
            return path

    return daftar_path[0]


class AppPenyedia(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mini App - Data Penyedia")
        self.geometry("1000x600")
        self.minsize(800, 500)
        self.configure(padx=10, pady=10)

        try:
            self.state('zoomed')
        except Exception:
            pass

        self.dict_penyedia = {}
        self.data_semua = []
        self.data_semua_teks = []
        self._filter_json_job = None

        self.create_layout()
        self.load_data_awal()

    def create_layout(self):
        # === Frame Top Control ===
        frame_top = ttk.LabelFrame(self, text="  Kontrol Pencarian Penyedia  ", padding=(15, 15))
        frame_top.pack(side="top", fill="x", pady=(0, 15))

        # ===============================
        # BARIS 1: Filter tabel kiri
        # ===============================
        ttk.Label(frame_top, text="Cari Nama/NPWP:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky="w", padx=(0, 5)
        )

        self.ent_filter_penyedia = ttk.Entry(frame_top, width=35, font=('Arial', 11))
        self.ent_filter_penyedia.grid(row=0, column=1, padx=(0, 20), sticky="w")
        self.ent_filter_penyedia.bind("<KeyRelease>", self.filter_tabel_penyedia)

        ttk.Label(frame_top, text="Highlight Teks:", font=('Arial', 10, 'bold')).grid(
            row=0, column=2, sticky="w", padx=(0, 5)
        )

        self.ent_filter_internal = ttk.Entry(frame_top, width=30, font=('Arial', 11))
        self.ent_filter_internal.grid(row=0, column=3, padx=(0, 10), sticky="w")
        self.ent_filter_internal.bind("<Return>", lambda e: self.eksekusi_cari_internal_text())

        btn_cari_internal = ttk.Button(frame_top, text="⚡ Highlight", command=self.eksekusi_cari_internal_text)
        btn_cari_internal.grid(row=0, column=4, sticky="w")

        # ===============================
        # BARIS 2: Filter blok JSON
        # ===============================
        ttk.Label(frame_top, text="Filter Blok JSON:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky="w", padx=(0, 5), pady=(10, 0)
        )

        self.ent_filter_json = ttk.Entry(frame_top, width=35, font=('Arial', 11))
        self.ent_filter_json.grid(row=1, column=1, padx=(0, 10), pady=(10, 0), sticky="w")
        self.ent_filter_json.bind("<KeyRelease>", self.on_filter_json_key)
        self.ent_filter_json.bind("<Return>", lambda e: self.terapkan_filter_json())

        btn_filter_json = ttk.Button(frame_top, text="🔎 Filter Blok", command=self.terapkan_filter_json)
        btn_filter_json.grid(row=1, column=2, padx=(0, 20), pady=(10, 0), sticky="w")

        btn_semua_data = ttk.Button(frame_top, text="📄 Tampilkan Semua Data", command=self.tampilkan_semua_data)
        btn_semua_data.grid(row=1, column=3, pady=(10, 0), sticky="w")

        self.lbl_info_json = ttk.Label(frame_top, text="", foreground="#1d4ed8", font=('Arial', 9))
        self.lbl_info_json.grid(row=1, column=4, padx=(15, 0), pady=(10, 0), sticky="w")

        # === Main Split Frame ===
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True)

        # Sisi Kiri: Tabel Daftar Penyedia
        frame_kiri = ttk.LabelFrame(paned, text="  Daftar Penyedia  ", padding=5)
        paned.add(frame_kiri, weight=1)

        self.tree = ttk.Treeview(frame_kiri, columns=("no", "nama", "npwp"), show="headings", selectmode="browse")
        self.tree.heading("no", text="No.")
        self.tree.heading("nama", text="Nama Penyedia")
        self.tree.heading("npwp", text="NPWP / Kode")
        self.tree.column("no", width=40, anchor="center")
        self.tree.column("nama", width=250, anchor="w")
        self.tree.column("npwp", width=150, anchor="w")

        scroll_tree = ttk.Scrollbar(frame_kiri, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scroll_tree.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll_tree.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.tampilkan_detail_json)

        # Sisi Kanan: Pratinjau JSON
        frame_kanan = ttk.LabelFrame(paned, text="  Detail JSON Lengkap  ", padding=5)
        paned.add(frame_kanan, weight=3)

        self.txt_viewer = tk.Text(frame_kanan, wrap="none", font=("Consolas", 10), bg="#f8fafc", fg="#0f172a")
        self.txt_viewer.bind("<Key>", self.kunci_teks)

        scroll_y = ttk.Scrollbar(frame_kanan, orient="vertical", command=self.txt_viewer.yview)
        scroll_x = ttk.Scrollbar(frame_kanan, orient="horizontal", command=self.txt_viewer.xview)
        self.txt_viewer.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_x.pack(side="bottom", fill="x")
        self.txt_viewer.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        self.txt_viewer.tag_config('highlight_search', background='#eab308', foreground='#000000')

    def update_info_json(self, pesan):
        self.lbl_info_json.configure(text=pesan)

    def tampilkan_semua_data(self):
        """
        Menampilkan semua data penyedia sebagai list JSON di panel kanan.
        """
        if self._filter_json_job is not None:
            try:
                self.after_cancel(self._filter_json_job)
            except Exception:
                pass
            self._filter_json_job = None

        self.ent_filter_json.delete(0, tk.END)
        self.ent_filter_internal.delete(0, tk.END)
        self.txt_viewer.tag_remove('highlight_search', '1.0', tk.END)
        self.txt_viewer.delete(1.0, tk.END)

        if not self.data_semua:
            self.txt_viewer.insert(tk.END, "Tidak ada data untuk ditampilkan.")
            self.update_info_json("0 blok data")
            return

        pretty_json = json.dumps(self.data_semua, indent=4, ensure_ascii=False, sort_keys=True)
        self.txt_viewer.insert(tk.END, pretty_json)
        self.update_info_json(f"Menampilkan {len(self.data_semua)} blok data")

    def on_filter_json_key(self, event=None):
        """
        Debounce filter JSON agar tidak terlalu sering diproses saat mengetik.
        """
        if self._filter_json_job is not None:
            try:
                self.after_cancel(self._filter_json_job)
            except Exception:
                pass

        self._filter_json_job = self.after(250, self.terapkan_filter_json)

    def terapkan_filter_json(self):
        """
        Menampilkan hanya blok data { ... } yang mengandung teks filter.
        """
        self._filter_json_job = None
        keyword = self.ent_filter_json.get().strip()

        if not keyword:
            self.tampilkan_semua_data()
            return

        keyword_lower = keyword.lower()
        hasil = []

        for item, teks_item in zip(self.data_semua, self.data_semua_teks):
            if keyword_lower in teks_item:
                hasil.append(item)

        self.ent_filter_internal.delete(0, tk.END)
        self.txt_viewer.tag_remove('highlight_search', '1.0', tk.END)
        self.txt_viewer.delete(1.0, tk.END)

        if hasil:
            pretty_json = json.dumps(hasil, indent=4, ensure_ascii=False, sort_keys=True)
            self.txt_viewer.insert(tk.END, pretty_json)
            self.update_info_json(f"Ditemukan {len(hasil)} blok dari {len(self.data_semua)}")
        else:
            self.txt_viewer.insert(tk.END, f"Tidak ada blok data yang mengandung '{keyword}'.")
            self.update_info_json("0 blok ditemukan")

    def load_data_awal(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        self.dict_penyedia.clear()
        self.data_semua.clear()
        self.data_semua_teks.clear()

        data_list = []
        file_path = ambil_path_file_json()

        try:
            if not os.path.isfile(file_path):
                base_dir = os.path.dirname(os.path.abspath(__file__))
                parent_dir = os.path.dirname(base_dir)

                raise FileNotFoundError(
                    "File JSON tidak ditemukan.\n\n"
                    f"Path yang dicoba 1:\n{os.path.join(base_dir, NAMA_FOLDER_DATA, NAMA_FILE_JSON)}\n\n"
                    f"Path yang dicoba 2:\n{os.path.join(parent_dir, NAMA_FOLDER_DATA, NAMA_FILE_JSON)}\n\n"
                    "Buat folder bernama 'data_master' lalu letakkan file 'kamus_penyedia.json' di dalamnya."
                )

            with open(file_path, "r", encoding="utf-8-sig") as file_json:
                data_json = json.load(file_json)

            if isinstance(data_json, list):
                data_list = data_json

            elif isinstance(data_json, dict):
                extracted_list = None

                for key in ("data", "result", "items", "records", "rows"):
                    if key in data_json and isinstance(data_json[key], list):
                        extracted_list = data_json[key]
                        break

                if extracted_list is None:
                    for value in data_json.values():
                        if isinstance(value, list):
                            extracted_list = value
                            break

                if extracted_list is None:
                    raise ValueError("JSON object tidak berisi daftar/list data penyedia.")

                data_list = extracted_list

            else:
                raise ValueError("Format data JSON harus berupa list/array atau object yang berisi list.")

            data_list = [row for row in data_list if isinstance(row, dict)]

        except FileNotFoundError as e:
            messagebox.showerror("File Tidak Ditemukan", str(e))
            self.txt_viewer.delete(1.0, tk.END)
            self.txt_viewer.insert(tk.END, "File data belum ditemukan.\nSilakan siapkan folder 'data_master' dan file 'kamus_penyedia.json'.")
            self.update_info_json("File tidak ditemukan")
            return

        except json.JSONDecodeError as e:
            messagebox.showerror(
                "Error JSON",
                f"Format file JSON tidak valid.\n\nFile:\n{file_path}\n\nDetail error:\n{str(e)}"
            )
            self.txt_viewer.delete(1.0, tk.END)
            self.txt_viewer.insert(tk.END, "Format file JSON tidak valid.")
            self.update_info_json("JSON error")
            return

        except Exception as e:
            messagebox.showerror("Error Sistem", f"Terjadi kesalahan saat membaca file:\n{str(e)}")
            self.txt_viewer.delete(1.0, tk.END)
            self.txt_viewer.insert(tk.END, "Terjadi kesalahan saat membaca file.")
            self.update_info_json("Error")
            return

        # Urutkan data berdasarkan nama penyedia A-Z
        sorted_data = sorted(
            data_list,
            key=lambda x: str(x.get("nama_penyedia", "") or "").lower()
        )

        self.data_semua = sorted_data

        # Precompute teks JSON untuk filter blok
        for item in sorted_data:
            try:
                teks_item = json.dumps(item, ensure_ascii=False).lower()
            except Exception:
                teks_item = str(item).lower()
            self.data_semua_teks.append(teks_item)

        # Isi tabel kiri
        for idx, item in enumerate(sorted_data, 1):
            key_id = f"item_{idx}"
            self.dict_penyedia[key_id] = item

            nama = item.get("nama_penyedia") or "-"
            npwp = item.get("npwp_penyedia") or item.get("kode_penyedia") or "-"

            self.tree.insert("", "end", values=(idx, nama, npwp), tags=(key_id,))

        # Tampilkan petunjuk awal di panel kanan
        self.txt_viewer.delete(1.0, tk.END)
        self.txt_viewer.insert(
            tk.END,
            f"Berhasil memuat {len(sorted_data)} data penyedia.\n\n"
            "1. Klik salah satu penyedia di tabel kiri untuk melihat detail satu blok data.\n"
            "2. Klik tombol '📄 Tampilkan Semua Data' untuk menampilkan seluruh data.\n"
            "3. Gunakan 'Filter Blok JSON' untuk menampilkan hanya blok { ... } yang mengandung teks tertentu."
        )

        self.update_info_json(f"{len(sorted_data)} blok data dimuat")

    def filter_tabel_penyedia(self, event=None):
        keyword = self.ent_filter_penyedia.get().lower().strip()

        for i in self.tree.get_children():
            self.tree.delete(i)

        counter = 1
        for key_id, item in self.dict_penyedia.items():
            nama = str(item.get("nama_penyedia", "") or "").lower()
            npwp = str(item.get("npwp_penyedia", "") or "").lower()
            kode = str(item.get("kode_penyedia", "") or "").lower()

            if keyword in nama or keyword in npwp or keyword in kode:
                nama_display = item.get("nama_penyedia") or "-"
                npwp_display = item.get("npwp_penyedia") or item.get("kode_penyedia") or "-"
                self.tree.insert("", "end", values=(counter, nama_display, npwp_display), tags=(key_id,))
                counter += 1

    def tampilkan_detail_json(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return

        key_id = self.tree.item(selected[0], "tags")[0]
        data_item = self.dict_penyedia.get(key_id, {})

        if self._filter_json_job is not None:
            try:
                self.after_cancel(self._filter_json_job)
            except Exception:
                pass
            self._filter_json_job = None

        self.ent_filter_json.delete(0, tk.END)
        self.ent_filter_internal.delete(0, tk.END)
        self.txt_viewer.tag_remove('highlight_search', '1.0', tk.END)

        self.txt_viewer.delete(1.0, tk.END)

        pretty_json = json.dumps(data_item, indent=4, ensure_ascii=False, sort_keys=True)
        self.txt_viewer.insert(tk.END, pretty_json)

        self.update_info_json("Menampilkan 1 blok data")

    def eksekusi_cari_internal_text(self):
        self.txt_viewer.tag_remove('highlight_search', '1.0', tk.END)
        keyword = self.ent_filter_internal.get().strip()
        if not keyword:
            return

        posisi_index = '1.0'
        jumlah_cocok = 0

        while True:
            posisi_index = self.txt_viewer.search(keyword, posisi_index, nocase=True, stopindex=tk.END)
            if not posisi_index:
                break

            jumlah_cocok += 1
            akhir_posisi = f"{posisi_index}+{len(keyword)}c"
            self.txt_viewer.tag_add('highlight_search', posisi_index, akhir_posisi)

            if jumlah_cocok == 1:
                self.txt_viewer.see(posisi_index)

            posisi_index = akhir_posisi

        if jumlah_cocok == 0:
            messagebox.showinfo("Pencarian", f"Kata '{keyword}' tidak ditemukan di detail JSON ini.")

    def kunci_teks(self, event):
        # Mengizinkan Ctrl+C (Copy) dan Panah Navigasi, text box bersifat Read-Only
        if (event.state & 4 and event.keysym.lower() == 'c') or event.keysym in ('Up', 'Down', 'Left', 'Right', 'Prior', 'Next'):
            return None
        return "break"


if __name__ == "__main__":
    import traceback

    print("[INFO] Aplikasi data_penyedia.py dimulai")

    try:
        print("[INFO] Membuat window aplikasi...")
        app = AppPenyedia()
        print("[INFO] Window berhasil dibuat. Membuka mainloop...")
        app.mainloop()
        print("[INFO] Aplikasi ditutup.")

    except Exception:
        print("[ERROR] Terjadi exception:")
        traceback.print_exc()

        try:
            with open("error_data_penyedia.log", "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
        except Exception:
            pass

        input("Tekan Enter untuk keluar...")