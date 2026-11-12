import json
import os
import re
import tkinter as tk
from tkinter import ttk, messagebox

# ==========================================
# KONFIGURASI LOKASI FILE JSON
# ==========================================
NAMA_FOLDER_DATA = "data_master"
NAMA_FILE_JSON = "kamus_penyedia.json"

# ==========================================
# TEMA WARNA
# ==========================================
C = {
    "bg": "#eef2f7",
    "card": "#ffffff",
    "header": "#0f172a",
    "header_sub": "#94a3b8",
    "aksen": "#2563eb",
    "aksen_hover": "#1d4ed8",
    "netral": "#e2e8f0",
    "netral_hover": "#cbd5e1",
    "teks": "#0f172a",
    "muted": "#64748b",
    "garis": "#e2e8f0",
    "zebra": "#f8fafc",
    "seleksi": "#dbeafe",
    "viewer_bg": "#0b1220",
    "viewer_fg": "#e2e8f0",
    "j_key": "#7dd3fc",
    "j_str": "#86efac",
    "j_num": "#fbbf24",
    "j_bool": "#f472b6",
    "j_null": "#94a3b8",
}
FONT = "Segoe UI"


def ambil_path_file_json():
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


# ==========================================
# HELPER MEMBACA FIELD (tahan terhadap sufiks _v1/_v5/_v6)
# ==========================================
RE_BADAN = re.compile(
    r'(^|[\s,(])(PT|CV|UD|PD|PO|FA|FIRMA|KOPERASI|KOP|YAYASAN|PERUM|PERSERO|BUMDES)(?=[\s.,)]|$)',
    re.I,
)


def _rapikan(v):
    if v is None:
        return ""
    s = " ".join(str(v).split())
    return "" if s.lower() in ("-", "none", "null") else s


def varian_field(item, base):
    """Daftar (nama_key, nomor_versi|None) untuk base, base_v1, base_v5, base_v6, ..."""
    pola = re.compile(re.escape(base) + r"(?:_v(\d+))?")
    hasil = []
    for k in item:
        m = pola.fullmatch(k)
        if m:
            hasil.append((k, m.group(1)))
    hasil.sort(key=lambda x: (x[1] is not None, int(x[1] or 0)))
    return hasil


def ambil_nama(item):
    langsung = _rapikan(item.get("nama_tampil"))
    if langsung:
        return langsung

    kandidat = []
    for k, _ in varian_field(item, "nama_penyedia"):
        n = _rapikan(item.get(k))
        if n and n not in kandidat:
            kandidat.append(n)
    for n in kandidat:
        if RE_BADAN.search(n):
            return n
    return kandidat[0] if kandidat else "-"


def _digits(v):
    return re.sub(r"\D", "", str(v or ""))


def _cari_npwp(item, panjang):
    keys = [f"npwp_{panjang}_tampil", f"npwp_{panjang}"]
    keys += [k for k in item if k.startswith(f"npwp_{panjang}_") and k not in keys]
    keys += ["npwp_penyedia"]
    for k in keys:
        d = _digits(item.get(k))
        if len(d) == panjang and len(set(d)) > 1:
            return d
    return ""


def format_npwp15(d):
    if len(d) != 15:
        return d or "-"
    return f"{d[0:2]}.{d[2:5]}.{d[5:8]}.{d[8]}-{d[9:12]}.{d[12:15]}"


def format_npwp16(d):
    return d or "-"


def ambil_kode(item):
    for k in ("epur_v6", "epur_v5", "kd_penyedia_v1", "kode_penyedia", "kd_penyedia"):
        if item.get(k):
            return str(item[k])
    return ""


def tampil_nilai(v):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return "-"
    if isinstance(v, (list, dict)):
        return json.dumps(v, ensure_ascii=False)
    return str(v)


# ---------- IZIN USAHA ----------
def ambil_izin(item):
    """Return list (jenis, kualifikasi, nomor)."""
    hasil = []
    for k in sorted(item):
        if not (k == "izin_usaha" or re.fullmatch(r"izin_usaha_v\d+", k)):
            continue
        raw = item[k]
        if isinstance(raw, str):
            try:
                raw = json.loads(raw)
            except ValueError:
                continue
        if isinstance(raw, list):
            for it in raw:
                if isinstance(it, dict):
                    hasil.append((
                        str(it.get("jenis_izin_usaha") or "-"),
                        str(it.get("kualifikasi_usaha") or "-"),
                        str(it.get("nomor_izin_usaha") or "-"),
                    ))
    return hasil


# ---------- KBLI ----------
def _split_id(v):
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in re.split(r"[;,]", str(v or "")) if p.strip()]


def _split_nama(v):
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    s = str(v or "")
    bagian = s.split(";") if ";" in s else re.split(r",(?!\s)", s)
    return [p.strip() for p in bagian if p.strip()]


def _pilih_varian(item, base):
    for k, _ in varian_field(item, base):
        if _rapikan(item.get(k)) or isinstance(item.get(k), (list, tuple)):
            return item[k]
    return None


def ambil_kbli(item):
    """Return (rows[(kode, uraian)], catatan)."""
    ids = _split_id(_pilih_varian(item, "kbli_id"))
    names = _split_nama(_pilih_varian(item, "kbli_name"))
    rows, catatan = [], ""

    if ids or names:
        if ids and names and len(ids) != len(names):
            catatan = f"jumlah kode ({len(ids)}) berbeda dengan uraian ({len(names)})"
        for i in range(max(len(ids), len(names))):
            rows.append((ids[i] if i < len(ids) else "", names[i] if i < len(names) else ""))

    for k in sorted(item):
        if k.startswith("kbli") and not re.fullmatch(r"kbli_(id|name)(_v\d+)?", k):
            for el in _split_nama(item[k]):
                if re.fullmatch(r"\d{4,6}", el):
                    rows.append((el, f"({k})"))
                else:
                    rows.append(("", f"{el}  ({k})"))

    unik, lihat = [], set()
    for r in rows:
        if r not in lihat:
            lihat.add(r)
            unik.append(r)
    return unik, catatan


# ---------- RINGKASAN ----------
GRUP_RINGKAS = [
    ("Identitas", [
        ("bentuk_usaha", "Bentuk Usaha"),
        ("jenis_perusahaan", "Jenis Perusahaan"),
        ("status_pelaku_usaha", "Status Pelaku Usaha"),
        ("status_aktif", "Status Aktif"),
    ]),
    ("Alamat & Kontak", [
        ("alamat_penyedia", "Alamat"),
        ("kabupaten", "Kabupaten / Kota"),
        ("provinsi", "Provinsi"),
        ("kodepos", "Kode Pos"),
        ("email", "Email"),
        ("telepon", "Telepon"),
    ]),
    ("Status & Verifikasi", [
        ("status_npwp", "Status NPWP"),
        ("status_kswp", "Status KSWP"),
        ("status_umkk", "Status UMKK"),
        ("setuju_publikasi_data", "Setuju Publikasi Data"),
        ("tgl_daftar_sikap", "Tgl Daftar Sikap"),
        ("tgl_persetujuan_verifikasi_daftar_sikap", "Tgl Verifikasi Sikap"),
    ]),
]

KODE_ROWS = [
    ("epur_v6", "Kode e-Purchasing v6"),
    ("epur_v5", "Kode e-Purchasing v5"),
    ("kd_penyedia_v1", "Kode Sikap (v1)"),
    ("kode_penyedia", "Kode Penyedia"),
    ("kd_penyedia", "Kode Penyedia (arsip)"),
    ("nib", "NIB"),
]


def nilai_gabung(item, base, used):
    var = varian_field(item, base)
    if not var:
        return None
    unik = []
    for k, s in var:
        used.add(k)
        t = tampil_nilai(item[k])
        if t not in [u[1] for u in unik]:
            unik.append((s, t))
    isi = [u for u in unik if u[1] != "-"]
    if not isi:
        return "-"
    if len(isi) == 1:
        return isi[0][1]
    return "\n".join(f"{t}   [v{s}]" if s else t for s, t in isi)


class AppPenyedia(tk.Tk):
    BATAS_FOLD = 1500

    def __init__(self):
        super().__init__()
        self.title("Data Penyedia")
        self.geometry("1200x700")
        self.minsize(900, 550)
        self.configure(bg=C["bg"])
        try:
            self.state("zoomed")
        except Exception:
            pass

        self.rows = []
        self.skema = {}   # semua parameter yang pernah muncul -> nilai default bila kosong
        self.tab = {}
        self.folds = []
        self._teks_json = ""
        self._mode_multi = False
        self._filter_json_job = None

        self.setup_style()
        self.create_layout()
        self.load_data_awal()

    # ------------------------------------------------------------------
    # STYLE
    # ------------------------------------------------------------------
    def setup_style(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except Exception:
            pass

        s.configure("TFrame", background=C["bg"])
        s.configure("Card.TFrame", background=C["card"], relief="flat")
        s.configure("Header.TFrame", background=C["header"])
        s.configure("TLabel", background=C["card"], foreground=C["teks"], font=(FONT, 10))
        s.configure("Field.TLabel", background=C["card"], foreground=C["muted"], font=(FONT, 9, "bold"))
        s.configure("Title.TLabel", background=C["header"], foreground="#ffffff", font=(FONT, 16, "bold"))
        s.configure("Sub.TLabel", background=C["header"], foreground=C["header_sub"], font=(FONT, 9))
        s.configure("Badge.TLabel", background="#1e293b", foreground="#e2e8f0",
                    font=(FONT, 10, "bold"), padding=(12, 6))
        s.configure("Info.TLabel", background=C["card"], foreground=C["aksen"], font=(FONT, 9, "bold"))
        s.configure("Judul.TLabel", background=C["card"], foreground=C["teks"], font=(FONT, 14, "bold"))
        s.configure("Muted.TLabel", background=C["card"], foreground=C["muted"], font=(FONT, 9))
        s.configure("CardTitle.TLabel", background=C["card"], foreground=C["teks"], font=(FONT, 11, "bold"))

        s.configure("TEntry", fieldbackground="#ffffff", bordercolor=C["garis"],
                    lightcolor=C["garis"], darkcolor=C["garis"], padding=6)
        s.map("TEntry", bordercolor=[("focus", C["aksen"])], lightcolor=[("focus", C["aksen"])],
              darkcolor=[("focus", C["aksen"])])

        s.configure("Aksen.TButton", background=C["aksen"], foreground="#ffffff",
                    font=(FONT, 9, "bold"), padding=(14, 7), borderwidth=0, focusthickness=0)
        s.map("Aksen.TButton", background=[("active", C["aksen_hover"])])
        s.configure("Netral.TButton", background=C["netral"], foreground=C["teks"],
                    font=(FONT, 9, "bold"), padding=(14, 7), borderwidth=0, focusthickness=0)
        s.map("Netral.TButton", background=[("active", C["netral_hover"])])

        s.configure("Treeview", background="#ffffff", fieldbackground="#ffffff",
                    foreground=C["teks"], rowheight=30, borderwidth=0, font=(FONT, 10))
        s.map("Treeview", background=[("selected", C["seleksi"])],
              foreground=[("selected", C["teks"])])
        s.configure("Treeview.Heading", background=C["header"], foreground="#ffffff",
                    font=(FONT, 10, "bold"), padding=(8, 8), borderwidth=0, relief="flat")
        s.map("Treeview.Heading", background=[("active", "#1e293b")])

        for orient in ("Vertical", "Horizontal"):
            s.configure(f"{orient}.TScrollbar", background="#94a3b8", troughcolor="#e2e8f0",
                        bordercolor="#e2e8f0", lightcolor="#94a3b8", darkcolor="#94a3b8",
                        arrowcolor="#334155", arrowsize=16, gripcount=0)
            s.map(f"{orient}.TScrollbar", background=[("active", C["aksen"]), ("pressed", C["aksen_hover"])])
        s.configure("TPanedwindow", background=C["bg"])

        s.configure("TNotebook", background=C["card"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        s.configure("TNotebook.Tab", background=C["netral"], foreground=C["muted"],
                    padding=(16, 8), font=(FONT, 10, "bold"), borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", C["aksen"]), ("active", C["netral_hover"])],
              foreground=[("selected", "#ffffff")])

    # ------------------------------------------------------------------
    # LAYOUT
    # ------------------------------------------------------------------
    def create_layout(self):
        header = ttk.Frame(self, style="Header.TFrame", padding=(22, 14))
        header.pack(side="top", fill="x")

        kiri = ttk.Frame(header, style="Header.TFrame")
        kiri.pack(side="left")
        ttk.Label(kiri, text="Data Penyedia", style="Title.TLabel").pack(anchor="w")
        ttk.Label(kiri, text="Kamus penyedia gabungan e-Katalog (S6 + S7) dan Sikap",
                  style="Sub.TLabel").pack(anchor="w")
        self.lbl_total = ttk.Label(header, text="0 penyedia", style="Badge.TLabel")
        self.lbl_total.pack(side="right")

        body = ttk.Frame(self, padding=(14, 12, 14, 8))
        body.pack(fill="both", expand=True)

        # ---------- Kartu kontrol ----------
        card = ttk.Frame(body, style="Card.TFrame", padding=(18, 14))
        card.pack(side="top", fill="x", pady=(0, 12))

        ttk.Label(card, text="Cari Nama / NPWP", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, text="Highlight Teks (tab JSON)", style="Field.TLabel").grid(row=0, column=1, sticky="w", padx=(18, 0))
        ttk.Label(card, text="Filter Blok JSON", style="Field.TLabel").grid(row=0, column=3, sticky="w", padx=(18, 0))

        self.ent_filter_penyedia = ttk.Entry(card, width=34, font=(FONT, 11))
        self.ent_filter_penyedia.grid(row=1, column=0, sticky="we", pady=(4, 0))
        self.ent_filter_penyedia.bind("<KeyRelease>", self.filter_tabel_penyedia)

        self.ent_filter_internal = ttk.Entry(card, width=26, font=(FONT, 11))
        self.ent_filter_internal.grid(row=1, column=1, sticky="we", padx=(18, 6), pady=(4, 0))
        self.ent_filter_internal.bind("<Return>", lambda e: self.eksekusi_cari_internal_text())

        ttk.Button(card, text="Highlight", style="Aksen.TButton",
                   command=self.eksekusi_cari_internal_text).grid(row=1, column=2, pady=(4, 0))

        self.ent_filter_json = ttk.Entry(card, width=26, font=(FONT, 11))
        self.ent_filter_json.grid(row=1, column=3, sticky="we", padx=(18, 6), pady=(4, 0))
        self.ent_filter_json.bind("<KeyRelease>", self.on_filter_json_key)
        self.ent_filter_json.bind("<Return>", lambda e: self.terapkan_filter_json())

        ttk.Button(card, text="Filter Blok", style="Aksen.TButton",
                   command=self.terapkan_filter_json).grid(row=1, column=4, pady=(4, 0))
        ttk.Button(card, text="Semua Data", style="Netral.TButton",
                   command=self.tampilkan_semua_data).grid(row=1, column=5, padx=(8, 0), pady=(4, 0))

        ttk.Button(card, text="Reset Filter", style="Netral.TButton",
                   command=self.reset_filter).grid(row=1, column=6, padx=(8, 0), pady=(4, 0))

        self.lbl_info_json = ttk.Label(card, text="", style="Info.TLabel")
        self.lbl_info_json.grid(row=1, column=7, padx=(18, 0), pady=(4, 0), sticky="w")

        card.columnconfigure(0, weight=2)
        card.columnconfigure(1, weight=1)
        card.columnconfigure(3, weight=1)

        # ---------- Split ----------
        paned = ttk.PanedWindow(body, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True)

        # === Kiri: daftar ===
        kartu_kiri = ttk.Frame(paned, style="Card.TFrame", padding=(12, 10))
        paned.add(kartu_kiri, weight=2)

        bar_kiri = ttk.Frame(kartu_kiri, style="Card.TFrame")
        bar_kiri.pack(fill="x", pady=(0, 8))
        ttk.Label(bar_kiri, text="Daftar Penyedia", style="CardTitle.TLabel").pack(side="left")
        self.lbl_hitung = ttk.Label(bar_kiri, text="", style="Muted.TLabel")
        self.lbl_hitung.pack(side="right")

        wadah_tree = ttk.Frame(kartu_kiri, style="Card.TFrame")
        wadah_tree.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(wadah_tree, columns=("no", "nama", "npwp15", "npwp16"),
                                 show="headings", selectmode="browse")
        self.tree.heading("no", text="No.")
        self.tree.heading("nama", text="Nama Penyedia", anchor="w")
        self.tree.heading("npwp15", text="NPWP 15", anchor="w")
        self.tree.heading("npwp16", text="NPWP 16", anchor="w")
        self.tree.column("no", width=50, anchor="center", stretch=False)
        self.tree.column("nama", width=270, anchor="w")
        self.tree.column("npwp15", width=150, anchor="w", stretch=False)
        self.tree.column("npwp16", width=150, anchor="w", stretch=False)
        self.tree.tag_configure("ganjil", background=C["zebra"])
        self.tree.tag_configure("genap", background="#ffffff")

        scroll_tree = ttk.Scrollbar(wadah_tree, orient="vertical", command=self.tree.yview)
        scroll_tree_x = ttk.Scrollbar(wadah_tree, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_tree.set, xscrollcommand=scroll_tree_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_tree.grid(row=0, column=1, sticky="ns")
        scroll_tree_x.grid(row=1, column=0, sticky="ew")
        wadah_tree.rowconfigure(0, weight=1)
        wadah_tree.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self.tampilkan_detail_json)

        # === Kanan: detail bertab ===
        kartu_kanan = ttk.Frame(paned, style="Card.TFrame", padding=(14, 10))
        paned.add(kartu_kanan, weight=3)

        judul_box = ttk.Frame(kartu_kanan, style="Card.TFrame")
        judul_box.pack(fill="x", pady=(0, 8))
        self.lbl_judul = ttk.Label(judul_box, text="Detail Penyedia", style="Judul.TLabel")
        self.lbl_judul.pack(anchor="w")
        self.lbl_sub = ttk.Label(judul_box, text="Pilih penyedia di tabel kiri", style="Muted.TLabel")
        self.lbl_sub.pack(anchor="w")

        self.nb = ttk.Notebook(kartu_kanan)
        self.nb.pack(fill="both", expand=True)

        self.buat_tab_ringkasan()
        self.t_izin = self.buat_tab_tabel(
            "izin",
            [("jenis", "Jenis Izin Usaha", 300, True),
             ("kual", "Kualifikasi", 110, False),
             ("nomor", "Nomor Izin", 230, False)],
            kolom_salin=2,
        )
        self.t_kbli = self.buat_tab_tabel(
            "kbli",
            [("kode", "Kode KBLI", 110, False),
             ("uraian", "Uraian KBLI", 500, True)],
            kolom_salin=0,
        )
        self.buat_tab_json()

        self.nb.add(self.t_ringkas, text="Ringkasan")
        self.nb.add(self.t_izin, text="Izin Usaha")
        self.nb.add(self.t_kbli, text="KBLI")
        self.nb.add(self.t_json, text="JSON")

    # ---------- tab Ringkasan ----------
    def buat_tab_ringkasan(self):
        self.t_ringkas = ttk.Frame(self.nb, style="Card.TFrame", padding=(0, 10, 0, 0))
        self.txt_ringkas = tk.Text(
            self.t_ringkas, wrap="word", font=(FONT, 10), bg=C["card"], fg=C["teks"],
            relief="flat", borderwidth=0, padx=6, pady=4, cursor="arrow",
            tabs=("280",), selectbackground=C["seleksi"],
            selectforeground=C["teks"], spacing1=3, spacing3=3,
        )
        sc = ttk.Scrollbar(self.t_ringkas, orient="vertical", command=self.txt_ringkas.yview)
        self.txt_ringkas.configure(yscrollcommand=sc.set)
        sc.pack(side="right", fill="y")
        self.txt_ringkas.pack(side="left", fill="both", expand=True)
        self.txt_ringkas.bind("<Key>", self.kunci_teks)
        self.txt_ringkas.tag_config("grup", font=(FONT, 11, "bold"), foreground=C["aksen"],
                                    lmargin1=6, lmargin2=6, spacing1=14, spacing3=4)
        # lmargin hanya valid sebagai opsi tag. Baris dimulai dari label, jadi margin
        # baris pertama/pembungkus diatur di tag 'lbl'; lanjutan nilai multi-baris di tag 'val'.
        self.txt_ringkas.tag_config("lbl", foreground=C["muted"], font=(FONT, 10, "bold"),
                                    lmargin1=14, lmargin2=280)
        self.txt_ringkas.tag_config("val", foreground=C["teks"], lmargin1=280, lmargin2=280)
        self.txt_ringkas.tag_config("nama", font=(FONT, 12, "bold"), foreground=C["teks"],
                                    lmargin1=14, lmargin2=14)
        self.txt_ringkas.tag_config("pesan", foreground=C["muted"], lmargin1=10, lmargin2=10, spacing3=6)

    # ---------- tab tabel generik ----------
    def buat_tab_tabel(self, kunci, kolom, kolom_salin=0):
        wrap = ttk.Frame(self.nb, style="Card.TFrame", padding=(0, 10, 0, 0))

        bar = ttk.Frame(wrap, style="Card.TFrame")
        bar.pack(fill="x", pady=(0, 8))
        ttk.Label(bar, text="Filter:", style="Field.TLabel").pack(side="left")
        ent = ttk.Entry(bar, width=30, font=(FONT, 10))
        ent.pack(side="left", padx=8)
        lbl = ttk.Label(bar, text="", style="Muted.TLabel")
        lbl.pack(side="right")
        ttk.Label(bar, text="(klik 2x pada baris untuk menyalin)", style="Muted.TLabel").pack(side="right", padx=12)

        box = ttk.Frame(wrap, style="Card.TFrame")
        box.pack(fill="both", expand=True)
        ids = ["no"] + [k[0] for k in kolom]
        tree = ttk.Treeview(box, columns=ids, show="headings", selectmode="browse")
        tree.heading("no", text="No.")
        tree.column("no", width=50, anchor="center", stretch=False)
        for kid, judul, lebar, stretch in kolom:
            tree.heading(kid, text=judul, anchor="w")
            tree.column(kid, width=lebar, anchor="w", stretch=stretch)
        tree.tag_configure("ganjil", background=C["zebra"])
        tree.tag_configure("genap", background="#ffffff")
        sc = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        tree.configure(yscroll=sc.set)
        sc.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)

        self.tab[kunci] = {"tree": tree, "ent": ent, "lbl": lbl, "rows": [], "salin": kolom_salin}
        ent.bind("<KeyRelease>", lambda e, k=kunci: self.refresh_tab(k))
        tree.bind("<Double-1>", lambda e, k=kunci: self.salin_baris(k))
        return wrap

    def refresh_tab(self, kunci):
        t = self.tab[kunci]
        kata = t["ent"].get().strip().lower()
        for i in t["tree"].get_children():
            t["tree"].delete(i)
        n = 0
        for vals in t["rows"]:
            if kata and kata not in " ".join(map(str, vals)).lower():
                continue
            n += 1
            t["tree"].insert("", "end", values=(n, *vals),
                             tags=("ganjil" if n % 2 == 0 else "genap",))
        t["lbl"].configure(text=f"{n} dari {len(t['rows'])}")

    def isi_tab(self, kunci, rows):
        t = self.tab[kunci]
        t["rows"] = rows
        t["ent"].delete(0, tk.END)
        self.refresh_tab(kunci)

    def salin_baris(self, kunci):
        t = self.tab[kunci]
        sel = t["tree"].selection()
        if not sel:
            return
        vals = t["tree"].item(sel[0], "values")
        teks = str(vals[t["salin"] + 1])
        if teks:
            self.clipboard_clear()
            self.clipboard_append(teks)
            self.update_info_json(f"Disalin: {teks}")

    # ---------- tab JSON ----------
    def buat_tab_json(self):
        self.t_json = ttk.Frame(self.nb, style="Card.TFrame", padding=(0, 10, 0, 0))

        bar = ttk.Frame(self.t_json, style="Card.TFrame")
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="Buka Semua", style="Netral.TButton",
                   command=lambda: self.atur_semua_fold(True)).pack(side="left")
        ttk.Button(bar, text="Tutup Semua", style="Netral.TButton",
                   command=lambda: self.atur_semua_fold(False)).pack(side="left", padx=6)
        ttk.Button(bar, text="Salin JSON", style="Netral.TButton",
                   command=self.salin_json).pack(side="left")
        ttk.Label(bar, text="Bagian bersarang (mis. izin usaha) dilipat. Klik penanda ▶ untuk membuka.",
                  style="Muted.TLabel").pack(side="right")

        wadah = ttk.Frame(self.t_json, style="Card.TFrame")
        wadah.pack(fill="both", expand=True)

        self.txt_viewer = tk.Text(
            wadah, wrap="none", font=("Consolas", 10),
            bg=C["viewer_bg"], fg=C["viewer_fg"], insertbackground="#ffffff",
            selectbackground="#334155", relief="flat", padx=12, pady=10, borderwidth=0,
        )
        self.txt_viewer.bind("<Key>", self.kunci_teks)

        scroll_y = ttk.Scrollbar(wadah, orient="vertical", command=self.txt_viewer.yview)
        scroll_x = ttk.Scrollbar(wadah, orient="horizontal", command=self.txt_viewer.xview)
        self.txt_viewer.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_x.pack(side="bottom", fill="x")
        scroll_y.pack(side="right", fill="y")
        self.txt_viewer.pack(side="left", fill="both", expand=True)

        for t, w in (("j_key", "j_key"), ("j_str", "j_str"), ("j_num", "j_num"),
                     ("j_bool", "j_bool"), ("j_null", "j_null")):
            self.txt_viewer.tag_config(t, foreground=C[w])
        self.txt_viewer.tag_config("highlight_search", background="#eab308", foreground="#000000")

    # ------------------------------------------------------------------
    # VIEWER JSON + PEWARNAAN + LIPATAN
    # ------------------------------------------------------------------
    RE_KV = re.compile(r'^(\s*)("(?:[^"\\]|\\.)*")(\s*:\s*)(.*?)(,?)$')
    RE_STR_ONLY = re.compile(r'^(\s*)("(?:[^"\\]|\\.)*")(,?)$')

    def reset_folds(self):
        for f in self.folds:
            try:
                self.txt_viewer.tag_delete(f["tag"], f["mtag"])
            except Exception:
                pass
        self.folds = []

    def tampil_teks(self, teks, warnai=True, fold=None):
        """fold = indent dasar key yang boleh dilipat (4 = satu blok, 8 = daftar blok)."""
        self.reset_folds()
        self._teks_json = teks
        self.txt_viewer.tag_remove("highlight_search", "1.0", tk.END)
        self.txt_viewer.delete("1.0", tk.END)
        self.txt_viewer.insert(tk.END, teks)
        for t in ("j_key", "j_str", "j_num", "j_bool", "j_null"):
            self.txt_viewer.tag_remove(t, "1.0", tk.END)
        if warnai:
            self.warnai_json(teks)
        if fold:
            self.buat_fold(teks, fold)

    def warnai_json(self, teks):
        baris = teks.split("\n")
        if len(baris) > 25000:
            return
        tag_add = self.txt_viewer.tag_add

        def tag_nilai(no, awal, nilai):
            v = nilai.strip()
            if not v:
                return
            if v.startswith('"'):
                t = "j_str"
            elif v in ("true", "false"):
                t = "j_bool"
            elif v == "null":
                t = "j_null"
            elif re.match(r'^-?\d', v):
                t = "j_num"
            else:
                return
            tag_add(t, f"{no}.{awal}", f"{no}.{awal + len(v)}")

        for no, ln in enumerate(baris, 1):
            m = self.RE_KV.match(ln)
            if m:
                indent, key, sep, val, _ = m.groups()
                a = len(indent)
                tag_add("j_key", f"{no}.{a}", f"{no}.{a + len(key)}")
                tag_nilai(no, a + len(key) + len(sep), val)
                continue
            m = self.RE_STR_ONLY.match(ln)
            if m:
                indent, s, _ = m.groups()
                a = len(indent)
                tag_add("j_str", f"{no}.{a}", f"{no}.{a + len(s)}")

    def buat_fold(self, teks, indent_dasar):
        baris = teks.split("\n")
        if len(baris) > 60000:
            return
        stack, pasangan = [], []
        for no, ln in enumerate(baris, 1):
            s = ln.strip()
            if not s:
                continue
            if s[0] in "]}":
                if stack:
                    pasangan.append((stack.pop(), no))
            elif s[-1] in "[{":
                stack.append(no)

        kandidat = []
        for h, c in pasangan:
            ln = baris[h - 1]
            indent = len(ln) - len(ln.lstrip())
            if indent == indent_dasar and c - h > 1 and ln.lstrip().startswith('"'):
                kandidat.append((h, c, indent))

        if len(kandidat) > self.BATAS_FOLD:
            self.update_info_json("Data besar: lipatan dinonaktifkan")
            return

        t = self.txt_viewer
        for k, (h, c, indent) in enumerate(sorted(kandidat)):
            tag, mtag = f"fold_{k}", f"foldm_{k}"
            jml = 0
            for ln in baris[h:c - 1]:
                st = ln.strip()
                if st and len(ln) - len(ln.lstrip()) == indent + 4 and st[0] not in "]}":
                    jml += 1
            satuan = "item" if baris[h - 1].rstrip().endswith("[") else "field"

            t.insert(f"{h}.end", f"  ▶ {jml} {satuan}", (mtag,))
            t.tag_add(tag, f"{h}.end", f"{c}.end")
            t.tag_config(tag, elide=True)
            t.tag_config(mtag, foreground="#93c5fd", background="#1e293b",
                         font=("Consolas", 9, "bold"))
            t.tag_bind(mtag, "<Button-1>", lambda e, i=k: self._klik_fold(i))
            t.tag_bind(mtag, "<Enter>", lambda e: t.configure(cursor="hand2"))
            t.tag_bind(mtag, "<Leave>", lambda e: t.configure(cursor="xterm"))
            self.folds.append({"tag": tag, "mtag": mtag, "tutup": True})

    def _klik_fold(self, i):
        self.toggle_fold(i)
        return "break"

    def toggle_fold(self, i, buka=None):
        f = self.folds[i]
        if buka is None:
            buka = f["tutup"]
        if f["tutup"] == (not buka):
            return
        t = self.txt_viewer
        t.tag_configure(f["tag"], elide=not buka)
        f["tutup"] = not buka
        r = t.tag_ranges(f["mtag"])
        if r:
            a = str(r[0])
            t.delete(f"{a}+2c", f"{a}+3c")
            t.insert(f"{a}+2c", "▼" if buka else "▶", (f["mtag"],))

    def atur_semua_fold(self, buka):
        for i in range(len(self.folds)):
            self.toggle_fold(i, buka)

    def salin_json(self):
        if not self._teks_json:
            return
        self.clipboard_clear()
        self.clipboard_append(self._teks_json)
        self.update_info_json("JSON disalin ke clipboard")

    def update_info_json(self, pesan):
        self.lbl_info_json.configure(text=pesan)

    def batal_job_filter(self):
        if self._filter_json_job is not None:
            try:
                self.after_cancel(self._filter_json_job)
            except Exception:
                pass
            self._filter_json_job = None

    # ------------------------------------------------------------------
    # RINGKASAN
    # ------------------------------------------------------------------
    def tulis_ringkas(self, fn):
        t = self.txt_ringkas
        t.delete("1.0", tk.END)
        fn(t)
        t.yview_moveto(0)

    def isi_pesan(self, teks):
        self.tulis_ringkas(lambda t: t.insert(tk.END, teks, "pesan"))

    def hitung_skema(self, data_list):
        """Kumpulkan SEMUA parameter dari seluruh penyedia. list > dict > nilai biasa."""
        peringkat = {}
        for item in data_list:
            for k, v in item.items():
                p = 2 if isinstance(v, list) else 1 if isinstance(v, dict) else 0
                if p > peringkat.get(k, -1):
                    peringkat[k] = p
        self.skema = {k: ([] if p == 2 else {} if p == 1 else None)
                      for k, p in sorted(peringkat.items())}

    def lengkapi(self, item):
        """Salinan item dengan semua parameter; yang tidak dimiliki penyedia diisi kosong."""
        hasil = {}
        for k, default in self.skema.items():
            if k in item:
                hasil[k] = item[k]
            else:
                hasil[k] = list(default) if isinstance(default, list) else \
                    dict(default) if isinstance(default, dict) else default
        for k, v in item.items():   # pengaman bila ada key di luar skema
            hasil.setdefault(k, v)
        return hasil

    def isi_ringkasan(self, r, n_izin, n_kbli):
        item = r["item"]
        used = set()

        def tulis(t):
            def grup(judul):
                t.insert(tk.END, judul + "\n", "grup")

            def baris(label, nilai):
                t.insert(tk.END, label + "\t", "lbl")
                t.insert(tk.END, nilai + "\n", "val")

            t.insert(tk.END, r["nama"] + "\n", "nama")
            sumber = nilai_gabung(item, "nama_penyedia", used)
            if sumber and sumber != "-" and sumber != r["nama"]:
                baris("Nama di sumber data", sumber)

            grup("Kelengkapan Data")
            baris("Izin Usaha", f"{n_izin} izin  (lihat tab Izin Usaha)" if n_izin else "-")
            baris("KBLI", f"{n_kbli} klasifikasi  (lihat tab KBLI)" if n_kbli else "-")

            grup("Kode & NPWP")
            baris("NPWP 16", format_npwp16(r["n16"]))
            baris("NPWP 15", format_npwp15(r["n15"]))
            for k in item:
                if k.startswith("npwp"):
                    used.add(k)
            for key, label in KODE_ROWS:
                baris(label, tampil_nilai(item.get(key)))
                used.add(key)

            for judul, daftar in GRUP_RINGKAS:
                isi = []
                for base, label in daftar:
                    v = nilai_gabung(item, base, used)
                    isi.append((label, "-" if v is None else v))
                grup(judul)
                for label, v in isi:
                    baris(label, v)

            lain = []
            for k in sorted(set(self.skema) | set(item)):
                if k in used or k.startswith(("kbli", "izin_usaha", "npwp", "nama_penyedia", "nama_tampil")):
                    continue
                lain.append((k.replace("_", " ").title(), tampil_nilai(item.get(k))))
            if lain:
                grup("Lainnya")
                for label, v in lain:
                    baris(label, v)

        self.tulis_ringkas(tulis)

    # ------------------------------------------------------------------
    # MODE TAMPILAN
    # ------------------------------------------------------------------
    def reset_tab_detail(self, pesan):
        self.isi_pesan(pesan)
        self.isi_tab("izin", [])
        self.isi_tab("kbli", [])
        self.nb.tab(self.t_izin, text="Izin Usaha")
        self.nb.tab(self.t_kbli, text="KBLI")

    def mode_multi(self, judul, sub, teks, warnai=True):
        self._mode_multi = True
        self.lbl_judul.configure(text=judul)
        self.lbl_sub.configure(text=sub)
        self.reset_tab_detail("Mode banyak blok data. Lihat tab JSON.\n"
                              "Pilih satu penyedia di tabel kiri untuk melihat ringkasan, izin usaha dan KBLI.")
        self.tampil_teks(teks, warnai=warnai, fold=8 if warnai else None)
        self.nb.select(self.t_json)

    def tampilkan_semua_data(self):
        self.batal_job_filter()
        self.ent_filter_json.delete(0, tk.END)
        self.ent_filter_internal.delete(0, tk.END)

        if not self.rows:
            self.mode_multi("Semua Data", "Tidak ada data", "Tidak ada data untuk ditampilkan.", warnai=False)
            self.update_info_json("0 blok data")
            return

        data = [self.lengkapi(r["item"]) for r in self.rows]
        self.mode_multi("Semua Data", "Seluruh blok data penyedia",
                        json.dumps(data, indent=4, ensure_ascii=False, sort_keys=True))
        self.update_info_json(f"Menampilkan {len(data)} blok data")

    def reset_filter(self):
        """Kosongkan semua kolom filter dan kembalikan tampilan ke kondisi awal."""
        self.batal_job_filter()
        for ent in (self.ent_filter_penyedia, self.ent_filter_internal, self.ent_filter_json):
            ent.delete(0, tk.END)
        for kunci in ("izin", "kbli"):
            self.tab[kunci]["ent"].delete(0, tk.END)

        self.tree.selection_remove(self.tree.selection())
        self.isi_tabel(range(len(self.rows)))
        if self.tree.get_children():
            self.tree.yview_moveto(0)

        self._mode_multi = False
        self.lbl_judul.configure(text="Detail Penyedia")
        self.lbl_sub.configure(text="Pilih penyedia di tabel kiri")
        self.reset_tab_detail(getattr(self, "pesan_awal", ""))
        self.tampil_teks("", warnai=False)
        self.nb.select(self.t_ringkas)
        self.update_info_json("Filter direset")

    def on_filter_json_key(self, event=None):
        self.batal_job_filter()
        self._filter_json_job = self.after(250, self.terapkan_filter_json)

    def terapkan_filter_json(self):
        self._filter_json_job = None
        keyword = self.ent_filter_json.get().strip()
        if not keyword:
            self.tampilkan_semua_data()
            return

        kw = keyword.lower()
        hasil = [self.lengkapi(r["item"]) for r in self.rows if kw in r["cari_json"]]
        self.ent_filter_internal.delete(0, tk.END)

        if hasil:
            self.mode_multi(f"Filter: {keyword}", "Blok data yang mengandung teks tersebut",
                            json.dumps(hasil, indent=4, ensure_ascii=False, sort_keys=True))
            self.update_info_json(f"Ditemukan {len(hasil)} blok dari {len(self.rows)}")
        else:
            self.mode_multi(f"Filter: {keyword}", "Tidak ada hasil",
                            f"Tidak ada blok data yang mengandung '{keyword}'.", warnai=False)
            self.update_info_json("0 blok ditemukan")

    def isi_tabel(self, daftar_index):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for urut, idx in enumerate(daftar_index, 1):
            r = self.rows[idx]
            self.tree.insert(
                "", "end", iid=str(idx),
                values=(urut, r["nama"], format_npwp15(r["n15"]), format_npwp16(r["n16"])),
                tags=("ganjil" if urut % 2 == 0 else "genap",),
            )
        self.lbl_hitung.configure(text=f"{self.tree.get_children().__len__()} dari {len(self.rows)}")

    def filter_tabel_penyedia(self, event=None):
        keyword = self.ent_filter_penyedia.get().lower().strip()
        kw_digit = re.sub(r"\D", "", keyword)
        if not keyword:
            self.isi_tabel(range(len(self.rows)))
            return
        hasil = []
        for idx, r in enumerate(self.rows):
            if keyword in r["nama"].lower() or keyword in r["kode"].lower():
                hasil.append(idx)
            elif kw_digit and (kw_digit in r["n15"] or kw_digit in r["n16"]):
                hasil.append(idx)
        self.isi_tabel(hasil)

    def tampilkan_detail_json(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return

        r = self.rows[int(selected[0])]
        item = r["item"]
        self.batal_job_filter()
        self.ent_filter_json.delete(0, tk.END)
        self.ent_filter_internal.delete(0, tk.END)

        izin = ambil_izin(item)
        kbli, catatan = ambil_kbli(item)

        sub = []
        if r["n16"]:
            sub.append(f"NPWP 16: {r['n16']}")
        if r["n15"]:
            sub.append(f"NPWP 15: {format_npwp15(r['n15'])}")
        self.lbl_judul.configure(text=r["nama"])
        self.lbl_sub.configure(text="   |   ".join(sub) if sub else "NPWP tidak tersedia")

        self.isi_ringkasan(r, len(izin), len(kbli))
        self.isi_tab("izin", izin)
        self.isi_tab("kbli", kbli)
        self.nb.tab(self.t_izin, text=f"Izin Usaha ({len(izin)})")
        self.nb.tab(self.t_kbli, text=f"KBLI ({len(kbli)})")
        self.tampil_teks(json.dumps(self.lengkapi(item), indent=4, ensure_ascii=False, sort_keys=True), fold=4)

        if self._mode_multi:
            self._mode_multi = False
            self.nb.select(self.t_ringkas)

        info = "Menampilkan 1 blok data"
        if catatan:
            info += f"  |  Catatan KBLI: {catatan}"
        self.update_info_json(info)

    def eksekusi_cari_internal_text(self):
        t = self.txt_viewer
        t.tag_remove("highlight_search", "1.0", tk.END)
        keyword = self.ent_filter_internal.get().strip()
        if not keyword:
            return
        self.nb.select(self.t_json)

        posisi = "1.0"
        jumlah = 0
        pertama = None
        while True:
            posisi = t.search(keyword, posisi, nocase=True, stopindex=tk.END)
            if not posisi:
                break
            akhir = f"{posisi}+{len(keyword)}c"
            if any(n.startswith("foldm_") for n in t.tag_names(posisi)):
                posisi = akhir
                continue
            jumlah += 1
            t.tag_add("highlight_search", posisi, akhir)
            # buka lipatan yang berisi kecocokan
            for i, f in enumerate(self.folds):
                if f["tutup"]:
                    rg = t.tag_ranges(f["tag"])
                    if rg and t.compare(posisi, ">=", rg[0]) and t.compare(posisi, "<=", rg[1]):
                        self.toggle_fold(i, True)
            if pertama is None:
                pertama = posisi
            posisi = akhir

        if jumlah == 0:
            messagebox.showinfo("Pencarian", f"Kata '{keyword}' tidak ditemukan di JSON ini.")
        else:
            t.see(pertama)
            self.update_info_json(f"{jumlah} kecocokan di-highlight")

    def kunci_teks(self, event):
        if (event.state & 4 and event.keysym.lower() in ("c", "a")) or \
                event.keysym in ("Up", "Down", "Left", "Right", "Prior", "Next", "Home", "End"):
            return None
        return "break"

    # ------------------------------------------------------------------
    # LOAD DATA
    # ------------------------------------------------------------------
    def load_data_awal(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.rows = []

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

            with open(file_path, "r", encoding="utf-8-sig") as f:
                data_json = json.load(f)

            if isinstance(data_json, list):
                data_list = data_json
            elif isinstance(data_json, dict):
                extracted = None
                for key in ("data", "result", "items", "records", "rows"):
                    if isinstance(data_json.get(key), list):
                        extracted = data_json[key]
                        break
                if extracted is None:
                    for v in data_json.values():
                        if isinstance(v, list):
                            extracted = v
                            break
                if extracted is None:
                    raise ValueError("JSON object tidak berisi daftar/list data penyedia.")
                data_list = extracted
            else:
                raise ValueError("Format data JSON harus berupa list/array atau object yang berisi list.")

            data_list = [row for row in data_list if isinstance(row, dict)]

        except FileNotFoundError as e:
            messagebox.showerror("File Tidak Ditemukan", str(e))
            self.isi_pesan("File data belum ditemukan.\nSilakan siapkan folder 'data_master' dan file 'kamus_penyedia.json'.")
            self.update_info_json("File tidak ditemukan")
            return
        except json.JSONDecodeError as e:
            messagebox.showerror("Error JSON", f"Format file JSON tidak valid.\n\nFile:\n{file_path}\n\nDetail:\n{e}")
            self.isi_pesan("Format file JSON tidak valid.")
            self.update_info_json("JSON error")
            return
        except Exception as e:
            messagebox.showerror("Error Sistem", f"Terjadi kesalahan saat membaca file:\n{e}")
            self.isi_pesan("Terjadi kesalahan saat membaca file.")
            self.update_info_json("Error")
            return

        self.hitung_skema(data_list)

        rows = []
        for item in data_list:
            try:
                cari_json = json.dumps(item, ensure_ascii=False).lower()
            except Exception:
                cari_json = str(item).lower()
            rows.append({
                "item": item,
                "nama": ambil_nama(item),
                "n15": _cari_npwp(item, 15),
                "n16": _cari_npwp(item, 16),
                "kode": ambil_kode(item),
                "cari_json": cari_json,
            })

        rows.sort(key=lambda r: (r["nama"] == "-", r["nama"].lower()))
        self.rows = rows

        self.isi_tabel(range(len(self.rows)))
        self.lbl_total.configure(text=f"{len(self.rows)} penyedia")

        self.pesan_awal = (
            f"Berhasil memuat {len(rows)} data penyedia.\n\n"
            "1. Klik penyedia di tabel kiri. Detailnya tampil di empat tab: Ringkasan, Izin Usaha, KBLI dan JSON.\n"
            "2. Tab Izin Usaha dan KBLI berisi tabel yang bisa difilter. Klik 2x pada baris untuk menyalin nomor/kode.\n"
            "3. Di tab JSON, bagian bersarang dilipat. Klik penanda ▶ untuk membuka.\n"
            "4. Tombol 'Semua Data' dan 'Filter Blok' menampilkan banyak blok sekaligus di tab JSON."
        )
        self.isi_pesan(self.pesan_awal)
        self.update_info_json(f"{len(rows)} blok data dimuat")


if __name__ == "__main__":
    import traceback

    print("[INFO] Aplikasi data_penyedia.py dimulai")
    try:
        app = AppPenyedia()
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