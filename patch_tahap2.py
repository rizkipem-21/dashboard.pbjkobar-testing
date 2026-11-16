# ======================================================
# patch_tahap2.py  (JALANKAN SEKALI dari folder utama proyek, lalu hapus)
#   python patch_tahap2.py
# 1) data_tepra.html     : tombol Excel memakai SheetJS dari 5 JSON
# 2) download_json.py    : buang backup_harian & import shutil (tidak terpakai)
# Aman: jika ada penanda yang tidak cocok, TIDAK ada file yang diubah.
# ======================================================
import os, re, sys, shutil

ROOT = os.path.dirname(os.path.abspath(__file__))

def cari(nama):
    hasil = []
    for d, subs, files in os.walk(ROOT):
        subs[:] = [s for s in subs if s not in ('.git', 'node_modules', 'arsip_lokal', 'arsip_json')]
        if nama in files:
            hasil.append(os.path.join(d, nama))
    return hasil

def baca(p):
    with open(p, 'r', encoding='utf-8-sig', newline='') as f:
        return f.read()

def gagal(msg):
    print("[GAGAL] " + msg + "\nTidak ada file yang diubah.")
    sys.exit(1)

html_list = cari('data_tepra.html')
dl_list = cari('download_json.py')
if len(html_list) != 1: gagal(f"data_tepra.html ditemukan {len(html_list)} kali: {html_list}")
if len(dl_list) != 1: gagal(f"download_json.py ditemukan {len(dl_list)} kali: {dl_list}")
p_html, p_dl = html_list[0], dl_list[0]
html, dl = baca(p_html), baca(p_dl)

def sekali(teks, s, nama):
    if teks.count(s) != 1:
        gagal(f"Penanda '{nama}' ditemukan {teks.count(s)} kali (harus 1).")

# ---------- 1) data_tepra.html ----------
TAG_TAILWIND = '<script src="https://cdn.tailwindcss.com"></script>'
TAG_XLSX = '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>'
AWAL_FUNGSI = 'function downloadExcelHarian() {'
AKHIR_FUNGSI = '// INISIALISASI'
for s, n in [(TAG_TAILWIND, 'tag tailwind'), (AWAL_FUNGSI, 'fungsi downloadExcelHarian'), (AKHIR_FUNGSI, 'komentar INISIALISASI')]:
    sekali(html, s, n)
if TAG_XLSX in html: gagal("SheetJS sudah terpasang (skrip sudah pernah dijalankan?).")

FUNGSI_BARU = r"""// ==========================================
// DOWNLOAD EXCEL (dibuat di browser dari 5 JSON, satu sheet per tab)
// ==========================================
const TAB_EXCEL = [
  { file: "rekap_tepra_0_50m",     sheet: "Non strategis 0 sd 50 jt",    lengkap: false },
  { file: "rekap_tepra_50_200m",   sheet: "Non strategis 50jt sd 200jt", lengkap: true },
  { file: "rekap_tepra_200m_2v5m", sheet: "Strategis > 200jt sd 2,5m",   lengkap: true },
  { file: "rekap_tepra_2v5m_50m",  sheet: "Strategis > 2,5m sd 50m",     lengkap: true },
  { file: "rekap_tepra_total",     sheet: "Total Keseluruhan Paket",     lengkap: true }
];

function kolomExcel(lengkap) {
  const k = [
    ["DATA LAPORAN", "Data Laporan_Paket", "Data Laporan_Pagu"],
    [lengkap ? "PEMILIHAN/PELAKSANAAN" : "DATA REALISASI (RUP)", "Pemilihan_Paket", "Pemilihan_Pagu"],
    [lengkap ? "HASIL PEMILIHAN" : "DATA REALISASI", "Hasil_Paket", "Hasil_Pagu"]
  ];
  if (lengkap) {
    k.push(["KONTRAK", "Kontrak_Paket", "Kontrak_Pagu"]);
    k.push(["SERAH TERIMA", "SerahTerima_Paket", "SerahTerima_Pagu"]);
  }
  k.push(["BELUM PENGADAAN", "Belum_Paket", "Belum_Pagu"]);
  return k;
}

function buatSheetExcel(data, lengkap) {
  const kol = kolomExcel(lengkap);
  const header = ["NO", "SATUAN KERJA"];
  kol.forEach(k => { header.push(k[0] + " - PAKET"); header.push(k[0] + " - Rp"); });
  header.push("PROGRES (%)");

  const aoa = [header];
  let nomor = 1;
  data.forEach(row => {
    const isTotal = row["Satuan Kerja"] === "TOTAL KESELURUHAN";
    const baris = [isTotal ? "" : nomor++, row["Satuan Kerja"]];
    kol.forEach(k => { baris.push(Number(row[k[1]]) || 0); baris.push(Number(row[k[2]]) || 0); });
    baris.push(row["Data Laporan_Paket"] ? Number(row["Progres_Persen"] || 0) / 100 : "-");
    aoa.push(baris);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const lastCol = header.length - 1;
  for (let r = 1; r < aoa.length; r++) {
    for (let c = 2; c <= lastCol; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
      if (cell && cell.t === "n") cell.z = (c === lastCol) ? "0.0%" : "#,##0";
    }
  }
  ws["!cols"] = header.map((h, i) => ({ wch: i === 0 ? 5 : (i === 1 ? 38 : (i === lastCol ? 12 : 16)) }));
  return ws;
}

async function downloadExcelHarian() {
  const btn = document.querySelector(".btn-excel");
  const originalText = btn.innerText;
  btn.innerText = "⏳ Menyiapkan...";
  btn.style.opacity = "0.7";
  btn.style.pointerEvents = "none";

  try {
    if (typeof XLSX === "undefined") throw new Error("Pustaka Excel (SheetJS) gagal dimuat. Periksa koneksi internet lalu muat ulang halaman.");

    const hasil = await Promise.all(TAB_EXCEL.map(async (t) => {
      try {
        const res = await fetch(`../data/${currentYear}/${t.file}_${currentYear}.json?t=${Date.now()}`);
        if (!res.ok) return null;
        const data = await res.json();
        return (Array.isArray(data) && data.length) ? data : null;
      } catch (e) { return null; }
    }));

    const wb = XLSX.utils.book_new();
    let jumlah = 0;
    TAB_EXCEL.forEach((t, i) => {
      if (!hasil[i]) return;
      XLSX.utils.book_append_sheet(wb, buatSheetExcel(hasil[i], t.lengkap), t.sheet);
      jumlah++;
    });
    if (!jumlah) throw new Error("Data Tepra tahun " + currentYear + " belum tersedia.");

    const today = new Date();
    const tgl = today.getFullYear() + "-" + String(today.getMonth() + 1).padStart(2, "0") + "-" + String(today.getDate()).padStart(2, "0");
    XLSX.writeFile(wb, `Tepra Kobar_Progres PBJ tahun ${currentYear} (${tgl}).xlsx`);
  } catch (e) {
    alert(e.message || "Gagal membuat file Excel.");
  } finally {
    btn.innerText = originalText;
    btn.style.opacity = "1";
    btn.style.pointerEvents = "auto";
  }
}

"""

i, j = html.index(AWAL_FUNGSI), html.index(AKHIR_FUNGSI)
if i >= j: gagal("Urutan fungsi downloadExcelHarian dan INISIALISASI tidak sesuai.")
html_baru = html[:i] + FUNGSI_BARU + html[j:]
html_baru = html_baru.replace(TAG_TAILWIND, TAG_TAILWIND + "\n" + TAG_XLSX, 1)
if "../output/tepra" in html_baru: gagal("Masih ada rujukan ../output/tepra di HTML.")

# ---------- 2) download_json.py ----------
AWAL_BK = "def backup_harian(output_path, filename):"
AKHIR_BK = "def tulis_json_aman(output_path, data):"
sekali(dl, AWAL_BK, 'def backup_harian')
sekali(dl, AKHIR_BK, 'def tulis_json_aman')
a, b = dl.index(AWAL_BK), dl.index(AKHIR_BK)
if a >= b: gagal("Urutan backup_harian dan tulis_json_aman tidak sesuai.")
dl_baru = dl[:a] + dl[b:]
dl_baru = re.sub(r'^import shutil[ \t]*\r?\n', '', dl_baru, flags=re.M)
dl_baru = re.sub(r'^# - Backup harian ke arsip_json[^\r\n]*\r?\n', '', dl_baru, flags=re.M)
dl_baru = dl_baru.replace("# FILE: BACKUP HARIAN & TULIS AMAN", "# FILE: TULIS AMAN", 1)
for sisa in ("backup_harian(", "shutil."):
    if sisa in dl_baru: gagal(f"Masih ada '{sisa}' di download_json.py setelah pemotongan.")
try:
    compile(dl_baru, 'download_json.py', 'exec')
except SyntaxError as e:
    gagal(f"download_json.py hasil akhir error sintaks (baris {e.lineno}).")

# ---------- tulis (cadangan dulu) ----------
for p, t in [(p_html, html_baru), (p_dl, dl_baru)]:
    shutil.copyfile(p, p + '.bak')
    with open(p, 'w', encoding='utf-8', newline='') as f:
        f.write(t)
    print(f"[OK] {os.path.relpath(p, ROOT)}  (cadangan: .bak)")
print("Selesai. Hapus file .bak setelah uji berhasil.")