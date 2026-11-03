// ============================================================================
// JUDUL  : status_logic.js
// FUNGSI : Menentukan KATEGORI status paket (Belum Proses / Sedang Berjalan /
//          Sudah Selesai) dari kolom "Sumber" dan "Status".
// DIPAKAI: Dashboard Utama dan Dashboard Pengadaan (fungsi global getKategoriStatus)
// CATATAN: Logika kategori SAMA dengan versi sebelumnya (angka dashboard tidak berubah).
//          Perbedaan: aman terhadap data kosong, dan baris "Tidak Cocok" kini
//          ditangani eksplisit lewat satu pengaturan di bawah.
// ============================================================================

// Baris dengan Status "Tidak Cocok" (Kode Paket tidak ditemukan di tabel sumber).
// Saat ini masuk "Sedang Berjalan" (sama seperti perilaku sebelumnya).
// Ubah ke "Belum Proses" jika ingin sama dengan perlakuan TEPRA.
const KATEGORI_TIDAK_COCOK = "Sedang Berjalan";

function getKategoriStatus(row) {
  const sumber = String((row && row["Sumber"]) ?? "").toLowerCase();
  const s      = String((row && row["Status"]) ?? "").toLowerCase();

  // Pengaman khusus untuk paket konsolidasi yang masih menyisakan pagu
  if (s.includes("sisa pagu")) return "Sedang Berjalan";

  // Sumber 1 & 1_2 = Pengumuman RUP (belum realisasi)
  if (sumber.includes("sumber 1")) return "Belum Proses";

  // Paket yang Kode Paket-nya tidak ditemukan di tabel sumber (kolom Sumber kosong)
  if (s.includes("tidak cocok")) return KATEGORI_TIDAK_COCOK;

  // Sumber 2 (Non Tender) & Sumber 5 (Tender)
  // Kinerja Dinilai dan BAPBAST dianggap Sudah Selesai.
  if (sumber.includes("sumber 2") || sumber.includes("sumber 5")) {
    if (s.includes("kinerja dinilai") || s.includes("bapbast")) return "Sudah Selesai";
    return "Sedang Berjalan";
  }

  // Sumber 3 (Pencatatan Non Tender) & Sumber 4 (Pencatatan Swakelola)
  if (sumber.includes("sumber 3") || sumber.includes("sumber 4")) {
    if (s.includes("paket selesai") || s.includes("selesai")) return "Sudah Selesai";
    return "Sedang Berjalan";
  }

  // Sumber 6 (E-Purchasing V6)
  if (sumber.includes("sumber 6")) {
    if (s.includes("payment outside") || s.includes("completed")) return "Sudah Selesai";
    return "Sedang Berjalan";
  }

  // Sumber 7 (E-Purchasing V5)
  if (sumber.includes("sumber 7")) {
    if (s.includes("paket selesai") || s.includes("selesai")) return "Sudah Selesai";
    return "Sedang Berjalan";
  }

  // Fallback pengaman untuk status secara umum
  if (s.includes("kinerja dinilai") || s.includes("bapbast") || s.includes("payment") || s.includes("completed")) return "Sudah Selesai";
  if (s.includes("pengumuman rup") || s === "" || s === "-") return "Belum Proses";

  return "Sedang Berjalan";
}