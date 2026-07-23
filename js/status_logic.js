// File: js/status_logic.js
// Fungsi ini digunakan secara global oleh Dashboard Utama dan Dashboard Pengadaan

function getKategoriStatus(row) {
  const sumber = (row["Sumber"] || "").toLowerCase();
  const s = (row["Status"] || "").toLowerCase();

  if (sumber.includes("sumber 1")) return "Belum Proses";
  
  // LOGIKA UNTUK SUMBER 2 (NON TENDER) & SUMBER 5 (TENDER)
  if (sumber.includes("sumber 2") || sumber.includes("sumber 5")) {
    // HANYA BAPBAST yang dianggap Sudah Selesai. 
    if (s.includes("bapbast")) return "Sudah Selesai";
    return "Sedang Berjalan"; 
  }
  
  if (sumber.includes("sumber 3") || sumber.includes("sumber 4")) {
    if (s.includes("paket selesai") || s.includes("selesai")) return "Sudah Selesai";
    return "Sedang Berjalan";
  }
  
  if (sumber.includes("sumber 6")) {
    if (s.includes("payment outside") || s.includes("completed")) return "Sudah Selesai";
    return "Sedang Berjalan";
  }
  
  if (sumber.includes("sumber 7")) {
    if (s.includes("paket selesai") || s.includes("selesai")) return "Sudah Selesai";
    return "Sedang Berjalan";
  }

  // Fallback pengaman untuk status secara umum
  if (s.includes("bapbast") || s.includes("payment") || s.includes("completed")) return "Sudah Selesai";
  if (s.includes("pengumuman rup") || s === "" || s === "-") return "Belum Proses";
  
  return "Sedang Berjalan";
}