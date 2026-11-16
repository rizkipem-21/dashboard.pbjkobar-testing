# ======================================================
# arsip_tepra.py
# Membaca 5 JSON Tepra yang sudah jadi, lalu membuat Excel berformat
# di arsip_lokal/tepra/{tahun}/ (satu file per hari, tidak dipangkas).
# ======================================================

import os
import sys
import json
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if not os.path.exists(BASE_DIR) or os.path.basename(BASE_DIR) != 'dashboard.pbjkobar-testing':
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))

tahun = datetime.now().year

# --- Log harian: logs/tepra/, simpan 30 hari. Semua print() ikut ditulis ke file ---
_LOG_DIR = os.path.join(BASE_DIR, 'logs', 'tepra')
os.makedirs(_LOG_DIR, exist_ok=True)
_LOG_FILE = os.path.join(_LOG_DIR, f"log_tepra_{datetime.now().strftime('%Y-%m-%d')}.txt")
import time as _t
for _f in os.listdir(_LOG_DIR):
    if _f.startswith('log_tepra_') and _f.endswith('.txt'):
        try:
            if os.path.getmtime(os.path.join(_LOG_DIR, _f)) < _t.time() - 30 * 86400:
                os.remove(os.path.join(_LOG_DIR, _f))
        except OSError: pass
_print_asli = print
def print(*args, **kwargs):
    _print_asli(*args, **kwargs)
    try:
        with open(_LOG_FILE, 'a', encoding='utf-8') as _fh:
            _fh.write(" ".join(str(a) for a in args) + '\n')
    except OSError: pass
print(f"\n=== {os.path.basename(__file__)} mulai {datetime.now().strftime('%H:%M:%S')} ===")


# =======================================================================================
# HELPER PERSENTASE PROGRES (Hasil Pemilihan PKT / Data Laporan PKT)
# =======================================================================================
def nilai_persen(row):
    return row['Progres_Persen'] / 100 if row['Data Laporan_Paket'] else "-"


# =======================================================================================
# FUNGSI STYLE EXCEL KHUSUS: 0 SD 50 JUTA (11 Kolom)
# =======================================================================================
def style_excel_sheet_0_50(ws, df_rekap, title_text):
    font_normal = Font(name='Arial', size=9)
    header_font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
    header_font_italic = Font(name='Arial', bold=True, italic=True, size=9, color='FFFFFF')
    header_fill = PatternFill('solid', start_color='1F4E79')

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center')
    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Judul
    ws.merge_cells('A1:K1'); ws.merge_cells('A2:K2')
    ws['A1'] = title_text
    bln = {1:'JANUARI', 2:'FEBRUARI', 3:'MARET', 4:'APRIL', 5:'MEI', 6:'JUNI', 7:'JULI', 8:'AGUSTUS', 9:'SEPTEMBER', 10:'OKTOBER', 11:'NOVEMBER', 12:'DESEMBER'}
    ws['A2'] = f"KABUPATEN KOTAWARINGIN BARAT S/D {datetime.now().day} {bln[datetime.now().month]} {datetime.now().year}"
    ws['A1'].font = ws['A2'].font = Font(name='Arial', bold=True, size=11)
    ws['A1'].alignment = ws['A2'].alignment = align_center

    # Header Tabel
    ws.merge_cells('A4:A7'); ws['A4'] = "NO"
    ws.merge_cells('B4:B7'); ws['B4'] = "SATUAN KERJA"
    ws.merge_cells('C4:D5'); ws['C4'] = "DATA LAPORAN"
    ws.merge_cells('E4:F5'); ws['E4'] = "DATA REALISASI (RUP)"
    ws.merge_cells('G4:H5'); ws['G4'] = "DATA REALISASI"
    ws.merge_cells('I4:J5'); ws['I4'] = "BELUM PENGADAAN"
    ws.merge_cells('K4:K7'); ws['K4'] = "PROGRES\n(%)"

    ws.merge_cells('C6:C7'); ws['C6'] = "JUMLAH\nPAKET"
    ws.merge_cells('D6:D7'); ws['D6'] = "JUMLAH PAGU\n(Rp)"
    ws.merge_cells('E6:E7'); ws['E6'] = "JUMLAH\nPAKET"
    ws.merge_cells('F6:F7'); ws['F6'] = "JUMLAH PAGU\n(Rp)"
    ws.merge_cells('G6:G7'); ws['G6'] = "JUMLAH\nPAKET"
    ws.merge_cells('H6:H7'); ws['H6'] = "JUMLAH PAGU\n(Rp)"
    ws.merge_cells('I6:I7'); ws['I6'] = "JUMLAH\nPAKET"
    ws.merge_cells('J6:J7'); ws['J6'] = "JUMLAH PAGU\n(Rp)"
    ws.row_dimensions[6].height = 30

    numbers = ['1','2','3','4','5','6','7','8','9','10','11=7/3']
    for col_idx, val in enumerate(numbers, start=1): ws.cell(row=8, column=col_idx, value=val)

    for r in range(4, 9):
        for c in range(1, 12):
            cell = ws.cell(row=r, column=c)
            cell.font = header_font_italic if r == 8 else header_font
            cell.fill, cell.alignment, cell.border = header_fill, align_center, border_thin

    if df_rekap.empty: return

    # Isi Data
    row_num = 9
    for idx, row in df_rekap.iterrows():
        is_total = row['Satuan Kerja'] == 'TOTAL KESELURUHAN'
        data_row = [
            "" if is_total else row_num - 8, row['Satuan Kerja'],
            row['Data Laporan_Paket'] if row['Data Laporan_Paket'] else "-", row['Data Laporan_Pagu'] if row['Data Laporan_Pagu'] else "-",
            row['Pemilihan_Paket'] if row['Pemilihan_Paket'] else "-", row['Pemilihan_Pagu'] if row['Pemilihan_Pagu'] else "-",
            row['Hasil_Paket'] if row['Hasil_Paket'] else "-", row['Hasil_Pagu'] if row['Hasil_Pagu'] else "-",
            row['Belum_Paket'] if row['Belum_Paket'] else "-", row['Belum_Pagu'] if row['Belum_Pagu'] else "-",
            nilai_persen(row)
        ]

        for col_idx, val in enumerate(data_row, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            if is_total:
                cell.font, cell.fill = header_font, header_fill
            else:
                cell.font = font_normal
                cell.fill = PatternFill('solid', start_color='FFFFFF') if row_num % 2 != 0 else PatternFill('solid', start_color='DCE6F1')

            cell.border = border_thin
            if col_idx == 2: cell.alignment = align_center if is_total else align_left
            elif col_idx == 11 and isinstance(val, (int, float)):
                cell.number_format, cell.alignment = '0.0%', align_center
            elif col_idx in [4, 6, 8, 10] and isinstance(val, (int, float)):
                cell.number_format, cell.alignment = '#,##0', align_right
            else:
                cell.alignment = align_center
                if isinstance(val, (int, float)): cell.number_format = '#,##0'

        ws.row_dimensions[row_num].height = 45
        row_num += 1

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 35
    for col in ['C', 'E', 'G', 'I']: ws.column_dimensions[col].width = 8
    for col in ['D', 'F', 'H', 'J']: ws.column_dimensions[col].width = 16
    ws.column_dimensions['K'].width = 10


# =======================================================================================
# FUNGSI STYLE EXCEL STANDAR: (15 Kolom)
# =======================================================================================
def style_excel_sheet(ws, df_rekap, title_text):
    font_normal = Font(name='Arial', size=9)
    header_font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
    header_font_italic = Font(name='Arial', bold=True, italic=True, size=9, color='FFFFFF')
    header_fill = PatternFill('solid', start_color='1F4E79')

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center')
    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Judul
    ws.merge_cells('A1:O1'); ws.merge_cells('A2:O2')
    ws['A1'] = title_text
    bln = {1:'JANUARI', 2:'FEBRUARI', 3:'MARET', 4:'APRIL', 5:'MEI', 6:'JUNI', 7:'JULI', 8:'AGUSTUS', 9:'SEPTEMBER', 10:'OKTOBER', 11:'NOVEMBER', 12:'DESEMBER'}
    ws['A2'] = f"KABUPATEN KOTAWARINGIN BARAT S/D {datetime.now().day} {bln[datetime.now().month]} {datetime.now().year}"
    ws['A1'].font = ws['A2'].font = Font(name='Arial', bold=True, size=11)
    ws['A1'].alignment = ws['A2'].alignment = align_center

    # Header Tabel
    ws.merge_cells('A4:A7'); ws['A4'] = "NO"
    ws.merge_cells('B4:B7'); ws['B4'] = "SATUAN KERJA"
    ws.merge_cells('C4:D4'); ws['C4'] = "DATA LAPORAN"
    ws.merge_cells('E4:L4'); ws['E4'] = "PROSES PENGADAAN"
    ws.merge_cells('M4:N4'); ws['M4'] = "BELUM PENGADAAN"
    ws.merge_cells('O4:O7'); ws['O4'] = "PROGRES\n(%)"
    ws.merge_cells('C5:C7'); ws['C5'] = "JUMLAH\nPAKET"
    ws.merge_cells('D5:D7'); ws['D5'] = "JUMLAH PAGU\n(Rp)"
    ws.merge_cells('E5:L5'); ws['E5'] = "SUDAH PENGADAAN"
    ws.merge_cells('M5:M7'); ws['M5'] = "PAKET"
    ws.merge_cells('N5:N7'); ws['N5'] = "Rp"
    ws.merge_cells('E6:F6'); ws['E6'] = "PEMILIHAN/\nPELAKSANAAN"
    ws.merge_cells('G6:H6'); ws['G6'] = "HASIL PEMILIHAN"
    ws.merge_cells('I6:J6'); ws['I6'] = "KONTRAK"
    ws.merge_cells('K6:L6'); ws['K6'] = "SERAH TERIMA"
    ws.row_dimensions[6].height = 30

    headers_row7 = ['PAKET', 'Rp', 'PAKET', 'Rp', 'PAKET', 'Rp', 'PAKET', 'Rp']
    for col_idx, val in enumerate(headers_row7, start=5): ws.cell(row=7, column=col_idx, value=val)

    numbers = ['1','2','3','4','5','6','7','8','9','10','11','12','13=3-5','14=4-6','15=7/3']
    for col_idx, val in enumerate(numbers, start=1): ws.cell(row=8, column=col_idx, value=val)

    for r in range(4, 9):
        for c in range(1, 16):
            cell = ws.cell(row=r, column=c)
            cell.font = header_font_italic if r == 8 else header_font
            cell.fill, cell.alignment, cell.border = header_fill, align_center, border_thin

    if df_rekap.empty: return

    # Isi Data
    row_num = 9
    for idx, row in df_rekap.iterrows():
        is_total = row['Satuan Kerja'] == 'TOTAL KESELURUHAN'
        data_row = [
            "" if is_total else row_num - 8, row['Satuan Kerja'],
            row['Data Laporan_Paket'] if row['Data Laporan_Paket'] else "-", row['Data Laporan_Pagu'] if row['Data Laporan_Pagu'] else "-",
            row['Pemilihan_Paket'] if row['Pemilihan_Paket'] else "-", row['Pemilihan_Pagu'] if row['Pemilihan_Pagu'] else "-",
            row['Hasil_Paket'] if row['Hasil_Paket'] else "-", row['Hasil_Pagu'] if row['Hasil_Pagu'] else "-",
            row['Kontrak_Paket'] if row['Kontrak_Paket'] else "-", row['Kontrak_Pagu'] if row['Kontrak_Pagu'] else "-",
            row['SerahTerima_Paket'] if row['SerahTerima_Paket'] else "-", row['SerahTerima_Pagu'] if row['SerahTerima_Pagu'] else "-",
            row['Belum_Paket'] if row['Belum_Paket'] else "-", row['Belum_Pagu'] if row['Belum_Pagu'] else "-",
            nilai_persen(row)
        ]

        for col_idx, val in enumerate(data_row, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            if is_total:
                cell.font, cell.fill = header_font, header_fill
            else:
                cell.font = font_normal
                cell.fill = PatternFill('solid', start_color='FFFFFF') if row_num % 2 != 0 else PatternFill('solid', start_color='DCE6F1')

            cell.border = border_thin
            if col_idx == 2: cell.alignment = align_center if is_total else align_left
            elif col_idx == 15 and isinstance(val, (int, float)):
                cell.number_format, cell.alignment = '0.0%', align_center
            elif col_idx in [4, 6, 8, 10, 12, 14] and isinstance(val, (int, float)):
                cell.number_format, cell.alignment = '#,##0', align_right
            else:
                cell.alignment = align_center
                if isinstance(val, (int, float)): cell.number_format = '#,##0'

        ws.row_dimensions[row_num].height = 45
        row_num += 1

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 35
    for col in ['C', 'E', 'G', 'I', 'K', 'M']: ws.column_dimensions[col].width = 8
    for col in ['D', 'F', 'H', 'J', 'L', 'N']: ws.column_dimensions[col].width = 16
    ws.column_dimensions['O'].width = 10


# =======================================================================================
# PROSES UTAMA
# =======================================================================================
T = "PROSES PENGADAAN BARANG DAN JASA "
SHEETS = [
    ('rekap_tepra_0_50m',     "Non strategis 0 sd 50 jt",    style_excel_sheet_0_50, T + "PAKET NON STRATEGIS (0 S/D 50 JUTA)"),
    ('rekap_tepra_50_200m',   "Non strategis 50jt sd 200jt", style_excel_sheet,      T + "PAKET NON STRATEGIS (>50 JUTA - ≤200 JUTA)"),
    ('rekap_tepra_200m_2v5m', "Strategis > 200jt sd 2,5m",   style_excel_sheet,      T + "PAKET STRATEGIS (> Rp200 JUTA S/D 2,5 M)"),
    ('rekap_tepra_2v5m_50m',  "Strategis > 2,5m sd 50m",     style_excel_sheet,      T + "PAKET STRATEGIS (≥Rp2,5M sd 50M)"),
    ('rekap_tepra_total',     "Total Keseluruhan Paket",     style_excel_sheet,      T + "TOTAL KESELURUHAN PAKET (NON STRATEGIS DAN STRATEGIS)"),
]


def buat_arsip():
    wb = Workbook()
    for i, (nama, judul_sheet, fungsi, judul) in enumerate(SHEETS):
        path_json = os.path.join(BASE_DIR, 'data', str(tahun), f'{nama}_{tahun}.json')
        if not os.path.exists(path_json):
            print(f"[!] JSON tidak ada: {path_json}")
            return False
        with open(path_json, 'r', encoding='utf-8') as f:
            df = pd.DataFrame(json.load(f))
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = judul_sheet
        fungsi(ws, df, judul)

    tgl = datetime.now().strftime("%Y-%m-%d")
    folder = os.path.join(BASE_DIR, 'arsip_lokal', 'tepra', str(tahun))
    os.makedirs(folder, exist_ok=True)
    nama_file = f'Tepra Kobar_Progres PBJ tahun {tahun} ({tgl}).xlsx'

    try:
        wb.save(os.path.join(folder, nama_file))
    except PermissionError:
        print(f"[!] File {nama_file} sedang terbuka di Excel. Tutup lalu jalankan ulang.")
        return False

    print(f"Arsip Tepra tersimpan: {nama_file}")
    return True


if __name__ == '__main__':
    try:
        sukses = buat_arsip()
    except Exception as e:
        print(f"[!] Arsip Tepra error: {e}")
        sukses = False
    sys.exit(0 if sukses else 1)