# ======================================================
# generate_tepra.py
# Membuat 5 file JSON rekap Tepra (untuk web). Tanpa Excel.
# Excel berformat dibuat terpisah oleh arsip_tepra.py (lokal).
# ======================================================

import os
import json
import pandas as pd
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if not os.path.exists(BASE_DIR) or os.path.basename(BASE_DIR) != 'dashboard.pbjkobar-testing':
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))

sys.path.append(BASE_DIR)

tahun_n  = datetime.now().year

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

SET_PEMILIHAN = {'bapbast', 'berlangsung', 'completed', 'kontrak', 'non tender selesai', 'on process', 'paket sedang berjalan', 'paket selesai', 'payment outside system', 'spmkspp', 'sppbj'}
SET_KONTRAK = {'bapbast', 'completed', 'kontrak', 'on process', 'paket sedang berjalan', 'paket selesai', 'payment outside system', 'spmkspp'}
SET_SERAH_TERIMA = {'bapbast', 'completed', 'paket selesai', 'payment outside system'}


def safe_float(val):
    try: return float(str(val).replace(',', ''))
    except: return 0.0


def get_data_rekap(df_source, min_pagu, max_pagu):
    # Ambil daftar seluruh Satuan Kerja dari sumber utama agar tidak ada yang terlewat
    semua_satker = sorted(df_source['Satuan Kerja'].astype(str).str.strip().unique())

    # Filter data sesuai rentang Pagu
    df = df_source[(df_source['Nilai Pagu RUP'] > min_pagu) & (df_source['Nilai Pagu RUP'] <= max_pagu)].copy()

    if not df.empty:
        df['Status_Clean'] = df['Status'].astype(str).str.strip().str.lower()
        df['Satuan Kerja'] = df['Satuan Kerja'].astype(str).str.strip()
        grouped = df.groupby('Satuan Kerja')
    else:
        grouped = None

    rekap = []
    for satker in semua_satker:
        if grouped is not None and satker in grouped.groups:
            group = grouped.get_group(satker)
            mask_pemilihan = group['Status_Clean'].isin(SET_PEMILIHAN)
            paket_pemilihan = mask_pemilihan.sum()
            pagu_pemilihan = group.loc[mask_pemilihan, 'Nilai Pagu RUP'].sum()
            hasil_pemilihan = group.loc[mask_pemilihan, 'Nilai Hasil Pemilihan'].sum()

            mask_kontrak = group['Status_Clean'].isin(SET_KONTRAK)
            paket_kontrak = mask_kontrak.sum()
            pagu_kontrak = group.loc[mask_kontrak, 'Nilai Hasil Pemilihan'].sum()

            mask_serah_terima = group['Status_Clean'].isin(SET_SERAH_TERIMA)
            paket_serah_terima = mask_serah_terima.sum()
            pagu_serah_terima = group.loc[mask_serah_terima, 'Nilai Hasil Pemilihan'].sum()

            mask_belum = ~mask_pemilihan
            paket_belum = mask_belum.sum()
            pagu_belum = group.loc[mask_belum, 'Nilai Pagu RUP'].sum()
        else:
            paket_pemilihan = pagu_pemilihan = hasil_pemilihan = 0
            paket_kontrak = pagu_kontrak = 0
            paket_serah_terima = pagu_serah_terima = 0
            paket_belum = pagu_belum = 0

        rekap.append({
            'Satuan Kerja': satker,
            'Data Laporan_Paket': int(paket_pemilihan + paket_belum),
            'Data Laporan_Pagu': pagu_pemilihan + pagu_belum,
            'Pemilihan_Paket': int(paket_pemilihan),
            'Pemilihan_Pagu': pagu_pemilihan,
            'Hasil_Paket': int(paket_pemilihan),
            'Hasil_Pagu': hasil_pemilihan,
            'Kontrak_Paket': int(paket_kontrak),
            'Kontrak_Pagu': pagu_kontrak,
            'SerahTerima_Paket': int(paket_serah_terima),
            'SerahTerima_Pagu': pagu_serah_terima,
            'Belum_Paket': int(paket_belum),
            'Belum_Pagu': pagu_belum
        })

    df_rekap = pd.DataFrame(rekap).sort_values('Satuan Kerja')
    total_row = {'Satuan Kerja': 'TOTAL KESELURUHAN'}
    for col in df_rekap.columns:
        if col != 'Satuan Kerja':
            total_row[col] = df_rekap[col].sum()

    df_out = pd.concat([df_rekap, pd.DataFrame([total_row])], ignore_index=True)
    # Persentase dihitung SETELAH baris TOTAL dibuat, agar total = persentase tertimbang
    df_out['Progres_Persen'] = (df_out['Hasil_Paket'] / df_out['Data Laporan_Paket'].replace(0, float('nan')) * 100).round(1).fillna(0)
    return df_out


def process_tepra(tahun):
    input_json = os.path.join(BASE_DIR, 'data', str(tahun), f'rekap_pengadaan_{tahun}.json')
    if not os.path.exists(input_json):
        print(f"[!] JSON rekap tidak ada: {input_json}")
        return False
    try:
        with open(input_json, 'r', encoding='utf-8') as f: data = json.load(f)
    except Exception as e:
        print(f"[!] Gagal baca JSON rekap: {e}")
        return False

    df = pd.DataFrame(data)
    if df.empty:
        print("[!] JSON rekap kosong")
        return False
    df['Nilai Pagu RUP'] = df['Nilai Pagu RUP'].apply(safe_float)
    df['Nilai Hasil Pemilihan'] = df['Nilai Hasil Pemilihan'].apply(safe_float)
    df['Satuan Kerja'] = df['Satuan Kerja'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip().str.upper()

    # Batas bawah -1 agar paket bernilai 0 ikut terbawa
    hasil = {
        f'rekap_tepra_0_50m_{tahun}.json':     get_data_rekap(df, -1, 50000000),
        f'rekap_tepra_50_200m_{tahun}.json':   get_data_rekap(df, 50000000, 200000000),
        f'rekap_tepra_200m_2v5m_{tahun}.json': get_data_rekap(df, 200000000, 2500000000),
        f'rekap_tepra_2v5m_50m_{tahun}.json':  get_data_rekap(df, 2500000000, 50000000000),
        f'rekap_tepra_total_{tahun}.json':     get_data_rekap(df, -1, 50000000000),
    }

    out_dir_data = os.path.join(BASE_DIR, 'data', str(tahun))
    os.makedirs(out_dir_data, exist_ok=True)
    for nama_file, df_hasil in hasil.items():
        with open(os.path.join(out_dir_data, nama_file), "w", encoding="utf-8") as f:
            json.dump(df_hasil.to_dict(orient='records'), f, ensure_ascii=False, indent=2)

    print(f"TEPRA {tahun} Sukses (5 JSON)")
    return True


def kirim_telegram(pesan):
    try:
        import requests, config_rahasia
        requests.post(f"https://api.telegram.org/bot{config_rahasia.BOT_TOKEN}/sendMessage",
                      data={"chat_id": config_rahasia.CHAT_ID, "text": pesan}, timeout=30)
    except Exception as e:
        print(f"[!] Telegram gagal: {e}")


if __name__ == '__main__':
    try:
        sukses = process_tepra(tahun_n)
    except Exception as e:
        print(f"[!] Tepra error: {e}")
        sukses = False
    kirim_telegram(f"TEPRA {tahun_n}: " + ("sukses, 5 JSON diperbarui" if sukses else "GAGAL, cek folder logs/tepra"))
    sys.exit(0 if sukses else 1)