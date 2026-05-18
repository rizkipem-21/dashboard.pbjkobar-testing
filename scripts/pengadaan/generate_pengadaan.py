# ======================================================
# PAKET PENGADAAN - MULTI TAHUN (n, n-1, n-2)
# ======================================================

import pandas as pd
import json
import re
import os
import shutil
from datetime import datetime
import warnings
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ======================================================
# KONFIGURASI TAHUN OTOMATIS
# ======================================================
BASE_DIR   = r'D:\rup-2026-inaproc'
tahun_n    = datetime.now().year       # tahun berjalan
tahun_n1   = tahun_n - 1
tahun_n2   = tahun_n - 2
daftar_tahun = [tahun_n2, tahun_n1, tahun_n]  # urut dari lama ke baru

def format_tgl(val):
    """Konversi 2026-04-22T00:00:00.000000Z atau 2026-04-22 menjadi 22-04-2026"""
    if not val or (not isinstance(val, str) and pd.isna(val)):
        return ""
    try:
        # Ambil hanya bagian tanggal yyyy-mm-dd
        tgl = str(val).strip()[:10]
        parts = tgl.split('-')
        if len(parts) == 3:
            return f"{parts[2]}-{parts[1]}-{parts[0]}"
        return ""
    except:
        return ""

# ======================================================
# FUNGSI UTAMA PER TAHUN
# ======================================================
def process_tahun(tahun):

    # Cek skip untuk tahun n-2
    data_dir    = os.path.join(BASE_DIR, 'data', str(tahun))
    output_json = os.path.join(data_dir, f'rekap_pengadaan_{tahun}.json')
    if tahun == tahun_n2 and os.path.exists(output_json):
        print(f'\n[SKIP] Tahun {tahun} sudah final -> skip generate')
        return

    print(f'\n{"="*55}')
    print(f'  GENERATE TAHUN {tahun}')
    print(f'{"="*55}')

    # ======================================================
    # PATH SUMBER DATA
    # ======================================================
    def p(nama): return os.path.join(data_dir, f'Legacy_{nama}_{tahun}.json')
    sumber1     = p('rup_paket-penyedia-terumumkan')
    sumber1_2   = p('rup_paket-swakelola-terumumkan')
    sumber2     = p('tender_non-tender-pengumuman')
    sumber2_1   = p('tender_non-tender-selesai')
    sumber2_2   = p('tender_non-tender-ekontrak-sppbj')
    sumber2_3   = p('tender_non-tender-ekontrak-kontrak')
    sumber2_4   = p('tender_non-tender-ekontrak-spmkspp')
    sumber2_5   = p('tender_non-tender-ekontrak-bapbast')
    sumber3     = p('tender_pencatatan-non-tender')
    sumber4     = p('tender_pencatatan-swakelola')
    sumber5     = p('tender_pengumuman')
    sumber5_1   = p('tender_tender-selesai')
    sumber5_1_1 = p('tender_tender-selesai-nilai')
    sumber5_2   = p('tender_tender-ekontrak-sppbj')
    sumber5_3   = p('tender_tender-ekontrak-kontrak')
    sumber5_4   = p('tender_tender-ekontrak-spmkspp')
    sumber5_5   = p('tender_tender-ekontrak-bapbast')
    sumber6     = p('ekatalog_paket-e-purchasing')
    sumber7     = p('ekatalog-archive_paket-e-purchasing')

    # ======================================================
    # LOAD MASTER KAMUS PENYEDIA (OFFLINE)
    # ======================================================
    path_kamus = os.path.join(BASE_DIR, 'data', 'master', 'kamus_penyedia.json')
    map_offline_penyedia = {}
    if os.path.exists(path_kamus):
        try:
            with open(path_kamus, 'r', encoding='utf-8') as f:
                kamus_list = json.load(f)
                if isinstance(kamus_list, list):
                    for item in kamus_list:
                        nama = item.get('nama_penyedia', "")
                        # Map kode_penyedia (S6) dan kd_penyedia (S7) ke nama
                        if item.get('kode_penyedia'):
                            map_offline_penyedia[str(item['kode_penyedia'])] = nama
                        if item.get('kd_penyedia'):
                            map_offline_penyedia[str(item['kd_penyedia'])] = nama
        except Exception as e:
            print(f"Gagal load kamus penyedia: {e}")

    # ======================================================
    # LOAD JSON
    # ======================================================
    def load_json(path):
        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                data = json.load(f)
                if isinstance(data, list):
                    return pd.json_normalize(data)
                if isinstance(data, dict):
                    for k in ['data','items','results']:
                        if k in data and isinstance(data[k], list):
                            return pd.json_normalize(data[k])
                return pd.json_normalize(data)
        except:
            return pd.DataFrame()

    print(f"Load semua sumber tahun {tahun}...")

    df1     = load_json(sumber1)
    df1_2   = load_json(sumber1_2)
    df2     = load_json(sumber2)
    df2_1   = load_json(sumber2_1)
    df2_2   = load_json(sumber2_2)
    df2_3   = load_json(sumber2_3)
    df2_4   = load_json(sumber2_4)
    df2_5   = load_json(sumber2_5)
    df3     = load_json(sumber3)
    df4     = load_json(sumber4)
    df5     = load_json(sumber5)
    df5_1   = load_json(sumber5_1)
    df5_1_1 = load_json(sumber5_1_1)
    df5_2   = load_json(sumber5_2)
    df5_3   = load_json(sumber5_3)
    df5_4   = load_json(sumber5_4)
    df5_5   = load_json(sumber5_5)
    df6     = load_json(sumber6)
    df7     = load_json(sumber7)

    # ======================================================
    # FILTER
    # ======================================================
    if not df2.empty and 'status_nontender' in df2.columns:
        df2 = df2[df2['status_nontender'] != 'Gagal/Batal']

    if not df5.empty and 'status_tender' in df5.columns:
        df5 = df5[df5['status_tender'] != 'Gagal/Batal']

    # ======================================================
    # SET KD
    # ======================================================
    def get_set(df, col):
        if df.empty or col not in df.columns:
            return set()
        return set(df[col].astype(str).str.split(';').explode().str.strip())

    set_selesai   = get_set(df2_1,'kd_nontender')
    set_sppbj     = get_set(df2_2,'kd_nontender')
    set_kontrak   = get_set(df2_3,'kd_nontender')
    set_spmkspp   = get_set(df2_4,'kd_nontender')
    set_bapbast   = get_set(df2_5,'kd_nontender')

    set_t_selesai = get_set(df5_1,'kd_tender')
    set_t_sppbj   = get_set(df5_2,'kd_tender')
    set_t_kontrak = get_set(df5_3,'kd_tender')
    set_t_spmkspp = get_set(df5_4,'kd_tender')
    set_t_bapbast = get_set(df5_5,'kd_tender')

    # ======================================================
    # MAP NILAI KONTRAK (FIX MULTI KD)
    # ======================================================
    map_nt_kontrak = {}
    if not df2_1.empty and 'nilai_kontrak' in df2_1.columns:
        for _, r in df2_1.iterrows():
            kd_list = str(r.get('kd_nontender')).split(';')
            for k in kd_list:
                k = k.strip()
                if k:
                    map_nt_kontrak[k] = r.get('nilai_kontrak')

    map_nt_pdn = {}
    if not df2_1.empty and 'nilai_pdn_kontrak' in df2_1.columns:
        for _, r in df2_1.iterrows():
            kd_list = str(r.get('kd_nontender')).split(';')
            for k in kd_list:
                k = k.strip()
                if k:
                    map_nt_pdn[k] = r.get('nilai_pdn_kontrak')

    map_nt_umk = {}
    if not df2_1.empty and 'nilai_umk_kontrak' in df2_1.columns:
        for _, r in df2_1.iterrows():
            kd_list = str(r.get('kd_nontender')).split(';')
            for k in kd_list:
                k = k.strip()
                if k:
                    map_nt_umk[k] = r.get('nilai_umk_kontrak')

    map_t_kontrak = {}
    if not df5_1_1.empty and 'nilai_kontrak' in df5_1_1.columns:
        for _, r in df5_1_1.iterrows():
            kd_list = str(r.get('kd_tender')).split(';')
            for k in kd_list:
                k = k.strip()
                if k:
                    map_t_kontrak[k] = r.get('nilai_kontrak')

    map_t_pdn = {}
    if not df5_3.empty and 'nilai_pdn_kontrak' in df5_3.columns:
        for _, r in df5_3.iterrows():
            kd_list = str(r.get('kd_tender')).split(';')
            for k in kd_list:
                k = k.strip()
                if k:
                    map_t_pdn[k] = r.get('nilai_pdn_kontrak')

    map_t_umk = {}
    if not df5_3.empty and 'nilai_umk_kontrak' in df5_3.columns:
        for _, r in df5_3.iterrows():
            kd_list = str(r.get('kd_tender')).split(';')
            for k in kd_list:
                k = k.strip()
                if k:
                    map_t_umk[k] = r.get('nilai_umk_kontrak')

    # ======================================================
    # MAP TANGGAL KONTRAK
    # ======================================================
    map_nt_tgl_kontrak = {}
    if not df2_3.empty and 'tgl_kontrak' in df2_3.columns:
        for _, r in df2_3.iterrows():
            kd_list = str(r.get('kd_nontender')).split(';')
            for k in kd_list:
                k = k.strip()
                if k:
                    map_nt_tgl_kontrak[k] = r.get('tgl_kontrak')

    map_t_tgl_kontrak = {}
    if not df5_3.empty and 'tgl_kontrak' in df5_3.columns:
        for _, r in df5_3.iterrows():
            kd_list = str(r.get('kd_tender')).split(';')
            for k in kd_list:
                k = k.strip()
                if k:
                    map_t_tgl_kontrak[k] = r.get('tgl_kontrak')

    # ======================================================
    # MAP NAMA PENYEDIA (SUMBER 2 & 5 DARI JSON LOKAL)
    # ======================================================
    map_nt_penyedia = {}
    if not df2_3.empty and 'nama_penyedia' in df2_3.columns:
        for _, r in df2_3.iterrows():
            kd_list = str(r.get('kd_nontender')).split(';')
            for k in kd_list:
                k = k.strip()
                if k:
                    map_nt_penyedia[k] = r.get('nama_penyedia')

    map_t_penyedia = {}
    if not df5_3.empty and 'nama_penyedia' in df5_3.columns:
        for _, r in df5_3.iterrows():
            kd_list = str(r.get('kd_tender')).split(';')
            for k in kd_list:
                k = k.strip()
                if k:
                    map_t_penyedia[k] = r.get('nama_penyedia')

    # ======================================================
    # STANDARD KD RUP (RAW + EXPLODE)
    # ======================================================
    def split_kd_list(x):
        try:
            if pd.isna(x):
                return []
            return [int(i.strip()) for i in str(x).split(';') if i.strip().isdigit()]
        except:
            return []

    def standardize_kd_rup(df, col):
        if df.empty or col not in df.columns:
            return df
        df[col+'_raw'] = df[col]
        df[col+'_list'] = df[col].apply(split_kd_list)
        df[col] = df[col].apply(lambda x: split_kd_list(x)[0] if len(split_kd_list(x))>0 else None)
        return df

    df1   = standardize_kd_rup(df1,  'kd_rup')
    df1_2 = standardize_kd_rup(df1_2,'kd_rup')
    df2   = standardize_kd_rup(df2,  'kd_rup')
    df3   = standardize_kd_rup(df3,  'kd_rup')
    df4   = standardize_kd_rup(df4,  'kd_rup')
    df5   = standardize_kd_rup(df5,  'kd_rup')
    df6   = standardize_kd_rup(df6,  'rup_code')
    df7   = standardize_kd_rup(df7,  'kd_rup')

    # ======================================================
    # MAP PAGU SUMBER 1 & 1_2
    # ======================================================
    map_pagu_s1   = df1.set_index('kd_rup')['pagu']   if not df1.empty   else {}
    map_pagu_s1_2 = df1_2.set_index('kd_rup')['pagu'] if not df1_2.empty else {}

    def get_pagu_multi(kd_list, tipe='s1'):
        total = 0
        if not isinstance(kd_list, list):
            return None
        for k in kd_list:
            try:
                if tipe=='s1' and k in map_pagu_s1:
                    total += map_pagu_s1[k]
                if tipe=='s1_2' and k in map_pagu_s1_2:
                    total += map_pagu_s1_2[k]
            except:
                pass
        return total if total!=0 else None

    # ======================================================
    # CLEAN ILLEGAL CHAR
    # ======================================================
    def clean_illegal_chars(df):
        return df.map(lambda x: re.sub(r'[\x00-\x1F]', '', str(x)) if isinstance(x,str) else x)

    # ======================================================
    # MAPPING SUMBER 1 (DENGAN PENYAMAAN TIPE DATA INTEGER)
    # ======================================================
    if not df1.empty:
        # Bersihkan data NaN lalu convert ke integer utuh
        df1_clean = df1.dropna(subset=['kd_rup']).copy()
        df1_clean['kd_rup'] = df1_clean['kd_rup'].apply(lambda x: int(float(str(x).strip())))
        df1_map = df1_clean.set_index('kd_rup')
    else:
        df1_map = pd.DataFrame()

    def get_s1(kd, col):
        try:
            if pd.isna(kd) or df1_map.empty:
                return None
            # Samakan tipe data input ke integer
            kd_match = int(float(str(kd).strip()))
            
            if kd_match in df1_map.index:
                val = df1_map.loc[kd_match, col]
                return None if pd.isna(val) else val
            return None
        except:
            return None

    # ======================================================
    # SUMBER 2
    # ======================================================
    data_s2=[]
    for _,r in df2.iterrows():
        kd=r.get('kd_rup')
        kd_list=r.get('kd_rup_list')
        kd_nt_list = [i.strip() for i in str(r.get('kd_nontender')).split(';')] if pd.notna(r.get('kd_nontender')) else []

        status = r.get('status_nontender')
        for k in kd_nt_list:
            if k in set_bapbast:
                status='BAPBAST'; break
            elif k in set_spmkspp:
                status='SPMKSPP'; break
            elif k in set_kontrak:
                status='Kontrak'; break
            elif k in set_sppbj:
                status='SPPBJ'; break
            elif k in set_selesai:
                status='Non Tender Selesai'; break

        nilai_hasil=None
        found=False
        for k in kd_nt_list:
            if k in map_nt_kontrak:
                nilai_hasil=map_nt_kontrak[k]; found=True; break
        if not found:
            nilai_hasil="N/A"
        elif nilai_hasil is not None and not isinstance(nilai_hasil, str) and pd.isna(nilai_hasil):
            nilai_hasil=""

        nilai_pdn="N/A"
        found=False
        for k in kd_nt_list:
            if k in map_nt_pdn:
                nilai_pdn = map_nt_pdn[k]
                if nilai_pdn is not None and not isinstance(nilai_pdn, str) and pd.isna(nilai_pdn):
                    nilai_pdn=""
                found=True
                break

        nilai_umk="N/A"
        found=False
        for k in kd_nt_list:
            if k in map_nt_umk:
                nilai_umk = map_nt_umk[k]
                if nilai_umk is not None and not isinstance(nilai_umk, str) and pd.isna(nilai_umk):
                    nilai_umk=""
                found=True
                break

        pagu = get_pagu_multi(kd_list,'s1')

        data_s2.append({
            'Kode RUP':r.get('kd_rup_raw'),
            'Satuan Kerja':r.get('nama_satker'),
            'Nama Paket':r.get('nama_paket'),
            'Metode Pengadaan':r.get('mtd_pemilihan'),
            'Jenis Pengadaan':r.get('jenis_pengadaan'),
            'Sumber Dana':r.get('sumber_dana'),
            'PDN':get_s1(kd,'status_pdn'),
            'UKM':get_s1(kd,'status_ukm'),
            'Nilai Pagu RUP':pagu,
            'Nilai Hasil Pemilihan':nilai_hasil,
            'Tanggal Kontrak':format_tgl(next((map_nt_tgl_kontrak[k] for k in kd_nt_list if k in map_nt_tgl_kontrak), "")),
            'Nama Penyedia':next((map_nt_penyedia[k] for k in kd_nt_list if k in map_nt_penyedia), ""),
            'Status':status,
            'Kode Paket':r.get('kd_nontender'),
            'Nilai HPS':r.get('hps'),
            'Nilai PDN':nilai_pdn,
            'Nilai UMK':nilai_umk,
            'Metode':'Non Tender',
            'Sumber':'Sumber 2'
        })

    df_s2=pd.DataFrame(data_s2)

    # ======================================================
    # SUMBER 3
    # ======================================================
    data_s3=[]
    for _,r in df3.iterrows():
        kd=r.get('kd_rup')
        kd_list=r.get('kd_rup_list')
        pagu = get_pagu_multi(kd_list,'s1')

        nilai_hasil = r.get('total_realisasi')
        if pd.isna(nilai_hasil):
            nilai_hasil=""

        nilai_pdn = r.get('nilai_pdn_pct')
        nilai_umk = r.get('nilai_umk_pct')

        data_s3.append({
            'Kode RUP':r.get('kd_rup_raw'),
            'Satuan Kerja':r.get('nama_satker'),
            'Nama Paket':r.get('nama_paket'),
            'Metode Pengadaan':r.get('mtd_pemilihan'),
            'Jenis Pengadaan':r.get('kategori_pengadaan'),
            'Sumber Dana':r.get('sumber_dana'),
            'PDN':get_s1(kd,'status_pdn'),
            'UKM':get_s1(kd,'status_ukm'),
            'Nilai Pagu RUP':pagu,
            'Nilai Hasil Pemilihan':nilai_hasil,
            'Tanggal Kontrak':format_tgl(r.get('tgl_selesai_paket','')),
            'Nama Penyedia':"",
            'Status':r.get('status_nontender_pct_ket'),
            'Kode Paket':r.get('kd_nontender_pct'),
            'Nilai HPS':pd.NA,
            'Nilai PDN':nilai_pdn,
            'Nilai UMK':nilai_umk,
            'Metode':'Pencatatan Non Tender',
            'Sumber':'Sumber 3'
        })

    df_s3=pd.DataFrame(data_s3)

    # ======================================================
    # SUMBER 4
    # ======================================================
    data_s4=[]
    swakelola_map = df1_2.set_index('kd_rup')['tipe_swakelola'] if not df1_2.empty else {}

    for _,r in df4.iterrows():
        kd=r.get('kd_rup')
        kd_list=r.get('kd_rup_list')

        jenis = f"Swakelola {int(swakelola_map[kd])}" if kd in swakelola_map else "N/A"
        pagu = get_pagu_multi(kd_list,'s1_2')

        nilai_hasil = r.get('total_realisasi')
        if pd.isna(nilai_hasil):
            nilai_hasil=""

        nilai_pdn = r.get('nilai_pdn_pct')
        nilai_umk = r.get('nilai_umk_pct')

        data_s4.append({
            'Kode RUP':r.get('kd_rup_raw'),
            'Satuan Kerja':r.get('nama_satker'),
            'Nama Paket':r.get('nama_paket'),
            'Metode Pengadaan':'Swakelola',
            'Jenis Pengadaan':jenis,
            'Sumber Dana':r.get('sumber_dana'),
            'PDN':"PDN" if r.get('nilai_pdn_pct',0)!=0 else "Tidak",
            'UKM':"UKM" if r.get('nilai_umk_pct',0)!=0 else "Tidak",
            'Nilai Pagu RUP':pagu,
            'Nilai Hasil Pemilihan':nilai_hasil,
            'Tanggal Kontrak':format_tgl(r.get('tgl_selesai_paket','')),
            'Nama Penyedia':"",
            'Status':r.get('status_swakelola_pct_ket'),
            'Kode Paket':r.get('kd_swakelola_pct'),
            'Nilai HPS':pd.NA,
            'Nilai PDN':nilai_pdn,
            'Nilai UMK':nilai_umk,
            'Metode':'Pencatatan Swakelola',
            'Sumber':'Sumber 4'
        })

    df_s4=pd.DataFrame(data_s4)

    # ======================================================
    # TAMBAHAN SUMBER 1_2
    # ======================================================
    data_s1_2=[]
    set_s4_kd = set(df4['kd_rup_raw'].astype(str).str.split(';').explode().str.strip()) \
    .union(set(pd.Series([i for sub in df4['kd_rup_list'] if isinstance(sub,list) for i in sub]).astype(str))) if not df4.empty else set()

    for _,r in df1_2.iterrows():
        kd=r.get('kd_rup')

        if str(kd) not in set_s4_kd:
            jenis = f"Swakelola {int(swakelola_map[kd])}" if kd in swakelola_map else "N/A"

            data_s1_2.append({
                'Kode RUP':kd,
                'Satuan Kerja':r.get('nama_satker'),
                'Nama Paket':r.get('nama_paket'),
                'Metode Pengadaan':'Swakelola',
                'Jenis Pengadaan':jenis,
                'Sumber Dana':None,
                'PDN':None,
                'UKM':None,
                'Nilai Pagu RUP':r.get('pagu'),
                'Nilai Hasil Pemilihan':"",
                'Tanggal Kontrak':"",
                'Nama Penyedia':"",
                'Status':'Pengumuman RUP',
                'Kode Paket':pd.NA,
                'Nilai HPS':pd.NA,
                'Nilai PDN':pd.NA,
                'Nilai UMK':pd.NA,
                'Metode':'Swakelola',
                'Sumber':'Sumber 1_2'
            })

    df_s1_2=pd.DataFrame(data_s1_2)

    # ======================================================
    # SUMBER 5
    # ======================================================
    data_s5=[]
    for _,r in df5.iterrows():
        kd=r.get('kd_rup')
        kd_list=r.get('kd_rup_list')

        kd_t_list = [i.strip() for i in str(r.get('kd_tender')).split(';')] if pd.notna(r.get('kd_tender')) else []

        status = r.get('status_tender')
        for k in kd_t_list:
            if k in set_t_bapbast:
                status='BAPBAST'; break
            elif k in set_t_spmkspp:
                status='SPMKSPP'; break
            elif k in set_t_kontrak:
                status='Kontrak'; break
            elif k in set_t_sppbj:
                status='SPPBJ'; break
            elif k in set_t_selesai:
                status='Tender Selesai'; break

        nilai_hasil=None
        found=False
        for k in kd_t_list:
            if k in map_t_kontrak:
                nilai_hasil=map_t_kontrak[k]; found=True; break
        if not found:
            nilai_hasil="N/A"
        elif nilai_hasil is not None and not isinstance(nilai_hasil, str) and pd.isna(nilai_hasil):
            nilai_hasil=""

        nilai_pdn="N/A"
        found=False
        for k in kd_t_list:
            if k in map_t_pdn:
                nilai_pdn = map_t_pdn[k]
                if nilai_pdn is not None and not isinstance(nilai_pdn, str) and pd.isna(nilai_pdn):
                    nilai_pdn=""
                found=True
                break

        nilai_umk="N/A"
        found=False
        for k in kd_t_list:
            if k in map_t_umk:
                nilai_umk = map_t_umk[k]
                if nilai_umk is not None and not isinstance(nilai_umk, str) and pd.isna(nilai_umk):
                    nilai_umk=""
                found=True
                break

        pagu = get_pagu_multi(kd_list,'s1')

        # --- LOGIKA BARU UNTUK MEMISAHKAN TENDER & SELEKSI ---
        mtd_pemilihan = str(r.get('mtd_pemilihan', '')).strip()
        if mtd_pemilihan.lower() == 'seleksi':
            kategori_metode = 'Seleksi'
        elif mtd_pemilihan.lower() == 'tender cepat':
            kategori_metode = 'Tender Cepat'
        else:
            kategori_metode = 'Tender'

        data_s5.append({
            'Kode RUP':r.get('kd_rup_raw'),
            'Satuan Kerja':r.get('nama_satker'),
            'Nama Paket':r.get('nama_paket'),
            'Metode Pengadaan':r.get('mtd_pemilihan'),
            'Jenis Pengadaan':r.get('jenis_pengadaan'),
            'Sumber Dana':r.get('sumber_dana'),
            'PDN':get_s1(kd,'status_pdn'),
            'UKM':get_s1(kd,'status_ukm'),
            'Nilai Pagu RUP':pagu,
            'Nilai Hasil Pemilihan':nilai_hasil,
            'Tanggal Kontrak':format_tgl(next((map_t_tgl_kontrak[k] for k in kd_t_list if k in map_t_tgl_kontrak), "")),
            'Nama Penyedia':next((map_t_penyedia[k] for k in kd_t_list if k in map_t_penyedia), ""),
            'Status':status,
            'Kode Paket':r.get('kd_tender'),
            'Nilai HPS':r.get('hps'),
            'Nilai PDN':nilai_pdn,
            'Nilai UMK':nilai_umk,
            'Metode':kategori_metode,
            'Sumber':'Sumber 5'
        })

    df_s5=pd.DataFrame(data_s5)

    # ======================================================
    # SUMBER 6
    # ======================================================
    data_s6=[]
    for _,r in df6.iterrows():
        kd=r.get('rup_code')
        kd_list=r.get('rup_code_list')
        pagu = get_pagu_multi(kd_list,'s1')

        nilai_hasil = r.get('total')
        if pd.isna(nilai_hasil):
            nilai_hasil=""

        # PENYAMAAN TIPE DATA (String to Integer) & CEK MATCH KE SUMBER 1
        try:
            kd_match = int(float(str(kd).strip()))
            is_match = not df1_map.empty and kd_match in df1_map.index
        except:
            kd_match = None
            is_match = False

        if is_match:
            status_pdn_s1 = str(df1_map.loc[kd_match, 'status_pdn']).strip().upper()
            status_ukm_s1 = str(df1_map.loc[kd_match, 'status_ukm']).strip().upper()

            if status_pdn_s1 == 'PDN':
                nilai_pdn_val = nilai_hasil
            else:
                nilai_pdn_val = 0

            if status_ukm_s1 == 'UKM':
                nilai_umk_val = nilai_hasil
            else:
                nilai_umk_val = 0
        else:
            nilai_pdn_val = "N/A"
            nilai_umk_val = "N/A"

        # MENGAMBIL NAMA PENYEDIA DARI KAMUS OFFLINE
        kode_p = str(r.get('kode_penyedia', ""))
        nama_p = map_offline_penyedia.get(kode_p, kode_p) if kode_p and kode_p != "None" else ""

        data_s6.append({
            'Kode RUP':r.get('rup_code_raw'),
            'Satuan Kerja':r.get('nama_satker'),
            'Nama Paket':r.get('rup_name'),
            'Metode Pengadaan':'E-Purchasing',
            'Jenis Pengadaan':get_s1(kd,'jenis_pengadaan'),
            'Sumber Dana':r.get('funding_source'),
            'PDN':get_s1(kd,'status_pdn'),
            'UKM':get_s1(kd,'status_ukm'),
            'Nilai Pagu RUP':pagu,
            'Nilai Hasil Pemilihan':nilai_hasil,
            'Tanggal Kontrak':"",
            'Nama Penyedia':nama_p,
            'Status':r.get('status'),
            'Kode Paket':r.get('order_id'),
            'Nilai HPS':pd.NA,
            'Nilai PDN':nilai_pdn_val,
            'Nilai UMK':nilai_umk_val,
            'Metode':'E-Purchasing V6',
            'Sumber':'Sumber 6'
        })

    df_s6=pd.DataFrame(data_s6)

    # ======================================================
    # SUMBER 7
    # ======================================================
    data_s7=[]
    for _,r in df7.iterrows():
        kd=r.get('kd_rup')
        kd_list=r.get('kd_rup_list')
        pagu = get_pagu_multi(kd_list,'s1')

        nilai_hasil = r.get('total_harga')
        if pd.isna(nilai_hasil):
            nilai_hasil=""

        # PENYAMAAN TIPE DATA (String to Integer) & CEK MATCH KE SUMBER 1
        try:
            kd_match = int(float(str(kd).strip()))
            is_match = not df1_map.empty and kd_match in df1_map.index
        except:
            kd_match = None
            is_match = False

        if is_match:
            status_pdn_s1 = str(df1_map.loc[kd_match, 'status_pdn']).strip().upper()
            status_ukm_s1 = str(df1_map.loc[kd_match, 'status_ukm']).strip().upper()

            if status_pdn_s1 == 'PDN':
                nilai_pdn_val = nilai_hasil
            else:
                nilai_pdn_val = 0

            if status_ukm_s1 == 'UKM':
                nilai_umk_val = nilai_hasil
            else:
                nilai_umk_val = 0
        else:
            nilai_pdn_val = "N/A"
            nilai_umk_val = "N/A"

        # MENGAMBIL NAMA PENYEDIA DARI KAMUS OFFLINE
        kode_p = str(r.get('kd_penyedia', ""))
        nama_p = map_offline_penyedia.get(kode_p, kode_p) if kode_p and kode_p != "None" else ""

        data_s7.append({
            'Kode RUP':r.get('kd_rup_raw'),
            'Satuan Kerja':r.get('nama_satker') if pd.notna(r.get('nama_satker')) else get_s1(kd,'nama_satker'),
            'Nama Paket':r.get('nama_paket'),
            'Metode Pengadaan':'E-Purchasing',
            'Jenis Pengadaan':get_s1(kd,'jenis_pengadaan'),
            'Sumber Dana':r.get('nama_sumber_dana'),
            'PDN':get_s1(kd,'status_pdn'),
            'UKM':get_s1(kd,'status_ukm'),
            'Nilai Pagu RUP':pagu,
            'Nilai Hasil Pemilihan':nilai_hasil,
            'Tanggal Kontrak':"",
            'Nama Penyedia':nama_p,
            'Status':r.get('paket_status_str'),
            'Kode Paket':r.get('kd_paket'),
            'Nilai HPS':pd.NA,
            'Nilai PDN':nilai_pdn_val,
            'Nilai UMK':nilai_umk_val,
            'Metode':'E-Purchasing V5',
            'Sumber':'Sumber 7'
        })

    df_s7=pd.DataFrame(data_s7)

    # ======================================================
    # TAMBAHAN SUMBER 1 (DUAL FIX)
    # ======================================================
    def get_all_kd(df, col):
        if df.empty or col not in df.columns:
            return set()
        raw = set(df[col].astype(str).str.split(';').explode().str.strip())
        lst = set(pd.Series([i for sub in df[col.replace('_raw','_list')] if isinstance(sub,list) for i in sub]).astype(str)) if col.replace('_raw','_list') in df.columns else set()
        return raw.union(lst)

    set_all = get_all_kd(df2,'kd_rup_raw') \
    .union(get_all_kd(df3,'kd_rup_raw')) \
    .union(get_all_kd(df4,'kd_rup_raw')) \
    .union(get_all_kd(df5,'kd_rup_raw')) \
    .union(get_all_kd(df6,'rup_code_raw')) \
    .union(get_all_kd(df7,'kd_rup_raw'))

    data_s1=[]
    for _,r in df1.iterrows():
        kd=r.get('kd_rup')

        if str(kd) not in set_all:
            data_s1.append({
                'Kode RUP':kd,
                'Satuan Kerja':r.get('nama_satker'),
                'Nama Paket':r.get('nama_paket'),
                'Metode Pengadaan':r.get('metode_pengadaan'),
                'Jenis Pengadaan':r.get('jenis_pengadaan'),
                'Sumber Dana':None,
                'PDN':'PDN' if r.get('status_pdn')=='PDN' else 'Non-PDN',
                'UKM':'UKM' if r.get('status_ukm')=='UKM' else 'Non-UKM',
                'Nilai Pagu RUP':r.get('pagu'),
                'Nilai Hasil Pemilihan':"",
                'Tanggal Kontrak':"",
                'Nama Penyedia':"",
                'Status':'Pengumuman RUP',
                'Kode Paket':pd.NA,
                'Nilai HPS':pd.NA,
                'Nilai PDN':pd.NA,
                'Nilai UMK':pd.NA,
                'Metode':r.get('metode_pengadaan'),
                'Sumber':'Sumber 1'
            })

    df_s1=pd.DataFrame(data_s1)

    # ======================================================
    # FILTER FINAL (ANTI DUPLIKAT AMAN)
    # ======================================================
    def safe_col(df, col):
        if df.empty or col not in df.columns:
            return pd.Series([], dtype=object)
        raw = df[col].astype(str)
        lst = pd.Series([str(i) for sub in df.get(col.replace('Kode RUP','kd_rup_list'),[]) if isinstance(sub,list) for i in sub])
        return pd.concat([raw,lst],ignore_index=True)

    df_s1 = df_s1[~safe_col(df_s1,'Kode RUP').isin(pd.concat([
        safe_col(df_s2,'Kode RUP'),
        safe_col(df_s3,'Kode RUP'),
        safe_col(df_s5,'Kode RUP'),
        safe_col(df_s6,'Kode RUP'),
        safe_col(df_s7,'Kode RUP')
    ]))]

    df_s1_2 = df_s1_2[~safe_col(df_s1_2,'Kode RUP').isin(safe_col(df_s4,'Kode RUP'))]

    # ======================================================
    # CLEAN & GABUNG
    # ======================================================
    final_df = pd.concat([df_s2,df_s3,df_s4,df_s1_2,df_s5,df_s6,df_s7,df_s1],ignore_index=True)
    final_df = clean_illegal_chars(final_df)

    # ======================================================
    # SUSUN KOLOM (Kolom Versi Dihapus)
    # ======================================================
    cols = [
        'Kode RUP','Satuan Kerja','Nama Paket','Metode Pengadaan','Jenis Pengadaan',
        'Sumber Dana','PDN','UKM','Nilai Pagu RUP','Nilai Hasil Pemilihan',
        'Tanggal Kontrak','Nama Penyedia','Status','Kode Paket','Nilai HPS',
        'Nilai PDN','Nilai UMK','Metode','Sumber'
    ]

    final_df = final_df[cols]

    # ======================================================
    # BERSIHKAN NaN AGAR JSON VALID
    # ======================================================
    # Kolom PDN & UKM: None/NaN artinya join tidak match → "N/A"
    # Kolom lainnya: None/NaN → ""
    def fill_na_value(series, na_value):
        return series.apply(
            lambda x: na_value if (x is None or (not isinstance(x, str) and pd.isna(x))) else x
        )

    final_df['PDN'] = fill_na_value(final_df['PDN'], "N/A")
    final_df['UKM'] = fill_na_value(final_df['UKM'], "N/A")

    final_df = final_df.fillna("")

    # ======================================================
    # SIMPAN JSON
    # ======================================================
    os.makedirs(data_dir, exist_ok=True)
    output_json = os.path.join(data_dir, f'rekap_pengadaan_{tahun}.json')

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(final_df.to_dict(orient='records'), f, ensure_ascii=False, indent=2)

    print(f"JSON tersimpan: {output_json}")

    # ======================================================
    # SIMPAN EXCEL MASTER
    # ======================================================
    print(f"Generate Excel tahun {tahun}...")

    # Ambil tahun dari df1, fallback ke 2026
    tahun_label = str(df1['tahun_anggaran'].iloc[0]) if not df1.empty and 'tahun_anggaran' in df1.columns else str(tahun)
    tgl   = datetime.now().strftime('%Y-%m-%d')
    nama_file = f'Paket Pengadaan Tahun {tahun_label} ({tgl}) (Api Gateway Legacy).xlsx'

    output_dir   = os.path.join(BASE_DIR, 'output', 'pengadaan', str(tahun))
    output_excel = os.path.join(output_dir, nama_file)

    # Buat folder jika belum ada
    os.makedirs(output_dir, exist_ok=True)

    # Tanggal sama  → to_excel otomatis timpa file yang ada
    # Tanggal berbeda → file baru dibuat, file lama tetap tersimpan sebagai history

    # Kolom angka yang perlu diformat rupiah di Excel
    kolom_angka = ['Nilai Pagu RUP', 'Nilai Hasil Pemilihan', 'Nilai HPS', 'Nilai PDN', 'Nilai UMK']

    # Kolom yang murni angka (boleh coerce ke NaN jika kosong)
    # Nilai Pagu RUP dan Nilai HPS selalu angka atau kosong — aman dikonversi
    kolom_angka_murni = ['Nilai Pagu RUP', 'Nilai HPS']

    # Kolom yang bisa berisi "N/A" — jangan dicoerce, tampilkan apa adanya
    # Nilai Hasil Pemilihan, Nilai PDN, Nilai UMK bisa berisi "N/A" jika join tidak match
    kolom_angka_mixed = ['Nilai Hasil Pemilihan', 'Nilai PDN', 'Nilai UMK']

    # Simpan dulu pakai pandas (cepat), lalu format pakai openpyxl
    excel_df = final_df.copy()

    # Konversi hanya kolom murni angka
    for col in kolom_angka_murni:
        if col in excel_df.columns:
            excel_df[col] = pd.to_numeric(excel_df[col], errors='coerce')

    # Kolom mixed: konversi ke angka hanya jika nilainya bukan "N/A" dan bukan kosong
    # Jika "N/A" → tetap string "N/A", jika angka → jadi float, jika "" → tetap ""
    def safe_numeric(val):
        if val == "N/A" or val == "":
            return val
        try:
            return float(val)
        except (ValueError, TypeError):
            return val

    for col in kolom_angka_mixed:
        if col in excel_df.columns:
            excel_df[col] = excel_df[col].apply(safe_numeric)

    excel_df.to_excel(output_excel, index=False, sheet_name='Pengadaan')

    # ======================================================
    # FORMAT EXCEL DENGAN OPENPYXL
    # ======================================================
    wb = load_workbook(output_excel)
    ws = wb['Pengadaan']

    # --- Warna & style header ---
    header_fill   = PatternFill('solid', start_color='1F4E79')   # biru tua
    header_font   = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    header_align  = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # --- Style data ---
    data_font     = Font(name='Arial', size=10)
    data_align_l  = Alignment(vertical='center', wrap_text=False)
    data_align_c  = Alignment(horizontal='center', vertical='center')
    data_align_r  = Alignment(horizontal='right', vertical='center')

    # --- Warna baris selang-seling ---
    fill_putih    = PatternFill('solid', start_color='FFFFFF')
    fill_biru_muda= PatternFill('solid', start_color='DCE6F1')

    # --- Border tipis ---
    thin = Side(style='thin', color='BFBFBF')
    border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

    # --- Format rupiah ---
    fmt_rupiah = '#,##0'

    # Lebar kolom default per nama kolom (Kolom Versi sudah dihapus)
    lebar_kolom = {
        'Kode RUP'              : 18,
        'Satuan Kerja'          : 38,
        'Nama Paket'            : 50,
        'Metode Pengadaan'      : 22,
        'Jenis Pengadaan'       : 32,
        'Sumber Dana'           : 14,
        'PDN'                   : 10,
        'UKM'                   : 10,
        'Nilai Pagu RUP'        : 20,
        'Nilai Hasil Pemilihan' : 20,
        'Tanggal Kontrak'       : 18,
        'Nama Penyedia'         : 35,
        'Status'                : 22,
        'Kode Paket'            : 20,
        'Nilai HPS'             : 20,
        'Nilai PDN'             : 18,
        'Nilai UMK'             : 18,
        'Metode'                : 22,
        'Sumber'                : 12,
    }

    # Terapkan lebar kolom
    for i, col in enumerate(final_df.columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = lebar_kolom.get(col, 15)

    # Format header (baris 1)
    for cell in ws[1]:
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = header_align
        cell.border    = border_thin

    ws.row_dimensions[1].height = 32

    # Format baris data
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
        fill = fill_putih if row_idx % 2 == 0 else fill_biru_muda
        for cell in row:
            col_name = final_df.columns[cell.column - 1]
            cell.font   = data_font
            cell.fill   = fill
            cell.border = border_thin

            if col_name in kolom_angka:
                cell.number_format = fmt_rupiah
                cell.alignment     = data_align_r
            elif col_name in ('PDN','UKM','Sumber Dana','Metode','Sumber','Status'):
                cell.alignment = data_align_c
            else:
                cell.alignment = data_align_l

    # Freeze baris header
    ws.freeze_panes = 'A2'

    # Auto filter
    ws.auto_filter.ref = ws.dimensions

    wb.save(output_excel)

    print("Excel tersimpan :", output_excel)

    # ======================================================
    # COPY MASTER EXCEL KE FOLDER DATA (untuk web)
    # ======================================================
    master_excel = os.path.join(data_dir, f'master_pengadaan_{tahun}.xlsx')
    shutil.copy2(output_excel, master_excel)
    print(f"Master Excel   : {master_excel}")

    print(f'\nSELESAI TAHUN {tahun} | Total data: {len(final_df)}')
    print(f'JSON  : {output_json}')
    print(f'Excel : {output_excel}')
    return len(final_df)


# ======================================================
# MAIN - LOOP SEMUA TAHUN
# ======================================================
if __name__ == '__main__':
    print("="*55)
    print(f"  PAKET PENGADAAN MULTI TAHUN")
    print(f"  n={tahun_n} | n-1={tahun_n1} | n-2={tahun_n2}")
    print("="*55)

    total_semua = 0
    for t in daftar_tahun:
        hasil = process_tahun(t)
        if hasil:
            total_semua += hasil

    # Simpan last update 1x setelah semua tahun selesai
    with open(os.path.join(BASE_DIR, 'data', 'last-update-pengadaan.txt'), 'w') as f:
        f.write(datetime.now().strftime("%d %B %Y %H:%M WIB"))

    print("\n" + "="*55)
    print(f"  SEMUA TAHUN SELESAI | Total seluruh data: {total_semua}")
    print("="*55)