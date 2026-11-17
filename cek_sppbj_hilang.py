import json, glob, os, sys
import pandas as pd

t = int(sys.argv[1]) if len(sys.argv) > 1 else 2025
d = f'data/{t}/'

def kode(x):
    return str(x.get('kd_nontender') or x.get('kd_tender'))

n = json.load(open(f'{d}rekap_pengadaan_{t}.json', encoding='utf-8'))
o = json.load(open(f'{d}rekap_pengadaan_{t}.json.bak_sebelum', encoding='utf-8'))
so = {str(r['Kode Paket']): r['Status'] for r in o if r['Kode Paket']}

# paket yang dulu SPPBJ/Kontrak, sekarang mundur ke "Selesai" (sama dengan regresi2)
reg = {str(r['Kode Paket']): r for r in n
       if so.get(str(r['Kode Paket'])) in ('SPPBJ', 'Kontrak') and 'Selesai' in r['Status']}

# endpoint baru
baru = {}
for f in ('v1_tender_non-tender-ekontrak', 'v1_tender_tender-ekontrak'):
    for x in json.load(open(f'{d}{f}_{t}.json', encoding='utf-8')):
        baru[kode(x)] = x

# semua file *ekontrak* (Legacy dan v1 lama), diindeks per file
per_file = {}
for p in sorted(glob.glob(f'{d}*ekontrak*.json')):
    nama = os.path.basename(p)
    try:
        per_file[nama] = {kode(x): x for x in json.load(open(p, encoding='utf-8'))}
    except Exception as e:
        print(nama, 'gagal baca', e)

hilang = set(reg) - set(baru)          # tidak ada di endpoint baru
kosong = set(reg) & set(baru)          # ada di endpoint baru, tahap kosong

def baris(k):
    r = {'Kode Paket': k, 'Status Lama': so.get(k), 'Status Baru': reg[k]['Status']}
    for f, idx in per_file.items():
        r['ada: ' + f] = k in idx
    # data SPPBJ dari file Legacy sppbj (non-tender / tender)
    for f, idx in per_file.items():
        if 'sppbj' in f and k in idx:
            for a, b in idx[k].items():
                r[f'sppbj_{a}'] = b
            break
    if k in baru:
        for a, b in baru[k].items():
            r[f'baru_{a}'] = b
    for a, b in reg[k].items():          # kolom rekap baru, untuk identifikasi paket
        r.setdefault(f'rekap_{a}', b)
    return r

df_h = pd.DataFrame([baris(k) for k in sorted(hilang)])
df_k = pd.DataFrame([baris(k) for k in sorted(kosong)])

ringkas = pd.DataFrame({
    'Keterangan': ['Mundur ke Selesai', 'Hilang di endpoint baru', 'Ada di endpoint baru (tahap kosong)']
                  + [f'Hilang, ada di {f}' for f in per_file],
    'Jumlah': [len(reg), len(hilang), len(kosong)]
              + [len(hilang & set(idx)) for idx in per_file.values()]})

out = f'{d}cek_sppbj_hilang_{t}.xlsx'
with pd.ExcelWriter(out) as w:
    ringkas.to_excel(w, sheet_name='Ringkasan', index=False)
    df_h.to_excel(w, sheet_name='Hilang di endpoint baru', index=False)
    df_k.to_excel(w, sheet_name='Ada tapi tahap kosong', index=False)

print(ringkas.to_string(index=False))
print('Excel:', out)