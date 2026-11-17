import os, json
os.chdir(r'D:\dashboard.pbjkobar-testing')
def muat(p):
    with open(p, encoding='utf-8') as f: return json.load(f)

d = muat('data/2026/v1_tender_tender-ekontrak_2026.json')
d = d.get('data', d) if isinstance(d, dict) else d
for r in d:
    if str(r.get('kd_tender')) == '10115455000':
        print({k: v for k, v in r.items() if v not in (None, '') and any(s in k for s in ('tgl', 'status', 'kd_', 'no_', 'addendum', 'versi'))})
        print('---')

for r in muat('data/2026/rekap_pengadaan_2026.json'):
    if str(r.get('Kode Paket')) == '10115455000':
        print('REKAP:', {k: v for k, v in r.items() if v and ('SPPBJ' in k or 'SPMK' in k or 'Status' in k or 'Kontrak' in k)})