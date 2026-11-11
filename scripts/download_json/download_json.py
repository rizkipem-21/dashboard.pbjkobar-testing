# ======================================================
# download_json.py
# DOWNLOAD JSON TERPUSAT (RUP + PENGADAAN) - SATU KALI JALAN
# - Menggabungkan url_rup.txt + url_pengadaan.txt, URL yang sama (nama file hasil sama) hanya diunduh 1x
# - Aturan lewati tahun dinilai per sumber:
#     RUP       : tahun n-1 & n-2 dilewati jika rekap_rup_{tahun}.json sudah ada
#     PENGADAAN : hanya tahun n-2 yang dilewati jika rekap_pengadaan_{tahun}.json / file JSON-nya sudah ada
#     URL dipakai kedua sumber -> tetap diunduh selama salah satu sumber masih membutuhkannya
# - File ditulis lewat file sementara (.tmp) lalu diganti -> tidak pernah setengah jadi
# - Backup harian ke arsip_json dengan percobaan ulang; gagal backup TIDAK menghentikan proses
# - Log sendiri (tools/log_download.txt) & notifikasi Telegram sendiri (CHAT_ID_DOWNLOAD)
# - Kode keluar: 0 = normal (boleh ada file gagal, file lama dipertahankan)
#                2 = SEMUA unduhan gagal (induk .bat sebaiknya melewati generate)
#                1 = skrip crash / konfigurasi salah
# ======================================================

import os
import re
import sys
import json
import time
import shutil
import traceback
import warnings
import requests
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta, timezone

warnings.filterwarnings('ignore')

# ======================================================
# KONFIGURASI
# ======================================================
KIRIM_RINGKASAN_SUKSES = True   # True = kirim Telegram juga saat semua lancar

def cari_base_dir():
    """Naik dari lokasi skrip sampai menemukan token.txt (skrip boleh diletakkan di folder mana pun)"""
    d = os.path.dirname(os.path.abspath(__file__))
    for _ in range(4):
        if os.path.exists(os.path.join(d, 'token.txt')): return d
        d = os.path.dirname(d)
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE_DIR = cari_base_dir()
sys.path.append(BASE_DIR)
try:
    import config_rahasia
except ImportError:
    config_rahasia = None

URL_RUP       = os.path.join(BASE_DIR, 'scripts', 'rup', 'url_rup.txt')
URL_PENGADAAN = os.path.join(BASE_DIR, 'scripts', 'pengadaan', 'url_pengadaan.txt')
LOG_FILE      = os.path.join(BASE_DIR, 'tools', 'log_download.txt')
HASIL_FILE    = os.path.join(BASE_DIR, 'tools', 'hasil_download.json')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

tahun_n  = datetime.now().year
tahun_n1 = tahun_n - 1
tahun_n2 = tahun_n - 2
daftar_tahun = [tahun_n, tahun_n1, tahun_n2]

try:
    with open(os.path.join(BASE_DIR, 'token.txt'), 'r', encoding='utf-8') as f:
        TOKEN = f.read().strip()
    if not TOKEN:
        print("ERROR: File token.txt kosong! Silakan isi dengan token terbaru.")
        sys.exit(1)
except FileNotFoundError:
    print("ERROR: File token.txt tidak ditemukan di folder project!")
    sys.exit(1)

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept": "application/json, text/plain, */*"
}

daftar_error    = []   # unduhan gagal
daftar_kosong   = []   # API membalas kosong, file lama dipertahankan
daftar_peringatan = [] # backup gagal, URL bermasalah, dll
stats = {'ok': 0, 'gagal': 0, 'kosong': 0, 'skip': 0}
STATUS_SKIP = {400, 403, 404, 405, 410, 422}   # permanen: tidak dicoba ulang
MAKS_COBA, MAKS_COBA_RINGAN, BATAS_BERUNTUN = 5, 1, 3
STREAK_FILE = os.path.join(BASE_DIR, 'tools', 'gagal_beruntun.json')
streak = {}
TOKEN_BERMASALAH = False

def muat_streak():
    global streak
    try:
        with open(STREAK_FILE, 'r', encoding='utf-8') as f: streak = json.load(f)
    except Exception: streak = {}

def simpan_streak():
    try:
        with open(STREAK_FILE, 'w', encoding='utf-8') as f: json.dump(streak, f, ensure_ascii=False, indent=1)
    except OSError: pass

# ======================================================
# LOG (ke layar + file sendiri). Baris sebagian (end=" ") digabung dulu agar log utuh
# ======================================================
_baris_tertunda = ""

def log_print(*args, **kwargs):
    global _baris_tertunda
    msg = " ".join(str(a) for a in args)
    if kwargs.get('end') == " ":
        print(msg, end=" ", flush=True)
        _baris_tertunda += msg + " "
        return
    print(msg, flush=True)
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(_baris_tertunda + msg + '\n')
    except OSError:
        pass
    _baris_tertunda = ""

# ======================================================
# HELPER UMUM
# ======================================================
def get_waktu_indonesia():
    tz_wib = timezone(timedelta(hours=7))
    s = datetime.now(tz_wib)
    bulan = {1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April', 5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'}
    return f"{s.day} {bulan[s.month]} {s.year} | {s.strftime('%H.%M')} WIB"

def alasan_http(resp):
    peta = {
        400: "Permintaan tidak valid (parameter URL salah)",
        401: "Token tidak valid / kedaluwarsa (cek token.txt)",
        403: "Akses ditolak: token tidak punya izin ke endpoint ini",
        404: "Endpoint/data tidak ditemukan (cek URL)",
        408: "Waktu permintaan habis",
        429: "Terlalu banyak permintaan (kena batas rate limit)",
        500: "Error internal di server Inaproc",
        502: "Bad Gateway: server perantara Inaproc menerima respons buruk dari server utama",
        503: "Layanan Inaproc tidak tersedia (maintenance / kelebihan beban)",
        504: "Gateway Timeout: server Inaproc terlalu lama memproses (data besar / server sibuk)",
    }
    alasan = peta.get(resp.status_code, resp.reason or "Tidak diketahui")
    cuplikan = ""
    try:
        teks = re.sub(r"<(style|script)[^>]*>.*?</\1>", " ", resp.text or "", flags=re.S | re.I)
        teks = re.sub(r"<[^>]+>", " ", teks)
        teks = re.sub(r"\s+", " ", teks).strip()[:120]
        if teks: cuplikan = f" | Pesan server: {teks}"
    except Exception:
        pass
    return f"Status {resp.status_code} - {alasan}{cuplikan}"

def alasan_error(e):
    nama = type(e).__name__
    if isinstance(e, requests.exceptions.Timeout):
        return "Waktu tunggu habis (server tidak merespons dalam 150 detik)"
    if isinstance(e, requests.exceptions.SSLError):
        return f"Gagal koneksi aman/SSL ({str(e)[:80]})"
    if isinstance(e, requests.exceptions.ConnectionError):
        return f"Gagal tersambung ke server (internet putus / DNS / server mati) [{nama}]"
    if isinstance(e, ValueError):
        return "Respons server bukan JSON yang valid"
    return f"{nama}: {str(e)[:100]}"

def hitung_baris(data):
    if isinstance(data, list): return len(data)
    if isinstance(data, dict):
        for k in ('data', 'items', 'results'):
            if isinstance(data.get(k), list): return len(data[k])
        return 1
    return 0

def kirim_telegram(pesan):
    token = getattr(config_rahasia, 'BOT_TOKEN', None)
    chat  = getattr(config_rahasia, 'CHAT_ID_DOWNLOAD', None)
    if not token or not chat:
        log_print("⚠️ Telegram dilewati: BOT_TOKEN / CHAT_ID_DOWNLOAD belum ada di config_rahasia.py")
        return
    if len(pesan) > 4000: pesan = pesan[:4000] + "\n...[TERPOTONG]"
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    for percobaan in range(1, 4):
        try:
            res = requests.post(url, data={"chat_id": chat, "text": pesan}, timeout=30)
            if res.status_code != 200:
                log_print(f"🚨 GAGAL KIRIM TELEGRAM (STATUS {res.status_code}): {res.text}")
            return
        except Exception as e:
            if percobaan < 3:
                log_print(f"⚠️ Koneksi Telegram lambat. Mencoba ulang ({percobaan}/3)...")
                time.sleep(5)
            else:
                log_print(f"🚨 GAGAL KONEKSI KE TELEGRAM: {str(e)}")

# ======================================================
# FILE: BACKUP HARIAN & TULIS AMAN
# ======================================================
def backup_harian(output_path, filename):
    """Salin file lama ke arsip_json bila berasal dari hari sebelumnya. Gagal -> peringatan saja, proses lanjut."""
    if not os.path.exists(output_path): return
    tgl_file = datetime.fromtimestamp(os.path.getmtime(output_path)).date()
    if tgl_file >= datetime.now().date(): return

    m = re.search(r'(20\d{2})', filename)
    base_arsip = os.path.join(BASE_DIR, 'arsip_json', str(tgl_file.year), f"{tgl_file.month:02d}", f"{tgl_file.day:02d}")
    if m:
        f_arsip, folder_log = os.path.join(base_arsip, m.group(1)), f"{tgl_file}/{m.group(1)}"
    else:
        f_arsip, folder_log = base_arsip, f"{tgl_file}"

    err = ""
    for i in range(1, 6):
        try:
            os.makedirs(f_arsip, exist_ok=True)
            shutil.copy2(output_path, os.path.join(f_arsip, filename))
            log_print(f"[BACKUP] {filename} di-copy ke arsip {folder_log}")
            return
        except OSError as e:
            err = f"{type(e).__name__}: {str(e)[:80]}"
            log_print(f"[BACKUP] Gagal salin {filename} (percobaan {i}/5): {err}")
            time.sleep(2 * i)
    daftar_peringatan.append(f"⚠️ Backup {filename} GAGAL ({err}) - file lama tertimpa tanpa arsip")
    log_print(f"[BACKUP] {filename} dilewati setelah 5 percobaan, proses tetap dilanjutkan.")

def tulis_json_aman(output_path, data):
    """Tulis ke .tmp lalu os.replace -> file tujuan tidak pernah setengah jadi. Return (berhasil, pesan_error)"""
    tmp = output_path + ".tmp"
    try:
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError as e:
        return False, f"Gagal menulis file sementara: {type(e).__name__}: {str(e)[:80]}"
    err = ""
    for i in range(1, 6):
        try:
            os.replace(tmp, output_path)
            return True, ""
        except OSError as e:
            err = f"{type(e).__name__}: {str(e)[:80]}"
            time.sleep(2 * i)
    try: os.remove(tmp)
    except OSError: pass
    return False, f"Gagal mengganti file (terkunci proses lain?): {err}"

# ======================================================
# PENGGABUNGAN URL
# ======================================================
def baca_url(path):
    if not os.path.exists(path): return None
    with open(path, 'r', encoding='utf-8') as f:
        return [l.strip() for l in f if l.strip() and not l.strip().startswith(('#', '='))]

def kueri(u):
    return parse_qs(urlparse(u).query)

def gabung_url():
    """Return (daftar_entry_berurutan). Kunci = (tipe, base_name) = identik dengan nama file hasil."""
    entries, urut = {}, []
    for sumber, path in [('rup', URL_RUP), ('pengadaan', URL_PENGADAAN)]:
        urls = baca_url(path)
        if urls is None:
            msg = f"⚠️ File URL {sumber} tidak ditemukan: {path}"
            log_print(msg); daftar_peringatan.append(msg)
            continue
        for raw in urls:
            m = re.search(r'api/(?:v1|legacy)/(.*?)\?', raw)
            if not m:
                msg = f"⚠️ URL {sumber} dilewati (pola tidak dikenali): {raw[:90]}"
                log_print(msg); daftar_peringatan.append(msg)
                continue
            is_v1 = '/v1/' in raw
            tipe = "v1" if is_v1 else "Legacy"
            base = m.group(1).replace('/', '_')
            key = (tipe, base)
            if key in entries:
                entries[key]['sumber'].add(sumber)
                if kueri(raw) != kueri(entries[key]['url']):
                    msg = f"⚠️ URL berbeda untuk file yang sama ({tipe}_{base}); dipakai versi {entries[key]['asal_url'].upper()}"
                    log_print(msg); daftar_peringatan.append(msg)
            else:
                entries[key] = {'url': raw, 'tipe': tipe, 'base': base, 'is_v1': is_v1, 'sumber': {sumber}, 'asal_url': sumber}
                urut.append(key)
    return [entries[k] for k in urut]

def sumber_yang_butuh(e, tahun, output_path):
    """Sumber mana yang MASIH membutuhkan file ini untuk tahun tsb (kosong = boleh dilewati)"""
    data_dir = os.path.join(BASE_DIR, 'data', str(tahun))
    perlu = []
    if 'rup' in e['sumber']:
        lewati = (tahun != tahun_n) and os.path.exists(os.path.join(data_dir, f'rekap_rup_{tahun}.json'))
        if not lewati: perlu.append('rup')
    if 'pengadaan' in e['sumber']:
        lewati = (tahun == tahun_n2) and (os.path.exists(os.path.join(data_dir, f'rekap_pengadaan_{tahun}.json')) or os.path.exists(output_path))
        if not lewati: perlu.append('pengadaan')
    return perlu

# ======================================================
# UNDUH
# ======================================================
def get_dengan_ulang(url, label, maks):
    """Return (ok, data, err). Error permanen -> langsung berhenti; sementara -> coba ulang."""
    global TOKEN_BERMASALAH
    err = "Unknown Error"
    for i in range(1, maks + 1):
        tunggu = i * 5
        try:
            log_print(f"  {label} (Coba {i}/{maks})...", end=" ")
            resp = requests.get(url, headers=HEADERS, timeout=150)
            if resp.status_code == 200:
                data = resp.json()
                log_print("SUKSES")
                return True, data, ""
            err = alasan_http(resp)
            log_print(f"GAGAL ({err})")
            if resp.status_code == 401:
                TOKEN_BERMASALAH = True
                return False, None, err
            if resp.status_code in STATUS_SKIP:
                log_print("    -> Error permanen, tidak dicoba ulang.")
                return False, None, err
            if resp.status_code == 429:
                try: tunggu = min(int(resp.headers.get('Retry-After', 30 * i)), 120)
                except ValueError: tunggu = 30 * i
        except Exception as e:
            err = alasan_error(e)
            log_print(f"ERROR: {err}")
        if i < maks:
            log_print(f"    -> [!] Menunggu {tunggu} detik...")
            time.sleep(tunggu)
    return False, None, err

def unduh_v1(target_url, maks):
    all_data, cursor, req_count = [], None, 1
    while True:
        url_cursor = target_url if not cursor else f"{target_url}{'&' if '?' in target_url else '?'}cursor={cursor}"
        ok, resp_data, err = get_dengan_ulang(url_cursor, f"Request ke-{req_count}", maks)
        if not ok:
            log_print("  GAGAL TOTAL pada request ini. Berhenti ditarik.")
            return False, [], err
        if isinstance(resp_data, dict) and 'data' in resp_data:
            if resp_data['data'] is not None: all_data.extend(resp_data['data'])
        elif isinstance(resp_data, list):
            all_data.extend(resp_data)
        if isinstance(resp_data, dict) and isinstance(resp_data.get('meta'), dict) and resp_data['meta'].get('has_more'):
            cursor = resp_data['meta'].get('cursor')
            req_count += 1
            time.sleep(1)
        else:
            return True, all_data, ""

def unduh_legacy(target_url, maks):
    ok, data, err = get_dengan_ulang(target_url, "Percobaan", maks)
    return ok, data, err

def proses_url(e, tahun, perlu, output_path, filename):
    target_url = e['url'].replace('{tahun}', str(tahun))
    label = "+".join(s.upper() for s in ('rup', 'pengadaan') if s in perlu)
    tag = f"{label} {e['tipe'].upper()} ({tahun}) - {e['base']}"
    n_gagal = streak.get(filename, 0)
    maks = MAKS_COBA_RINGAN if n_gagal >= BATAS_BERUNTUN else MAKS_COBA

    log_print("")
    backup_harian(output_path, filename)
    info = f" (dicoba {maks}x: gagal beruntun {n_gagal} kali)" if maks < MAKS_COBA else ""
    log_print(f"DOWNLOAD [{e['tipe'].upper()}] [{label}]: {target_url}{info}")

    ok, data, err = unduh_v1(target_url, maks) if e['is_v1'] else unduh_legacy(target_url, maks)
    ada_lama = os.path.exists(output_path)

    if not ok:
        if not ada_lama:
            tulis_json_aman(output_path, [])
            log_print(f"  -> File {filename} belum ada, dibuat kosong [].")
        else:
            log_print(f"  -> [Aman] Proses terputus, mempertahankan file {filename} lama yang utuh.")
        streak[filename] = n_gagal + 1
        catatan = f" [gagal beruntun {n_gagal + 1}x]" if n_gagal + 1 >= 2 else ""
        daftar_error.append(f"❌ {tag} ({err}){catatan}")
        stats['gagal'] += 1
        return

    streak.pop(filename, None)   # berhasil -> hitungan beruntun direset

    if e['is_v1'] and len(data) == 0:
        if ada_lama:
            log_print(f"  -> [Aman] Data API kosong (null), mempertahankan file {filename} lama yang utuh.")
            daftar_kosong.append(f"ℹ️ {tag.replace('_', ' ')}")
            stats['kosong'] += 1
        else:
            tulis_json_aman(output_path, [])
            log_print("  -> File baru dibuat dengan daftar kosong [] (Total: 0 baris)")
            stats['ok'] += 1
        return

    berhasil, pesan = tulis_json_aman(output_path, data)
    if not berhasil:
        log_print(f"  -> [Aman] {pesan}. File lama dipertahankan.")
        daftar_error.append(f"❌ {tag} ({pesan})")
        stats['gagal'] += 1
    else:
        log_print(f"  -> Disimpan ke {filename} (Total: {hitung_baris(data)} baris)")
        stats['ok'] += 1

# ======================================================
# MAIN
# ======================================================
def main():
    waktu_mulai = time.time()
    log_print("\n" + "=" * 55)
    log_print(f"START DOWNLOAD JSON (RUP + PENGADAAN) {get_waktu_indonesia()}")
    log_print("=" * 55)

    entries = gabung_url()
    if not entries:
        log_print("ERROR: Tidak ada URL yang bisa diproses.")
        kirim_telegram("🚨 DOWNLOAD JSON GAGAL 🚨\n\nTidak ada URL yang terbaca dari url_rup.txt / url_pengadaan.txt.\n\n" + "\n".join(daftar_peringatan) + f"\n\n📅 {get_waktu_indonesia()}")
        return 1
    log_print(f"ℹ️ {len(entries)} endpoint unik digabung dari url_rup.txt + url_pengadaan.txt")
    muat_streak()

    for t in daftar_tahun:
        data_dir = os.path.join(BASE_DIR, 'data', str(t))
        os.makedirs(data_dir, exist_ok=True)
        log_print(f'\n{"=" * 55}\n   DOWNLOAD DATA TAHUN {t}\n{"=" * 55}')
        for e in entries:
            filename = f"{e['tipe']}_{e['base']}_{t}.json"
            output_path = os.path.join(data_dir, filename)
            perlu = sumber_yang_butuh(e, t, output_path)
            if not perlu:
                stats['skip'] += 1
                log_print(f"[SKIP] {filename} (sudah final)")
                continue
            proses_url(e, t, perlu, output_path, filename)
            if TOKEN_BERMASALAH: break
        if TOKEN_BERMASALAH: break

    durasi = str(timedelta(seconds=int(time.time() - waktu_mulai)))
    total_gagal_semua = (stats['gagal'] > 0 and stats['ok'] + stats['kosong'] == 0) or TOKEN_BERMASALAH
    if TOKEN_BERMASALAH:
        daftar_peringatan.append("🔑 Token tidak valid/kedaluwarsa (401). Unduhan dihentikan, perbarui token.txt.")
    simpan_streak()
    log_print("\n" + "=" * 55)
    log_print(f"RINGKASAN: Berhasil {stats['ok']} | Kosong {stats['kosong']} | Gagal {stats['gagal']} | Dilewati {stats['skip']} | Durasi {durasi}")
    log_print("=" * 55)

    try:
        with open(HASIL_FILE, 'w', encoding='utf-8') as f:
            json.dump({'waktu': get_waktu_indonesia(), 'durasi': durasi, 'ringkasan': stats, 'gagal_total': total_gagal_semua,
                       'error': daftar_error, 'kosong': daftar_kosong, 'peringatan': daftar_peringatan}, f, ensure_ascii=False, indent=2)
    except OSError:
        pass

    ada_masalah = bool(daftar_error or daftar_kosong or daftar_peringatan)
    if ada_masalah or KIRIM_RINGKASAN_SUKSES:
        judul = "🚨 LAPORAN DOWNLOAD JSON 🚨" if ada_masalah else "✅ DOWNLOAD JSON BERHASIL ✅"
        pesan = f"{judul}\n\n📦 Berhasil {stats['ok']} | Kosong {stats['kosong']} | Gagal {stats['gagal']} | Dilewati {stats['skip']}\n\n"
        if total_gagal_semua:
            pesan += "⛔ SEMUA UNDUHAN GAGAL. Generate sebaiknya tidak dijalankan.\n\n"
        if daftar_error:
            teks = "⚠️ GAGAL DOWNLOAD:\n" + "".join(f"{x.replace('_', ' ')}\n" for x in daftar_error)
            pesan += (teks[:2500] + "\n... [DIPOTONG, lihat log_download.txt]\n" if len(teks) > 2500 else teks) + "\n"
        if daftar_kosong:
            teks = "ℹ️ API KOSONG (file lama dipertahankan):\n" + "".join(f"{x}\n" for x in daftar_kosong)
            pesan += (teks[:800] + "...\n" if len(teks) > 800 else teks) + "\n"
        if daftar_peringatan:
            teks = "⚠️ PERINGATAN:\n" + "".join(f"{x}\n" for x in daftar_peringatan)
            pesan += (teks[:700] + "...\n" if len(teks) > 700 else teks) + "\n"
        pesan += f"⏱ Durasi: {durasi}\n📅 Waktu: {get_waktu_indonesia()}"
        kirim_telegram(pesan)

    log_print(f"\nPROSES DOWNLOAD SELESAI PADA {get_waktu_indonesia()}")
    return 2 if total_gagal_semua else 0

if __name__ == '__main__':
    try:
        kode = main()
    except Exception:
        tb = traceback.format_exc()
        log_print(tb)
        kirim_telegram("🚨 DOWNLOAD JSON CRASH 🚨\n\n" + tb[-1500:] + f"\n\n📅 {get_waktu_indonesia()}")
        sys.exit(1)
    sys.exit(kode)