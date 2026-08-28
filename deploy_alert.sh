#!/bin/bash

# 1. Ambil Token dan Chat ID langsung dari config_rahasia.py
TOKEN=$(python -c "import config_rahasia; print(config_rahasia.BOT_TOKEN)")
CHAT_ID=$(python -c "import config_rahasia; print(config_rahasia.CHAT_ID)")

# 2. Jalankan perintah deploy/build Anda
# Ubah 'npx wrangler deploy' dengan perintah asli Anda jika berbeda
npx wrangler deploy > log_proses.txt 2>&1
STATUS=$?

# 3. Tampilkan isi log ke layar Cloudflare
cat log_proses.txt

# 4. Jika status gagal, kirim ke Telegram
if [ $STATUS -ne 0 ]; then
    PESAN_ERROR=$(tail -n 15 log_proses.txt)
    
    # Gunakan variabel $TOKEN dan $CHAT_ID yang sudah diambil di atas
    curl -s -X POST "https://api.telegram.org/bot${TOKEN}/sendMessage" \
         -d chat_id="${CHAT_ID}" \
         --data-urlencode "text=🚨 ALERT: Deploy Cloudflare Gagal!

Log Terakhir:
$PESAN_ERROR"

    exit 1
fi