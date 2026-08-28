#!/bin/bash

# 1. Jalankan perintah deploy/build (gunakan 'perintah_fiktif' dulu untuk tes)
npx wrangler deploy > log_proses.txt 2>&1
STATUS=$?

# 2. Tampilkan isi log ke layar Cloudflare
cat log_proses.txt

# 3. Jika status gagal, kirim ke Telegram
if [ $STATUS -ne 0 ]; then
    PESAN_ERROR=$(tail -n 15 log_proses.txt)
    
    # Memanggil $BOT_TOKEN dan $CHAT_ID langsung dari Environment Variables Cloudflare
    curl -s -X POST "https://api.telegram.org/bot${BOT_TOKEN}/sendMessage" \
         -d chat_id="${CHAT_ID}" \
         --data-urlencode "text=🚨 ALERT: Deploy Cloudflare Gagal!

Log Terakhir:
$PESAN_ERROR"

    exit 1
fi