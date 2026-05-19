# 1. Gunakan base image Python resmi berbasis Linux Alpine yang sangat ringan
FROM python:3.11-alpine

# 2. Instal Nginx (Web Server) dan Cron (Otomatisasi Jadwal Linux)
RUN apk add --no-cache nginx dcron tzdata

# 3. Atur zona waktu ke Asia/Jakarta (WIB) agar jadwal penarikan data sinkron
ENV TZ=Asia/Jakarta

# 4. Tentukan folder kerja di dalam Linux Kontainer
WORKDIR /app

# 5. Salin file requirements atau langsung instal library Python yang dibutuhkan
RUN pip install --no-cache-dir requests pandas openpyxl

# 6. Salin seluruh file proyek lokal ke dalam kontainer Docker
COPY . .

# 7. Konfigurasi agar Nginx menyajikan file HTML di folder /app
RUN mkdir -p /run/nginx
COPY docker/nginx.conf /etc/nginx/nginx.conf

# 8. Daftarkan skrip cron untuk otomatisasi penarikan data (Setiap hari jam 06:00 WIB)
# 1. Jadwal RUP: Setiap hari (Jam 05:00-11:00 tiap jam, lalu 14:00, 17:00, 20:00)
RUN echo '0 5,6,7,8,9,10,11,14,17,20 * * * cd /app && python scripts/rup/generate_rup.py' > /etc/crontabs/root && \
# 2. Jadwal Pengadaan: Setiap Senin (1) dan Kamis (4) dengan jam yang sama
    echo '0 5,6,7,8,9,10,11,14,17,20 * * 1,4 cd /app && python scripts/pengadaan/generate_pengadaan.py' >> /etc/crontabs/root && \
# 3. Jadwal Pengadaan: Setiap Tanggal 1 (Awal Bulan) dengan jam yang sama
    echo '0 5,6,7,8,9,10,11,14,17,20 1 * * cd /app && python scripts/pengadaan/generate_pengadaan.py' >> /etc/crontabs/root && \
# 4. Jadwal Pengadaan: Setiap Akhir Bulan (Trik filter Python di tanggal 28-31)
    echo '0 5,6,7,8,9,10,11,14,17,20 28,29,30,31 * * cd /app && python -c "import datetime,calendar; d=datetime.date.today(); exit(0 if d.day==calendar.monthrange(d.year,d.month)[1] else 1)" && python scripts/pengadaan/generate_pengadaan.py' >> /etc/crontabs/root
    
# 9. Jalankan Nginx web server dan sistem Cron secara bersamaan saat kontainer aktif
EXPOSE 80
CMD ["sh", "-c", "crond && nginx -g 'daemon off;'"]