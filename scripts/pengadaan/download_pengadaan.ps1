# ==============================
# AUTO DOWNLOAD PENGADAAN INAPROC (MULTI TAHUN)
# ==============================

$baseDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataRootPath = Join-Path $baseDir "..\..\data"

# Buat folder data root jika belum ada
if (!(Test-Path $dataRootPath)) {
    New-Item -ItemType Directory -Path $dataRootPath | Out-Null
}

$token = "inprc7642391c38774272bf57ca25ac1d4544"

$headers = @{
    Authorization = "Bearer $token"
}

# ============================================
# HITUNG TAHUN n, n-1, n-2
# ============================================
$tahunN   = (Get-Date).Year        # 2026 (tahun berjalan)
$tahunN1  = $tahunN - 1            # 2025
$tahunN2  = $tahunN - 2            # 2024

# ============================================
# DAFTAR ENDPOINT (tanpa tahun, nanti diisi loop)
# ============================================
$endpoints = @(
    "rup/paket-penyedia-terumumkan",
    "rup/paket-swakelola-terumumkan",
    "tender/non-tender-ekontrak-bapbast",
    "tender/non-tender-ekontrak-kontrak",
    "tender/non-tender-ekontrak-spmkspp",
    "tender/non-tender-ekontrak-sppbj",
    "tender/non-tender-pengumuman",
    "tender/non-tender-selesai",
    "tender/pengumuman",
    "tender/tender-ekontrak-bapbast",
    "tender/tender-ekontrak-kontrak",
    "tender/tender-ekontrak-spmkspp",
    "tender/tender-ekontrak-sppbj",
    "tender/tender-selesai",
    "tender/tender-selesai-nilai",
    "tender/pencatatan-non-tender",
    "tender/pencatatan-swakelola",
    "ekatalog-archive/paket-e-purchasing",
    "ekatalog/paket-e-purchasing"
)

# ============================================
# FUNGSI DOWNLOAD DENGAN RETRY
# ============================================
function Download-WithRetry($url, $output) {

    $maxRetry = 5
    $success  = $false

    for ($i = 1; $i -le $maxRetry; $i++) {

        try {
            Write-Host "  Percobaan ke-$i..."

            $response = Invoke-RestMethod -Method GET -Uri $url -Headers $headers

            if ($null -ne $response) {
                $response | ConvertTo-Json -Depth 20 | Out-File -Encoding utf8 $output
                Write-Host "  SUKSES" -ForegroundColor Green
                $success = $true
                break
            }
        }
        catch {
            Write-Host "  Gagal percobaan ke-$i" -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }

    if (-not $success) {
        Write-Host "  GAGAL TOTAL -> buat file kosong" -ForegroundColor Red
        "[]" | Out-File -Encoding utf8 $output
    }
}

# ============================================
# LOOP PER TAHUN
# ============================================
$daftarTahun = @($tahunN, $tahunN1, $tahunN2)

foreach ($tahun in $daftarTahun) {

    Write-Host ""
    Write-Host "======================================" -ForegroundColor Cyan
    Write-Host "  TAHUN $tahun" -ForegroundColor Cyan
    Write-Host "======================================" -ForegroundColor Cyan

    # Subfolder per tahun: data4, data5, data6
    $dataPath = Join-Path $dataRootPath "$tahun"
    if (!(Test-Path $dataPath)) {
        New-Item -ItemType Directory -Path $dataPath | Out-Null
        Write-Host "  Folder dibuat: $dataPath" -ForegroundColor DarkGray
    }

    # Cek apakah tahun ini adalah n-2
    $isN2 = ($tahun -eq $tahunN2)

    foreach ($endpoint in $endpoints) {

        $url      = "https://data.inaproc.id/api/legacy/" + $endpoint + "?kode_klpd=D228&tahun=" + $tahun
        $baseName = ($endpoint -replace '/', '_')
        $filename = "Legacy_${baseName}_${tahun}.json"
        $output   = Join-Path $dataPath $filename

        # -----------------------------------------------
        # LOGIK SKIP untuk tahun n-2
        # File sudah ada → skip (data sudah final)
        # File belum ada → download 1x
        # -----------------------------------------------
        if ($isN2 -and (Test-Path $output)) {
            Write-Host ""
            Write-Host "SKIP (sudah final): $filename" -ForegroundColor DarkGray
            continue
        }

        Write-Host ""
        Write-Host "DOWNLOAD: $url" -ForegroundColor Yellow

        Download-WithRetry $url $output

        Write-Host "FILE -> $filename"
    }
}

Write-Host ""
Write-Host "SELESAI DOWNLOAD SEMUA DATA PENGADAAN" -ForegroundColor Cyan