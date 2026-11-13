const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const INAPROC = "https://data.inaproc.id/api/v1/ekatalog/produk-detail?kode_produk=";
const PESAN_MATI = "Layanan Detail Produk sedang tidak aktif (komputer kantor belum menyala). Coba lagi nanti.";

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }
  });

export async function onRequestGet({ request, env }) {
  const kode = new URL(request.url).searchParams.get("kode_produk") || "";
  if (!UUID.test(kode)) return json({ error: "kode_produk tidak valid" }, 400);

  let res;

  if (env.INAPROC_TOKEN) {
    // Mode lokal (wrangler dev): langsung ke Inaproc dari IP komputer ini
    res = await fetch(INAPROC + kode, {
      headers: {
        "Authorization": "Bearer " + env.INAPROC_TOKEN.trim(),
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json"
      }
    });
  } else {
    // Mode Cloudflare: lewat tunnel ke komputer kantor
    const base = env.TUNNEL ? await env.TUNNEL.get("tunnel_url") : null;
    if (!base || !env.PROXY_KEY) return json({ error: PESAN_MATI }, 503);
    try {
      res = await fetch(base + "/produk?kode_produk=" + kode, {
        headers: { "X-Proxy-Key": env.PROXY_KEY },
        signal: AbortSignal.timeout(25000)
      });
    } catch {
      return json({ error: PESAN_MATI }, 503);
    }
    // Tunnel mati menjawab HTML, sedangkan proxy selalu menjawab JSON
    if (!(res.headers.get("content-type") || "").includes("json")) {
      return json({ error: PESAN_MATI }, 503);
    }
  }

  return new Response(await res.text(), {
    status: res.status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }
  });
}