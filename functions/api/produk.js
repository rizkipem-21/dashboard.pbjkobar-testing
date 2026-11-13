const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }
  });

export async function onRequestGet({ request, env }) {
  const kode = new URL(request.url).searchParams.get("kode_produk") || "";
  if (!UUID.test(kode)) return json({ error: "kode_produk tidak valid" }, 400);
  if (!env.INAPROC_TOKEN) return json({ error: "Token belum diatur di Cloudflare" }, 500);

  const res = await fetch(
    "https://data.inaproc.id/api/v1/ekatalog/produk-detail?kode_produk=" + kode,
    {
      headers: {
        "Authorization": "Bearer " + env.INAPROC_TOKEN,
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json"
      }
    }
  );
  return new Response(res.body, {
    status: res.status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }
  });
}
