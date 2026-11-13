export async function onRequestGet({ request }) {
  const target = new URL(request.url).searchParams.get("url") || "";
  let u;
  try { u = new URL(target); } catch { return new Response("URL tidak valid", { status: 400 }); }

  // Hanya boleh katalog.inaproc.id, supaya tidak jadi proxy terbuka
  if (u.protocol !== "https:" || u.hostname !== "katalog.inaproc.id") {
    return new Response("Hanya link katalog.inaproc.id yang diizinkan", { status: 403 });
  }

  const res = await fetch(u.href, {
    headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)", "Accept": "text/html" }
  });
  return new Response(await res.text(), {
    status: res.status,
    headers: { "Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store" }
  });
}