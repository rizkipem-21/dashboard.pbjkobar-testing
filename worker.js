import * as halaman from "./functions/api/halaman.js";
import * as produk from "./functions/api/produk.js";

const routes = {
  "/api/halaman": halaman,
  "/api/produk": produk,
};

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const mod = routes[url.pathname];

    if (mod) {
      const method = request.method.toUpperCase();
      const handler =
        mod["onRequest" + method.charAt(0) + method.slice(1).toLowerCase()] ||
        mod.onRequest;
      if (handler) {
        return handler({ request, env, waitUntil: ctx.waitUntil.bind(ctx) });
      }
      return new Response("Method not allowed", { status: 405 });
    }

    return env.ASSETS.fetch(request);
  },
};