// Protège /BTS/ : réservé aux comptes actifs qui ont le droit « grilles-bts » (vérifié côté serveur).
const SB = "https://olqclfsywasktwirswpb.supabase.co";
const KEY = "sb_publishable_Qry6l1DGbG9IkLtGT7CF4w_OyP4rI8x";
const DROIT = "grilles-bts";

export async function onRequest(ctx) {
  const url = new URL(ctx.request.url);
  const m = (ctx.request.headers.get("Cookie") || "").match(/(?:^|;\s*)kep_at=([^;]+)/);
  let refus = "connexion";
  if (m) {
    const r = await fetch(SB + "/rest/v1/rpc/mes_droits", {
      method: "POST",
      headers: { apikey: KEY, Authorization: "Bearer " + m[1], "Content-Type": "application/json" },
      body: "{}",
    });
    if (r.ok) {
      const d = await r.json();
      const ok = d && d.actif && (DROIT === "admin" ? d.admin : (d.outils || []).includes(DROIT));
      if (ok) {
        const res = await ctx.next();
        const out = new Response(res.body, res);
        out.headers.set("Cache-Control", "private, no-store");
        // En-têtes de sécurité (mêmes valeurs que /_headers)
        out.headers.set("X-Frame-Options", "DENY");
        out.headers.set("Content-Security-Policy", "frame-ancestors 'none'; base-uri 'self'; object-src 'none'; form-action 'self'");
        out.headers.set("X-Content-Type-Options", "nosniff");
        out.headers.set("Referrer-Policy", "strict-origin-when-cross-origin");
        out.headers.set("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=()");
        out.headers.set("Strict-Transport-Security", "max-age=31536000");
        return out;
      }
      refus = d && d.actif ? "droit" : "suspendu";
    }
  }
  const back = refus === "connexion"
    ? url.origin + "/?next=" + encodeURIComponent(url.pathname)
    : url.origin + "/?refus=" + refus;
  return new Response(null, { status: 302, headers: { Location: back, "Cache-Control": "no-store" } });
}
