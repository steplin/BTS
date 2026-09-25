// Protège /livret/ (outil et modèles) : l'outil n'est servi qu'aux personnes connectées (jeton Supabase valide).
const SB = "https://olqclfsywasktwirswpb.supabase.co";
const KEY = "sb_publishable_Qry6l1DGbG9IkLtGT7CF4w_OyP4rI8x";

export async function onRequest(ctx) {
  const url = new URL(ctx.request.url);
  const cookie = ctx.request.headers.get("Cookie") || "";
  const m = cookie.match(/(?:^|;\s*)kep_at=([^;]+)/);
  if (m) {
    const r = await fetch(SB + "/auth/v1/user", {
      headers: { apikey: KEY, Authorization: "Bearer " + m[1] },
    });
    if (r.ok) {
      const res = await ctx.next();
      const out = new Response(res.body, res);
      out.headers.set("Cache-Control", "private, no-store");
      return out;
    }
  }
  const back = url.origin + "/?next=" + encodeURIComponent(url.pathname);
  return new Response(null, {
    status: 302,
    headers: { Location: back, "Cache-Control": "no-store" },
  });
}
