// /api/veille : agrège les flux RSS/Atom de la veille professionnelle (sources dans la table veille_sources).
// Réservé au personnel connecté : la liste des sources est lue avec le jeton de la personne (RLS est_personnel).
// Chaque flux est mis en cache 2 heures côté Cloudflare : les actualités se mettent à jour d'elles-mêmes.
const SB = "https://olqclfsywasktwirswpb.supabase.co";
const KEY = "sb_publishable_Qry6l1DGbG9IkLtGT7CF4w_OyP4rI8x";
const CACHE_S = 7200, MAX_JOURS = 90, MAX_PAR_SOURCE = 40;

const ENT = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ", rsquo: "’", lsquo: "‘", ldquo: "“", rdquo: "”", hellip: "…", ndash: "–", mdash: "—", laquo: "«", raquo: "»", eacute: "é", egrave: "è", agrave: "à", ccedil: "ç", ecirc: "ê", ocirc: "ô", icirc: "î", ucirc: "û", euro: "€", oelig: "œ" };
export function decode(t) {
  return String(t || "")
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1")
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d))
    .replace(/&([a-z]+);/gi, (m, n) => ENT[n.toLowerCase()] ?? m);
}
const texte = (t) => decode(decode(t).replace(/<[^>]+>/g, " ")).replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
const tag = (x, n) => { const m = x.match(new RegExp(`<${n}(?:\\s[^>]*)?>([\\s\\S]*?)</${n}>`, "i")); return m ? m[1] : ""; };
const sansAccent = (t) => String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();

export function lireFlux(xml) {
  const items = [];
  const blocs = xml.match(/<item[\s>][\s\S]*?<\/item>/gi) || xml.match(/<entry[\s>][\s\S]*?<\/entry>/gi) || [];
  for (const b of blocs) {
    let url = texte(tag(b, "link"));
    if (!url) { const l = b.match(/<link[^>]*rel=["']alternate["'][^>]*href=["']([^"']+)/i) || b.match(/<link[^>]*href=["']([^"']+)/i); url = l ? decode(l[1]) : ""; }
    if (!url) url = texte(tag(b, "guid"));
    const d = texte(tag(b, "pubDate") || tag(b, "published") || tag(b, "updated") || tag(b, "dc:date"));
    const date = d && !isNaN(Date.parse(d)) ? new Date(d).toISOString() : null;
    let resume = texte(tag(b, "description") || tag(b, "summary") || tag(b, "content:encoded") || tag(b, "content"));
    if (resume.length > 320) resume = resume.slice(0, 300).replace(/\s+\S*$/, "") + "…";
    const titre = texte(tag(b, "title"));
    if (titre && /^https?:\/\//.test(url)) items.push({ titre, url, date, resume });
  }
  return items;
}

async function flux(src, ctx) {
  const cache = caches.default, cle = new Request("https://cache.veille/" + encodeURIComponent(src.url));
  let r = await cache.match(cle);
  if (!r) {
    const ctl = new AbortController(); const to = setTimeout(() => ctl.abort(), 8000);
    try {
      const f = await fetch(src.url, { signal: ctl.signal, headers: { "User-Agent": "Mozilla/5.0 (veille pédagogique UFA Kerplouz LaSalle)", Accept: "application/rss+xml, application/atom+xml, application/xml, text/xml;q=0.9, */*;q=0.5" }, cf: { cacheTtl: CACHE_S } });
      if (!f.ok) throw new Error("réponse " + f.status);
      const xml = await f.text();
      if (!/<(rss|feed|rdf:RDF)[\s>]/i.test(xml)) throw new Error("ce n'est pas un flux RSS");
      r = new Response(xml, { headers: { "Cache-Control": "public, max-age=" + CACHE_S, "X-Lu-Le": new Date().toISOString() } });
      ctx.waitUntil(cache.put(cle, r.clone()));
    } finally { clearTimeout(to); }
  }
  const lu = r.headers.get("X-Lu-Le");
  let items = lireFlux(await r.text());
  const mots = (src.filtre || "").split(",").map((m) => sansAccent(m.trim())).filter(Boolean);
  if (mots.length) items = items.filter((i) => { const t = sansAccent(i.titre + " " + i.resume); return mots.some((m) => t.includes(m)); });
  const limite = Date.now() - MAX_JOURS * 864e5;
  items = items.filter((i) => !i.date || Date.parse(i.date) >= limite).slice(0, MAX_PAR_SOURCE);
  return { items: items.map((i) => ({ ...i, source: src.nom, site: src.site, theme: src.theme, filieres: src.filieres || [] })), lu };
}

export async function onRequestGet(ctx) {
  const m = (ctx.request.headers.get("Cookie") || "").match(/(?:^|;\s*)kep_at=([^;]+)/);
  const json = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "private, no-store" } });
  if (!m) return json({ erreur: "connexion" }, 401);
  const rs = await fetch(SB + "/rest/v1/veille_sources?actif=is.true&order=ordre,nom", { headers: { apikey: KEY, Authorization: "Bearer " + m[1] } });
  if (!rs.ok) return json({ erreur: "sources " + rs.status }, rs.status === 401 ? 401 : 502);
  const sources = await rs.json();
  const res = await Promise.allSettled(sources.map((s) => flux(s, ctx)));
  const items = [], etat = [];
  res.forEach((x, i) => {
    const s = sources[i];
    if (x.status === "fulfilled") { items.push(...x.value.items); etat.push({ id: s.id, source: s.nom, ok: true, nb: x.value.items.length, lu: x.value.lu }); }
    else etat.push({ id: s.id, source: s.nom, ok: false, erreur: String(x.reason && x.reason.message || x.reason) });
  });
  const vus = new Set();
  const uniques = items.filter((i) => { const k = i.url.replace(/[?&]utm_[^&#]*/g, "").replace(/#.*$/, ""); if (vus.has(k)) return false; vus.add(k); return true; })
    .sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  return json({ maj: new Date().toISOString(), items: uniques, sources: etat });
}
