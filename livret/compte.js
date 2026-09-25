// Compte Kerplouz (Supabase) : même session que la page d'accueil. Ne reçoit jamais la liste
// d'apprentis : seulement la correspondance groupe → modèle (outil "livret").
window.Compte = (() => {
  "use strict";
  const SB = "https://olqclfsywasktwirswpb.supabase.co", KEY = "sb_publishable_Qry6l1DGbG9IkLtGT7CF4w_OyP4rI8x";
  const SKEY = "kep.session", OUTIL = "livret";
  let session = null;
  const ls = {
    get(k) { try { return JSON.parse(localStorage.getItem(k) || "null"); } catch (e) { return null; } },
    set(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} },
  };
  async function api(path, { method = "GET", body, auth = true, prefer } = {}) {
    const h = { apikey: KEY, "Content-Type": "application/json" };
    if (auth && session) h.Authorization = "Bearer " + session.access_token;
    if (prefer) h.Prefer = prefer;
    const r = await fetch(SB + path, { method, headers: h, body: body ? JSON.stringify(body) : undefined });
    const txt = await r.text(); let data = null; try { data = txt ? JSON.parse(txt) : null; } catch (e) { data = txt; }
    if (!r.ok) {
      const e = new Error((data && (data.msg || data.message || data.error_description || data.error)) || ("Erreur " + r.status));
      e.status = r.status; throw e;
    }
    return data;
  }
  function cookie() {
    document.cookie = session
      ? "kep_at=" + session.access_token + "; Path=/; Max-Age=" + Math.max(60, session.expires_at - Math.floor(Date.now() / 1000)) + "; Secure; SameSite=Lax"
      : "kep_at=; Path=/; Max-Age=0; Secure; SameSite=Lax";
  }
  function stocker(t) {
    session = {
      access_token: t.access_token, refresh_token: t.refresh_token,
      expires_at: t.expires_at ? +t.expires_at : Math.floor(Date.now() / 1000) + (+t.expires_in || 3600),
      user: t.user ? { id: t.user.id, email: t.user.email } : (session && session.user) || null,
    };
    ls.set(SKEY, session); cookie();
  }
  async function frais() {
    if (!session) return false;
    if (session.expires_at - 60 > Date.now() / 1000) return true;
    try {
      stocker(await api("/auth/v1/token?grant_type=refresh_token", { method: "POST", body: { refresh_token: session.refresh_token }, auth: false }));
      return true;
    } catch (e) { session = null; ls.set(SKEY, null); cookie(); return false; }
  }
  async function init() {
    session = ls.get(SKEY);
    if (!(await frais())) return false;
    try {
      const u = await api("/auth/v1/user");
      session.user = { id: u.id, email: u.email }; ls.set(SKEY, session); cookie();
    } catch (e) { if (e.status === 401 || e.status === 403) { deconnecter(); return false; } }
    setInterval(frais, 5 * 60 * 1000);
    return true;
  }
  async function lire() {
    if (!(await frais()) || !session.user) return null;
    const r = await api("/rest/v1/reglages?select=donnees&outil=eq." + OUTIL + "&user_id=eq." + session.user.id);
    return r && r[0] ? r[0].donnees : null;
  }
  async function enregistrer(donnees) {
    if (!(await frais()) || !session.user) throw new Error("session expirée, reconnectez-vous");
    await api("/rest/v1/reglages?on_conflict=user_id,outil", {
      method: "POST", prefer: "resolution=merge-duplicates,return=minimal",
      body: { user_id: session.user.id, outil: OUTIL, donnees, updated_at: new Date().toISOString() },
    });
  }
  function deconnecter() {
    const t = session && session.access_token;
    session = null; ls.set(SKEY, null); cookie();
    if (t) fetch(SB + "/auth/v1/logout", { method: "POST", headers: { apikey: KEY, Authorization: "Bearer " + t } }).catch(() => {});
  }
  return { init, lire, enregistrer, deconnecter, get email() { return session && session.user && session.user.email; } };
})();
