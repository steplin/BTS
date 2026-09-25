// Lire et écrire des fichiers .zip sans bibliothèque.
//
// Un .docx et un .xlsx sont des zip ; le publipostage doit en ouvrir un, y remplacer un
// fichier XML, et le réécrire — puis empaqueter le tout dans un zip de livrets.
// Les navigateurs savent compresser et décompresser nativement (Compression/DecompressionStream) :
// il ne manquait que l'enveloppe zip, qui tient en quelques dizaines de lignes.

(function (global) {
  "use strict";

  const TABLE_CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(octets) {
    let c = 0xffffffff;
    for (let i = 0; i < octets.length; i++) c = TABLE_CRC[(c ^ octets[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  async function deflate(octets) {
    const flux = new Blob([octets]).stream().pipeThrough(new CompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(flux).arrayBuffer());
  }
  async function inflate(octets) {
    const flux = new Blob([octets]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(flux).arrayBuffer());
  }

  // ------------------------------------------------------------------ lecture
  async function ouvrir(source) {
    const buffer = source instanceof ArrayBuffer ? source : await source.arrayBuffer();
    const vue = new DataView(buffer);
    const octets = new Uint8Array(buffer);
    let fin = -1;
    for (let i = octets.length - 22; i >= 0 && i > octets.length - 66000; i--) {
      if (vue.getUint32(i, true) === 0x06054b50) { fin = i; break; }
    }
    if (fin < 0) throw new Error("archive illisible");
    const nb = vue.getUint16(fin + 10, true);
    let pos = vue.getUint32(fin + 16, true);

    const entrees = [];
    for (let n = 0; n < nb; n++) {
      if (vue.getUint32(pos, true) !== 0x02014b50) break;
      const methode = vue.getUint16(pos + 10, true);
      const taille = vue.getUint32(pos + 20, true);
      const lgNom = vue.getUint16(pos + 28, true);
      const lgExtra = vue.getUint16(pos + 30, true);
      const lgCom = vue.getUint16(pos + 32, true);
      const local = vue.getUint32(pos + 42, true);
      const nom = new TextDecoder().decode(octets.subarray(pos + 46, pos + 46 + lgNom));
      const lgNomL = vue.getUint16(local + 26, true);
      const lgExtraL = vue.getUint16(local + 28, true);
      const debut = local + 30 + lgNomL + lgExtraL;
      entrees.push({ nom, methode, brut: octets.subarray(debut, debut + taille) });
      pos += 46 + lgNom + lgExtra + lgCom;
    }

    const parNom = {};
    entrees.forEach((e) => { parNom[e.nom] = e; });
    return {
      noms: entrees.map((e) => e.nom),
      async octets(nom) {
        const e = parNom[nom];
        if (!e) return null;
        return e.methode === 0 ? e.brut : inflate(e.brut);
      },
      async texte(nom) {
        const o = await this.octets(nom);
        return o ? new TextDecoder("utf-8").decode(o) : null;
      },
    };
  }

  // ------------------------------------------------------------------ écriture
  // entrees : [{ nom, donnees: Uint8Array | string, compresser?: bool }]
  async function ecrire(entrees, typeMime) {
    const morceaux = [], central = [];
    let position = 0;

    for (const e of entrees) {
      const nomOctets = new TextEncoder().encode(e.nom);
      // Entrée déjà compressée (préparée une fois, recopiée dans chaque livret) :
      // { nom, precompresse: { corps, crc, taille } }.
      let corps, crc, compresser, taille;
      if (e.precompresse) {
        ({ corps, crc, taille } = e.precompresse);
        compresser = true;
      } else {
        const donnees = typeof e.donnees === "string"
          ? new TextEncoder().encode(e.donnees)
          : new Uint8Array(e.donnees);
        compresser = e.compresser !== false;
        corps = compresser ? await deflate(donnees) : donnees;
        crc = crc32(donnees);
        taille = donnees.length;
      }

      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);
      local.setUint16(4, 20, true);              // version minimale
      local.setUint16(6, 0x0800, true);          // noms en UTF-8
      local.setUint16(8, compresser ? 8 : 0, true);
      local.setUint16(10, 0, true);              // heure
      local.setUint16(12, 0x0021, true);         // date (1er janvier 1980)
      local.setUint32(14, crc, true);
      local.setUint32(18, corps.length, true);
      local.setUint32(22, taille, true);
      local.setUint16(26, nomOctets.length, true);
      local.setUint16(28, 0, true);
      morceaux.push(new Uint8Array(local.buffer), nomOctets, corps);

      const entete = new DataView(new ArrayBuffer(46));
      entete.setUint32(0, 0x02014b50, true);
      entete.setUint16(4, 20, true);
      entete.setUint16(6, 20, true);
      entete.setUint16(8, 0x0800, true);
      entete.setUint16(10, compresser ? 8 : 0, true);
      entete.setUint16(12, 0, true);
      entete.setUint16(14, 0x0021, true);
      entete.setUint32(16, crc, true);
      entete.setUint32(20, corps.length, true);
      entete.setUint32(24, taille, true);
      entete.setUint16(28, nomOctets.length, true);
      entete.setUint32(42, position, true);
      central.push(new Uint8Array(entete.buffer), nomOctets);

      position += 30 + nomOctets.length + corps.length;
    }

    const tailleCentral = central.reduce((t, m) => t + m.length, 0);
    const queue = new DataView(new ArrayBuffer(22));
    queue.setUint32(0, 0x06054b50, true);
    queue.setUint16(8, entrees.length, true);
    queue.setUint16(10, entrees.length, true);
    queue.setUint32(12, tailleCentral, true);
    queue.setUint32(16, position, true);

    return new Blob([...morceaux, ...central, new Uint8Array(queue.buffer)],
      { type: typeMime || "application/zip" });
  }

  global.ZIP_SIMPLE = { ouvrir, ecrire, crc32, deflate };
})(typeof window !== "undefined" ? window : globalThis);
