// Publipostage : une liste d'apprentis (CSV) + un modèle de livret → un livret par apprenti.
//
// RÈGLE ABSOLUE : la liste d'apprentis ne sort pas du navigateur. Elle est lue ici, fusionnée
// ici, et le .zip est écrit ici. Aucun appel réseau, et surtout aucun appel à Claude, ne doit
// jamais recevoir une ligne de cette liste.

(function (global) {
  "use strict";

  // Les alias sont testés dans l'ordre : le premier qui correspond à une colonne gagne, et à
  // la lecture d'une ligne on prend la première colonne non vide (portable absent → fixe).
  // Les intitulés Ypareo bruts sont reconnus tels quels, sans renommer une seule colonne.
  const ALIAS = {
    groupe_ypareo: ["nomgroupeapprenant", "groupeapprenant", "nomgroupe"],
    nom: ["nomapprenant", "nom", "name"],
    prenom: ["prenomapprenant", "prenom", "firstname"],
    telephone: ["portableapprenant", "telephoneapprenant", "portable", "telephone", "tel"],
    mail: ["emailapprenant", "mail", "email", "courriel"],
    entreprise_nom: ["nomentreprise", "entreprisenom", "entreprise", "raisonsociale", "enseigneentreprise"],
    entreprise_adresse: ["entrepriseadresse", "adresseentreprise", "adr1entreprise"],
    entreprise_telephone: ["telephoneentreprise", "portableentreprise", "entreprisetelephone", "telentreprise"],
    entreprise_mail: ["emailentreprise", "entreprisemail", "mailentreprise"],
    ma_nom: ["nomma", "manom", "maitreapprentissagenom"],
    ma_prenom: ["prenomma", "maprenom", "maitreapprentissageprenom"],
    ma_telephone: ["portablema", "telephonema", "tel1ma", "tel2ma", "matelephone", "telma"],
    ma_mail: ["emailma", "mamail", "mailma"],
    representant_nom: ["nomresplegalapprenant", "nomrepresentantlegal", "nomresplegal"],
    representant_prenom: ["prenomresplegalapprenant", "prenomrepresentantlegal", "prenomresplegal"],
    representant_adresse: ["adresserepresentantlegal", "adr1courrier"],
    representant_telephone: ["portablecourrier", "telephonecourrier", "telrepresentantlegal"],
    representant_mail: ["emailcourrier", "mailrepresentantlegal"],
    contrat_debut: ["datedebcontrat", "datedebutcontrat", "debutcontrat"],
    contrat_fin: ["datefincontrat", "fincontrat"],
  };

  // Champs reconstitués depuis plusieurs colonnes (Ypareo éclate l'adresse).
  const COMPOSITES = {
    entreprise_adresse: [
      [["adr1entreprise", "entrepriseadresse", "adresseentreprise"]],
      [["adr2entreprise"]],
      [["adr3entreprise"]],
      [["adr4entreprise"]],
      [["cpentreprise"], ["villeentreprise"]],
    ],
    representant_adresse: [
      [["adr1courrier", "adresserepresentantlegal"]],
      [["adr2courrier"]],
      [["cpcourrier"], ["villecourrier"]],
    ],
  };

  // Les jetons posés dans le modèle. Les champs "classe" viennent du fichier de classe,
  // les champs "apprenti" de la liste.
  const JETONS = {
    nom: "NOM", prenom: "PRENOM", telephone: "TELEPHONE", mail: "MAIL",
    entreprise_nom: "ENTREPRISE_NOM", entreprise_adresse: "ENTREPRISE_ADRESSE",
    entreprise_telephone: "ENTREPRISE_TELEPHONE", entreprise_mail: "ENTREPRISE_MAIL",
    ma_nom: "MA_NOM", ma_prenom: "MA_PRENOM", ma_telephone: "MA_TELEPHONE", ma_mail: "MA_MAIL",
    representant_nom: "REPRESENTANT_NOM", representant_prenom: "REPRESENTANT_PRENOM",
    representant_adresse: "REPRESENTANT_ADRESSE", representant_telephone: "REPRESENTANT_TELEPHONE",
    representant_mail: "REPRESENTANT_MAIL",
    contrat_debut: "CONTRAT_DEBUT", contrat_fin: "CONTRAT_FIN",
    formateur_nom: "FORMATEUR_NOM", formateur_telephone: "FORMATEUR_TELEPHONE", formateur_mail: "FORMATEUR_MAIL",
    groupe: "GROUPE", annee: "ANNEE",
    responsable_ufa_nom: "RESPONSABLE_UFA_NOM", responsable_ufa_telephone: "RESPONSABLE_UFA_TELEPHONE",
    responsable_ufa_mail: "RESPONSABLE_UFA_MAIL",
    cvs_nom: "CVS_NOM", cvs_telephone: "CVS_TELEPHONE", cvs_mail: "CVS_MAIL",
    referent_handicap: "REFERENT_HANDICAP", referent_handicap_telephone: "REFERENT_HANDICAP_TELEPHONE",
    referent_handicap_mail: "REFERENT_HANDICAP_MAIL",
    referent_mobilite: "REFERENT_MOBILITE", referent_mobilite_telephone: "REFERENT_MOBILITE_TELEPHONE",
    referent_mobilite_mail: "REFERENT_MOBILITE_MAIL",
  };

  // Champs cochables à l'écran, groupés comme ils le sont dans le livret.
  const GROUPES = [
    ["L'apprenti·e", [
      ["nom", "NOM", true], ["prenom", "Prénom", true],
      ["telephone", "Téléphone", false], ["mail", "Mail", false],
    ]],
    ["L'entreprise", [
      ["entreprise_nom", "Raison sociale", true], ["entreprise_adresse", "Adresse", true],
      ["entreprise_telephone", "Téléphone", false], ["entreprise_mail", "Mail", false],
    ]],
    ["Le maître d'apprentissage", [
      ["ma_nom", "NOM", true], ["ma_prenom", "Prénom", true],
      ["ma_telephone", "Téléphone", false], ["ma_mail", "Mail", false],
    ]],
    // Décoché par défaut pour toutes les classes (demande explicite) : contrairement aux
    // autres groupes, ces champs ne sont pas automatiquement reportés au publipostage — il
    // faut les cocher à la main quand on veut vraiment les inclure.
    ["Le représentant légal (mineur·e·s)", [
      ["representant_nom", "NOM", false], ["representant_prenom", "Prénom", false],
      ["representant_adresse", "Adresse", false],
      ["representant_telephone", "Téléphone", false], ["representant_mail", "Mail", false],
    ]],
    ["Le contrat d'apprentissage", [
      ["contrat_debut", "Date de début", true], ["contrat_fin", "Date de fin", true],
    ]],
  ];

  const normaliser = (s) => String(s || "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().replace(/[^a-z0-9]/g, "");

  // ------------------------------------------------------------------- CSV
  // Séparateur deviné (« ; » en France, « , » ailleurs), guillemets gérés, BOM retiré.
  function lireCSV(texte) {
    texte = texte.replace(/^﻿/, "");
    const premiere = texte.split(/\r?\n/)[0] || "";
    const sep = (premiere.match(/;/g) || []).length >= (premiere.match(/,/g) || []).length ? ";" : ",";
    const lignes = [];
    let champ = "", ligne = [], entreGuillemets = false;
    for (let i = 0; i < texte.length; i++) {
      const c = texte[i];
      if (entreGuillemets) {
        if (c === '"' && texte[i + 1] === '"') { champ += '"'; i++; }
        else if (c === '"') entreGuillemets = false;
        else champ += c;
      } else if (c === '"') entreGuillemets = true;
      else if (c === sep) { ligne.push(champ); champ = ""; }
      else if (c === "\n") { ligne.push(champ); lignes.push(ligne); ligne = []; champ = ""; }
      else if (c !== "\r") champ += c;
    }
    if (champ !== "" || ligne.length) { ligne.push(champ); lignes.push(ligne); }
    return lignes.filter((l) => l.some((v) => String(v).trim() !== ""));
  }

  function premiereValeur(ligne, index, alias) {
    for (const a of alias) {
      const cols = index[a];
      if (!cols) continue;
      for (const k of cols) {
        const v = String(ligne[k] || "").trim();
        if (v) return v;
      }
    }
    return "";
  }

  function valeurComposite(ligne, index, morceaux) {
    const parties = morceaux.map((groupe) =>
      groupe.map((alias) => premiereValeur(ligne, index, alias)).filter(Boolean).join(" ")
    ).filter(Boolean);
    return parties.join(", ");
  }

  // Rend { apprentis: [{champ: valeur}], colonnesReconnues: [champ] }
  function lireListe(texte) {
    const lignes = lireCSV(texte);
    if (lignes.length < 2) throw new Error("la liste ne contient pas de ligne d'apprenti");
    const entetes = lignes[0].map(normaliser);
    const index = {};
    entetes.forEach((e, i) => { (index[e] = index[e] || []).push(i); });

    const reconnues = [];
    Object.keys(ALIAS).forEach((champ) => {
      const trouve = ALIAS[champ].some((a) => index[a])
        || (COMPOSITES[champ] || []).some((g) => g.some((al) => al.some((a) => index[a])));
      if (trouve) reconnues.push(champ);
    });

    const apprentis = lignes.slice(1).map((ligne) => {
      const a = {};
      Object.keys(ALIAS).forEach((champ) => {
        let v = premiereValeur(ligne, index, ALIAS[champ]);
        if (COMPOSITES[champ]) {
          const compose = valeurComposite(ligne, index, COMPOSITES[champ]);
          if (compose.length > v.length) v = compose;
        }
        a[champ] = v;
      });
      return a;
    });
    // Les lignes sans nom ni prénom sont écartées mais comptées : l'écran les signale.
    const sansNom = apprentis.filter((a) => !a.nom && !a.prenom);
    return { apprentis: apprentis.filter((a) => a.nom || a.prenom), sansNom, colonnesReconnues: reconnues };
  }

  // --------------------------------------------------------------- fusion
  const echapperXML = (s) => String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  const nomFichier = (a, base) => {
    const propre = (s) => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[^A-Za-z0-9]+/g, "_").replace(/^_|_$/g, "");
    return `${base}_${propre(a.nom)}_${propre(a.prenom)}.docx`;
  };

  // Remplace les jetons {{NOM}}, {{PRENOM}}… dans le XML d'un document Word, y compris
  // quand Word a fractionné un jeton entre plusieurs <w:t> — ce qui arrive dès qu'un
  // document est rouvert puis réenregistré (vérification orthographique, changement de
  // langue…, redécoupent les "runs" même sans aucune modification visible). C'est le cas
  // du "modèle déposé" ou "validé" : un livret fabriqué par ce générateur garde ses jetons
  // intacts (chacun est un seul <w:t>), mais un livret repassé par Word ne l'est plus
  // forcément. Un simple split/join sur le texte brut du XML rate ces jetons coupés — d'où
  // le passage par le texte concaténé de tous les <w:t>, plutôt que par le XML brut.
  function fusionnerXML(xml, valeurs, champsRetenus) {
    // Balise ouvrante <w:t> ou <w:t xml:space="preserve"> — jamais la forme vide <w:t/>, que
    // Word écrit parfois : l'accepter ferait avaler tout le XML jusqu'au </w:t> suivant.
    const RE_T = /<w:t(?:\s[^>]*[^\/>])?>([\s\S]*?)<\/w:t>/g;
    const runs = []; // { xmlStart, xmlEnd, catStart, catEnd, texte } — un par <w:t>
    let concat = "";
    let m;
    while ((m = RE_T.exec(xml))) {
      const texte = m[1];
      const xmlStart = m.index + m[0].indexOf(">") + 1; // juste après la balise ouvrante
      runs.push({ xmlStart, xmlEnd: xmlStart + texte.length, catStart: concat.length, catEnd: concat.length + texte.length, texte });
      concat += texte;
    }
    if (!runs.length) return xml;

    const editsParRun = new Map();
    const ajouterEdit = (i, from, to, repl) => {
      if (!editsParRun.has(i)) editsParRun.set(i, []);
      editsParRun.get(i).push({ from, to, repl });
    };
    // Recherche linéaire : quelques milliers de runs au plus dans un livret, sans incidence.
    const runDe = (pos) => {
      for (let i = 0; i < runs.length; i++) if (pos >= runs[i].catStart && pos < runs[i].catEnd) return i;
      return runs.length - 1;
    };

    const RE_JETON = /\{\{([A-Z0-9_]+)\}\}/g;
    while ((m = RE_JETON.exec(concat))) {
      const nomJeton = m[1];
      const champ = Object.keys(JETONS).find((c) => JETONS[c] === nomJeton);
      // Un jeton inconnu est vidé lui aussi : aucun {{…}} ne doit rester visible dans un livret.
      const retenu = champ && (!champsRetenus || champsRetenus.has(champ));
      const valeur = retenu ? echapperXML(valeurs[champ] || "") : "";
      const debut = m.index, fin = m.index + m[0].length;
      const iDebut = runDe(debut), iFin = runDe(fin - 1);
      if (iDebut === iFin) {
        ajouterEdit(iDebut, debut - runs[iDebut].catStart, fin - runs[iDebut].catStart, valeur);
      } else {
        // Jeton coupé entre plusieurs runs : la valeur entière va dans le premier run
        // concerné, les runs suivants perdent juste leur part du jeton (texte vidé).
        ajouterEdit(iDebut, debut - runs[iDebut].catStart, runs[iDebut].texte.length, valeur);
        for (let i = iDebut + 1; i < iFin; i++) ajouterEdit(i, 0, runs[i].texte.length, "");
        ajouterEdit(iFin, 0, fin - runs[iFin].catStart, "");
      }
    }
    if (!editsParRun.size) return xml;

    // Reconstruction : on recopie le XML d'origine tel quel (balises, espaces, runs non
    // concernés), en substituant seulement le texte des runs édités.
    let sortie = "";
    let curseur = 0;
    runs.forEach((run, i) => {
      const edits = editsParRun.get(i);
      sortie += xml.slice(curseur, run.xmlStart);
      if (edits) {
        edits.sort((a, b) => a.from - b.from);
        let texteCurseur = 0, nouveauTexte = "";
        edits.forEach((e) => {
          nouveauTexte += run.texte.slice(texteCurseur, e.from) + e.repl;
          texteCurseur = e.to;
        });
        nouveauTexte += run.texte.slice(texteCurseur);
        sortie += nouveauTexte;
      } else {
        sortie += run.texte;
      }
      curseur = run.xmlEnd;
    });
    sortie += xml.slice(curseur);
    return sortie;
  }

  // modele : Blob du .docx contenant les jetons {{NOM}}, {{PRENOM}}…
  // valeurs : { champ: valeur } pour UN apprenti, champs de classe compris
  // champsRetenus : Set des champs à reporter ; un champ absent laisse la ligne vide
  async function fusionner(modele, valeurs, champsRetenus) {
    const zip = await ZIP_SIMPLE.ouvrir(modele);
    const entrees = [];
    for (const nom of zip.noms) {
      if (nom === "word/document.xml") {
        const xml = await zip.texte(nom);
        entrees.push({ nom, donnees: fusionnerXML(xml, valeurs, champsRetenus) });
      } else {
        entrees.push({ nom, donnees: await zip.octets(nom) });
      }
    }
    return ZIP_SIMPLE.ecrire(entrees,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  }

  // Ouvre un modèle une seule fois : ses fichiers sont décompressés et gardés, seul
  // word/document.xml sera réécrit pour chaque apprenti. Rend aussi les jetons trouvés.
  async function preparer(modele) {
    const zip = await ZIP_SIMPLE.ouvrir(modele);
    const autres = [];
    let xml = null;
    for (const nom of zip.noms) {
      if (nom === "word/document.xml") xml = await zip.texte(nom);
      else {
        const donnees = await zip.octets(nom);
        autres.push({ nom, precompresse: {
          corps: await ZIP_SIMPLE.deflate(donnees), crc: ZIP_SIMPLE.crc32(donnees), taille: donnees.length,
        } });
      }
    }
    if (xml === null) throw new Error("ce fichier n'est pas un document Word (word/document.xml absent)");
    const RE_T = /<w:t(?:\s[^>]*[^\/>])?>([\s\S]*?)<\/w:t>/g;
    let texte = "", m;
    while ((m = RE_T.exec(xml))) texte += m[1];
    const jetons = [...new Set((texte.match(/\{\{[A-Z0-9_]+\}\}/g) || []))];
    return { zip, xml, autres, jetons, texte };
  }

  // Un livret (Uint8Array .docx) à partir d'un modèle préparé. Les autres parties du paquet
  // (pieds de page, styles, images, réglages) sont recopiées octet pour octet ; l'ordre des
  // entrées est celui du modèle, [Content_Types].xml en tête.
  async function produire(prep, valeurs, champsRetenus) {
    const entrees = prep.zip.noms.map((nom) => nom === "word/document.xml"
      ? { nom, donnees: fusionnerXML(prep.xml, valeurs, champsRetenus) }
      : prep.autres.find((e) => e.nom === nom));
    const blob = await ZIP_SIMPLE.ecrire(entrees,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    return new Uint8Array(await blob.arrayBuffer());
  }

  // Rend un Blob .zip contenant un livret par apprenti.
  async function publiposter(modele, apprentis, champsClasse, champsRetenus, base, avancement) {
    const entrees = [];
    for (let i = 0; i < apprentis.length; i++) {
      const valeurs = Object.assign({}, champsClasse, apprentis[i]);
      const livret = await fusionner(modele, valeurs, champsRetenus);
      // Un .docx est déjà compressé : on le range tel quel dans le zip de sortie.
      entrees.push({
        nom: nomFichier(apprentis[i], base),
        donnees: new Uint8Array(await livret.arrayBuffer()),
        compresser: false,
      });
      if (avancement) avancement(i + 1, apprentis.length);
    }
    return ZIP_SIMPLE.ecrire(entrees);
  }

  // Jetons à passer au générateur pour produire le modèle : chaque champ vaut son jeton.
  function apprentiJetons() {
    const a = {};
    for (const [champ, jeton] of Object.entries(JETONS)) a[champ] = "{{" + jeton + "}}";
    return a;
  }

  global.PUBLIPOSTAGE = { preparer, produire, normaliser, nomFichier, lireCSV, ALIAS, JETONS, GROUPES, lireListe, fusionner, fusionnerXML, publiposter, apprentiJetons };
})(typeof window !== "undefined" ? window : globalThis);
