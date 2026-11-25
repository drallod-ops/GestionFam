// Gestion SCI : serveur de la web app familiale (script lié au tableau « Gestion SCI »).
// Règle : l'app n'écrit que dans Tâches, Notes, Journal, Appareils, Profils, le dossier « 00 - Contexte »,
// et, pour Steve seulement, dans les cases sans formule de la fiche locataire (Référentiel, 5. Dettes colonne E).
// Toutes les fonctions internes finissent par « _ » : google.script.run ne peut appeler que api() et doGet().

const DOSSIER_PHOTOS = '1EW0brRLpssAL9TDRJDNaGxhloFy_HrzY'; // 00 - Contexte
const STEVE = 'Steve';
const DROITS = ['mois', 'dossiers', 'envoi', 'controle', 'photos', 'conseils', 'simple'];
const SCI_NOM = { CLS: 'SCI CLSV', CGF: 'SCI CGF', PERSO: 'Hors SCI' };
const STATUTS = ['À faire', 'En cours', 'Fait'];
const MOIS_FR = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.token = String((e && e.parameter && e.parameter.t) || '').replace(/[^A-Za-z0-9-]/g, '').slice(0, 64);
  return t.evaluate().setTitle('Gestion SCI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// Point d'entrée unique de l'interface. Les dates voyagent en texte 'AAAA-MM-JJ' (JSON), jamais en objet Date.
function api(f, argsJson) {
  const API = { init: init_, demanderAcces: demanderAcces_, taches: taches_, leMois: leMois_, conseils: conseils_,
    echeances: echeances_, locataires: locataires_, setTache: setTache_, creer: creer_, photo: photo_, addNote: addNote_,
    marquer: marquer_, admin: admin_, globale: globale_, droit: droit_, statutAppareil: statutAppareil_, annuler: annuler_,
    setFiche: setFiche_ };
  if (!Object.prototype.hasOwnProperty.call(API, f)) throw new Error('Action inconnue.');
  const r = API[f].apply(null, JSON.parse(argsJson || '[]'));
  return JSON.stringify(r === undefined ? null : r);
}

// ---------- Outils ----------
function ss_() { return SpreadsheetApp.getActive(); }
function sh_(n) { const s = ss_().getSheetByName(n); if (!s) throw new Error('Onglet introuvable : ' + n); return s; }
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function lire_(nom, row, nc) { const s = sh_(nom), n = s.getLastRow() - row + 1; return n > 0 ? s.getRange(row, 1, n, nc).getValues() : []; }
function asDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number' && v > 20000) return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5 + 12 * 36e5);
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(v).trim());
  return m ? Utilities.parseDate(m[3] + '-' + m[2] + '-' + m[1], tz_(), 'yyyy-MM-dd') : null;
}
function ymd_(v) { const d = asDate_(v); return d ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : ''; }
function versDate_(s) {
  s = String(s).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('Date invalide.');
  return Utilities.parseDate(s, tz_(), 'yyyy-MM-dd'); // minuit dans le fuseau du tableau
}
function num_(v) { return typeof v === 'number' ? v : (parseFloat(String(v).replace(/[\s\u00a0\u202f€]/g, '').replace(',', '.')) || 0); }
function txt_(v) { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; } // jamais de formule tapée
function aujourdhui_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function moisCourant_() { return aujourdhui_().slice(0, 7); }
function addMois_(mk, k) { const y = +mk.slice(0, 4), m = +mk.slice(5, 7) - 1 + k, d = y * 12 + m; return Math.floor(d / 12) + '-' + String(d % 12 + 1).padStart(2, '0'); }
function moisFr_(mk) { return MOIS_FR[+mk.slice(5, 7) - 1] + ' ' + mk.slice(0, 4); }
function enTexte_(v) { return v instanceof Date ? ymd_(v) : String(v == null ? '' : v); }
function verrou_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }

// ---------- Accès et droits ----------
function profils_() {
  const s = sh_('Profils'), v = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 2 + DROITS.length).getValues();
  const head = v[0].map(x => String(x).trim()), noms = [], droits = {}, rows = {};
  DROITS.forEach(k => droits[k] = []);
  v.slice(1).forEach((r, i) => {
    const p = String(r[0]).trim(); if (!p) return;
    noms.push(p); rows[p] = i + 2;
    DROITS.forEach(k => { const c = head.indexOf(k); if (c >= 0 && String(r[c]).trim().toLowerCase() === 'oui') droits[k].push(p); });
  });
  return { noms, droits, head, rows };
}
function appareil_(t) {
  if (!t) return null;
  const v = lire_('Appareils', 2, 5);
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === t) return { row: i + 2, prenom: String(v[i][1]).trim(), statut: String(v[i][2]).trim() };
  return null;
}
function auth_(t, as, steveSeul) {
  const a = appareil_(t);
  if (!a || a.statut !== 'Autorisé') throw new Error('Accès refusé. Rouvre l’app depuis ton lien personnel.');
  const P = profils_();
  if (P.noms.indexOf(a.prenom) < 0) throw new Error('Profil inconnu : demande à Steve.');
  const isSteve = a.prenom === STEVE;
  if (steveSeul && !isSteve) throw new Error('Réservé à Steve.');
  const profil = isSteve && as && P.noms.indexOf(as) >= 0 ? as : a.prenom; // seul Steve peut « voir l'app comme »
  return { moi: a.prenom, isSteve, profil, P, steve: profil === STEVE,
    qui: profil === a.prenom ? a.prenom : a.prenom + ' (en tant que ' + profil + ')',
    peut: (cle, p) => P.droits[cle].indexOf(p || profil) >= 0 };
}
function journal_(qui, action, onglet, ligne, colonne, avant, apres) {
  sh_('Journal').appendRow([new Date(), qui, action, onglet, txt_(ligne), txt_(colonne), txt_(avant), txt_(apres), '']);
}
function journalLu_() {
  return lire_('Journal', 2, 9).map((r, i) => ({ row: i + 2, d: r[0], le: r[0] instanceof Date ? Utilities.formatDate(r[0], tz_(), 'dd/MM HH:mm') : String(r[0]),
    jour: ymd_(r[0]), qui: String(r[1]), action: String(r[2]), onglet: String(r[3]), ligne: String(r[4]), colonne: String(r[5]),
    avant: String(r[6]), apres: String(r[7]), annule: r[8] !== '' })).filter(j => j.action);
}

// ---------- Lecture du tableau ----------
function locs_() {
  return sh_('Référentiel').getRange(11, 1, 30, 22).getValues().map((r, i) => {
    const sortie = r[19];
    return { row: i + 11, id: String(r[0]).trim(), sci: String(r[1]).trim(), bien: String(r[2]), nom: String(r[4]),
      mail: String(r[6]).trim(), tel: String(r[7]).trim(), bail: ymd_(r[8]), loyer: num_(r[9]), charges: num_(r[10]), du: num_(r[11]),
      depot: r[13] === '' ? null : num_(r[13]), dpe: String(r[14]).trim().toUpperCase(), gestion: String(r[20]).trim(), surface: r[21] === '' ? '' : num_(r[21]),
      statut: sortie === '' ? 'actif' : String(sortie).trim() === 'Ancien bail' ? 'recouvrement' : 'historique' };
  }).filter(l => /^L\d+$/.test(l.id));
}
function banque_() {
  return lire_('Banque', 5, 7).map(r => { const d = asDate_(r[0]); return d ? { jour: ymd_(d), m: ymd_(r[6] instanceof Date ? r[6] : d).slice(0, 7),
    montant: num_(r[2]), sci: String(r[3]).trim(), id: String(r[4]).trim() } : null; }).filter(Boolean);
}
function dettes_() {
  const o = {};
  sh_('5. Dettes').getRange(7, 1, 30, 10).getValues().forEach((r, i) => {
    const id = String(r[0]).trim();
    if (id) o[id] = { row: i + 7, base: r[4], dette: r[7] === '' ? null : num_(r[7]) };
  });
  return o;
}
function suivi_() {
  const s = sh_('2. Suivi du mois'), k1 = s.getRange(1, 11).getValue(), rows = {};
  s.getRange(11, 1, 30, 10).getValues().forEach(r => { const id = String(r[9]).trim(); if (id) rows[id] = { du: num_(r[3]), recu: num_(r[4]), action: String(r[7]) }; });
  return { mk: ymd_(k1).slice(0, 7), rows };
}
function revisions_() {
  const o = {};
  sh_('Révisions IRL').getRange(5, 1, 36, 10).getValues().forEach(r => { const id = String(r[0]).trim(); if (id) o[id] = { date: ymd_(r[2]), statut: String(r[9]) }; });
  return o;
}
function notes_() {
  const o = {};
  lire_('Notes', 2, 4).forEach(r => { const s = String(r[2]); if (!s || !String(r[3])) return;
    (o[s] = o[s] || []).unshift({ le: ymd_(r[0]), qui: String(r[1]), texte: String(r[3]) }); });
  return o;
}
function tachesLues_() {
  return lire_('Tâches', 2, 11).map((r, i) => ({ row: i + 2, id: String(r[0]).trim(), dossier: String(r[1]), tache: String(r[2]), resp: String(r[3]).trim(),
    ech: ymd_(r[4]), statut: STATUTS.indexOf(String(r[5]).trim()) >= 0 ? String(r[5]).trim() : 'À faire', loc: String(r[6]), rem: String(r[7]), faite: ymd_(r[10]) }))
    .filter(t => t.id);
}
function modeles_() {
  const o = {};
  lire_('Modèles de tâches', 2, 5).forEach(r => { const m = String(r[0]).trim(); if (!m || !String(r[2]).trim()) return;
    (o[m] = o[m] || []).push({ n: num_(r[1]), etape: String(r[2]).trim(), resp: String(r[3]).trim(), rem: String(r[4]).trim() }); });
  Object.keys(o).forEach(m => o[m].sort((a, b) => a.n - b.n));
  return o;
}
// Seule l'étape la plus basse non faite d'un dossier est « en cours » (même règle que l'interface).
function actuelles_(T) {
  const min = {};
  T.forEach(t => { const n = /^(\d+)\.\s/.exec(t.tache); if (n && t.statut !== 'Fait') min[t.dossier] = Math.min(min[t.dossier] || 1e9, +n[1]); });
  return T.filter(t => { if (t.statut === 'Fait') return false; const n = /^(\d+)\.\s/.exec(t.tache); return !n || +n[1] === min[t.dossier]; });
}
function envois_(J) {
  const m = {};
  J.forEach(j => { if (j.action === 'Envoyé' && j.onglet === 'Le mois' && !j.annule) m[j.ligne + '|' + j.colonne] = { qui: j.qui, le: j.jour }; });
  return m;
}
function pub_(l) { return { id: l.id, nom: l.nom, sci: l.sci, bien: l.bien, du: l.du, loyer: l.loyer, charges: l.charges, dpe: l.dpe, gestion: l.gestion }; }
function echeance_(l) { return /century/i.test(l.gestion) ? 25 : 10; } // jour à partir duquel un loyer non reçu est en retard
// Couleur d'un mois : vert payé, orange partiel, rouge impayé, gris à venir ou pas de données.
function couleur_(l, du, recu, reste, mk, cur, debut) {
  if (!debut || mk < debut || (l.bail && l.bail.slice(0, 7) > mk)) return 'gris';
  if (l.statut === 'recouvrement') return recu > 0 ? 'vert' : 'gris';
  if (reste <= 0.005) return 'vert';
  if (recu > 0) return 'orange';
  if (mk === cur && +aujourdhui_().slice(8, 10) < echeance_(l) && reste <= du + 0.005) return 'gris';
  return 'rouge';
}

// ---------- Accès ----------
function init_(t) {
  const a = appareil_(t);
  if (!a) return { state: 'nouveau' };
  if (a.statut !== 'Autorisé') return { state: a.statut };
  const P = profils_();
  if (P.noms.indexOf(a.prenom) < 0) return { state: 'Refusé' };
  return { state: 'ok', prenom: a.prenom, isSteve: a.prenom === STEVE, profils: P.noms, droits: P.droits };
}
function demanderAcces_(nom) {
  nom = String(nom || '').trim();
  const P = profils_();
  if (P.noms.indexOf(nom) < 0) throw new Error('Prénom inconnu. Écris-le exactement comme dans la liste (ex. : Laetitia).');
  return verrou_(() => {
    const enAttente = lire_('Appareils', 2, 5).filter(r => String(r[2]) === 'En attente').length;
    if (enAttente >= 10) throw new Error('Trop de demandes en attente. Demande à Steve.'); // limite les demandes en masse
    const token = Utilities.getUuid();
    sh_('Appareils').appendRow([token, nom, 'En attente', new Date(), '']);
    journal_(nom, 'Demande d’accès', 'Appareils', nom, 'Statut', '', 'En attente');
    return ScriptApp.getService().getUrl() + '?t=' + token;
  });
}

// ---------- Tâches et dossiers ----------
function taches_(t, as) {
  const A = auth_(t, as), p = A.profil, T = tachesLues_(), L = locs_();
  const nomLoc = id => (L.find(l => l.id === id) || {}).nom || id;
  const nomTache = id => { const x = T.find(y => y.id === id); return x ? x.tache.replace(/^\d+\.\s*/, '') : 'une tâche'; };
  const famille = journalLu_().reverse().filter(j => !j.annule && j.qui.split(' ')[0] !== p).map(j => {
    const q = j.qui.split(' ')[0];
    if (j.action === 'Modification' && j.colonne === 'Statut' && j.apres === 'Fait') return { txt: q + ' a terminé « ' + nomTache(j.ligne) + ' »', le: j.le.slice(0, 5) };
    if (j.action === 'Note') return { txt: q + ' a laissé une note sur ' + nomLoc(j.ligne), le: j.le.slice(0, 5) };
    if (j.action === 'Envoyé') return { txt: q + ' a envoyé le courrier de ' + nomLoc(j.ligne), le: j.le.slice(0, 5) };
    return null;
  }).filter(Boolean).slice(0, 5);
  return { profil: p, isSteve: A.steve, profils: A.P.noms, taches: T.map(x => { const o = Object.assign({}, x); delete o.row; return o; }),
    notes: notes_(), famille, modeles: modeles_(), locataires: L.filter(l => l.statut !== 'historique').map(l => l.nom) };
}
const CHAMPS_ = { 'Responsable': 4, 'Échéance': 5, 'Statut': 6, 'Remarques': 8 };
const CLES_ = { 4: 'resp', 5: 'ech', 6: 'statut', 8: 'rem' };
function setTache_(t, as, id, champ, val) {
  const A = auth_(t, as), p = A.profil, col = CHAMPS_[champ];
  if (!col) throw new Error('Champ non modifiable.');
  val = String(val == null ? '' : val).trim().slice(0, 500);
  return verrou_(() => {
    const x = tachesLues_().find(y => y.id === id);
    if (!x) throw new Error('Tâche introuvable.');
    const aMoi = x.resp === p || x.resp === 'Tous';
    if (champ === 'Responsable') {
      if (A.P.noms.concat('Claude', 'Tous', '').indexOf(val) < 0) throw new Error('Responsable inconnu.');
      if (!A.steve && !(!x.resp && val === p)) throw new Error('Cette tâche est attribuée à ' + (x.resp || 'personne') + '.');
    } else if (champ === 'Échéance') {
      if (!A.steve) throw new Error('Seul Steve change les dates.');
      if (val) versDate_(val);
    } else if (champ === 'Statut') {
      if (STATUTS.indexOf(val) < 0) throw new Error('Statut inconnu.');
      if (!A.steve && !aMoi) throw new Error('Cette tâche est attribuée à ' + (x.resp || 'personne') + '.');
    } else if (!A.steve && !aMoi) throw new Error('Cette tâche est attribuée à ' + (x.resp || 'personne') + '.');
    const s = sh_('Tâches'), avant = x[CLES_[col]];
    if (avant === val) return;
    s.getRange(x.row, col).setValue(col === 5 ? (val ? versDate_(val) : '') : txt_(val));
    if (col === 5) s.getRange(x.row, 5).setNumberFormat('dd/mm/yyyy');
    if (col === 6) s.getRange(x.row, 11).setValue(val === 'Fait' ? new Date() : '').setNumberFormat('dd/mm/yyyy');
    journal_(A.qui, 'Modification', 'Tâches', id, champ, avant, val);
  });
}
function creer_(t, as, n) {
  const A = auth_(t, as);
  if (!A.peut('dossiers') && !A.steve) throw new Error('Tu n’as pas le droit de créer un dossier.');
  n = n || {};
  const dossier = String(n.dossier || '').trim().slice(0, 120);
  if (!dossier) throw new Error('Indique le nom du dossier.');
  const M = modeles_(), ok = A.P.noms.concat('Claude', 'Tous', '');
  let etapes;
  if (n.modele) {
    if (!M[n.modele]) throw new Error('Modèle inconnu.');
    etapes = M[n.modele].map(e => ({ tache: e.n + '. ' + e.etape, resp: e.resp, rem: e.rem }));
  } else {
    const tache = String(n.tache || '').trim().slice(0, 200);
    if (!tache) throw new Error('Écris la tâche.');
    etapes = [{ tache, resp: String(n.resp || ''), rem: '' }];
  }
  etapes.forEach(e => { if (ok.indexOf(e.resp) < 0) e.resp = ''; });
  const loc = String(n.loc || '').slice(0, 120), ech = n.ech ? versDate_(n.ech) : '';
  return verrou_(() => {
    const s = sh_('Tâches');
    let max = tachesLues_().reduce((a, x) => Math.max(a, +((/^T(\d+)$/.exec(x.id) || [0, 0])[1])), 0);
    const rows = etapes.map(e => ['T' + String(++max).padStart(3, '0'), txt_(dossier), txt_(e.tache), e.resp, ech, 'À faire', txt_(loc), txt_(e.rem), new Date(), A.qui, '']);
    const r0 = s.getLastRow() + 1;
    s.getRange(r0, 1, rows.length, 11).setValues(rows);
    s.getRange(r0, 5, rows.length, 1).setNumberFormat('dd/mm/yyyy');
    s.getRange(r0, 9, rows.length, 1).setNumberFormat('dd/mm/yyyy');
    rows.forEach(r => journal_(A.qui, 'Création', 'Tâches', r[0], '', '', r[2]));
    return rows.length;
  });
}
function photo_(t, as, id, b64, mime) {
  const A = auth_(t, as);
  if (!A.peut('photos')) throw new Error('Tu n’as pas le droit d’envoyer des photos.');
  if (mime !== 'image/jpeg' || typeof b64 !== 'string' || !b64 || b64.length > 12e6) throw new Error('Photo refusée (format ou taille).');
  const x = tachesLues_().find(y => y.id === id);
  if (!x) throw new Error('Tâche introuvable.');
  const nom = ('Photo_' + x.dossier + '_' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd_HHmmss') + '.jpg').replace(/[\\/:*?"<>|]/g, '-');
  DriveApp.getFolderById(DOSSIER_PHOTOS).createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', nom));
  verrou_(() => journal_(A.qui, 'Photo', 'Drive', id, nom, '', ''));
}
function addNote_(t, as, sujet, texte) {
  const A = auth_(t, as);
  sujet = String(sujet || '').trim().slice(0, 200); texte = String(texte || '').trim().slice(0, 3000);
  if (!sujet) throw new Error('Sujet manquant.');
  if (!texte) throw new Error('Écris ta note.');
  verrou_(() => { sh_('Notes').appendRow([new Date(), A.qui, txt_(sujet), txt_(texte)]); journal_(A.qui, 'Note', 'Notes', sujet, '', '', texte.slice(0, 200)); });
}

// ---------- Le mois ----------
function leMois_(t, as, mk) {
  const A = auth_(t, as), p = A.profil, cur = moisCourant_();
  if (!A.peut('mois') && !A.steve) throw new Error('Tu n’as pas accès au mois.');
  if (mk && !/^\d{4}-\d{2}$/.test(mk)) throw new Error('Mois invalide.');
  mk = mk && mk < cur ? mk : cur;
  const L = locs_().filter(l => l.statut === 'actif'), B = banque_(), E = envois_(journalLu_());
  const recu = (id, m) => B.filter(b => b.id === id && b.m === m).reduce((a, b) => a + b.montant, 0);
  const debut = B.reduce((a, b) => !a || b.m < a ? b.m : a, '');
  const base = { mk, mois: moisFr_(mk), releve: B.reduce((a, b) => b.jour > a ? b.jour : a, ''), moi: p,
    envoyeur: A.P.droits.envoi.join(' ou ') || 'personne', notes: notes_() };
  if (mk < cur) { // mois passé : calculé depuis Banque, sans toucher au choix du mois du tableau
    return Object.assign(base, { courant: false, passe: true, controle: false, peutEnvoyer: false,
      locs: L.map(l => { const r = recu(l.id, mk), c = couleur_(l, l.du, r, l.du - r, mk, cur, debut);
        return Object.assign(pub_(l), { recu: r, dette: null, couleur: c, action: c === 'gris' ? 'pas de données' : c === 'vert' ? 'Quittance' : 'Relance', envoye: E[l.id + '|' + mk] || null }); }) });
  }
  // Mois en cours : couleurs tirées de « 2. Suivi du mois » et « 5. Dettes » si le tableau affiche bien ce mois
  const S = suivi_(), D = dettes_(), aJour = S.mk === cur;
  return Object.assign(base, { courant: aJour, passe: false, controle: A.peut('controle'), peutEnvoyer: A.peut('envoi'),
    locs: L.map(l => {
      const s = aJour ? S.rows[l.id] : null, d = aJour ? D[l.id] : null;
      const du = s ? s.du : l.du, r = s ? s.recu : recu(l.id, cur);
      const dette = d && d.base !== '' ? d.dette : null; // dette passée pas encore saisie : inconnue
      const c = couleur_(l, du, r, dette != null ? dette : du - r, cur, cur, debut);
      return Object.assign(pub_(l), { du, recu: r, dette, couleur: c,
        action: s ? s.action : (r >= du ? 'Quittance' : r > 0 ? 'Reçu partiel + relance' : 'Relance à décider'), envoye: E[l.id + '|' + cur] || null });
    }) });
}
function marquer_(t, as, id, mk) {
  const A = auth_(t, as);
  if (!A.peut('envoi')) throw new Error('Les envois sont marqués par ' + (A.P.droits.envoi.join(' ou ') || 'personne') + '.');
  if (!/^\d{4}-\d{2}$/.test(String(mk)) || mk > moisCourant_()) throw new Error('Mois invalide.');
  if (!locs_().some(l => l.id === id && l.statut === 'actif')) throw new Error('Locataire inconnu.');
  verrou_(() => { if (!envois_(journalLu_())[id + '|' + mk]) journal_(A.qui, 'Envoyé', 'Le mois', id, mk, '', 'Envoyé'); });
}

// ---------- Locataires, échéances, conseils ----------
function locataires_(t, as) {
  const A = auth_(t, as);
  if (!A.peut('dossiers') && !A.steve) throw new Error('Tu n’as pas accès aux fiches.');
  const cur = moisCourant_(), B = banque_(), R = revisions_(), S = suivi_(), D = dettes_();
  const debut = B.reduce((a, b) => !a || b.m < a ? b.m : a, '');
  const locs = locs_().filter(l => l.statut !== 'historique').map(l => {
    const histo = [];
    for (let i = 11; i >= 0; i--) {
      const m = addMois_(cur, -i), r = B.filter(b => b.id === l.id && b.m === m).reduce((a, b) => a + b.montant, 0);
      histo.push(couleur_(l, l.du, r, l.du - r, m, cur, debut));
    }
    const d = D[l.id] || {};
    return Object.assign(pub_(l), { sciNom: SCI_NOM[l.sci] || l.sci, statut: l.statut, surface: l.surface, bail: l.bail, depot: l.depot,
      tel: l.tel, mail: l.mail, histo, revision: l.bail ? prochaineRev_((R[l.id] || {}).date, aujourdhui_()) : '',
      dette: S.mk === cur && d.base !== '' && d.base !== undefined ? d.dette : null, dette31: d.base === undefined ? null : d.base === '' ? '' : num_(d.base) });
  });
  return { notes: notes_(), locs, modifiable: A.isSteve && A.steve };
}
// Champs que Steve peut corriger depuis la fiche : [onglet, colonne, type]
const FICHE_ = { mail: ['Référentiel', 7, 'texte'], tel: ['Référentiel', 8, 'texte'], entree: ['Référentiel', 9, 'date'],
  depot: ['Référentiel', 14, 'nombre'], dpe: ['Référentiel', 15, 'dpe'], gestion: ['Référentiel', 21, 'texte'],
  surface: ['Référentiel', 22, 'nombre'], dette31: ['5. Dettes', 5, 'nombre'] };
function convertir_(type, val) {
  const s = String(val == null ? '' : val).trim();
  if (s === '') return '';
  if (type === 'date') return versDate_(s);
  if (type === 'nombre') { const n = parseFloat(s.replace(/[\s\u00a0\u202f€]/g, '').replace(',', '.')); if (isNaN(n)) throw new Error('Nombre invalide.'); return Math.round(n * 100) / 100; }
  if (type === 'dpe') { if (!/^[A-G]$/i.test(s)) throw new Error('DPE : une lettre de A à G.'); return s.toUpperCase(); }
  return txt_(s.slice(0, 200));
}
function celluleFiche_(id, champ) {
  const f = FICHE_[champ];
  if (!f) throw new Error('Champ non modifiable.');
  const row = f[0] === '5. Dettes' ? (dettes_()[id] || {}).row : (locs_().find(l => l.id === id) || {}).row;
  if (!row) throw new Error('Locataire introuvable dans « ' + f[0] + ' ».');
  const cell = sh_(f[0]).getRange(row, f[1]);
  if (cell.getFormula()) throw new Error('Cette case contient une formule : modification refusée.');
  return { f, cell };
}
function setFiche_(t, id, champ, val) {
  const A = auth_(t, null, true);
  return verrou_(() => {
    const c = celluleFiche_(id, champ), avant = enTexte_(c.cell.getValue()), v = convertir_(c.f[2], val);
    if (enTexte_(v) === avant) return;
    c.cell.setValue(v);
    if (c.f[2] === 'date') c.cell.setNumberFormat('dd/mm/yyyy');
    journal_(A.qui, 'Fiche', c.f[0], id, champ, avant, enTexte_(v));
  });
}
function prochaineRev_(d, auj) { while (d && d < auj) d = (+d.slice(0, 4) + 1) + d.slice(4); return d || ''; } // révision annuelle
function echeances_(t, as) {
  const A = auth_(t, as);
  if (!A.peut('dossiers') && !A.steve) throw new Error('Tu n’as pas accès aux échéances.');
  const auj = aujourdhui_(), R = revisions_(), ev = [], an = +auj.slice(0, 4);
  const prochaine = mmjj => (an + '-' + mmjj) >= auj ? an + '-' + mmjj : (an + 1) + '-' + mmjj;
  locs_().filter(l => l.statut === 'actif').forEach(l => {
    const r = R[l.id];
    const rd = prochaineRev_(r && r.date, auj);
    if (rd && l.bail) ev.push(/[FG]/.test(l.dpe) ? { d: rd, i: '🔒', t: 'Pas de révision (DPE ' + l.dpe + ')', s: l.nom + ' : loyer gelé' }
      : { d: rd, i: '📈', t: 'Révision de loyer', s: l.nom + ', ' + l.bien + (l.dpe ? '' : ' (DPE à renseigner)') });
    if (l.bail) { // bail reconduit tacitement tous les 3 ans
      let y = +l.bail.slice(0, 4), fin;
      do { y += 3; fin = y + l.bail.slice(4); } while (fin <= auj);
      ev.push({ d: fin, i: '📄', t: 'Fin de bail (reconduction tacite sauf congé)', s: l.nom + ', ' + l.bien });
    }
  });
  ev.push({ d: prochaine('01-15'), i: '🛡️', t: 'Attestations d’assurance des locataires', s: 'Lancer le modèle « Attestations annuelles »' });
  ev.push({ d: prochaine('05-15'), i: '🧾', t: 'Déclarations 2072 des deux SCI', s: 'Claude prépare, date limite à vérifier' });
  ev.push({ d: prochaine('05-15'), i: '🏠', t: '1 place Alphonse Daudet : déclaration des parents', s: 'Hors SCI : déclaré directement, pas dans la 2072' });
  ev.push({ d: prochaine('10-01'), i: '🧮', t: 'Régularisation des charges', s: 'Lancer le modèle, photos des avis de taxe foncière' });
  const t0 = versDate_(auj).getTime();
  return ev.map(e => Object.assign(e, { j: Math.round((versDate_(e.d).getTime() - t0) / 864e5) }))
    .filter(e => e.j >= 0 && e.j <= 365).sort((a, b) => a.j - b.j);
}
function conseils_(t) {
  const A = auth_(t, null);
  if (!A.peut('conseils')) throw new Error('Tu n’as pas accès aux conseils.');
  const g = {};
  lire_('Conseils', 2, 4).forEach(r => {
    if (!String(r[3]).trim()) return;
    const k = r[1] instanceof Date ? ymd_(r[1]).slice(0, 7) : String(r[1]);
    const x = g[k] = g[k] || { k, mois: /^\d{4}-\d{2}$/.test(k) ? moisFr_(k).replace(/^./, c => c.toUpperCase()) : k, le: '', items: [] };
    x.le = ymd_(r[0]) > x.le ? ymd_(r[0]) : x.le;
    x.items.push([String(r[2]), String(r[3])]);
  });
  return Object.keys(g).sort().reverse().map(k => g[k]);
}

// ---------- Pilotage (Steve) ----------
function globale_(t) {
  auth_(t, null, true);
  const an = aujourdhui_().slice(0, 4), cur = moisCourant_(), B = banque_(), L = locs_();
  const som = sci => B.filter(b => b.sci === sci && b.montant > 0 && /^L\d+$/.test(b.id) && b.jour.slice(0, 4) === an).reduce((a, b) => a + b.montant, 0);
  const retard = {}, auj = aujourdhui_();
  actuelles_(tachesLues_()).forEach(x => { if (x.ech && x.ech < auj) { const k = x.resp || 'Sans responsable'; retard[k] = (retard[k] || 0) + 1; } });
  const D = dettes_(), courant = suivi_().mk === cur;
  const impayes = courant ? Object.keys(D).filter(id => D[id].dette > 0.005).map(id => ({
    nom: ((L.find(l => l.id === id) || {}).nom || id) + (D[id].base === '' ? ' (dette passée à saisir)' : ''), dette: D[id].dette })).sort((a, b) => b.dette - a.dette) : [];
  return { retard, courant, encaisse: [['SCI CLSV', som('CLS')], ['SCI CGF', som('CGF')], ['Hors SCI', som('PERSO')]], impayes };
}
function admin_(t) {
  auth_(t, null, true);
  const appareils = lire_('Appareils', 2, 5).map((r, i) => ({ row: i + 2, prenom: String(r[1]), statut: String(r[2]), le: ymd_(r[3]) }))
    .filter(a => a.statut === 'En attente' || a.statut === 'Autorisé');
  const journal = journalLu_().reverse().slice(0, 30).map(j => ({ row: j.row, le: j.le, qui: j.qui, action: j.action, onglet: j.onglet,
    ligne: j.ligne, colonne: j.colonne, avant: j.avant, apres: j.apres, annule: j.annule,
    annulable: !j.annule && ((j.action === 'Modification' && j.onglet === 'Tâches') || (j.action === 'Envoyé' && j.onglet === 'Le mois') || j.action === 'Fiche') }));
  return { appareils, journal };
}
function droit_(t, cle, p, oui) {
  const A = auth_(t, null, true);
  if (DROITS.indexOf(cle) < 0) throw new Error('Droit inconnu.');
  if (A.P.noms.indexOf(p) < 0) throw new Error('Profil inconnu.');
  return verrou_(() => {
    const P = profils_(), col = P.head.indexOf(cle) + 1;
    if (!col) throw new Error('Colonne « ' + cle + ' » absente de Profils.');
    const avant = P.droits[cle].indexOf(p) >= 0 ? 'oui' : '';
    if (avant !== (oui ? 'oui' : '')) {
      sh_('Profils').getRange(P.rows[p], col).setValue(oui ? 'oui' : '');
      journal_(A.qui, 'Droit ' + (oui ? 'donné' : 'retiré'), 'Profils', p, cle, avant, oui ? 'oui' : '');
    }
    return profils_().droits;
  });
}
function statutAppareil_(t, row, st) {
  const A = auth_(t, null, true);
  if (['Autorisé', 'Refusé', 'Révoqué'].indexOf(st) < 0) throw new Error('Statut inconnu.');
  row = Math.floor(Number(row));
  return verrou_(() => {
    const s = sh_('Appareils');
    if (!(row >= 2 && row <= s.getLastRow())) throw new Error('Ligne introuvable.');
    const r = s.getRange(row, 1, 1, 5).getValues()[0];
    s.getRange(row, 3).setValue(st);
    s.getRange(row, 5).setValue(new Date());
    journal_(A.qui, 'Accès : ' + st, 'Appareils', String(r[1]), 'Statut', String(r[2]), st);
  });
}
function annuler_(t, row) {
  const A = auth_(t, null, true);
  return verrou_(() => {
    const j = journalLu_().find(x => x.row === Math.floor(Number(row)));
    if (!j) throw new Error('Action introuvable.');
    if (j.annule) throw new Error('Déjà annulé.');
    if (j.action === 'Modification' && j.onglet === 'Tâches') {
      const x = tachesLues_().find(y => y.id === j.ligne), col = CHAMPS_[j.colonne];
      if (!x || !col) throw new Error('Tâche introuvable.');
      if (String(x[CLES_[col]]) !== j.apres) throw new Error('La valeur a changé depuis : annulation impossible.');
      const s = sh_('Tâches');
      s.getRange(x.row, col).setValue(col === 5 ? (j.avant ? versDate_(j.avant) : '') : txt_(j.avant));
      if (col === 6) s.getRange(x.row, 11).setValue(j.avant === 'Fait' ? new Date() : '');
    } else if (j.action === 'Fiche') {
      const c = celluleFiche_(j.ligne, j.colonne);
      if (enTexte_(c.cell.getValue()) !== enTexte_(convertir_(c.f[2], j.apres))) throw new Error('La valeur a changé depuis : annulation impossible.');
      c.cell.setValue(convertir_(c.f[2], j.avant));
    } else if (!(j.action === 'Envoyé' && j.onglet === 'Le mois')) throw new Error('Cette action ne peut pas être annulée.');
    sh_('Journal').getRange(j.row, 9).setValue(new Date());
    journal_(A.qui, 'Annulation', j.onglet, j.ligne, j.colonne, j.apres, j.avant);
  });
}
