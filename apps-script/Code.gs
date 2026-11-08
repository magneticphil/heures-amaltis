/**
 * Heures — Amaltis (Philippe Brunelle)
 * Pointage entrée/sortie, cumul de la paie, feuille de temps Excel remplie automatiquement
 * et brouillon Gmail pour Marc un samedi sur deux.
 *
 * Feuilles du classeur : « Pointages », « Paies », « Paramètres ».
 * Écran : GitHub Pages, qui appelle doPost({fn, args}).
 */

const TZ = 'America/Toronto';
const F_POINT = 'Pointages', F_PAIES = 'Paies', F_PARAM = 'Paramètres', F_CONGE = 'Congés';
const H_CONGE = ['ID', 'Date', 'Heures', 'Requête', 'Modifié le'];
const H_POINT = ['ID', 'Date', 'Entrée', 'Sortie', 'Requête', 'Modifié le'];
const H_PAIES = ['Début', 'Fin', 'Banque reportée', 'Heures normales', 'Heures fériées', 'Total', 'Nouvelle banque', 'Fichier', 'Brouillon', 'Créé le'];
const JOURS = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
const MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/* ====================== Porte d'entrée ====================== */

// Seules ces fonctions sont joignables depuis l'écran.
function fns_() {
  return { etat, connexion, getData, pointer, ajouterQuart, modifierQuart, supprimerQuart, definirConge, preparerPaie };
}

function doGet() {
  return ContentService.createTextOutput('Heures — Amaltis : service en ligne.');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const f = Object.prototype.hasOwnProperty.call(fns_(), req.fn) ? fns_()[req.fn] : null;
    if (!f) throw new Error('demande invalide');
    out = { ok: true, data: f.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ====================== Sécurité : code personnel ====================== */

function hash_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'heures-amaltis:' + s, Utilities.Charset.UTF_8));
}

/** Vrai si aucun code n'est encore défini (premier lancement). */
function etat() {
  return { codeDefini: !!PropertiesService.getScriptProperties().getProperty('CODE') };
}

function verifier_(code) {
  const cache = CacheService.getScriptCache();
  const echecs = Number(cache.get('echecs') || 0);
  if (echecs >= 5) throw new Error('Trop d’essais. Réessayez dans 15 minutes.');
  const attendu = PropertiesService.getScriptProperties().getProperty('CODE');
  if (!attendu) throw new Error('Code non défini');
  if (!code || hash_(String(code)) !== attendu) {
    cache.put('echecs', String(echecs + 1), 900);
    throw new Error('Code invalide');
  }
}

/** Premier lancement : le code choisi devient le code de l'app. Ensuite : vérification. */
function connexion(code) {
  code = String(code || '').trim();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('CODE')) {
    if (!/^\S{4,40}$/.test(code)) throw new Error('Choisissez un code d’au moins 4 caractères.');
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try { if (!props.getProperty('CODE')) props.setProperty('CODE', hash_(code)); } finally { lock.releaseLock(); }
  }
  verifier_(code);
  return getData(code);
}

/* ====================== Paramètres et feuilles ====================== */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function params_() {
  const sh = ss_().getSheetByName(F_PARAM);
  const p = {};
  sh.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0]) p[r[0].trim()] = String(r[1]).trim(); });
  return {
    reference: p['Début de paie de référence'],          // un dimanche de début de paie (aaaa-mm-jj)
    payees: Number(String(p['Heures payées par paie']).replace(',', '.')) || 80,
    courriel: p['Courriel de Marc'],
    dossier: p['Dossier Drive (ID)'],
    modele: p['Modèle Excel (ID)'],
    nom: p['Nom de l’employé'] || 'Philippe Brunelle'
  };
}

/** À lancer une fois depuis l'éditeur : crée les feuilles, les paramètres et le déclencheur du samedi. */
function installer() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  let sh = ss.getSheetByName(F_POINT) || ss.insertSheet(F_POINT);
  if (sh.getLastRow() === 0) sh.appendRow(H_POINT);
  sh.getRange('A:F').setNumberFormat('@'); sh.setFrozenRows(1);
  sh = ss.getSheetByName(F_PAIES) || ss.insertSheet(F_PAIES);
  if (sh.getLastRow() === 0) {
    sh.appendRow(H_PAIES);
    // Dernière paie faite à la main (25 oct. au 7 nov. 2026) : point de départ de la banque.
    sh.appendRow(['2026-10-25', '2026-11-07', 38, 96.5, 0, 96.5, 54.5,
      'https://drive.google.com/file/d/1qPbtIG_yq9tbUsDua95gwqmYbJBmIXM2/view', '(préparée à la main)', '']);
  }
  sh.getRange('A:B').setNumberFormat('@'); sh.setFrozenRows(1);
  sh = ss.getSheetByName(F_PARAM) || ss.insertSheet(F_PARAM);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 7, 2).setNumberFormat('@').setValues([
      ['Paramètre', 'Valeur'],
      ['Début de paie de référence', '2026-11-08'],
      ['Heures payées par paie', '80'],
      ['Courriel de Marc', 'info@amaltis.ca'],
      ['Dossier Drive (ID)', '1rytgkOOI_biM1Eo6l-I-JdLOPCe4YKXD'],
      ['Modèle Excel (ID)', '1YYOC27qFU83wPacZJpaZoUuMTn9m1G_M'],
      ['Nom de l’employé', 'Philippe Brunelle']]);
    sh.setFrozenRows(1); sh.autoResizeColumns(1, 2);
  }
  const vide = ss.getSheetByName('Feuille 1') || ss.getSheetByName('Sheet1') || ss.getSheetByName('Feuille1');
  if (vide && ss.getSheets().length > 3) ss.deleteSheet(vide);
  // Déclencheur : chaque samedi vers 23 h, la paie qui se termine ce jour-là est préparée.
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'finDePaieAuto').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('finDePaieAuto').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(23).inTimezone(TZ).create();
  return 'Installation terminée';
}

/* ====================== Dates ====================== */

const d2s_ = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
function s2utc_(s) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s); if (!m) throw new Error('Date invalide'); return Date.UTC(+m[1], +m[2] - 1, +m[3]); }
function addDays_(s, n) { return new Date(s2utc_(s) + n * 864e5).toISOString().slice(0, 10); }
function serial_(s) { return Math.round((s2utc_(s) - Date.UTC(1899, 11, 30)) / 864e5); }
function dow_(s) { return new Date(s2utc_(s)).getUTCDay(); }
function hm2min_(h) { const m = /^(\d{1,2}):(\d{2})$/.exec(String(h || '').trim()); if (!m || +m[1] > 23 || +m[2] > 59) throw new Error('Heure invalide'); return +m[1] * 60 + +m[2]; }
function aujourdhui_() { return d2s_(new Date()); }

/** Début (dimanche) de la paie qui contient la date s. */
function debutPaie_(s, ref) {
  const n = Math.floor((s2utc_(s) - s2utc_(ref)) / 864e5);
  return addDays_(ref, n - (((n % 14) + 14) % 14));
}

/* Jours fériés applicables au Québec — Pâques : Vendredi saint */
function paques_(y) {
  const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25),
    g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4,
    l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
  return Date.UTC(y, Math.floor((h + l - 7 * m + 114) / 31) - 1, ((h + l - 7 * m + 114) % 31) + 1);
}
function feries_(y) {
  const iso = t => new Date(t).toISOString().slice(0, 10), D = (m, d) => Date.UTC(y, m - 1, d), J = 864e5;
  const wd = t => new Date(t).getUTCDay();
  const lundiAvant = t => t - ((wd(t) + 6) % 7) * J;                // lundi qui précède ou égale t
  const premierLundi = m => { const t = D(m, 1); return t + ((8 - wd(t)) % 7) * J; };
  const dimSuivant = t => wd(t) === 0 ? t + J : t;
  return {
    [iso(D(1, 1))]: 'Jour de l’An',
    [iso(paques_(y) - 2 * J)]: 'Vendredi saint',
    [iso(lundiAvant(D(5, 24)))]: 'Journée nationale des patriotes',
    [iso(dimSuivant(D(6, 24)))]: 'Fête nationale',
    [iso(dimSuivant(D(7, 1)))]: 'Fête du Canada',
    [iso(premierLundi(9))]: 'Fête du Travail',
    [iso(premierLundi(10) + 7 * J)]: 'Action de grâce',
    [iso(D(12, 25))]: 'Noël'
  };
}
function ferie_(s) { return feries_(+s.slice(0, 4))[s] || ''; }

/* ====================== Pointages ====================== */

function lirePointages_() {
  const sh = ss_().getSheetByName(F_POINT);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 6).getDisplayValues().map((r, i) => ({
    ligne: i + 2, id: r[0], date: r[1], entree: r[2], sortie: r[3], req: r[4]
  })).filter(q => q.id);
}

/** Congés pris dans la banque : une ligne par journée (heures payées à même la banque). */
function feuilleConges_() {
  const ss = ss_();
  let sh = ss.getSheetByName(F_CONGE);
  if (!sh) { sh = ss.insertSheet(F_CONGE); sh.appendRow(H_CONGE); sh.getRange('A:B').setNumberFormat('@'); sh.setFrozenRows(1); }
  return sh;
}
function lireConges_() {
  const sh = feuilleConges_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 4).getDisplayValues().map((r, i) => ({
    ligne: i + 2, id: r[0], date: r[1], heures: Number(String(r[2]).replace(',', '.')) || 0, req: r[3]
  })).filter(c => c.id && c.heures > 0);
}

/** Minutes d'un quart (une sortie avant l'entrée = passage de minuit). */
function minutes_(q) {
  if (!q.sortie) return 0;
  let m = hm2min_(q.sortie) - hm2min_(q.entree);
  if (m < 0) m += 1440;
  return m;
}
const arrondi_ = min => Math.round(min / 15) * 15 / 60;   // au quart d'heure le plus proche

function paie_(debut, quarts, p, conges) {
  conges = conges || [];
  const jours = [];
  for (let i = 0; i < 14; i++) {
    const date = addDays_(debut, i);
    const qs = quarts.filter(q => q.date === date).sort((a, b) => a.entree < b.entree ? -1 : 1);
    const min = qs.reduce((s, q) => s + minutes_(q), 0);
    const fer = ferie_(date);
    jours.push({
      date, jour: JOURS[dow_(date)], ferie: fer, heuresFeriees: fer ? 8 : 0,
      quarts: qs.map(q => ({ id: q.id, entree: q.entree, sortie: q.sortie, minutes: minutes_(q) })),
      minutes: min, heures: qs.length ? arrondi_(min) : null,
      conge: conges.filter(c => c.date === date).reduce((s, c) => s + c.heures, 0) || 0
    });
  }
  const normales = jours.reduce((s, j) => s + (j.heures || 0), 0);
  const feriees = jours.reduce((s, j) => s + j.heuresFeriees, 0);
  const banque = jours.reduce((s, j) => s + j.conge, 0);
  // travail = ce qui alimente la banque ; total = heures payées inscrites sur la feuille (avec les congés pris dans la banque)
  return { debut, fin: addDays_(debut, 13), jours, normales, feriees, banque, travail: normales + feriees,
    total: normales + feriees + banque, payees: p.payees };
}

function banqueReportee_(debut) {
  const veille = addDays_(debut, -1);
  const r = ss_().getSheetByName(F_PAIES).getDataRange().getDisplayValues().slice(1).filter(x => x[1] === veille).pop();
  return r ? Number(String(r[6]).replace(',', '.')) : null;
}

function getData(code) {
  verifier_(code);
  const p = params_(), auj = aujourdhui_();
  const quarts = lirePointages_(), conges = lireConges_();
  let debut = debutPaie_(auj, p.reference);
  if (debut < p.reference) debut = p.reference;   // avant la première paie de l'app : afficher celle-ci
  const paie = paie_(debut, quarts, p, conges);
  // Banque reportée : paie précédente enregistrée, sinon calcul à partir de la dernière connue
  let banque = banqueReportee_(debut);
  if (banque === null) {
    const prec = debutPaie_(addDays_(debut, -1), p.reference), b0 = banqueReportee_(prec);
    if (b0 !== null) { const pp = paie_(prec, quarts, p, conges); banque = b0 + pp.travail - p.payees; }
  }
  paie.banqueReportee = banque;
  paie.nouvelleBanque = banque === null ? null : banque + paie.travail - p.payees;
  const ouvert = quarts.filter(q => !q.sortie).sort((a, b) => (a.date + a.entree < b.date + b.entree ? 1 : -1))[0] || null;
  const hist = ss_().getSheetByName(F_PAIES).getDataRange().getDisplayValues().slice(1).slice(-8).reverse()
    .map(r => ({ debut: r[0], fin: r[1], banqueReportee: r[2], total: r[5], nouvelleBanque: r[6], fichier: r[7], brouillon: r[8] }));
  // Paie précédente (pour corriger des oublis avant l'envoi) : jours seulement
  const precedente = paie_(addDays_(debut, -14), quarts, p, conges);
  return {
    aujourdhui: auj, maintenant: Utilities.formatDate(new Date(), TZ, 'HH:mm'),
    paie, precedente: { debut: precedente.debut, fin: precedente.fin, jours: precedente.jours, total: precedente.total },
    ouvert: ouvert && { id: ouvert.id, date: ouvert.date, entree: ouvert.entree },
    historique: hist, courriel: p.courriel, reference: p.reference
  };
}

function avecVerrou_(f) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try { return f(); } finally { lock.releaseLock(); }
}

/** Requête déjà traitée (renvoi depuis la file d'envoi du téléphone) ? */
function dejaFait_(req) { return req ? lirePointages_().filter(q => q.req.split(' ').indexOf(req) >= 0)[0] : null; }

/** type : 'entree' ou 'sortie' ; date et heure : celles du téléphone au moment du toucher. */
function pointer(code, type, date, heure, req) {
  verifier_(code);
  s2utc_(date); hm2min_(heure);
  return avecVerrou_(() => {
    if (dejaFait_(req)) return getData(code);
    const sh = ss_().getSheetByName(F_POINT), now = new Date().toISOString();
    const ouverts = lirePointages_().filter(q => !q.sortie);
    if (type === 'entree') {
      if (ouverts.length) throw new Error('Déjà entré (' + ouverts[0].date + ' à ' + ouverts[0].entree + '). Pointez la sortie d’abord.');
      sh.appendRow([Utilities.getUuid().slice(0, 8), date, heure, '', req || '', now]);
    } else if (type === 'sortie') {
      if (!ouverts.length) throw new Error('Aucune entrée en cours : ajoutez le quart à la main.');
      const q = ouverts[ouverts.length - 1];
      const ecart = (s2utc_(date) - s2utc_(q.date)) / 864e5;
      if (ecart < 0 || ecart > 1 || (ecart === 1 && hm2min_(heure) >= hm2min_(q.entree)))
        throw new Error('Sortie trop éloignée de l’entrée (' + q.date + ' à ' + q.entree + '). Corrigez le quart à la main.');
      if (ecart === 0 && hm2min_(heure) < hm2min_(q.entree)) throw new Error('La sortie est avant l’entrée');
      sh.getRange(q.ligne, 4, 1, 3).setValues([[heure, (q.req ? q.req + ' ' : '') + (req || ''), now]]);
    } else throw new Error('demande invalide');
    SpreadsheetApp.flush();
    return getData(code);
  });
}

function ajouterQuart(code, date, entree, sortie, req) {
  verifier_(code);
  s2utc_(date); hm2min_(entree); if (sortie) hm2min_(sortie);
  return avecVerrou_(() => {
    if (dejaFait_(req)) return getData(code);
    ss_().getSheetByName(F_POINT).appendRow([Utilities.getUuid().slice(0, 8), date, entree, sortie || '', req || '', new Date().toISOString()]);
    SpreadsheetApp.flush();
    return getData(code);
  });
}

function modifierQuart(code, id, date, entree, sortie) {
  verifier_(code);
  s2utc_(date); hm2min_(entree); if (sortie) hm2min_(sortie);
  return avecVerrou_(() => {
    const q = lirePointages_().filter(x => x.id === id)[0];
    if (!q) throw new Error('Quart introuvable');
    ss_().getSheetByName(F_POINT).getRange(q.ligne, 2, 1, 3).setValues([[date, entree, sortie || '']]);
    ss_().getSheetByName(F_POINT).getRange(q.ligne, 6).setValue(new Date().toISOString());
    SpreadsheetApp.flush();
    return getData(code);
  });
}

/** Fixe les heures de congé prises dans la banque pour une journée (0 = retirer). */
function definirConge(code, date, heures, req) {
  verifier_(code);
  s2utc_(date);
  heures = Math.round(Number(heures) * 4) / 4;
  if (!(heures >= 0 && heures <= 24)) throw new Error('Heures de congé invalides');
  return avecVerrou_(() => {
    const sh = feuilleConges_();
    const lignes = lireConges_().filter(c => c.date === date).map(c => c.ligne).sort((a, b) => b - a);
    lignes.forEach(l => sh.deleteRow(l));
    if (heures > 0) sh.appendRow([Utilities.getUuid().slice(0, 8), date, heures, req || '', new Date().toISOString()]);
    SpreadsheetApp.flush();
    return getData(code);
  });
}

function supprimerQuart(code, id) {
  verifier_(code);
  return avecVerrou_(() => {
    const q = lirePointages_().filter(x => x.id === id)[0];
    if (!q) throw new Error('Quart introuvable');
    ss_().getSheetByName(F_POINT).deleteRow(q.ligne);
    return getData(code);
  });
}

/* ====================== Fin de paie : feuille Excel + brouillon ====================== */

/** Déclencheur du samedi 23 h : prépare toute paie terminée qui ne l'a pas encore été. */
function finDePaieAuto() {
  const p = params_(), auj = aujourdhui_();
  const faites = ss_().getSheetByName(F_PAIES).getDataRange().getDisplayValues().slice(1).map(r => r[0]);
  let debut = debutPaie_(auj, p.reference);
  if (addDays_(debut, 13) !== auj) debut = addDays_(debut, -14);   // dernière paie terminée
  // rattrapage éventuel d'une paie oubliée (déclencheur manqué)
  const aFaire = [];
  for (let d = debut, n = 0; n < 4 && faites.indexOf(d) < 0 && addDays_(d, 13) <= auj; d = addDays_(d, -14), n++) aFaire.unshift(d);
  aFaire.forEach(d => preparer_(d));
}

/** Depuis l'app : prépare (ou refait) la paie qui commence à « debut ». */
function preparerPaie(code, debut) {
  verifier_(code);
  s2utc_(debut);
  return avecVerrou_(() => { const r = preparer_(debut, true); const d = getData(code); d.prepare = r; return d; });
}

function preparer_(debut, refaire) {
  const p = params_();
  if (debutPaie_(debut, p.reference) !== debut) throw new Error('Date de début de paie invalide');
  if (debut < p.reference) throw new Error('Cette paie a été faite à la main avant l’app : elle ne peut pas être refaite ici.');
  if (!p.modele) throw new Error('Modèle Excel non configuré (Paramètres)');
  const quarts = lirePointages_();
  const paie = paie_(debut, quarts, p, lireConges_());
  const banque = banqueReportee_(debut);
  if (banque === null) throw new Error('Banque reportée introuvable : la paie précédente n’est pas enregistrée.');
  paie.banqueReportee = banque;
  paie.nouvelleBanque = banque + paie.travail - p.payees;
  const ouverts = quarts.filter(q => !q.sortie && q.date >= debut && q.date <= paie.fin);

  const nom = 'feuille_de_temps_' + debut + '_au_' + paie.fin + '.xlsx';
  const blob = remplirModele_(DriveApp.getFileById(p.modele).getBlob(), paie, nom);
  const dossier = DriveApp.getFolderById(p.dossier);
  const anciens = dossier.getFilesByName(nom);
  while (anciens.hasNext()) anciens.next().setTrashed(true);
  const fichier = dossier.createFile(blob);

  const periode = jourMois_(debut) + ' au ' + jourMois_(paie.fin, true);
  const fmt = v => String(Math.round(v * 100) / 100).replace('.', ',');
  const corps = [
    'Bonjour Marc,', '',
    'Voici ma feuille de temps pour la période du ' + periode + '.', '',
    'Heures travaillées : ' + fmt(paie.normales) + ' h',
    paie.feriees ? 'Heures fériées : ' + fmt(paie.feriees) + ' h' : null,
    paie.banque ? 'Heures prises dans la banque : ' + fmt(paie.banque) + ' h' : null,
    'Total : ' + fmt(paie.total) + ' h (' + fmt(p.payees) + ' h payées)',
    'Banque d’heures : ' + fmt(banque) + ' h reportées, ' + fmt(paie.nouvelleBanque) + ' h à reporter', '',
    'Merci,', p.nom
  ].filter(x => x !== null).join('\n');
  const objet = (ouverts.length ? '[À VÉRIFIER : sortie manquante] ' : '') + 'Feuille de temps – ' + p.nom + ' – ' + periode;
  const brouillon = GmailApp.createDraft(p.courriel, objet, corps, { attachments: [fichier.getBlob()] });

  const sh = ss_().getSheetByName(F_PAIES);
  const ligne = [debut, paie.fin, banque, paie.normales, paie.feriees, paie.total, paie.nouvelleBanque,
    fichier.getUrl(), 'Brouillon Gmail ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), new Date().toISOString()];
  const rows = sh.getDataRange().getDisplayValues();
  const i = rows.findIndex((r, k) => k > 0 && r[0] === debut);
  if (i > 0) sh.getRange(i + 1, 1, 1, ligne.length).setValues([ligne]); else sh.appendRow(ligne);
  return { fichier: fichier.getUrl(), brouillon: brouillon.getId(), nouvelleBanque: paie.nouvelleBanque, sortieManquante: ouverts.length > 0 };
}

/** Essai depuis l'éditeur : fabrique la feuille de la paie en cours dans Drive (fichier TEST_), sans brouillon. */
function testerModele() {
  const p = params_(), debut = debutPaie_(aujourdhui_(), p.reference);
  const paie = paie_(debut, lirePointages_(), p, lireConges_());
  paie.banqueReportee = banqueReportee_(debut) || 0;
  const nom = 'TEST_feuille_de_temps_' + debut + '.xlsx';
  const f = DriveApp.getFolderById(p.dossier).createFile(remplirModele_(DriveApp.getFileById(p.modele).getBlob(), paie, nom));
  Logger.log(f.getId() + ' ' + f.getSize());
  return f.getId();
}

function jourMois_(s, annee) {
  const d = +s.slice(8, 10), m = +s.slice(5, 7);
  return (d === 1 ? '1er' : d) + ' ' + MOIS[m - 1] + (annee ? ' ' + s.slice(0, 4) : '');
}

/* ---------- Remplissage du modèle .xlsx (mêmes styles, logo et formules) ---------- */

function remplirModele_(modeleBlob, paie, nom) {
  modeleBlob = modeleBlob.copyBlob().setContentType('application/zip');
  const parts = Utilities.unzip(modeleBlob);
  const i = parts.findIndex(b => b.getName() === 'xl/worksheets/sheet1.xml');
  if (i < 0) throw new Error('Modèle Excel illisible');
  const xml = remplirFeuilleXml(parts[i].getDataAsString('UTF-8'), paie);
  parts[i] = Utilities.newBlob(xml, 'application/xml', 'xl/worksheets/sheet1.xml');
  return Utilities.zip(parts, nom).setContentType(XLSX).setName(nom);
}

/** Pure (testable hors Google) : écrit les valeurs de la paie dans sheet1.xml du modèle. */
function remplirFeuilleXml(xml, paie) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const num = v => String(Math.round(v * 10000) / 10000);
  function set(ref, val) {
    const re = new RegExp('<c r="' + ref + '"(?:\\s[^>]*?)?(?:/>|>[\\s\\S]*?</c>)');
    const m = re.exec(xml);
    if (!m) throw new Error('Cellule absente du modèle : ' + ref);
    const s = (/\ss="(\d+)"/.exec(m[0]) || [])[1];
    const f = (/<f>[\s\S]*?<\/f>/.exec(m[0]) || [''])[0];
    const sa = s ? ' s="' + s + '"' : '';
    let c;
    if (val === null || val === '') c = f ? '<c r="' + ref + '"' + sa + ' t="str">' + f + '<v></v></c>' : '<c r="' + ref + '"' + sa + '/>';
    else if (typeof val === 'number') c = '<c r="' + ref + '"' + sa + '>' + f + '<v>' + num(val) + '</v></c>';
    else c = '<c r="' + ref + '"' + sa + (f ? ' t="str">' + f + '<v>' + esc(val) + '</v>' : ' t="inlineStr"><is><t>' + esc(val) + '</t></is>') + '</c>';
    xml = xml.slice(0, m.index) + c + xml.slice(m.index + m[0].length);
  }
  const s0 = serial_(paie.debut);
  set('H4', s0); set('H5', s0 + 13);
  paie.jours.forEach((j, k) => {
    const r = 8 + k;
    set('B' + r, j.jour); set('C' + r, s0 + k);
    set('D' + r, j.heures === null ? null : j.heures);
    set('E' + r, j.heuresFeriees ? 8 : null);
    set('F' + r, j.conge ? j.conge : null);
    set('H' + r, (j.heures || 0) + j.heuresFeriees + (j.conge || 0));
  });
  set('D22', paie.normales); set('E22', paie.feriees); set('F22', paie.banque || 0); set('G22', 0); set('H22', paie.total);
  // même formule que la feuille : banque reportée + travaillées + fériées − payées (vide si rien n'est inscrit)
  set('I22', paie.total ? paie.banqueReportee + paie.travail - paie.payees : null);
  set('E23', paie.payees); set('G23', paie.banqueReportee);
  return xml;
}

if (typeof module !== 'undefined') module.exports = { remplirFeuilleXml, paie_, feries_, debutPaie_, addDays_, arrondi_, minutes_ };
