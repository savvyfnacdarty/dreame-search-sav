/**
 * EXTRACTION COMPLÈTE D'UN DRIVE → feuille "Export" (Chemin / Nom / Type / Lien)
 * Version PAR TRANCHES : pause à 5 min, reprise automatique là où on s'était arrêté.
 *
 * Produit exactement le format lu par la page Recherche_GEM_TV :
 *   | Chemin | Nom | Type | Lien |
 *   - Type = "Dossier (racine)" | "Dossier" | "Fichier"
 *   - Chemin = chemin complet du dossier parent (vide pour la racine)
 *
 * MISE EN PLACE :
 *  1. https://script.google.com -> Nouveau projet -> colle ce code.
 *  2. (Option) renseigne SPREADSHEET_URL ci-dessous. Sinon un classeur est créé
 *     automatiquement au 1er lancement (l'URL apparaît dans les journaux + par e-mail).
 *  3. Exécute  demarrerExport()   -> lance une extraction NEUVE (repart de zéro),
 *     puis se relance seule toutes les minutes jusqu'à la fin.
 *  4. Pour reprendre une extraction interrompue sans repartir de zéro : exportDrive().
 *  5. Pour tout remettre à zéro manuellement : reinitialiserExport().
 *
 * Après la fin : télécharge la feuille en CSV
 *   (Fichier -> Télécharger -> Valeurs séparées par des virgules)
 *   OU exécute  exporterCsvDansDrive()  pour générer directement le .csv dans le Drive.
 */

// ====== À PERSONNALISER ======
const ROOT_FOLDER_ID = '1qDVkgIRe9wfkwm9f4sL1kRQQUNJXLzYq'; // dossier racine à extraire
const SPREADSHEET_URL = ''; // laisse vide pour créer un classeur automatiquement
const OUTPUT_SHEET_NAME = 'Export';
const CSV_FILE_NAME = 'Recherche_GEM_TV_catalogue.csv'; // nom du CSV généré par exporterCsvDansDrive()
const BUDGET_MS = 5 * 60 * 1000; // durée max d'une tranche (5 min ; marge sous la limite de 6 min)
const FLUSH_EVERY = 1500;        // écrit dans la feuille tous les N lignes accumulées
const ENVOYER_EMAIL = true;      // e-mail récap en fin d'extraction
const EMAIL_DESTINATAIRE = 'jerome.leconte@fnacdarty.com';
// =============================

// Feuilles cachées de travail
const FEUILLE_QUEUE = '_export_queue';

/**
 * Fonction principale : traite UNE tranche (≤ 5 min) puis se relance si besoin.
 * Types d'éléments dans la file :
 *   {t:'F',  id,  path} -> dossier à ouvrir (démarre le listage sous-dossiers + fichiers)
 *   {t:'SF', tok, path} -> reprise du listage des SOUS-DOSSIERS (jeton de continuation)
 *   {t:'FF', tok, path} -> reprise du listage des FICHIERS (jeton de continuation)
 */
function exportDrive() {
  const DEBUT = Date.now();
  const props = PropertiesService.getScriptProperties();
  supprimerTriggers_('_continuerExport');

  const ss = getSpreadsheet_();
  const shOut = getOutputSheet_(ss);
  const shQueue = getFeuilleCachee_(ss, FEUILLE_QUEUE);

  // Démarrage d'une nouvelle extraction si aucune n'est en cours
  const enCours = props.getProperty('export_enCours') === 'true';
  if (!enCours) {
    shOut.clearContents();
    shOut.getRange(1, 1, 1, 4).setValues([['Chemin', 'Nom', 'Type', 'Lien']]);
    shOut.setFrozenRows(1);
    const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
    const rootName = root.getName();
    shOut.getRange(2, 1, 1, 4).setValues([['', rootName, 'Dossier (racine)', root.getUrl()]]);
    ecrireQueue_(shQueue, [{ t: 'F', id: ROOT_FOLDER_ID, path: rootName }]);
    props.setProperty('export_enCours', 'true');
    props.setProperty('export_nbFichiers', '0');
    props.setProperty('export_nbDossiers', '0');
  }

  const queue = lireQueue_(shQueue);
  let buffer = [];
  let nbF = 0, nbD = 0;

  const budgetDepasse = function () { return (Date.now() - DEBUT) >= BUDGET_MS; };

  while (queue.length > 0 && !budgetDepasse()) {
    const item = queue.shift();

    if (item.t === 'F') {
      let folder;
      try { folder = DriveApp.getFolderById(item.id); }
      catch (e) { Logger.log('Dossier inaccessible ' + item.id + ' : ' + e.message); continue; }
      // On planifie le listage des sous-dossiers puis des fichiers (repris via jetons)
      queue.push({ t: 'SF', tok: folder.getFolders().getContinuationToken(), path: item.path });
      queue.push({ t: 'FF', tok: folder.getFiles().getContinuationToken(), path: item.path });

    } else if (item.t === 'SF') {
      let it;
      try { it = DriveApp.continueFolderIterator(item.tok); }
      catch (e) { Logger.log('Jeton sous-dossier invalide : ' + e.message); continue; }
      while (it.hasNext()) {
        if (budgetDepasse()) { queue.unshift({ t: 'SF', tok: it.getContinuationToken(), path: item.path }); break; }
        const sf = it.next();
        const childPath = item.path + '/' + sf.getName();
        buffer.push([item.path, sf.getName(), 'Dossier', sf.getUrl()]);
        nbD++;
        queue.push({ t: 'F', id: sf.getId(), path: childPath });
      }

    } else if (item.t === 'FF') {
      let it;
      try { it = DriveApp.continueFileIterator(item.tok); }
      catch (e) { Logger.log('Jeton fichier invalide : ' + e.message); continue; }
      while (it.hasNext()) {
        if (budgetDepasse()) { queue.unshift({ t: 'FF', tok: it.getContinuationToken(), path: item.path }); break; }
        const f = it.next();
        buffer.push([item.path, f.getName(), 'Fichier', f.getUrl()]);
        nbF++;
      }
    }

    // Écriture périodique pour limiter la mémoire
    if (buffer.length >= FLUSH_EVERY) { flush_(shOut, buffer); buffer = []; }
  }

  // Écriture de fin de tranche
  if (buffer.length) flush_(shOut, buffer);
  ecrireQueue_(shQueue, queue);

  // Cumul des compteurs
  nbF += parseInt(props.getProperty('export_nbFichiers') || '0', 10);
  nbD += parseInt(props.getProperty('export_nbDossiers') || '0', 10);
  props.setProperty('export_nbFichiers', String(nbF));
  props.setProperty('export_nbDossiers', String(nbD));

  // Pas fini -> reprise automatique dans 1 min
  if (queue.length > 0) {
    planifierContinuation_();
    Logger.log('Tranche terminée. Reste ' + queue.length + ' élément(s) en file. '
      + 'Cumulé : ' + nbD + ' dossier(s), ' + nbF + ' fichier(s). Reprise auto dans 1 min.');
    return;
  }

  // Extraction terminée
  props.setProperty('export_enCours', 'false');
  props.deleteProperty('export_nbFichiers');
  props.deleteProperty('export_nbDossiers');
  Logger.log('✅ EXTRACTION TERMINÉE : ' + nbD + ' dossier(s), ' + nbF + ' fichier(s). '
    + 'Feuille "' + OUTPUT_SHEET_NAME + '" du classeur : ' + ss.getUrl());
  if (ENVOYER_EMAIL) {
    MailApp.sendEmail({
      to: EMAIL_DESTINATAIRE,
      subject: 'Extraction Drive terminée : ' + nbD + ' dossiers, ' + nbF + ' fichiers',
      htmlBody: '<p>Extraction complète du Drive terminée.</p>'
        + '<p>' + nbD + ' dossier(s) · ' + nbF + ' fichier(s).</p>'
        + '<p>Classeur : <a href="' + ss.getUrl() + '">ouvrir</a> (feuille "' + OUTPUT_SHEET_NAME + '").</p>'
    });
  }
}

/** Relancée par le déclencheur de reprise (ne pas exécuter à la main). */
function _continuerExport() { exportDrive(); }

/** Lance une extraction NEUVE (repart de zéro) puis enchaîne les tranches. */
function demarrerExport() {
  reinitialiserExport();
  exportDrive();
}

/* ===================== ÉCRITURE ===================== */

function flush_(shOut, buffer) {
  if (!buffer.length) return;
  const start = shOut.getLastRow() + 1;
  shOut.getRange(start, 1, buffer.length, 4).setValues(buffer);
}

/* ===================== SPREADSHEET / FEUILLES ===================== */

function getSpreadsheet_() {
  if (SPREADSHEET_URL) return SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('export_ss_id');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* recréé ci-dessous */ } }
  const ss = SpreadsheetApp.create('Export GEM_TV Drive');
  props.setProperty('export_ss_id', ss.getId());
  Logger.log('Nouveau classeur créé : ' + ss.getUrl());
  return ss;
}

function getOutputSheet_(ss) {
  let sh = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(OUTPUT_SHEET_NAME);
  return sh;
}

function getFeuilleCachee_(ss, nom) {
  let sh = ss.getSheetByName(nom);
  if (!sh) { sh = ss.insertSheet(nom); sh.hideSheet(); }
  return sh;
}

/* ===================== FILE D'ATTENTE ===================== */

function lireQueue_(sh) {
  const last = sh.getLastRow();
  if (last < 1) return [];
  const vals = sh.getRange(1, 1, last, 1).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0]) { try { out.push(JSON.parse(vals[i][0])); } catch (e) { /* ligne corrompue ignorée */ } }
  }
  return out;
}

function ecrireQueue_(sh, arr) {
  sh.clearContents();
  if (!arr.length) return;
  const data = arr.map(function (x) { return [JSON.stringify(x)]; });
  sh.getRange(1, 1, data.length, 1).setValues(data);
}

/* ===================== DÉCLENCHEURS ===================== */

function planifierContinuation_() {
  supprimerTriggers_('_continuerExport');
  ScriptApp.newTrigger('_continuerExport').timeBased().after(60 * 1000).create();
}

function supprimerTriggers_(handler) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
  });
}

/* ===================== OUTILS À LANCER À LA MAIN ===================== */

/** Remet l'extraction à zéro (état, file, compteurs). Le prochain exportDrive() repart du début. */
function reinitialiserExport() {
  const props = PropertiesService.getScriptProperties();
  ['export_enCours', 'export_nbFichiers', 'export_nbDossiers'].forEach(function (k) { props.deleteProperty(k); });
  supprimerTriggers_('_continuerExport');
  const ss = getSpreadsheet_();
  const q = ss.getSheetByName(FEUILLE_QUEUE);
  if (q) ss.deleteSheet(q);
  Logger.log('Extraction réinitialisée. Le prochain exportDrive() repartira de zéro.');
}

/** Test rapide d'accès au dossier racine et au classeur. */
function testAcces() {
  try {
    const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
    Logger.log('Dossier racine OK : ' + root.getName());
    const ss = getSpreadsheet_();
    Logger.log('Classeur OK : ' + ss.getName() + ' -> ' + ss.getUrl());
  } catch (e) { Logger.log('ERREUR : ' + e.message); }
}

/**
 * (Option) Génère/actualise un fichier CSV dans le Drive à partir de la feuille "Export".
 * Le CSV est créé à la racine du Drive (ou remplacé s'il existe déjà). Renvoie/loggue son URL.
 */
function exporterCsvDansDrive() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!sh) { Logger.log('Feuille "' + OUTPUT_SHEET_NAME + '" introuvable.'); return; }
  const values = sh.getDataRange().getValues();
  const csv = values.map(function (row) {
    return row.map(function (c) {
      let s = (c === null || c === undefined) ? '' : String(c);
      if (/[",\n\r]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
      return s;
    }).join(',');
  }).join('\r\n');
  const blob = Utilities.newBlob('﻿' + csv, 'text/csv', CSV_FILE_NAME); // BOM UTF-8 pour Excel
  const existing = DriveApp.getFilesByName(CSV_FILE_NAME);
  let file;
  if (existing.hasNext()) { file = existing.next(); file.setContent('﻿' + csv); }
  else { file = DriveApp.createFile(blob); }
  Logger.log('CSV écrit : ' + file.getUrl());
  return file.getUrl();
}
