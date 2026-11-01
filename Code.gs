/**
 * Stundenerfassung – Google Apps Script (Web-App, gebunden an die Tabelle «Stundenerfassung Assistenz»).
 * Tabellenblatt «Personen»: Spalte A = Token (geheimer Link-Teil), B = Name (genau wie in der Abrechnungs-App), C = aktiv (leer oder «ja»; «nein» sperrt)
 * Tabellenblatt «Stunden»: A = Name, B = Datum (JJJJ-MM-TT, Text), C = Dauer (Dezimalstunden), D = Geändert
 * Die Spalten A–C sind so benannt (Name | Datum | Dauer), dass die Abrechnungs-App die CSV-Datei direkt einlesen kann.
 * Tabellenblatt «Abschluss» (wird automatisch angelegt): A = Name, B = Monat (JJJJ-MM), C = Total, D = Zeitpunkt, E = Status («abgeschickt» oder «offen»), F = Anzahl Abschlüsse.
 *   Ein Monat ist für die Person gesperrt, solange der Status «abgeschickt» ist. Roger kann ihn öffnen, indem er den Status auf «offen» setzt oder die Zeile löscht.
 * Abruf durch die App: Aktion «export» mit Schlüssel (Skripteigenschaft ABRUF_KEY); liefert Stunden eines Monats als CSV-Text und den Abschluss-Status.
 * Beim Abschicken geht eine E-Mail an den Besitzer des Scripts (Konto, unter dem die Web-App läuft).
 * Ablage durch die App (POST, nur mit dem Schlüssel ABRUF_KEY). Hauptordner = Ordner der Stundentabelle (oder Skripteigenschaft HAUPTORDNER_ID):
 *   «ablegen»        legt PDFs nach Jahren ab: Hauptordner/JJJJ/[Kurzname]/ (Personen) bzw. Hauptordner/JJJJ/ (Jahresübersicht);
 *                    gleichnamige alte Dateien kommen in den Papierkorb. Jahresordner entstehen automatisch, Personenordner nur für Namen aus dem Blatt «Personen».
 *   «jahr_lesen»     liefert die Werte der Google Tabelle «Jahresübersicht Löhne Assistenz JJJJ» im Ordner JJJJ (einmalige Übernahme in die App).
 *   «jahr_schreiben» schreibt die Jahresübersicht in dieselbe Tabelle (gleicher Link, im Ordner JJJJ); weigert sich, wenn dabei Monate verloren gingen.
 *   Einmalig im Editor: «jahresordnerPruefen», dann «jahresordnerEinrichten» verschiebt die bestehenden Dateien in die Jahresordner.
 *   «entwurf»        legt Gmail-ENTWÜRFE mit PDF-Anhängen an.
 *   «senden»         sendet einen solchen Entwurf (nach Kontrolle und Rückfrage in der App). Nur an Adressen aus der
 *                    Skripteigenschaft ERLAUBTE_EMPFAENGER (durch Komma getrennt) oder an das eigene Konto, nur mit Betreff
 *                    «Lohnabrechnung …» und PDF-Anhängen. Ohne ERLAUBTE_EMPFAENGER sendet das Script an niemand anderen.
 *   «sicherung»      legt die verschlüsselte Sicherung der App («Sicherung Abrechnung JJJJ-MM-TT.json») im Hauptordner ab;
 *                    ältere Sicherungen kommen in den Papierkorb (nur die neueste bleibt). Nimmt nur verschlüsselte Sicherungen an.
 *   Mit test = true arbeitet alles nur im Unterordner «_TEST Beispieldaten» des Hauptordners (für erfundene Beispieldaten).
 */
var TZ = 'Europe/Zurich';
var SH_P = 'Personen';
var SH_S = 'Stunden';
var SH_A = 'Abschluss';
var MONATE = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];

function doGet(e) { return out_(handle_(e && e.parameter ? e.parameter : {})); }
function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (x) { p = (e && e.parameter) || {}; }
  return out_(handle_(p));
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

var ABLAGE = { ablegen: 1, jahr_lesen: 1, jahr_schreiben: 1, entwurf: 1, senden: 1, sicherung: 1 };

function handle_(p) {
  // Ablage-Aktionen der App: ohne Script-Sperre, damit die Erfassungsseite nie warten muss.
  if (ABLAGE[p.action]) {
    try {
      var err = checkKey_(String(p.k || ''));
      if (err) return { ok: false, error: err };
      if (p.action === 'ablegen') return ablegen_(p);
      if (p.action === 'jahr_lesen') return jahrLesen_(p);
      if (p.action === 'jahr_schreiben') return jahrSchreiben_(p);
      if (p.action === 'senden') return senden_(p);
      if (p.action === 'sicherung') return sicherung_(p);
      return entwurf_(p);
    } catch (x) { return { ok: false, error: 'Fehler: ' + x.message }; }
  }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (x) { return { ok: false, error: 'Gerade besetzt, bitte nochmals versuchen.' }; }
  try {
    if (p.action === 'export') return export_(String(p.k || ''), String(p.m || ''));
    var name = findPerson_(String(p.t || ''));
    if (!name) return { ok: false, error: 'Dieser Link ist ungültig. Bitte bei Roger nachfragen.' };
    if (p.action === 'set') return setHours_(name, String(p.date || ''), p.hours);
    if (p.action === 'submit') return submit_(name, String(p.m || ''));
    if (p.action === 'reopen') return reopen_(name, String(p.m || ''));
    return list_(name, String(p.m || ''));
  } catch (err) {
    return { ok: false, error: 'Fehler: ' + err.message };
  } finally { lock.releaseLock(); }
}

function sheet_(n) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(n);
  if (!s) throw new Error('Tabellenblatt «' + n + '» fehlt');
  return s;
}

function findPerson_(token) {
  if (token.length < 16) return null;
  var v = sheet_(SH_P).getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === token) {
      if (String(v[i][2]).trim().toLowerCase() === 'nein') return null;
      return String(v[i][1]).trim();
    }
  }
  return null;
}

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function prevMonth_(ym) {
  var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1;
  if (m < 1) { y--; m = 12; }
  return y + '-' + (m < 10 ? '0' : '') + m;
}
/** Bearbeitbar: laufender Monat und Vormonat. Ältere Monate nur noch lesen. */
function editable_(ym) { var t = today_().slice(0, 7); return ym === t || ym === prevMonth_(t); }

function datesOf_(name, ym) {
  var s = sheet_(SH_S), n = s.getLastRow(), res = {};
  if (n < 2) return res;
  var v = s.getRange(2, 1, n - 1, 3).getValues();
  for (var i = 0; i < v.length; i++) {
    var d = cellDate_(v[i][1]);
    if (String(v[i][0]).trim() === name && d.slice(0, 7) === ym) res[d] = Number(v[i][2]);
  }
  return res;
}
function cellDate_(c) {
  if (c instanceof Date) return Utilities.formatDate(c, TZ, 'yyyy-MM-dd');
  return String(c).trim();
}

function list_(name, ym) {
  if (!/^\d{4}-\d{2}$/.test(ym)) ym = today_().slice(0, 7);
  return { ok: true, name: name, month: ym, today: today_(), editable: editable_(ym), closed: closed_(name, ym), entries: datesOf_(name, ym) };
}

var MAX_H = 12;
/** Nur Dezimalstunden in Viertelstunden (Komma oder Punkt), 0 bis MAX_H. Leer = 0 (Eintrag löschen). */
function parseHours_(h) {
  if (h === '' || h === null || h === undefined) return 0;
  var s = String(h).trim().replace(',', '.');
  if (!s) return 0;
  if (!/^\d{1,2}(\.\d{1,2})?$/.test(s)) return NaN;
  var v = Number(s);
  if (!isFinite(v) || v < 0 || v > MAX_H) return NaN;
  if (Math.abs(v * 4 - Math.round(v * 4)) > 1e-9) return NaN;
  return v;
}

function setHours_(name, date, hours) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'Datum ungültig.' };
  var y = +date.slice(0, 4), m = +date.slice(5, 7), d = +date.slice(8, 10);
  var dt = new Date(y, m - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return { ok: false, error: 'Datum ungültig.' };
  if (!editable_(date.slice(0, 7))) return { ok: false, error: 'Dieser Monat ist abgeschlossen. Änderungen bitte Roger melden.' };
  if (date > today_()) return { ok: false, error: 'Für zukünftige Tage kann noch nichts eingetragen werden.' };
  if (closed_(name, date.slice(0, 7))) return { ok: false, error: 'Dieser Monat ist abgeschickt. Bitte zuerst «Nochmals bearbeiten» wählen.' };
  var h = parseHours_(hours);
  if (isNaN(h)) return { ok: false, error: 'Bitte Stunden in Viertelstunden eingeben (z. B. 4, 4.25, 4.5 oder 4.75), höchstens ' + MAX_H + ' Stunden pro Tag.' };
  var s = sheet_(SH_S), n = s.getLastRow(), row = -1;
  if (n >= 2) {
    var v = s.getRange(2, 1, n - 1, 3).getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]).trim() === name && cellDate_(v[i][1]) === date) { row = i + 2; break; }
    }
  }
  var now = Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm');
  if (h === 0) { if (row > 0) s.deleteRow(row); }
  else if (row > 0) { s.getRange(row, 3, 1, 2).setValues([[h, now]]); }
  else {
    var r = Math.max(n, 1) + 1;
    s.getRange(r, 2).setNumberFormat('@');
    s.getRange(r, 1, 1, 4).setValues([[name, date, h, now]]);
  }
  return { ok: true, date: date, hours: h };
}

function monthLabel_(ym) { return MONATE[+ym.slice(5, 7) - 1] + ' ' + ym.slice(0, 4); }
function cellMonth_(c) {
  if (c instanceof Date) return Utilities.formatDate(c, TZ, 'yyyy-MM');
  return String(c).trim();
}

/** Blatt «Abschluss» (wird bei Bedarf angelegt). */
function absSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SH_A);
  if (!s) {
    s = ss.insertSheet(SH_A);
    s.getRange(1, 1, 1, 6).setValues([['Name', 'Monat', 'Total', 'Zeitpunkt', 'Status', 'Anzahl']]);
    s.getRange('B:B').setNumberFormat('@');
  }
  return s;
}
function absFind_(name, ym) {
  var s = absSheet_(), n = s.getLastRow();
  if (n < 2) return null;
  var v = s.getRange(2, 1, n - 1, 6).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === name && cellMonth_(v[i][1]) === ym) {
      return { row: i + 2, status: String(v[i][4]).trim().toLowerCase(), count: Number(v[i][5]) || 0 };
    }
  }
  return null;
}
function closed_(name, ym) { var a = absFind_(name, ym); return !!a && a.status === 'abgeschickt'; }

function submit_(name, ym) {
  if (!/^\d{4}-\d{2}$/.test(ym)) return { ok: false, error: 'Monat ungültig.' };
  if (!editable_(ym)) return { ok: false, error: 'Dieser Monat ist abgeschlossen. Änderungen bitte Roger melden.' };
  var e = datesOf_(name, ym), total = 0, days = 0;
  Object.keys(e).forEach(function (k) { total += e[k]; days++; });
  var a = absFind_(name, ym);
  if (a && a.status === 'abgeschickt') return { ok: true, already: true, month: ym, total: total, days: days };
  var s = absSheet_(), now = Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm');
  var count = (a ? a.count : 0) + 1;
  if (a) { s.getRange(a.row, 3, 1, 4).setValues([[total, now, 'abgeschickt', count]]); }
  else {
    var r = Math.max(s.getLastRow(), 1) + 1;
    s.getRange(r, 2).setNumberFormat('@');
    s.getRange(r, 1, 1, 6).setValues([[name, ym, total, now, 'abgeschickt', count]]);
  }
  var mail = notify_(name, ym, total, days, count, now);
  return { ok: true, month: ym, total: total, days: days, mail: mail };
}

function reopen_(name, ym) {
  if (!/^\d{4}-\d{2}$/.test(ym)) return { ok: false, error: 'Monat ungültig.' };
  if (!editable_(ym)) return { ok: false, error: 'Dieser Monat ist abgeschlossen. Änderungen bitte Roger melden.' };
  var a = absFind_(name, ym);
  if (a && a.status === 'abgeschickt') absSheet_().getRange(a.row, 5, 1, 1).setValues([['offen']]);
  return { ok: true, month: ym };
}

/**
 * Abruf durch die Abrechnungs-App: liefert alle Stunden eines Monats als CSV-Text (Name,Datum,Dauer,Geändert)
 * und den Abschluss-Status pro Person. Nur mit dem geheimen Schlüssel aus den Skripteigenschaften (Eigenschaft ABRUF_KEY).
 */
function checkKey_(key) {
  var want = PropertiesService.getScriptProperties().getProperty('ABRUF_KEY') || '';
  if (want.length < 12) return 'Abruf ist nicht eingerichtet (Skripteigenschaft ABRUF_KEY fehlt).';
  if (key.length !== want.length || key !== want) return 'Abruf-Schlüssel ist falsch.';
  return '';
}

function export_(key, ym) {
  var kerr = checkKey_(key);
  if (kerr) return { ok: false, error: kerr };
  if (!/^\d{4}-\d{2}$/.test(ym)) return { ok: false, error: 'Monat ungültig.' };
  var s = sheet_(SH_S), n = s.getLastRow(), lines = ['Name,Datum,Dauer,Geändert'], rows = 0;
  var q = function (x) { x = String(x); return /[",\n]/.test(x) ? '"' + x.replace(/"/g, '""') + '"' : x; };
  if (n >= 2) {
    var v = s.getRange(2, 1, n - 1, 4).getValues();
    for (var i = 0; i < v.length; i++) {
      var d = cellDate_(v[i][1]);
      if (d.slice(0, 7) !== ym || !String(v[i][0]).trim()) continue;
      lines.push([q(String(v[i][0]).trim()), d, Number(v[i][2]), q(v[i][3] instanceof Date ? Utilities.formatDate(v[i][3], TZ, 'dd.MM.yyyy HH:mm') : v[i][3])].join(','));
      rows++;
    }
  }
  var abs = [], a = absSheet_(), an = a.getLastRow();
  if (an >= 2) {
    var av = a.getRange(2, 1, an - 1, 6).getValues();
    for (var j = 0; j < av.length; j++) {
      if (cellMonth_(av[j][1]) === ym) abs.push({ name: String(av[j][0]).trim(), status: String(av[j][4]).trim().toLowerCase(), total: Number(av[j][2]) || 0 });
    }
  }
  var pv = sheet_(SH_P).getDataRange().getValues(), act = [];
  for (var k = 1; k < pv.length; k++) {
    if (String(pv[k][1]).trim() && String(pv[k][2]).trim().toLowerCase() !== 'nein') act.push(String(pv[k][1]).trim());
  }
  return { ok: true, month: ym, rows: rows, csv: lines.join('\n') + '\n', abschluss: abs, aktiv: act };
}

/** E-Mail an den Besitzer des Scripts. Ein Fehler hier darf das Abschicken nicht verhindern. */
function notify_(name, ym, total, days, count, now) {
  try {
    var to = Session.getEffectiveUser().getEmail();
    if (!to) return false;
    var again = count > 1 ? ' (erneut)' : '';
    var subj = 'Stunden übermittelt: ' + name + ' ' + monthLabel_(ym) + again;
    var body = name + ' hat die Stunden für ' + monthLabel_(ym) + ' übermittelt' + again + '.\n\n' +
      'Total: ' + total.toFixed(2) + ' Stunden (' + days + (days === 1 ? ' Tag' : ' Tage') + ')\n' +
      'Zeitpunkt: ' + now + '\n\n' +
      'Die Einzelheiten stehen im Blatt «Stunden» der Tabelle «Stundenerfassung Assistenz».';
    MailApp.sendEmail(to, subj, body);
    return true;
  } catch (x) { return false; }
}

/** Einmalig im Script-Editor ausführen: holt die Berechtigungen (E-Mail-Hinweis, Drive-Ablage, Gmail-Entwürfe, Tabellen). */
function autorisieren() {
  MailApp.getRemainingDailyQuota();
  Session.getEffectiveUser().getEmail();
  Logger.log('Hauptordner: ' + hauptordner_().getName());
  GmailApp.getDrafts();
  Logger.log('Berechtigungen in Ordnung.');
}

/**
 * Einmalig im Script-Editor: bestehende Dateien in die Jahresordner verschieben.
 * Zuerst «jahresordnerPruefen» ausführen (zeigt im Protokoll nur, was geschehen würde), dann «jahresordnerEinrichten».
 * Verschoben werden nur PDFs mit dem Namen der App («JJJJ-MM Lohnabrechnung …», «JJJJ-MM Stundenübersicht …») aus den
 * Personenordnern, deren Name im Blatt «Personen» steht, sowie Jahresübersicht-Tabelle und -PDF aus dem Hauptordner.
 * Verschieben ändert keine Links. Alles andere bleibt, wo es ist, und wird im Protokoll aufgeführt.
 * Ein danach leerer alter Personenordner kommt in den Papierkorb (wiederherstellbar).
 */
function jahresordnerPruefen() { return jahresordner_(false); }
function jahresordnerEinrichten() { return jahresordner_(true); }
function jahresordner_(ausfuehren) {
  var base = hauptordner_(), namen = personenNamen_(), log = [], n = 0, bleibt = 0;
  var zeile = function (s) { log.push(s); Logger.log(s); };
  zeile(ausfuehren ? 'EINRICHTEN – Dateien werden verschoben:' : 'NUR PRÜFEN – es wird nichts verändert:');
  var ziel = function (year, person) {
    var jo = unterordner_(base, String(year), ausfuehren);
    if (!person) return jo;
    return jo ? unterordner_(jo, person, ausfuehren) : null;
  };
  var schiebe = function (f, year, person) {
    var name = nfc_(f.getName()), z = ziel(year, person), wohin = year + (person ? ' / ' + person : '');
    if (z && dateienNamens_(z, name).length) { zeile('  ÜBERSPRUNGEN (gibt es in «' + wohin + '» schon): ' + name); bleibt++; return; }
    if (ausfuehren) f.moveTo(z);
    zeile('  ' + name + '  →  ' + wohin); n++;
  };
  // Jahresübersichten im Hauptordner
  var it = base.getFiles();
  var hf = []; while (it.hasNext()) hf.push(it.next());
  hf.forEach(function (f) {
    var name = nfc_(f.getName()), m = name.match(/^Jahresübersicht Löhne Assistenz (\d{4})(\.pdf)?$/);
    if (m) schiebe(f, Number(m[1]), '');
  });
  // Personenordner
  var fit = base.getFolders(), ordner = [];
  while (fit.hasNext()) ordner.push(fit.next());
  ordner.forEach(function (fo) {
    var on = nfc_(fo.getName());
    if (RE_JAHR_ORDNER.test(on) || on === TEST_ORDNER) return;
    if (namen.indexOf(on) < 0) { zeile('Ordner «' + on + '»: kein Name aus dem Blatt «Personen» – bleibt unverändert.'); return; }
    zeile('Ordner «' + on + '»:');
    var fi = fo.getFiles(), liste = [];
    while (fi.hasNext()) liste.push(fi.next());
    liste.forEach(function (f) {
      var name = nfc_(f.getName()), y = RE_PERSON_PDF.test(name) ? jahrAusName_(name) : 0;
      if (y) schiebe(f, y, on); else { zeile('  BLEIBT (anderer Dateiname): ' + name); bleibt++; }
    });
    var sub = fo.getFolders(); while (sub.hasNext()) { zeile('  BLEIBT (Unterordner): ' + sub.next().getName()); bleibt++; }
    if (ausfuehren && !fo.getFiles().hasNext() && !fo.getFolders().hasNext()) { fo.setTrashed(true); zeile('  leerer Ordner «' + on + '» in den Papierkorb'); }
  });
  zeile((ausfuehren ? 'Verschoben: ' : 'Würde verschieben: ') + n + ' Datei(en). Nicht verschoben: ' + bleibt + '.');
  return { verschoben: n, bleibt: bleibt, log: log };
}

/* =====================  Ablage durch die Abrechnungs-App  ===================== */
var HAUPTORDNER_NAME = 'Assistenzpersonenabrechnungen';
var TEST_ORDNER = '_TEST Beispieldaten';
var MAX_PDF = 5 * 1024 * 1024;

function nfc_(s) { s = String(s == null ? '' : s).trim(); return typeof s.normalize === 'function' ? s.normalize('NFC') : s; }

/**
 * Hauptordner = Ordner, in dem die Tabelle «Stundenerfassung Assistenz» liegt
 * (oder Skripteigenschaft HAUPTORDNER_ID). Zur Sicherheit muss er «Assistenzpersonenabrechnungen» heissen.
 */
function hauptordner_() {
  var id = PropertiesService.getScriptProperties().getProperty('HAUPTORDNER_ID');
  var f;
  if (id) f = DriveApp.getFolderById(id);
  else {
    var it = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
    if (!it.hasNext()) throw new Error('Hauptordner nicht gefunden.');
    f = it.next();
  }
  if (nfc_(f.getName()) !== HAUPTORDNER_NAME) throw new Error('Der Ordner der Stundentabelle heisst nicht «' + HAUPTORDNER_NAME + '». Ablage abgebrochen.');
  return f;
}
function unterordner_(parent, name, create) {
  var want = nfc_(name), it = parent.getFolders(), found = [];
  while (it.hasNext()) { var f = it.next(); if (nfc_(f.getName()) === want) found.push(f); }
  if (found.length > 1) throw new Error('Den Ordner «' + name + '» gibt es mehrfach. Bitte im Drive bereinigen.');
  if (found.length === 1) return found[0];
  if (create) return parent.createFolder(want);
  return null;
}
/** Basisordner: echter Hauptordner oder (Test) dessen Unterordner «_TEST Beispieldaten». */
function basis_(test) {
  var h = hauptordner_();
  return test ? unterordner_(h, TEST_ORDNER, true) : h;
}
function dateienNamens_(folder, name) {
  var want = nfc_(name), it = folder.getFiles(), out = [];
  while (it.hasNext()) { var f = it.next(); if (nfc_(f.getName()) === want) out.push(f); }
  return out;
}

var RE_PERSON_PDF = /^\d{4}-\d{2} (Lohnabrechnung|Stundenübersicht) .+\.pdf$/;
var RE_JAHR_PDF = /^Jahresübersicht Löhne Assistenz \d{4}\.pdf$/;

var RE_SICHERUNG = /^Sicherung Abrechnung \d{4}-\d{2}-\d{2}\.json$/;
var RE_SICHERUNG_ALT = /^Sicherung Abrechnung .*\.json$/;
var MAX_SICHERUNG = 10 * 1024 * 1024;

/**
 * {name, inhalt, test}: verschlüsselte Sicherung der App im Hauptordner ablegen.
 * Nur Dateien im Format der App (v, salt, iter, iv, ct) – unverschlüsselte Daten werden abgelehnt.
 * Neue Datei zuerst anlegen, danach alle älteren Sicherungen im selben Ordner in den Papierkorb (nur die neueste bleibt).
 */
function sicherung_(p) {
  var name = nfc_(p.name), inhalt = String(p.inhalt || '');
  if (!RE_SICHERUNG.test(name)) return { ok: false, error: 'Unerwarteter Dateiname: ' + name };
  if (!inhalt || inhalt.length > MAX_SICHERUNG) return { ok: false, error: 'Sicherung fehlt oder ist zu gross.' };
  var o = null; try { o = JSON.parse(inhalt); } catch (x) { o = null; }
  var keys = o && typeof o === 'object' && !Array.isArray(o) ? Object.keys(o).sort().join(',') : '';
  if (keys !== 'ct,iter,iv,salt,v' || typeof o.ct !== 'string' || typeof o.iv !== 'string' || typeof o.salt !== 'string') {
    return { ok: false, error: 'Keine verschlüsselte Sicherung der App.' };
  }
  var folder = basis_(p.test === true);
  var neu = folder.createFile(Utilities.newBlob(inhalt, 'application/json', name));
  var alt = 0, it = folder.getFiles(), weg = [];
  while (it.hasNext()) { var f = it.next(); if (f.getId() !== neu.getId() && RE_SICHERUNG_ALT.test(nfc_(f.getName()))) weg.push(f); }
  weg.forEach(function (f) { f.setTrashed(true); alt++; });
  return { ok: true, name: name, ordner: folder.getName(), ersetzt: alt };
}

/*
 * Ablage nach Jahren:
 *   Hauptordner/JJJJ/                     Jahresübersicht (Tabelle und PDF) dieses Jahres
 *   Hauptordner/JJJJ/[Kurzname]/          PDFs der Person aus diesem Jahr
 * Das Jahr kommt aus dem Dateinamen («2027-01 Lohnabrechnung …», «Jahresübersicht Löhne Assistenz 2027.pdf»).
 * Jahresordner legt das Script selbst an. Einen Personenordner legt es nur an, wenn der Kurzname genau so im Blatt «Personen» steht
 * (im Testmodus immer, dort gibt es nur erfundene Personen). Dateien aus der früheren Ablage (Hauptordner/[Kurzname]/ bzw. Hauptordner)
 * mit gleichem Namen kommen beim Ersetzen ebenfalls in den Papierkorb, damit nichts doppelt liegt.
 */
var RE_JAHR_ORDNER = /^\d{4}$/;
function jahrAusName_(n) {
  var m = n.match(/^(\d{4})-\d{2} /) || n.match(/ (\d{4})\.pdf$/);
  var y = m ? Number(m[1]) : 0;
  return y >= 2020 && y <= 2100 ? y : 0;
}
/** Namen (Spalte B) im Blatt «Personen» der Stundentabelle. */
function personenNamen_() {
  var v = sheet_(SH_P).getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) { var n = nfc_(v[i][1]); if (n) out.push(n); }
  return out;
}

/**
 * {ordner: Kurzname oder '' für die Jahresübersicht, files: [{name, b64}], test}
 * Neue Datei zuerst anlegen, erst danach alte gleichnamige Dateien in den Papierkorb.
 */
function ablegen_(p) {
  var test = p.test === true, ordner = nfc_(p.ordner), files = p.files || [], year = 0;
  if (!files.length || files.length > 4) return { ok: false, error: 'Keine oder zu viele Dateien.' };
  for (var i = 0; i < files.length; i++) {
    var n = nfc_(files[i].name);
    if (ordner ? !RE_PERSON_PDF.test(n) : !RE_JAHR_PDF.test(n)) return { ok: false, error: 'Unerwarteter Dateiname: ' + n };
    if (!files[i].b64 || files[i].b64.length > MAX_PDF * 1.4) return { ok: false, error: 'Datei fehlt oder ist zu gross: ' + n };
    var y = jahrAusName_(n);
    if (!y) return { ok: false, error: 'Jahr im Dateinamen ungültig: ' + n };
    if (year && y !== year) return { ok: false, error: 'Dateien aus verschiedenen Jahren in einem Schritt: ' + n };
    year = y;
  }
  if (ordner && (RE_JAHR_ORDNER.test(ordner) || ordner === TEST_ORDNER)) return { ok: false, error: 'Ungültiger Ordnername: ' + ordner };
  // Erst alle Dateien prüfen, dann schreiben: bei einem Fehler wird nichts ersetzt und kein Ordner angelegt.
  var blobs = files.map(function (fl) {
    var name = nfc_(fl.name), bytes = Utilities.base64Decode(fl.b64);
    if (bytes.length < 5 || String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) !== '%PDF') throw new Error('Keine PDF-Datei: ' + name);
    return { name: name, bytes: bytes };
  });
  var base = basis_(test);
  var jahrOrdner = unterordner_(base, String(year), false);
  var folder = null, alterOrt = base;
  if (ordner) {
    folder = jahrOrdner ? unterordner_(jahrOrdner, ordner, false) : null;
    if (!folder && !test && personenNamen_().indexOf(ordner) < 0) {
      return { ok: false, error: 'Den Namen «' + ordner + '» gibt es im Blatt «Personen» der Stundentabelle nicht. Der Kurzname in der App muss genau gleich heissen wie dort. Es wurde nichts abgelegt.' };
    }
    alterOrt = unterordner_(base, ordner, false);
  }
  if (!jahrOrdner) jahrOrdner = base.createFolder(String(year));
  if (ordner && !folder) folder = jahrOrdner.createFolder(ordner);
  if (!ordner) folder = jahrOrdner;
  var res = [];
  blobs.forEach(function (bl) {
    var name = bl.name;
    var neu = folder.createFile(Utilities.newBlob(bl.bytes, 'application/pdf', name));
    var alt = 0;
    dateienNamens_(folder, name).concat(alterOrt ? dateienNamens_(alterOrt, name) : []).forEach(function (f) { if (f.getId() !== neu.getId()) { f.setTrashed(true); alt++; } });
    res.push({ name: name, id: neu.getId(), ersetzt: alt });
  });
  return { ok: true, ordner: ordner ? year + ' / ' + folder.getName() : String(year), files: res };
}

function jahrName_(year) { return 'Jahresübersicht Löhne Assistenz ' + year; }
function tabellenNamens_(folder, year) {
  var it = folder.getFilesByType(MimeType.GOOGLE_SHEETS), want = nfc_(jahrName_(year)), found = [];
  while (it.hasNext()) { var f = it.next(); if (nfc_(f.getName()) === want) found.push(f); }
  return found;
}
/**
 * Jahrestabelle im Jahresordner suchen, sonst noch an der früheren Stelle (direkt im Hauptordner).
 * Mit verschieben = true kommt eine Tabelle von der früheren Stelle in den Jahresordner (gleiche Datei, gleicher Link).
 */
function jahrTabelle_(base, year, verschieben) {
  var jo = unterordner_(base, String(year), false);
  var found = (jo ? tabellenNamens_(jo, year) : []).concat(tabellenNamens_(base, year));
  if (found.length > 1) throw new Error('Die Tabelle «' + jahrName_(year) + '» gibt es mehrfach. Bitte im Drive bereinigen.');
  if (!found.length) return null;
  if (verschieben && !(jo && tabellenNamens_(jo, year).length)) found[0].moveTo(jo || base.createFolder(String(year)));
  return SpreadsheetApp.openById(found[0].getId());
}
var MON3 = ['jan', 'feb', 'mar', 'apr', 'mai', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dez'];
/** Monat aus der ersten Spalte: «Januar», «Jan», «März 2026», «2026-03» oder ein Datum. Sonst 0. */
function monatsZeile_(v) {
  if (v instanceof Date) return v.getMonth() + 1;
  if (v && typeof v === 'object' && v.m) return Number(v.m);
  var s = nfc_(v).toLowerCase().replace('ä', 'a');
  var m = s.match(/^\d{4}-(\d{2})$/);
  if (m) return Number(m[1]) >= 1 && Number(m[1]) <= 12 ? Number(m[1]) : 0;
  for (var i = 0; i < 12; i++) if (s.slice(0, 3) === MON3[i] && /^[a-z]{3}/.test(s)) return i + 1;
  return 0;
}
function gefuellt_(rows) {
  var n = 0;
  rows.forEach(function (r) {
    if (!monatsZeile_(r[0])) return;
    for (var j = 1; j < r.length; j++) if (r[j] !== '' && r[j] !== null && Number(String(r[j]).replace(/[’'\s]/g, '')) !== 0) { n++; return; }
  });
  return n;
}

/** {year, test} -> Werte der ersten Tabellenseite (zum einmaligen Übernehmen in die App). */
function jahrLesen_(p) {
  var year = Number(p.year);
  if (!(year >= 2020 && year <= 2100)) return { ok: false, error: 'Jahr ungültig.' };
  var ss = jahrTabelle_(basis_(p.test === true), year, false);
  if (!ss) return { ok: true, found: false, name: jahrName_(year) };
  var v = ss.getSheets()[0].getDataRange().getValues().map(function (r) {
    return r.map(function (c) { return c instanceof Date ? { m: c.getMonth() + 1 } : c; });
  });
  return { ok: true, found: true, name: ss.getName(), values: v };
}

/**
 * {year, header: [..], rows: [[Monat, Werte…]] (12 Zeilen), test}
 * Schreibt die Jahresübersicht in die bestehende Tabelle (gleiche Datei, gleicher Link) oder legt sie neu an.
 * Sicherung: Hat die bestehende Tabelle mehr ausgefüllte Monate als die neuen Daten, wird nichts geschrieben.
 */
function jahrSchreiben_(p) {
  var year = Number(p.year), header = p.header || [], rows = p.rows || [];
  if (!(year >= 2020 && year <= 2100)) return { ok: false, error: 'Jahr ungültig.' };
  if (rows.length !== 12 || header.length < 3) return { ok: false, error: 'Daten der Jahresübersicht unvollständig.' };
  var w = header.length;
  for (var i = 0; i < 12; i++) {
    if (!rows[i] || rows[i].length !== w || monatsZeile_(rows[i][0]) !== i + 1) return { ok: false, error: 'Zeile ' + (i + 1) + ' der Jahresübersicht ungültig.' };
    for (var j = 1; j < w; j++) if (rows[i][j] !== '' && !isFinite(Number(rows[i][j]))) return { ok: false, error: 'Ungültiger Wert in der Jahresübersicht.' };
  }
  var base = basis_(p.test === true);
  var ss = jahrTabelle_(base, year, false), neu = false;
  if (ss) {
    var alt = gefuellt_(ss.getSheets()[0].getDataRange().getValues());
    var jetzt = gefuellt_(rows);
    if (alt > jetzt) return { ok: false, error: 'Die Tabelle «' + jahrName_(year) + '» enthält ' + alt + ' Monate, die App erst ' + jetzt + '. Bitte zuerst unter «Jahr» die Jahresdaten übernehmen. Es wurde nichts geändert.', alt: alt, jetzt: jetzt };
    jahrTabelle_(base, year, true); // liegt sie noch direkt im Hauptordner: in den Jahresordner (gleiche Datei, gleicher Link)
  } else {
    ss = SpreadsheetApp.create(jahrName_(year)); neu = true;
    DriveApp.getFileById(ss.getId()).moveTo(unterordner_(base, String(year), true));
  }
  var sh = ss.getSheets()[0];
  sh.clear();
  sh.getRange(1, 1, 1, w).setValues([header]).setFontWeight('bold');
  sh.getRange(2, 1, 12, w).setValues(rows);
  var tot = ['Total'];
  for (var c = 2; c <= w; c++) { var col = colName_(c); tot.push('=SUM(' + col + '2:' + col + '13)'); }
  sh.getRange(14, 1, 1, 1).setValue('Total');
  sh.getRange(14, 2, 1, w - 1).setFormulas([tot.slice(1)]);
  sh.getRange(14, 1, 1, w).setFontWeight('bold');
  sh.getRange(2, 2, 13, w - 1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(1);
  return { ok: true, neu: neu, name: ss.getName(), url: ss.getUrl() };
}
function colName_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

/**
 * {to, cc, subject, body, files: [{name, b64}], test}
 * Legt einen Gmail-ENTWURF an. Ein bestehender Entwurf mit genau gleichem Betreff wird ersetzt (keine Doppel).
 * Es wird nie etwas versendet.
 */
function entwurf_(p) {
  var test = p.test === true;
  var to = nfc_(p.to), cc = nfc_(p.cc), subject = nfc_(p.subject), body = String(p.body || ''), files = p.files || [];
  var mail = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/;
  if (!mail.test(to)) return { ok: false, error: 'E-Mail-Adresse der Person fehlt oder ist ungültig.' };
  if (cc && !mail.test(cc)) return { ok: false, error: 'CC-Adresse ist ungültig.' };
  var testOk = function (a) { return /@example\.com$/i.test(a) || a.toLowerCase() === eigene_(); };
  if (test && (!testOk(to) || (cc && !testOk(cc)))) return { ok: false, error: 'Im Testmodus nur Adressen @example.com oder die eigene Adresse.' };
  // Schon beim Entwurf prüfen, ob die Empfänger erlaubt sind: So merkt die App den Fehler, bevor irgendetwas gesendet wird.
  if (!test) {
    var erlaubt = adressen_(PropertiesService.getScriptProperties().getProperty('ERLAUBTE_EMPFAENGER') || ''), ich = eigene_();
    var alle = [to.toLowerCase()].concat(cc ? [cc.toLowerCase()] : []);
    for (var j = 0; j < alle.length; j++) {
      if (alle[j] !== ich && erlaubt.indexOf(alle[j]) < 0) return { ok: false, error: 'Adresse ' + alle[j] + ' ist nicht in der Liste der erlaubten Empfänger (Skripteigenschaft ERLAUBTE_EMPFAENGER) – kein Entwurf erstellt, nichts gesendet.' };
    }
  }
  if (test) subject = '[TEST] ' + subject;
  if (!/^(\[TEST\] )?Lohnabrechnung .+ \/ .+$/.test(subject)) return { ok: false, error: 'Betreff ungültig.' };
  if (!files.length || files.length > 2) return { ok: false, error: 'Anhänge fehlen.' };
  var att = files.map(function (fl) {
    var name = nfc_(fl.name);
    if (!RE_PERSON_PDF.test(name)) throw new Error('Unerwarteter Anhang: ' + name);
    var bytes = Utilities.base64Decode(fl.b64 || '');
    if (bytes.length < 5 || String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) !== '%PDF') throw new Error('Keine PDF-Datei: ' + name);
    return Utilities.newBlob(bytes, 'application/pdf', name);
  });
  // Gleichnamige Entwürfe gezielt suchen (schnell, auch bei vielen alten Entwürfen) und danach ersetzen.
  var ersetzt = 0, ids = {}, n = 0;
  GmailApp.search('in:drafts subject:"' + subject.replace(/"/g, '') + '"', 0, 20).forEach(function (t) {
    t.getMessages().forEach(function (m) { if (m.isDraft() && nfc_(m.getSubject()) === subject) { ids[m.getId()] = 1; n++; } });
  });
  if (n) GmailApp.getDrafts().forEach(function (d) { if (ids[d.getMessageId()]) { d.deleteDraft(); ersetzt++; } });
  var opt = { attachments: att };
  if (cc) opt.cc = cc;
  var d = GmailApp.createDraft(to, subject, body, opt);
  return { ok: true, id: d.getId(), ersetzt: ersetzt, subject: subject };
}

function eigene_() { return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); }
/** E-Mail-Adressen aus «Name <a@b>, c@d» */
function adressen_(s) {
  return String(s || '').split(',').map(function (x) { var m = x.match(/<([^>]+)>/); return (m ? m[1] : x).trim().toLowerCase(); }).filter(function (x) { return x; });
}

/**
 * {draftId, test} – sendet einen von der App erstellten Entwurf.
 * Prüft vorher: Empfänger erlaubt, Betreff «Lohnabrechnung …», nur PDF-Anhänge mit erwartetem Namen.
 * Testmodus: Adressen @example.com werden nicht wirklich gesendet (nur geprüft); an die eigene Adresse wird echt gesendet.
 */
function senden_(p) {
  var test = p.test === true, id = String(p.draftId || '');
  if (!id) return { ok: false, error: 'Entwurf fehlt.' };
  var d;
  try { d = GmailApp.getDraft(id); } catch (x) { d = null; }
  if (!d) return { ok: false, error: 'Entwurf nicht gefunden (schon gesendet oder in Gmail gelöscht?).' };
  var m = d.getMessage(), subject = nfc_(m.getSubject());
  if (!(test ? /^\[TEST\] Lohnabrechnung .+ \/ .+$/ : /^Lohnabrechnung .+ \/ .+$/).test(subject)) return { ok: false, error: 'Betreff passt nicht – nicht gesendet.' };
  var att = m.getAttachments();
  if (!att.length || att.length > 2) return { ok: false, error: 'Anhänge fehlen – nicht gesendet.' };
  for (var i = 0; i < att.length; i++) {
    if (!RE_PERSON_PDF.test(nfc_(att[i].getName())) || !/pdf/i.test(att[i].getContentType())) return { ok: false, error: 'Unerwarteter Anhang – nicht gesendet.' };
  }
  var an = adressen_(m.getTo()).concat(adressen_(m.getCc()), adressen_(m.getBcc ? m.getBcc() : ''));
  if (!adressen_(m.getTo()).length) return { ok: false, error: 'Kein Empfänger – nicht gesendet.' };
  var erlaubt = adressen_(PropertiesService.getScriptProperties().getProperty('ERLAUBTE_EMPFAENGER') || '');
  var ich = eigene_(), simuliert = false;
  for (var j = 0; j < an.length; j++) {
    var a = an[j];
    if (test && /@example\.com$/.test(a)) { simuliert = true; continue; }
    if (a !== ich && erlaubt.indexOf(a) < 0) return { ok: false, error: 'Adresse ' + a + ' ist nicht in der Liste der erlaubten Empfänger (Skripteigenschaft ERLAUBTE_EMPFAENGER) – nicht gesendet.' };
  }
  if (simuliert) return { ok: true, gesendet: false, simuliert: true, subject: subject };
  d.send();
  return { ok: true, gesendet: true, subject: subject };
}

/** Einmalig ausführen: legt die Tabellenblätter mit Kopfzeilen an. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p = ss.getSheetByName(SH_P) || ss.insertSheet(SH_P);
  if (p.getLastRow() < 1) p.getRange(1, 1, 1, 3).setValues([['Token', 'Name', 'aktiv']]);
  var s = ss.getSheetByName(SH_S) || ss.insertSheet(SH_S);
  if (s.getLastRow() < 1) s.getRange(1, 1, 1, 4).setValues([['Name', 'Datum', 'Dauer', 'Geändert']]);
  s.getRange('B:B').setNumberFormat('@');
  absSheet_();
}
