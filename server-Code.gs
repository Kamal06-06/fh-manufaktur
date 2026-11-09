/**
 * FH Manufaktur – Server (Google Apps Script)
 * Fragrance Hub Germany · Parfümkonzentrat: Chargen, SDB/IFRA, Rückverfolgung, Shop
 *
 * Einrichtung (einmalig):
 *  1. Projekt-Einstellungen → Skripteigenschaften:
 *       ADMIN_PIN      = Inhaber-Passwort (mind. 8 Zeichen, Buchstaben + Ziffern)
 *       AI_KEY         = API-Schlüssel für das automatische Lesen der Zertifikate
 *       SHOPIFY_SHOP   = z. B. fragranceshub.myshopify.com
 *       SHOPIFY_TOKEN  = Admin-API-Token (Rechte: read_orders, write_products, write_translations)
 *  2. Funktion «setup» einmal ausführen (legt Tabelle + Ordner an).
 *  3. Bereitstellen → Neue Bereitstellung → Web-App: Ausführen als «Ich», Zugriff «Jeder».
 */

const CODES_SHEET_DEFAULT = '1qETeGWZcHDYu7vdP0PxMVHgbyrUk9vv_TwQSL56fAk4';
const COLLECTIONS = ['products', 'batches', 'config'];
const CHUNK = 45000;
const P = PropertiesService.getScriptProperties();

/* ---------------- setup ---------------- */
function setup() {
  let id = P.getProperty('DATA_SHEET');
  let ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('FH Manufaktur – Daten');
  P.setProperty('DATA_SHEET', ss.getId());
  COLLECTIONS.concat(['users', 'log']).forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Tabelle1') || ss.getSheetByName('ورقة1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  const u = ss.getSheetByName('users');
  if (u.getLastRow() === 0) u.appendRow(['name', 'pinHash', 'role', 'active', 'created']);
  const l = ss.getSheetByName('log');
  if (l.getLastRow() === 0) l.appendRow(['Zeit', 'Benutzer', 'Aktion', 'Bereich', 'ID', 'Info']);
  if (!P.getProperty('FILES_FOLDER')) P.setProperty('FILES_FOLDER', DriveApp.createFolder('FH Manufaktur – Dateien').getId());
  if (!P.getProperty('CODES_SHEET')) P.setProperty('CODES_SHEET', CODES_SHEET_DEFAULT);
  Logger.log('OK – Daten: ' + ss.getUrl());
}

/* ---------------- http ---------------- */
function doGet() {
  return out({ ok: true, app: 'FH Manufaktur' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out({ err: 'BAD_REQUEST' }); }
  try {
    const me = auth(req.pin);
    const val = route(me, req.action, req.args || {});
    return out({ val: val });
  } catch (x) {
    return out({ err: String(x && x.message || x) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- auth ---------------- */
function hash(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'fhm:' + s, Utilities.Charset.UTF_8));
}

function auth(pin) {
  pin = String(pin || '');
  if (!pin) throw new Error('PIN');
  const c = CacheService.getScriptCache();
  const fails = +(c.get('fails') || 0);
  if (fails > 20) throw new Error('Zu viele Versuche');
  const admin = P.getProperty('ADMIN_PIN');
  if (admin && pin === admin) return { name: 'Inhaber', admin: true };
  const h = hash(pin);
  const rows = sh('users').getDataRange().getValues().slice(1);
  for (const r of rows) {
    if (r[1] === h && r[3] !== false && String(r[3]).toUpperCase() !== 'FALSE') return { name: String(r[0]), admin: r[2] === 'admin' };
  }
  c.put('fails', String(fails + 1), 600);
  throw new Error('PIN');
}

/* ---------------- router ---------------- */
function route(me, action, a) {
  switch (action) {
    case 'check': return { user: me.name, admin: me.admin };
    case 'list': return list(a.col);
    case 'set': return setDoc(me, a.col, a.id, a.data);
    case 'add': return setDoc(me, a.col, Utilities.getUuid().replace(/-/g, '').slice(0, 20), a.data);
    case 'del': return delDoc(me, a.col, a.id);
    case 'upload': return upload(me, a);
    case 'ai': return ai(a);
    case 'drive_search': return driveSearch(a);
    case 'drive_download': return driveDownload(a);
    case 'drive_create': return driveCreate(me, a);
    case 'drive_share': return driveShare(a);
    case 'drive_rename': return driveRename(a);
    case 'sheet_get': return sheetGet(a);
    case 'sheet_append': return sheetAppend(me, a);
    case 'shopify': return shopify(me, a);
    case 'users': return users(me, a);
    case 'log': return logList(me);
    default: throw new Error('Unbekannte Aktion: ' + action);
  }
}

/* ---------------- storage: one sheet per collection, JSON in chunks ---------------- */
function ss() { return SpreadsheetApp.openById(P.getProperty('DATA_SHEET')); }
function sh(n) { const s = ss().getSheetByName(n); if (!s) throw new Error('Setup fehlt: ' + n); return s; }
function colOk(c) { if (COLLECTIONS.indexOf(c) < 0) throw new Error('Bereich unbekannt'); return c; }

function list(col) {
  const s = sh(colOk(col));
  if (s.getLastRow() === 0) return [];
  return s.getDataRange().getValues().filter(r => r[0]).map(r => {
    let j = '';
    for (let i = 3; i < r.length; i++) j += r[i] || '';
    let d = {};
    try { d = JSON.parse(j); } catch (x) {}
    return { id: String(r[0]), data: d, updated: r[1] };
  });
}

function setDoc(me, col, id, data) {
  colOk(col);
  if (col === 'config' && !me.admin && id !== 'settings') throw new Error('NUR_INHABER');
  const json = JSON.stringify(data || {});
  const parts = [];
  for (let i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
  const row = [String(id), new Date(), me.name].concat(parts);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sh(col);
    const ids = s.getLastRow() ? s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0])) : [];
    const i = ids.indexOf(String(id));
    const r = i >= 0 ? i + 1 : s.getLastRow() + 1;
    const width = Math.max(s.getLastColumn(), row.length);
    const full = row.concat(new Array(width - row.length).fill(''));
    s.getRange(r, 1, 1, width).setValues([full]);
    if (col === 'batches' || i < 0) log(me, i < 0 ? 'neu' : 'geändert', col, id, data && (data.batchNo || data.name || '') + (data && data.status ? ' · ' + data.status : ''));
  } finally { lock.releaseLock(); }
  return { id: String(id) };
}

function delDoc(me, col, id) {
  colOk(col);
  if (!me.admin) throw new Error('NUR_INHABER');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sh(col);
    const ids = s.getRange(1, 1, Math.max(1, s.getLastRow()), 1).getValues().map(r => String(r[0]));
    const i = ids.indexOf(String(id));
    if (i >= 0) s.deleteRow(i + 1);
    log(me, 'gelöscht', col, id, '');
  } finally { lock.releaseLock(); }
  return true;
}

function log(me, what, col, id, info) {
  try { sh('log').appendRow([new Date(), me.name, what, col, id, String(info || '').slice(0, 200)]); } catch (x) {}
}

function logList(me) {
  if (!me.admin) throw new Error('NUR_INHABER');
  const s = sh('log');
  const n = s.getLastRow();
  if (n < 2) return [];
  const from = Math.max(2, n - 199);
  return s.getRange(from, 1, n - from + 1, 6).getValues().reverse();
}

/* ---------------- files ---------------- */
function folder() { return DriveApp.getFolderById(P.getProperty('FILES_FOLDER')); }

function upload(me, a) {
  const blob = Utilities.newBlob(Utilities.base64Decode(a.data), a.type || 'application/octet-stream', a.name || 'Datei');
  const f = folder().createFile(blob);
  log(me, 'Datei', 'files', f.getId(), a.name);
  return { id: f.getId(), url: f.getUrl() };
}

/* ---------------- automatic reading of certificates ---------------- */
function ai(a) {
  const key = P.getProperty('AI_KEY');
  if (!key) throw new Error('AI_FEHLT');
  const content = [];
  (a.images || []).forEach(im => content.push({ type: 'image', source: { type: 'base64', media_type: im.type || 'image/jpeg', data: im.data } }));
  content.push({ type: 'text', text: a.prompt + '\n\nAntworte nur mit gültigem JSON, ohne Erklärung.' });
  const body = { model: P.getProperty('AI_MODEL') || 'claude-sonnet-5-5', max_tokens: 16000, messages: [{ role: 'user', content: content }] };
  let res, code;
  for (let i = 0; i < 3; i++) {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' }, payload: JSON.stringify(body)
    });
    code = res.getResponseCode();
    if (code === 200) break;
    if (code !== 429 && code < 500) break;
    Utilities.sleep(3000 * (i + 1));
  }
  if (code !== 200) throw new Error('Lesedienst ' + code + ': ' + res.getContentText().slice(0, 200));
  const txt = (JSON.parse(res.getContentText()).content || []).filter(c => c.type === 'text').map(c => c.text).join('');
  const m = txt.replace(/```json|```/g, '').trim();
  const s = m.indexOf('{'), e = m.lastIndexOf('}');
  try { return JSON.parse(m.slice(s, e + 1)); } catch (x) { throw new Error('Antwort nicht lesbar'); }
}

/* ---------------- Google Drive ---------------- */
function driveSearch(a) {
  const it = DriveApp.searchFiles(a.query);
  const files = [];
  while (it.hasNext() && files.length < (a.pageSize || 20)) {
    const f = it.next();
    files.push({ id: f.getId(), title: f.getName(), mimeType: f.getMimeType(), modifiedTime: f.getLastUpdated().toISOString(), fileSize: String(f.getSize()), viewUrl: f.getUrl() });
  }
  return { files: files };
}

function driveDownload(a) {
  const f = DriveApp.getFileById(a.fileId);
  if (f.getSize() > 30 * 1024 * 1024) throw new Error('Datei zu groß');
  return { id: f.getId(), title: f.getName(), mimeType: f.getMimeType(), content: Utilities.base64Encode(f.getBlob().getBytes()) };
}

function driveCreate(me, a) {
  const blob = Utilities.newBlob(Utilities.base64Decode(a.base64Content), a.contentMimeType || 'application/pdf', a.title);
  const f = DriveApp.getFolderById(a.parentId).createFile(blob);
  log(me, 'Dokument', 'drive', f.getId(), a.title);
  return { id: f.getId(), viewUrl: f.getUrl() };
}

function driveShare(a) {
  const f = DriveApp.getFileById(a.fileId);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { shareableLink: f.getUrl() };
}

function driveRename(a) {
  DriveApp.getFileById(a.fileId).setName(a.title);
  return true;
}

/* ---------------- code table ---------------- */
function codesId() {
  return P.getProperty('CODES_SHEET') || CODES_SHEET_DEFAULT; // always the company code table, never another file
}

function sheetGet(a) {
  const s = SpreadsheetApp.openById(codesId(a)).getSheets()[0];
  return { values: s.getRange(a.range).getDisplayValues().filter(r => r.some(v => v !== '')) };
}

function sheetAppend(me, a) {
  const s = SpreadsheetApp.openById(codesId(a)).getSheets()[0];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const last = s.getRange('C:C').getValues().filter(r => r[0] !== '').length;
    s.getRange(last + 1, 3, a.values.length, a.values[0].length).setValues(a.values);
  } finally { lock.releaseLock(); }
  log(me, 'Code-Tabelle', 'codes', '', JSON.stringify(a.values).slice(0, 150));
  return true;
}

/* ---------------- Shopify ---------------- */
function shopify(me, a) {
  const shop = P.getProperty('SHOPIFY_SHOP'), token = P.getProperty('SHOPIFY_TOKEN');
  if (!shop || !token) throw new Error('SHOP_FEHLT');
  if (/^\s*mutation/.test(a.query) && !me.admin) throw new Error('NUR_INHABER');
  const res = UrlFetchApp.fetch('https://' + shop + '/admin/api/2025-10/graphql.json', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'X-Shopify-Access-Token': token }, payload: JSON.stringify({ query: a.query, variables: a.variables || {} })
  });
  const j = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200 || j.errors) throw new Error('Shopify: ' + JSON.stringify(j.errors || res.getResponseCode()).slice(0, 300));
  if (/^\s*mutation/.test(a.query)) log(me, 'Shopify', 'shop', '', a.query.slice(0, 60));
  return j;
}

/* ---------------- employees ---------------- */
function users(me, a) {
  if (!me.admin) throw new Error('NUR_INHABER');
  const s = sh('users');
  const rows = s.getDataRange().getValues();
  if (a.op === 'list') return rows.slice(1).map(r => ({ name: r[0], role: r[2], active: r[3] !== false && String(r[3]).toUpperCase() !== 'FALSE' }));
  if (a.op === 'save') {
    const name = String(a.name || '').trim();
    if (!name) throw new Error('Name fehlt');
    const pin = String(a.pin || '');
    if (pin && (pin.length < 6 || !/\d/.test(pin))) throw new Error('PIN: mind. 6 Zeichen mit Ziffern');
    const h = pin ? hash(pin) : null;
    if (h && rows.slice(1).some(r => r[1] === h && r[0] !== name)) throw new Error('Dieser PIN ist schon vergeben');
    const i = rows.findIndex((r, k) => k > 0 && r[0] === name);
    if (i > 0) {
      if (h) s.getRange(i + 1, 2).setValue(h);
      s.getRange(i + 1, 3, 1, 2).setValues([[a.role === 'admin' ? 'admin' : 'staff', a.active !== false]]);
    } else {
      if (!h) throw new Error('PIN fehlt');
      s.appendRow([name, h, a.role === 'admin' ? 'admin' : 'staff', true, new Date()]);
    }
    log(me, 'Mitarbeiter', 'users', name, a.active === false ? 'gesperrt' : 'gespeichert');
    return true;
  }
  throw new Error('op');
}
