/** ROUTINE TRACKER - backend (Google Apps Script + Google Sheets) */
const SCHEMA = {
  Schedule: ['id','week','day','name','start','end','category','desc','done','notes','reminder'],
  Todos: ['id','date','title','priority','deadline','category','notes','done'],
  Weeks: ['id','week','title','archived','review'],
  Templates: ['id','name','data'],
  Care: ['id','category','step','product','desc','order','frequency','duration','reminder','done'],
  Wishlist: ['id','name','category','brand','price','priority','reason','image','status'],
  Products: ['id','name','brand','category','image','purchaseDate','price','expiry','pao','openDate','qty','location','status','notes'],
  Vitamins: ['id','name','dosage','time','frequency','start','end','notes','doneDates'],
  Exercise: ['id','name','category','date','start','duration','desc','done','notes'],
  Plants: ['id','name','common','category','sci','source','image','sowDate','plantDate','location','polybag','medium','status','notes'],
  Fertilizing: ['id','plant','fertilizer','type','date','next','notes','done'],
  Watering: ['id','plant','date','frequency','notes','done'],
  Growth: ['id','plant','date','stage','height','image','notes'],
  Stationery: ['id','name','brand','category','qty','color','price','status','notes'],
  Reading: ['id','title','type','author','status','progress','rating','notes','image'],
  Shopping: ['id','item','category','qty','price','priority','done'],
  Settings: ['id','value']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Routine Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function sheet_(n) {
  if (!SCHEMA[n]) throw new Error('Unknown sheet: ' + n);
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n);
    s.getRange(1, 1, 1, SCHEMA[n].length).setValues([SCHEMA[n]]).setFontWeight('bold').setBackground('#F4DCE5');
    s.setFrozenRows(1);
    s.getRange(2, 1, Math.max(s.getMaxRows() - 1, 1), SCHEMA[n].length).setNumberFormat('@'); // keep dates/times as plain text
  }
  else if (s.getLastColumn() < SCHEMA[n].length) {   // add newly introduced columns to older sheets
    const w = SCHEMA[n].length;
    s.getRange(1, 1, 1, w).setValues([SCHEMA[n]]).setFontWeight('bold').setBackground('#F4DCE5');
    s.getRange(2, 1, Math.max(s.getMaxRows() - 1, 1), w).setNumberFormat('@');
  }
  return s;
}

function setup() { Object.keys(SCHEMA).forEach(sheet_); }   // run once manually

function list_(n, sh) {
  const s = sh || sheet_(n);
  const v = s.getDataRange().getValues();
  if (!v.length) return [];
  if (v[0].length < SCHEMA[n].length) sheet_(n);   // older sheet: adds the newly introduced columns
  v.shift();                                       // header row; column order follows SCHEMA
  const h = SCHEMA[n];
  return v.filter(r => r[0] !== '').map(r => {
    const o = {};
    h.forEach((k, i) => o[k] = r[i] === undefined || r[i] === null ? '' : String(r[i]));
    return o;
  });
}

function loadAll() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('This script is not attached to a Google Sheet. Open your Google Sheet and create the script from Extensions > Apps Script.');
  const names = Object.keys(SCHEMA), o = {}, errs = [], map = {};
  ss.getSheets().forEach(x => map[x.getName()] = x);   // one call instead of one lookup per sheet
  let cache = null, hit = {};
  try { cache = CacheService.getUserCache(); hit = cache.getAll(names.map(n => 't_' + n)) || {}; } catch (e) { cache = null; }
  names.forEach(n => {
    try { if (hit['t_' + n]) { o[n] = JSON.parse(hit['t_' + n]); return; } } catch (e) {}
    try {
      o[n] = map[n] ? list_(n, map[n]) : [];          // missing tab = empty; it is created on first save or by setup()
      if (cache && map[n]) { const j = JSON.stringify(o[n]); if (j.length < 90000) cache.put('t_' + n, j, 600); }
    } catch (e) { if (!o[n]) o[n] = []; errs.push(n + ': ' + e.message); }
  });
  if (errs.length) o._err = errs;
  return o;
}
function bust_(n) { try { CacheService.getUserCache().remove('t_' + n); } catch (e) {} }

/** Diagnostic: run from the editor and read the Execution log to see if every sheet loads. */
function testLoad() {
  const d = loadAll();
  Logger.log(Object.keys(SCHEMA).map(k => k + ': ' + (d[k] || []).length).join('\n'));
  if (d._err) Logger.log('ERRORS:\n' + d._err.join('\n'));
}

function save(n, o) { return saveMany(n, [o])[0]; }

function saveMany(n, arr) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sheet_(n), h = SCHEMA[n];
    const last = s.getLastRow();
    const ids = last > 1 ? s.getRange(2, 1, last - 1, 1).getValues().flat() : [];
    const adds = [];
    arr.forEach(o => {
      if (!o.id) o.id = Utilities.getUuid();
      const row = h.map(k => o[k] === undefined || o[k] === null ? '' : String(o[k]));
      const i = ids.indexOf(o.id);
      if (i >= 0) s.getRange(i + 2, 1, 1, h.length).setValues([row]);
      else adds.push(row);
    });
    if (adds.length) {
      const start = Math.max(s.getLastRow(), 1) + 1;
      s.getRange(start, 1, adds.length, h.length).setNumberFormat('@').setValues(adds);
    }
    bust_(n);
    return arr;
  } finally { lock.releaseLock(); }
}

function remove(n, id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sheet_(n), last = s.getLastRow();
    if (last < 2) return false;
    const i = s.getRange(2, 1, last - 1, 1).getValues().flat().indexOf(id);
    if (i >= 0) s.deleteRow(i + 2);
    bust_(n);
    return i >= 0;
  } finally { lock.releaseLock(); }
}

/** Run once from the editor to grant Google Drive permission (needed for plant photo uploads). */
function authorizeDrive() { DriveApp.getRootFolder(); }

/** Saves an uploaded photo (data URL) into Drive folder "Routine Tracker Images"; returns a viewable link. */
function uploadImage(dataUrl, name) {
  const m = /^data:(.+?);base64,(.+)$/.exec(dataUrl);
  if (!m) throw new Error('Invalid image data');
  const it = DriveApp.getFoldersByName('Routine Tracker Images');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Routine Tracker Images');
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (name || 'plant') + '-' + Date.now() + '.jpg'));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
}

/** Connectivity check used for troubleshooting. */
function ping() { return 'ok ' + new Date().toISOString(); }
