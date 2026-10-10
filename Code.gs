/**
 * Code.gs: backend API for the MyBudjet dashboard hosted on GitHub Pages (index.html).
 * Deploy as Web app: Execute as Me, Who has access: Anyone. Protected by the API_TOKEN script property.
 */

const DASH_SS_ID = '1AWH16EV239lPmdIAXTdz46d38uF0s9dJdvf-okukz8M';
const DASH_SHEET = 'Entries';      // the new responses table the dashboard is attached to
const DASH_SOURCE = 'Responses';   // copied once to create DASH_SHEET
const DASH_CB_COL = 'P';           // sheet column holding the extra cash back (your SUM(P2:P))

function dashColIndex_(letter) {   // 'A' -> 0, 'P' -> 15, 'AA' -> 26
  let n = 0; String(letter).toUpperCase().split('').forEach(function (ch) { n = n * 26 + (ch.charCodeAt(0) - 64); });
  return n - 1;
}

function dashNorm_(s) { return String(s).replace(/:/g, '').trim().toLowerCase(); }
function dashStr_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function dashNum_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/,/g, '').replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}
function dashIso_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const s = dashStr_(v);
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    const y = m[3].length === 2 ? '20' + m[3] : m[3];
    return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : '';
}

/** Returns the Entries sheet; creates it from Responses the first time. */
function dashSheet_(ss) {
  let sh = ss.getSheetByName(DASH_SHEET);
  if (sh) return sh;
  const src = ss.getSheetByName(DASH_SOURCE);
  if (src) {
    sh = src.copyTo(ss);
    sh.setName(DASH_SHEET);
    return sh;
  }
  sh = ss.insertSheet(DASH_SHEET);
  const headers = ['Timestamp', 'Withdrawn', 'Operation type', 'Deposited', 'Amount', 'Date', 'H', 'M',
    'SMS', 'Extras', 'Details', 'Fees', 'Amount', 'SMS', '1st Other Account', '2nd Other Account',
    'Category', 'Sub-category'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function dashColumns_(headerRow) {
  const hdr = headerRow.map(dashNorm_);
  const find = function (name, nth) {
    let n = 0;
    for (let i = 0; i < hdr.length; i++) {
      if (hdr[i] === name) { if (n === (nth || 0)) return i; n++; }
    }
    return -1;
  };
  let a1 = find('1st other acount'); if (a1 < 0) a1 = find('1st other account');
  let a2 = find('2nd other acount'); if (a2 < 0) a2 = find('2nd other account');
  return {
    ts: find('timestamp'), wd: find('withdrawn'), op: find('operation type'),
    dep: find('deposited'), amt: find('amount', 0), date: find('date'),
    h: find('h'), m: find('m'), sms: find('sms', 0), ex: find('extras'),
    det: find('details'), fee: find('fees'), amt2: find('amount', 1),
    sms2: find('sms', 1), url: find('form response edit url'),
    cat: find('category'), sub: find('sub-category'), o1: a1, o2: a2, count: hdr.length
  };
}

/* ---- Short cache of the sheet read (the slow part). Rows only; settings are always read fresh. ---- */
const DASH_CACHE_TTL = 60;   // seconds. Edits made straight in the sheet show up after at most this long, or at once with "Reload data".

function dashCachePut_(obj) {
  try {
    const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(obj), 'application/json')).getBytes());
    const CH = 90000, parts = Math.ceil(b64.length / CH);
    if (parts > 90) return;                              // too big for the cache: just skip it
    const pairs = { dashRows_n: String(parts) };
    for (let i = 0; i < parts; i++) pairs['dashRows_' + i] = b64.substr(i * CH, CH);
    CacheService.getScriptCache().putAll(pairs, DASH_CACHE_TTL);
  } catch (e) { /* caching is optional */ }
}
function dashCacheGet_() {
  try {
    const c = CacheService.getScriptCache(), n = parseInt(c.get('dashRows_n'), 10);
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push('dashRows_' + i);
    const got = c.getAll(keys); let b64 = '';
    for (let i = 0; i < n; i++) { const part = got['dashRows_' + i]; if (!part) return null; b64 += part; }
    const json = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString();
    return JSON.parse(json);
  } catch (e) { return null; }
}
function dashCacheClear_() {
  try {
    const c = CacheService.getScriptCache(), n = parseInt(c.get('dashRows_n'), 10) || 0, keys = ['dashRows_n'];
    for (let i = 0; i < n; i++) keys.push('dashRows_' + i);
    c.removeAll(keys);
  } catch (e) { /* ignore */ }
}

function dashReadSheet_() {
  const ss = SpreadsheetApp.openById(DASH_SS_ID);
  const sh = dashSheet_(ss);
  const tz = ss.getSpreadsheetTimeZone();
  const vals = sh.getDataRange().getValues();
  const cbIdx = dashColIndex_(DASH_CB_COL);

  if (!vals || vals.length === 0) {
    return { rows: [], cbCol: { letter: DASH_CB_COL, header: '' }, missing: [] };
  }

  const C = dashColumns_(vals[0]);
  const rows = [];
  for (let r = 1; r < vals.length; r++) {
    const v = vals[r];
    const g = function (i) { return i < 0 ? '' : v[i]; };
    if (!g(C.ts) && !g(C.amt)) continue;
    const tsv = g(C.ts);
    rows.push({
      r: r + 1,
      ts: tsv instanceof Date ? Utilities.formatDate(tsv, tz, 'yyyy-MM-dd HH:mm:ss') : dashStr_(tsv),
      wd: dashStr_(g(C.wd)), dep: dashStr_(g(C.dep)), op: dashStr_(g(C.op)),
      amt: dashNum_(g(C.amt)), date: dashIso_(g(C.date), tz) || dashIso_(tsv, tz),
      h: dashNum_(g(C.h)), m: dashNum_(g(C.m)), ex: dashStr_(g(C.ex)),
      det: dashStr_(g(C.det)), fee: dashNum_(g(C.fee)), amt2: dashNum_(g(C.amt2)),
      cbp: cbIdx < v.length ? dashNum_(v[cbIdx]) : 0,
      sms: dashStr_(g(C.sms)), url: dashStr_(g(C.url)),
      cat: dashStr_(g(C.cat)), sub: dashStr_(g(C.sub))
    });
  }
  const cbHeader = (vals[0] && cbIdx < vals[0].length) ? dashStr_(vals[0][cbIdx]) : '';
  return {
    rows: rows,
    cbCol: { letter: DASH_CB_COL, header: cbHeader },
    missing: Object.keys(C).filter(function (k) { return C[k] < 0; })
  };
}

/** force = true (the "Reload data" button) skips the cache. */
function dashGetData(force) {
  let base = force ? null : dashCacheGet_();
  if (!base) { base = dashReadSheet_(); dashCachePut_(base); }
  const raw = PropertiesService.getScriptProperties().getProperty('dashCfg');
  return { rows: base.rows, cbCol: base.cbCol, missing: base.missing, cfg: raw ? JSON.parse(raw) : {} };
}

function dashSaveCfg(json) {
  PropertiesService.getScriptProperties().setProperty('dashCfg', json);
  return true;
}

function dashSaveCats(items) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = dashSheet_(SpreadsheetApp.openById(DASH_SS_ID));
    let hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    let C = dashColumns_(hdr);
    if (C.cat < 0) { sh.getRange(1, sh.getLastColumn() + 1).setValue('Category'); }
    hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    C = dashColumns_(hdr);
    if (C.sub < 0) { sh.getRange(1, sh.getLastColumn() + 1).setValue('Sub-category'); }
    hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    C = dashColumns_(hdr);
    items.forEach(function (it) {
      sh.getRange(it.row, C.cat + 1, 1, 2).setValues([[it.cat || '', it.sub || '']]);
    });
    SpreadsheetApp.flush();
    dashCacheClear_();
    return items.length;
  } finally {
    lock.releaseLock();
  }
}

function dashAddEntry(d) {
  d = d || {};
  const amount = parseFloat(d.amount);
  const h = parseInt(d.h, 10), m = parseInt(d.m, 10);
  if (!(amount > 0)) throw new Error('Amount must be a number above 0.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.date))) throw new Error('Date is missing or invalid.');
  if (!(h >= 0 && h <= 23) || !(m >= 0 && m <= 59)) throw new Error('Hour or minute is invalid.');
  if (d.type === 'Other' && !d.withdrawn) throw new Error('Withdrawn account is missing.');
  if ((d.type === 'Incomes' || d.type === 'Refund' || d.operationType === 'Between Acounts') && !d.deposited)
    throw new Error('Deposited account is missing.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(DASH_SS_ID);
    const sh = dashSheet_(ss);
    const tz = ss.getSpreadsheetTimeZone();
    const lastCol = sh.getLastColumn();
    const C = dashColumns_(sh.getRange(1, 1, 1, lastCol).getValues()[0]);
    ['ts', 'amt', 'date', 'h', 'm'].forEach(function (k) {
      if (C[k] < 0) throw new Error('Header "' + k + '" was not found in sheet "' + DASH_SHEET + '".');
    });

    const now = new Date();
    const op = d.type === 'Other' ? (d.operationType || 'Expenses') : d.type;
    const cb = d.cbAmount !== '' && d.cbAmount !== undefined && d.cbAmount !== null ? parseFloat(d.cbAmount) : '';
    const row = new Array(lastCol).fill('');
    const set = function (k, v) { if (C[k] >= 0) row[C[k]] = v; };

    set('ts', now);
    set('wd', d.withdrawn || '');
    set('o1', d.other1 || '');
    set('op', op);
    set('dep', d.deposited || '');
    set('o2', d.other2 || '');
    set('amt', amount);
    set('date', Utilities.parseDate(d.date, tz, 'yyyy-MM-dd'));
    set('h', h);
    set('m', m);
    set('sms', d.sms || '');
    set('ex', d.extras || '');
    set('det', d.details || '');
    set('fee', d.type === 'Other' && d.fees !== '' ? parseFloat(d.fees) || 0 : '');
    set('amt2', cb === '' || isNaN(cb) ? '' : cb);
    set('sms2', d.cbSms || '');

    const pIdx = dashColIndex_(DASH_CB_COL);
    if (cb !== '' && !isNaN(cb) && pIdx !== C.amt2 && pIdx < lastCol && !sh.getRange(2, pIdx + 1).getFormula()) row[pIdx] = cb;

    const target = sh.getLastRow() + 1;
    sh.getRange(target, 1, 1, lastCol).setValues([row]);
    SpreadsheetApp.flush();
    dashCacheClear_();
    return { row: target, ts: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'), op: op };
  } finally {
    lock.releaseLock();
  }
}

const DASH_API_ = {
  dashGetData: dashGetData,
  dashSaveCfg: dashSaveCfg,
  dashSaveCats: dashSaveCats,
  dashAddEntry: dashAddEntry
};

function dashJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return dashJson_({ ok: true, msg: 'MyBudjet API is running.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
    if (!token || req.token !== token) return dashJson_({ ok: false, error: 'unauthorized' });
    const fn = DASH_API_[req.fn];
    if (!fn) return dashJson_({ ok: false, error: 'Unknown function: ' + req.fn });
    return dashJson_({ ok: true, data: fn.apply(null, req.args || []) });
  } catch (err) {
    return dashJson_({ ok: false, error: String(err && err.message || err) });
  }
}