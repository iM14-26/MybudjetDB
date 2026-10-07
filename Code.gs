/**
 * Code.gs: backend API for the MyBudjet dashboard hosted on GitHub Pages (index.html).
 * Deploy as Web app: Execute as Me, Who has access: Anyone. Protected by the API_TOKEN script property.
 *
 * The dashboard now reads its own responses table: the "Entries" sheet.
 * - First run: "Entries" is created automatically. If the spreadsheet has a "Responses"
 *   sheet it is copied (headers + rows); otherwise a new table with the standard headers is made.
 * - The "Add entry" tab of the dashboard appends new rows to "Entries".
 * - Category / Sub-category are saved in "Entries" (columns created if missing).
 *
 * The Google Form keeps writing to "Responses" only. Rows added there after the
 * first copy do NOT appear in "Entries" (see the note in the chat).
 */

const DASH_SS_ID = '1AWH16EV239lPmdIAXTdz46d38uF0s9dJdvf-okukz8M';
const DASH_SHEET = 'Entries';      // the new responses table the dashboard is attached to
const DASH_SOURCE = 'Responses';   // copied once to create DASH_SHEET

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
  if (src) {                      // copy the old table, rows and all
    sh = src.copyTo(ss);
    sh.setName(DASH_SHEET);
    return sh;
  }
  // No "Responses" sheet in this spreadsheet: start a fresh table with the standard headers
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

function dashGetData() {
  const ss = SpreadsheetApp.openById(DASH_SS_ID);
  const sh = dashSheet_(ss);
  const tz = ss.getSpreadsheetTimeZone();
  const vals = sh.getDataRange().getValues();
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
      sms: dashStr_(g(C.sms)), url: dashStr_(g(C.url)),
      cat: dashStr_(g(C.cat)), sub: dashStr_(g(C.sub))
    });
  }
  const raw = PropertiesService.getScriptProperties().getProperty('dashCfg');
  return { rows: rows, cfg: raw ? JSON.parse(raw) : {}, missing: Object.keys(C).filter(function (k) { return C[k] < 0; }) };
}

function dashSaveCfg(json) {
  PropertiesService.getScriptProperties().setProperty('dashCfg', json);
  return true;
}

/** items: [{row: <sheet row number>, cat: '...', sub: '...'}] */
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
    return items.length;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Appends one entry (from the dashboard's "Add entry" tab) to the Entries sheet.
 * Values are placed by header name, so column order does not matter.
 * d: {type, withdrawn, other1, fees, operationType, deposited, other2, amount,
 *     date 'yyyy-MM-dd', h, m, sms, extras, details, cbAmount, cbSms}
 * Returns {row, ts, op}.
 */
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
    const op = d.type === 'Other' ? (d.operationType || 'Expenses') : d.type;   // Incomes / Refund / Expenses / Between Acounts
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

    const target = sh.getLastRow() + 1;
    sh.getRange(target, 1, 1, lastCol).setValues([row]);
    SpreadsheetApp.flush();
    return { row: target, ts: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'), op: op };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Web API (called by index.html) ---------------- */

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
