/**
 * Code.gs: backend API for the MyBudjet dashboard hosted on GitHub Pages (index.html).
 * Deploy as Web app: Execute as Me, Who has access: Anyone. Protected by username + password login (see the Login section at the bottom).
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

/* ---------------- Login (username + password) ---------------- */
// Users live in the script property AUTH_USERS as {username: {salt, hash}}. Passwords are never stored, only salted hashes.
// A successful login returns a signed session token (valid AUTH_SESSION_DAYS days) that the page sends with every request.
// Passwords expire after AUTH_PW_DAYS (90) days; changing one starts a new period and signs out the other devices.
// Create / reset users by running addUserNow() from the Apps Script editor (instructions inside it).
const AUTH_SESSION_DAYS = 30;
const AUTH_MAX_FAILS = 5;          // wrong passwords allowed per username ...
const AUTH_LOCK_SECONDS = 900;     // ... then that username is locked for 15 minutes
const AUTH_ROUNDS = 400;
const AUTH_PW_DAYS = 90;           // a password expires 90 days after it was set or changed
const AUTH_WARN_DAYS = 7;          // (the page reminds daily during the last 7 days; the count is invisible before that)

function authProps_() { return PropertiesService.getScriptProperties(); }
function authUsers_() { try { return JSON.parse(authProps_().getProperty('AUTH_USERS') || '{}'); } catch (e) { return {}; } }
function authHash_(salt, pw) {
  let h = salt + ':' + pw;
  for (let i = 0; i < AUTH_ROUNDS; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h));
  return h;
}
function authSecret_() {
  const p = authProps_(); let s = p.getProperty('AUTH_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid(); p.setProperty('AUTH_SECRET', s); }
  return s;
}
function authSign_(payload) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, authSecret_())); }
/** When the password was last set. Users created before expiry existed start their 90 days the first time they are seen. */
function authChanged_(users, u) {
  const rec = users[u]; if (!rec) return 0;
  if (!rec.changed) { rec.changed = Date.now(); authProps_().setProperty('AUTH_USERS', JSON.stringify(users)); }
  return rec.changed;
}
function authDaysLeft_(changed) { return Math.ceil((changed + AUTH_PW_DAYS * 86400000 - Date.now()) / 86400000); }

/** scope 'full' = normal use; scope 'pw' = short token that may only change the password (issued when it has expired). */
function authMakeSession_(username, changed, scope, ttlMs) {
  const payload = [username, Date.now() + ttlMs, changed, scope].join('|');
  return Utilities.base64EncodeWebSafe(payload) + '.' + authSign_(payload);
}
/** Returns {user, scope, daysLeft, expired} for a valid token, otherwise null. A password change or removing the user ends older sessions. */
function authResolve_(token) {
  try {
    const parts = String(token || '').split('.'); if (parts.length !== 2) return null;
    const payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    if (authSign_(payload) !== parts[1]) return null;
    const f = payload.split('|'); if (f.length !== 4) return null;
    const user = f[0], users = authUsers_();
    if (!(+f[1] > Date.now()) || !users[user]) return null;
    const changed = authChanged_(users, user);
    if (String(changed) !== f[2]) return null;
    const left = authDaysLeft_(changed);
    return { user: user, scope: f[3], daysLeft: left, expired: left <= 0 };
  } catch (e) { return null; }
}
function authLockCheck_(cache, key) {
  if ((parseInt(cache.get(key), 10) || 0) >= AUTH_MAX_FAILS) throw new Error('Too many wrong attempts. Try again in 15 minutes.');
}
function authLogin_(username, password) {
  const u = String(username || '').trim().toLowerCase(), pw = String(password || '');
  if (!u || !pw) throw new Error('Enter your username and password.');
  const cache = CacheService.getScriptCache(), key = 'authfail_' + u.slice(0, 60);
  authLockCheck_(cache, key);
  const users = authUsers_(), rec = users[u];
  const hash = authHash_(rec ? rec.salt : 'none', pw);       // always hash, so unknown users take the same time
  if (!rec || hash !== rec.hash) {
    cache.put(key, String((parseInt(cache.get(key), 10) || 0) + 1), AUTH_LOCK_SECONDS);
    throw new Error('Wrong username or password.');
  }
  cache.remove(key);
  const changed = authChanged_(users, u), left = authDaysLeft_(changed);
  if (left <= 0) return { session: authMakeSession_(u, changed, 'pw', 15 * 60000), username: u, mustChange: true };
  return { session: authMakeSession_(u, changed, 'full', AUTH_SESSION_DAYS * 86400000), username: u, days: AUTH_SESSION_DAYS, daysLeft: left };
}
/** The signed-in user changes their own password. Starts a fresh 90-day period and signs out every other device. */
function authChange_(user, oldPw, newPw) {
  oldPw = String(oldPw || ''); newPw = String(newPw || '');
  const cache = CacheService.getScriptCache(), key = 'authfail_' + user.slice(0, 60);
  authLockCheck_(cache, key);
  const users = authUsers_(), rec = users[user];
  if (!rec) throw new Error('Unknown user.');
  if (authHash_(rec.salt, oldPw) !== rec.hash) {
    cache.put(key, String((parseInt(cache.get(key), 10) || 0) + 1), AUTH_LOCK_SECONDS);
    throw new Error('The current password is wrong.');
  }
  if (newPw.length < 8) throw new Error('The new password needs at least 8 characters.');
  if (newPw === oldPw) throw new Error('The new password must be different from the current one.');
  if (newPw.toLowerCase() === user) throw new Error('The password must not be the same as the username.');
  rec.salt = Utilities.getUuid(); rec.hash = authHash_(rec.salt, newPw); rec.changed = Date.now();
  authProps_().setProperty('AUTH_USERS', JSON.stringify(users));
  cache.remove(key);
  return { session: authMakeSession_(user, rec.changed, 'full', AUTH_SESSION_DAYS * 86400000), username: user, days: AUTH_SESSION_DAYS, daysLeft: AUTH_PW_DAYS };
}
function authSetUser_(username, password) {
  const u = String(username || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(u)) throw new Error('Username: 3-30 characters, letters, digits . _ - only.');
  if (String(password || '').length < 8) throw new Error('Password must be at least 8 characters.');
  const users = authUsers_(), salt = Utilities.getUuid();
  users[u] = { salt: salt, hash: authHash_(salt, String(password)), changed: Date.now() };
  authProps_().setProperty('AUTH_USERS', JSON.stringify(users));
  CacheService.getScriptCache().remove('authfail_' + u.slice(0, 60));
  return u;
}

/** RUN FROM THE EDITOR: type the new username and password below, press Run, then erase the password again. Running it for an existing username resets that password. */
function addUserNow() {
  const USERNAME = '';   // example: 'ahmed'
  const PASSWORD = '';   // at least 8 characters. Erase it after running.
  Logger.log('Saved user: ' + authSetUser_(USERNAME, PASSWORD));
}
/** RUN FROM THE EDITOR: removes a user (their open sessions stop working). */
function removeUserNow() {
  const USERNAME = '';
  const u = String(USERNAME).trim().toLowerCase(), users = authUsers_();
  if (!users[u]) throw new Error('No such user: ' + u);
  delete users[u]; authProps_().setProperty('AUTH_USERS', JSON.stringify(users));
  Logger.log('Removed user: ' + u);
}
/** RUN FROM THE EDITOR: shows the usernames that can sign in. */
function listUsersNow() { Logger.log(Object.keys(authUsers_()).join(', ') || '(no users yet)'); }

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents), args = req.args || [];
    if (req.fn === 'login') return dashJson_({ ok: true, data: authLogin_(args[0], args[1]) });
    const s = authResolve_(req.session);
    if (!s) return dashJson_({ ok: false, error: 'unauthorized' });
    if (req.fn === 'changePassword') return dashJson_({ ok: true, data: authChange_(s.user, args[0], args[1]) });
    if (s.scope !== 'full') return dashJson_({ ok: false, error: 'unauthorized' });
    if (s.expired) return dashJson_({ ok: false, error: 'password_expired', pw: s.daysLeft });
    const fn = DASH_API_[req.fn];
    if (!fn) return dashJson_({ ok: false, error: 'Unknown function: ' + req.fn });
    return dashJson_({ ok: true, data: fn.apply(null, args), pw: s.daysLeft });   // pw = days left; the page only shows it during the last 7
  } catch (err) {
    return dashJson_({ ok: false, error: String(err && err.message || err) });
  }
}
