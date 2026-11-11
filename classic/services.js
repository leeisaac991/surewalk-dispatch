/*
 * Google Apps Script services, for running the dispatch panel's own .gs files in the browser (in a Web Worker).
 * Only what Code.gs / GroupMe.gs / Crews.gs / Places.gs / Excel.gs use:
 *   PropertiesService  -> an in-memory copy the host page keeps in this browser's localStorage
 *   CacheService       -> in memory, with expiry
 *   UrlFetchApp        -> synchronous XMLHttpRequest (GroupMe's API allows browser requests)
 *   Utilities          -> dates in America/Chicago, ids, base64, a fingerprint hash
 *   LockService / Session / Logger / Maps / SpreadsheetApp -> harmless stand-ins
 * The sheet itself comes from Excel.gs (XlSheet_), fed with the host's snapshot of the workbook.
 */
/* eslint-disable no-unused-vars */
const __SW = { script: {}, changed: false };

/*
 * Clock. Apps Script ran on Google's servers (exact time); here "now" comes from the dispatcher's computer, while
 * GroupMe's timestamps come from GroupMe's servers. If the computer runs behind, a crew's "empty" can look newer than
 * the assignment made right after it, and they'd stay "available". So "now" is corrected forward by how far behind
 * this computer is — learned from GroupMe messages (nothing is posted in the future), plus a small margin.
 */
const __clock = { offset: 0 };
(() => {
  const RD = Date;
  class D extends RD {
    constructor(...a) { if (a.length) super(...a); else super(RD.now() + __clock.offset); }
    static now() { return RD.now() + __clock.offset; }
  }
  self.Date = D;
  __clock.learn = text => { // GroupMe JSON: the newest created_at (seconds) tells us the server's time
    let max = 0;
    String(text).replace(/"created_at":\s*(\d{9,11})/g, (m, s) => { if (+s > max) max = +s; return m; });
    if (!max) return;
    const behind = max * 1000 - RD.now(); // > 0: this computer is behind GroupMe
    if (behind > 1000 && behind + 1500 > __clock.offset) __clock.offset = Math.min(behind + 1500, 6 * 3600e3);
  };
})();

function __store(obj) {
  const S = {
    getProperty: k => (Object.prototype.hasOwnProperty.call(obj(), k) ? obj()[k] : null),
    setProperty: (k, v) => { obj()[k] = String(v); __SW.changed = true; return S; },
    deleteProperty: k => { if (k in obj()) { delete obj()[k]; __SW.changed = true; } return S; },
    getProperties: () => Object.assign({}, obj()),
    setProperties: o => { Object.keys(o).forEach(k => { obj()[k] = String(o[k]); }); __SW.changed = true; return S; },
    getKeys: () => Object.keys(obj()),
  };
  return S;
}
const PropertiesService = {
  getScriptProperties: () => __store(() => __SW.script),
  getUserProperties: () => __store(() => __SW.script), // Excel.gs keys per-dispatcher settings by install id anyway
  getDocumentProperties: () => null,
};

const __cache = {};
const __cacheApi = {
  get: k => { const x = __cache[k]; if (!x) return null; if (x.exp < Date.now()) { delete __cache[k]; return null; } return x.v; },
  put: (k, v, ttl) => { __cache[k] = { v: String(v), exp: Date.now() + (ttl || 600) * 1000 }; },
  remove: k => { delete __cache[k]; },
  getAll: ks => { const o = {}; ks.forEach(k => { const v = __cacheApi.get(k); if (v != null) o[k] = v; }); return o; },
  putAll: (o, ttl) => { Object.keys(o).forEach(k => __cacheApi.put(k, o[k], ttl)); },
  removeAll: ks => { ks.forEach(k => { delete __cache[k]; }); },
};
const CacheService = { getScriptCache: () => __cacheApi, getUserCache: () => __cacheApi, getDocumentCache: () => __cacheApi };
const LockService = { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {}, hasLock: () => true }),
  getUserLock() { return this.getScriptLock(); } };
const Session = { getActiveUser: () => ({ getEmail: () => '' }), getEffectiveUser: () => ({ getEmail: () => '' }),
  getScriptTimeZone: () => 'America/Chicago' };
const Logger = { log: (...a) => console.log(...a) };
const Maps = { newGeocoder: () => ({ setBounds() { return this; }, geocode: () => ({ status: 'ZERO_RESULTS', results: [] }) }) };
const SpreadsheetApp = { openById: () => { throw new Error('This panel works on the open Excel workbook.'); } };
const HtmlService = { createHtmlOutputFromFile: () => ({ setTitle() { return this; }, addMetaTag() { return this; }, getContent: () => '' }) };
const ContentService = { MimeType: { JSON: 'json', TEXT: 'text' }, createTextOutput: s => ({ s, setMimeType() { return this; } }) };

const UrlFetchApp = {
  fetch(url, opts) {
    opts = opts || {};
    const x = new XMLHttpRequest();
    x.open(String(opts.method || 'get').toUpperCase(), url, false); // synchronous, like UrlFetchApp
    const h = opts.headers || {};
    Object.keys(h).forEach(k => x.setRequestHeader(k, h[k]));
    if (opts.payload != null && opts.contentType) x.setRequestHeader('Content-Type', opts.contentType);
    try { x.send(opts.payload == null ? null : opts.payload); } catch (e) { throw new Error('Network error: ' + (e.message || e)); }
    if (x.status === 0) throw new Error('Network error (no response)');
    if (!opts.muteHttpExceptions && x.status >= 400) throw new Error(`Request failed for ${url} returned code ${x.status}`);
    if (x.status < 300) __clock.learn(x.responseText); // only GroupMe answers carry "created_at"
    return { getResponseCode: () => x.status, getContentText: () => x.responseText, getHeaders: () => ({}) };
  },
};

const Utilities = (() => {
  const pad = (n, w) => String(n).padStart(w || 2, '0');
  const fmts = {};
  function parts(d, tz) {
    const f = fmts[tz] || (fmts[tz] = new Intl.DateTimeFormat('en-US', { timeZone: tz || 'America/Chicago', year: 'numeric', month: 'numeric', day: 'numeric',
      hour: 'numeric', minute: 'numeric', second: 'numeric', weekday: 'short', hourCycle: 'h23' }));
    const o = {};
    f.formatToParts(d).forEach(p => { o[p.type] = p.value; });
    return { y: +o.year, M: +o.month, d: +o.day, H: +o.hour % 24, m: +o.minute, s: +o.second, E: o.weekday };
  }
  // Java SimpleDateFormat patterns, as used by the .gs files
  function formatDate(date, tz, pattern) {
    const p = parts(date instanceof Date ? date : new Date(date), tz);
    return String(pattern).replace(/'[^']*'|yyyy|yy|MM|M|dd|d|HH|H|hh|h|mm|m|ss|s|a|EEE|E/g, t => {
      switch (t) {
        case 'yyyy': return String(p.y); case 'yy': return pad(p.y % 100); case 'MM': return pad(p.M); case 'M': return String(p.M);
        case 'dd': return pad(p.d); case 'd': return String(p.d); case 'HH': return pad(p.H); case 'H': return String(p.H);
        case 'hh': return pad(((p.H + 11) % 12) + 1); case 'h': return String(((p.H + 11) % 12) + 1);
        case 'mm': return pad(p.m); case 'm': return String(p.m); case 'ss': return pad(p.s); case 's': return String(p.s);
        case 'a': return p.H < 12 ? 'AM' : 'PM'; case 'EEE': case 'E': return p.E;
        default: return t.slice(1, -1);
      }
    });
  }
  // wall-clock time in tz -> Date
  function wall(y, M, d, H, m, s, tz) {
    const guess = Date.UTC(y, M - 1, d, H, m, s || 0);
    let ms = guess;
    for (let i = 0; i < 2; i++) { const p = parts(new Date(ms), tz); ms = guess - (Date.UTC(p.y, p.M - 1, p.d, p.H, p.m, p.s) - ms); }
    return new Date(ms);
  }
  function parseDate(str, tz) {
    const m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/.exec(String(str).trim()) ||
      null;
    if (m) return wall(+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0), tz);
    const u = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/.exec(String(str).trim());
    if (u) return wall(+(u[3].length === 2 ? '20' + u[3] : u[3]), +u[1], +u[2], +(u[4] || 0), +(u[5] || 0), +(u[6] || 0), tz);
    throw new Error('Unparseable date: ' + str);
  }
  const bytes = s => Array.from(new TextEncoder().encode(String(s)));
  function digest(_alg, input) { // a 16-byte fingerprint (only used to notice changes)
    const b = typeof input === 'string' ? bytes(input) : input;
    let h1 = 0x811c9dc5, h2 = 0x01000193 ^ 0x5bd1e995, h3 = 0x9e3779b9, h4 = 0x7f4a7c15;
    for (let i = 0; i < b.length; i++) {
      h1 = Math.imul(h1 ^ b[i], 0x01000193); h2 = Math.imul(h2 ^ b[i], 0x5bd1e995);
      h3 = Math.imul(h3 + b[i], 0x85ebca6b); h4 = Math.imul(h4 ^ (b[i] + i), 0xc2b2ae35);
    }
    const out = [];
    [h1, h2, h3, h4].forEach(h => { for (let k = 0; k < 4; k++) out.push(((h >>> (k * 8)) & 255) << 24 >> 24); });
    return out;
  }
  return {
    formatDate, parseDate, computeDigest: digest,
    DigestAlgorithm: { MD5: 'MD5', SHA_256: 'SHA_256' }, Charset: { UTF_8: 'UTF-8' },
    getUuid: () => (self.crypto && crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2)),
    base64Encode: b => btoa(typeof b === 'string' ? unescape(encodeURIComponent(b)) : String.fromCharCode.apply(null, b.map(x => x & 255))),
    base64Decode: s => Array.from(atob(s), c => c.charCodeAt(0)),
    sleep: ms => { const end = Date.now() + Math.min(ms, 2000); while (Date.now() < end) { /* like Utilities.sleep */ } },
    newBlob: (data, type) => ({ getDataAsString: () => (typeof data === 'string' ? data : String.fromCharCode.apply(null, data)), type }),
  };
})();
