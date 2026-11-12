

const __SW = { script: {}, changed: false };

const __clock = { offset: 0 };
(() => {
  const RD = Date;
  class D extends RD {
    constructor(...a) { if (a.length) super(...a); else super(RD.now() + __clock.offset); }
    static now() { return RD.now() + __clock.offset; }
  }
  self.Date = D;
  __clock.learn = text => {
    let max = 0;
    String(text).replace(/"created_at":\s*(\d{9,11})/g, (m, s) => { if (+s > max) max = +s; return m; });
    if (!max) return;
    const behind = max * 1000 - RD.now();
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
  getUserProperties: () => __store(() => __SW.script),
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
    x.open(String(opts.method || 'get').toUpperCase(), url, false);
    const h = opts.headers || {};
    Object.keys(h).forEach(k => x.setRequestHeader(k, h[k]));
    if (opts.payload != null && opts.contentType) x.setRequestHeader('Content-Type', opts.contentType);
    try { x.send(opts.payload == null ? null : opts.payload); } catch (e) { throw new Error('Network error: ' + (e.message || e)); }
    if (x.status === 0) throw new Error('Network error (no response)');
    if (!opts.muteHttpExceptions && x.status >= 400) throw new Error(`Request failed for ${url} returned code ${x.status}`);
    if (x.status < 300) __clock.learn(x.responseText);
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
  function digest(_alg, input) {
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
    sleep: ms => { const end = Date.now() + Math.min(ms, 2000); while (Date.now() < end) {   } },
    newBlob: (data, type) => ({ getDataAsString: () => (typeof data === 'string' ? data : String.fromCharCode.apply(null, data)), type }),
  };
})();

;

const BUILD_ID = '2026-11-10 00:02:31';

;

const FILE_VERSION_PLACES = '2026-11-07a';

const PLACES = {

  'Jester':               ['jester', 'jester center', 'jester circle', 'jester east', 'jester e', 'east jester', 'jcl',
                          'j2', 'j2 dining', 'j2 dining hall', 'jester 2', 'jester j2', 'j1', 'j1 dining', 'j1 dining hall',
                          'jeast', 'j east', 'j-east'],
  'Jester West':          ['jwest', 'j west', 'j-west', 'jester west'],
  'Kinsolving':           ['kins', 'kinsolving', 'kiins', 'kinsoving', 'kinsolvinfg'],
  'San Jacinto':          ['san jac', 'sanjac', 'san jacinto', 'sacjac', 'san jac dorms', 'san jacinto dorms', 'sjc', 'cypress cafe', 'cypress', 'cypress bend cafe'],
  'Moore-Hill':           ['moore hill', 'moorehill', 'morehill', 'more hill', 'moore'],
  'Somerset Condos':      ['somerset', 'somerset condos', 'somerset condo', 'somerset condominiums', 'somerset condominium'],

  'Robbins Place':        ['robbins pl', 'robbins place', 'robinson pl', 'robinson pi', 'robbins pi', 'robinson place',
                           '1903 robbins pl', '1903 robbins place', '1903 robinson pl', '1903 robinson pi', '903 robinson pi', '903 robinson pl', '903 robbins pl'],
  'Stoneleigh Condominiums': ['stoneleigh', 'stoneleigh condominiums', 'stoneleigh condos', 'stoneleigh condo',
                           'stoneleigh condominiums on 2409 leon st', 'stoneleigh condominiums 2409 leon st'],
  'Phi Gamma Delta':      ['phi gamma delta', 'phi gam', 'phigam', 'fiji', 'fiji house', 'phi gamma delta house', 'phi gamma', 'phi gam house'],
  'Avon at 22nd':         ['avon', 'avon at 22nd', 'avon at 22nd st', 'avon at 22nd street', 'avon at 22', 'avon@22', 'avon@22nd',
                           'avon 22', 'avon 22nd', 'avon 22nd st', 'avon on 22nd', 'avon on 22nd st', 'avon apartments'],
  'Littlefield':          ['littlefield', 'little field', 'littlefield dorms'],
  'Duren':                ['duren', 'durin', 'duran', 'durren', 'duren hall', 'durin hall', 'duren dorm', 'duren dorms', 'durin dorms'],
  'Carothers':            ['carothers', 'carother', 'carothers dormitory'],
  'Roberts':              ['roberts', 'robert'],
  'Blanton':              ['blanton', 'blanton residend'],
  'Prather':              ['prather'],
  'Creekside':            ['creekside', 'creek side'],
  'Brackenridge':         ['brackenridge', 'breckenridge'],
  'Andrews':              ['andrews'],

  'Target (Dobie)':       ['target', 'target (dobie)', 'target dobie', 'dobie target', 'target at dobie', 'the target', 'target on guad', 'target on guadalupe', 'ut target', 'target store'],
  'Dobie Twenty21':       ['dobie', 'dobie 21', 'dobie21', 'dobie twenty21', 'dobie twenty 21', 'dobie twenty one', 'dobie center'],
  'Castilian':            ['castilian', 'castillian', 'castillion'],
  '2400 Nueces':          ['2400 nueces', '2400 neuces', '2400 nucess', '2400 nuecues', '2400 nueces complex'],
  '26 West':              ['26 west', '26west'],
  'Legacy on Rio':        ['legacy on rio', 'legacy', 'legacy rio'],
  'The Rambler':          ['rambler'],
  'Skyloft':              ['skyloft', 'sky loft'],
  'The Standard':         ['standard'],
  'The Nine at West Campus': ['nine at west campus', 'nine at west', 'nine west campus', '9 west campus', 'the nine'],
  'The Nine at Rio':      ['nine at rio'],
  'Villas on Rio':        ['villas on rio'],
  'Villas on Guadalupe':  ['villas on guad', 'villas on guadalupe', 'villas on gaud'],
  'Villas on Nueces':     ['villas on nueces', 'villas om nueces'],
  'Villas on 26th':       ['villas on 26th', 'villas on 26', 'villas on w 26th', 'villas on west 26th', 'villas 26th'],
  'Venue on Guadalupe':   ['venue on guad', 'venue on guadalupe'],
  'Hilltop':              ['hilltop', 'hill top'],
  'Lark':                 ['lark'],
  'Rise at West Campus':  ['rise', 'rise at west campus', 'rise at west', 'the rise'],
  'The Sinclair':         ['sinclair', 'sinclaire', 'sinclair aps', 'sinclair 2710 nueces'],
  'Axis West':            ['axis west', 'axis west campus', 'axis west campus 2505 longview'],
  '24 Longview':          ['24 longview', '24longview'],

  'Texan West Campus':    ['texan west campus', 'texan west', 'texan wc', 'the texan', 'texan apartments', 'the texan west campus',
                           'texan 26', 'texan26', '2616 salado', 'texan 26 2616 salado', 'texan 2616 salado'],
  'Crest at Pearl':       ['crest at pearl'],
  'Union on 24th':        ['union on 24th', 'union on 24', 'union 24th', 'uniton 24th', 'union of 24'],
  'Union on San Antonio': ['union on san antonio', 'union san antonio'],
  'The Block on 23rd':    ['block on 23rd', 'block on 23', 'block 23rd'],
  'The Block on 25th East': ['block on 25th east', 'block on 25 e', 'block on 25e'],
  'The Block on 25th West': ['block on 25th west', 'block on 25th w', 'block 25th w'],
  'The Block on 28th':    ['block on 28th'],
  'The Block on Leon':    ['block on leon', 'block at leon', 'bock on leon'],
  'The Block on Pearl North': ['block on pearl north', 'block on pearl n', 'block pearl north', 'north block on pearl'],
  'The Block on Pearl South': ['block on pearl south', 'block on pearl s', 'block pearl south', 'south block on pearl'],
  'The Block on Pearl':   ['block on pearl', 'block at pearl', 'block pearl'],
  'Signature 1909':       ['signature 1909'],
  'Inspire on 22nd':      ['inspire on 22nd', 'inspire on 2nnd'],
  '21 Pearl':             ['21 pearl'],
  '21 Rio':               ['21 rio'],
  'Yugo Rio':             ['yugo rio', 'yugo austin rio', 'yugo on rio'],
  'Yugo Waterloo':        ['yugo waterloo', 'yugo austin waterloo', 'waterloo yugo'],
  'Yugo Corner':          ['yugo corner', 'yugo austin corner', 'yugo the corner'],
  'Yugo Nueces':          ['yugo nueces', 'yugo austin nueces'],
  'Quarters Nueces':      ['quarters nueces', 'quarters on nueces', 'quarters at nueces', 'quarter nueces', 'quarters nueces house'],
  'Regents West at 24th': ['regents west at 24th', 'regents west 24th', 'regents west on 24th', 'regents on 24th', 'regents at 24th'],
  'Regents West at 26th': ['regents west at 26th', 'regents west on 26th', 'regents west on 26', 'regents at 26th',
                           'regents w 26th', 'regents on 26th', 'regents west in 26', 'regents west 26th'],
  'GrandMarc':            ['grandmarc', 'grand marc', 'grandmark'],
  'Noble 2500':           ['noble 2500'],
  'Sparq on Rio':         ['sparq on rio'],
  'Tradition on Rio':     ['tradition on rio', 'traditions on rio'],
  'Tanglewood North':     ['tanglewood north'],
  'Moontower':            ['moontower', 'moon tower'],
  '1883 Cameron':         ['1883 cameron', '1883 at cameron', '1883 cameron house'],
  '1883 Montgomery':      ['1883 montgomery', '1883 at montgomery'],
  'Twenty Two 15':        ['twenty two 15', 'twentytwo 15', '2215'],
  'Duval Villas':         ['duval villas', 'duval villa'],
  'Westridge':            ['westridge', 'west ridge', 'westridge condos'],
  'Miller Square & Spanish Trails': ['miller square and spanish trails', 'miller square and spanish trail'],

  'PCL':                  ['pcl', 'pcl library', 'perry castaneda library', 'perry castaneda',
                           'mccombs', 'mcombs', 'mccomb', 'mccombs school of business', 'mccombs business school', 'mccombs building',
                           'mccombs speedway', 'mccombs building pcl', 'mccombs pcl', 'pcl mccombs', 'gsb', 'cba',
                           'graduate school of business', 'graduate school of business mccombs next to dobie',
                           'mccombs school of business 2110 speedway austin tx 78705'],
  'Gregory Gym':          ['greg', 'greg gym', 'gregory gym', 'gregory'],
  'Rec Center':           ['rec', 'rec center', 'recsports', 'rec sports', 'ut rec center'],
  'Texas Union':          ['union', 'texas union'],
  'Butler School of Music': ['butler', 'butler school of music', 'butler school music', 'butler music', 'bsom', 'b.s.o.m', 'butler som', 'school of music'],
  'Moody CMA':            ['cma', 'moody cma'],

  'Moody (Center or DMC?)': ['moody'],
  'Winship':              ['winship', 'winship drama'],
  'Art Building':         ['art', 'art building'],
  'Welch':                ['welch'],
  'Burdine':              ['burdine'],
  'Sutton':               ['sutton'],
  'Belmont':              ['belmont', 'bellmont'],
  'Littlefield Fountain': ['littlefield fountain', 'little field fountain'],
  'Littlefield House':    ['littlefield house', 'little field house'],
  'UT Administration':    ['uta', 'ut administration'],
  'DKR Stadium':          ['dkr', 'dkr stadium'],
  'The Bubble':           ['bubble'],

  'Brazos Garage':        ['brazos', 'brazos garage', 'brg'],
  'Manor Garage':         ['manor', 'manor garage', 'manor gargae'],
  'San Jacinto Garage':   ['san jac garage', 'san jacinto garage'],
  'Sunchase Condos':      ['sunchase', 'sunchase condos', 'sunchase condo', 'sunchase condominiums', '3115 helms', '3115 helms st'],
  '21 Rio':               ['21 rio', '21 rio apartments', 'twenty one rio'],
  'Orange Tree Condos':   ['orange tree', 'orange tree condos', 'orange tree condominiums'],
  'St Thomas Condos':     ['st thomas', 'st thomas condo', 'st thomas condos', 'saint thomas condos'],
  'Overlook Condos':      ['overlook', 'overlook condos', 'overlook condos poplar st'],
  'Lenox Condos':         ['lenox', 'lenox condos'],
  'Lonestar Lofts':       ['lonestar lofts', 'lone star lofts'],
  'Miller Square':        ['miller square', 'miller square and spanish trails', 'miller square & spanish trails', 'spanish trails'],
  'Whitaker Fields':      ['whitaker', 'whitaker fields', 'whitaker field', 'whittaker', 'whittaker fields', 'whitaker intramural fields',
                          'witaker', 'witaker fields', 'whitaker feilds', 'wright witaker feilds', 'witaker feilds',
                          'alan wright fields', 'alan wright field', 'alan wright', 'wright fields', 'alan wright intramural fields'],
  'Scottish Rite Dormitory': ['srd', 'scottish rite', 'scottish rite dorm', 'scottish rite dormitory', 'scottish rite residence hall', 's r d'],
  'Guadalupe Garage':     ['guadalupe garage', 'gug', 'guad garage', 'guadalupe parking garage', 'gug garage'],
  'Health Center Garage': ['health center garage'],
  'East Campus Garage':   ['east campus garage'],
};

const PLACE_FILLER = new Set(['the', 'hall', 'dorm', 'dormitory', 'residence', 'apartments', 'apartment', 'apts',
  'apt', 'building', 'bldg', 'austin', 'tx', 'texas', 'st', 'street', 'unitedstates', 'usa', 'united', 'states']);

function placeNorm_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[’'`]/g, '').replace(/&/g, ' and ').replace(/\b787\d\d\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ').split(' ')
    .filter(w => w && !PLACE_FILLER.has(w)).join('');
}

let PLACE_INDEX_ = null;
function placeIndex_() {
  if (!PLACE_INDEX_) {
    PLACE_INDEX_ = {};
    Object.keys(PLACES).forEach(name => {
      [name].concat(PLACES[name]).forEach(v => { PLACE_INDEX_[placeNorm_(v)] = name; });
    });
  }
  return PLACE_INDEX_;
}

const dropNote_ = s => String(s || '').replace(/\([^)]*\)?/g, ' ').replace(/\s+/g, ' ').trim();

function canonPlace_(s) {
  const idx = placeIndex_();
  return idx[placeNorm_(s)] || idx[placeNorm_(dropNote_(s))] || fuzzyPlace_(placeNorm_(dropNote_(s) || s)) || null;
}

let FUZZY_CACHE_ = {};
function fuzzyPlace_(n) {
  if (!n || /\d/.test(n) || n.length < 6) return null;
  if (n in FUZZY_CACHE_) return FUZZY_CACHE_[n];
  const max = n.length >= 10 ? 2 : 1, idx = placeIndex_();
  let best = null, bestD = max + 1, tie = false;
  Object.keys(idx).forEach(k => {
    if (k.length < 6 || k[0] !== n[0] || /\d/.test(k) || Math.abs(k.length - n.length) > max) return;
    const d = editDist_(n, k, max);
    if (d < bestD) { bestD = d; best = idx[k]; tie = false; } else if (d === bestD && idx[k] !== best) tie = true;
  });
  return (FUZZY_CACHE_[n] = best && !tie ? best : null);
}

function editDist_(a, b, max) {
  let prev2 = null, prev = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    let rowMin = i;
    for (let j = 1; j <= b.length; j++) {
      let v = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      if (prev2 && i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) v = Math.min(v, prev2[j - 2] + 1);
      cur.push(v); if (v < rowMin) rowMin = v;
    }
    if (rowMin > max) return max + 1;
    prev2 = prev; prev = cur;
  }
  return prev[b.length];
}

function placeKey_(s) {
  const c = canonPlace_(s);
  return c ? c.toLowerCase() : placeNorm_(dropNote_(s) || s);
}

const ZONE_OF = {

  "2400 nueces": 'On', "andrews": 'On', "blanton": 'On', "brackenridge": 'On', "brazos garage": 'On',
  "butler school of music": 'On', "carothers": 'On', "cpe": 'On', "creekside": 'On', "dkr stadium": 'On',
  "dmc": 'On', "dobie twenty21": 'On', "target (dobie)": 'On', "duren": 'On', "ecg": 'On', "ecga": 'On', "eer": 'On', "etc": 'On',
  "gregory gym": 'On', "hcg": 'On', "health center garage": 'On', "jester": 'On',
  "jester west": 'On', "kinsolving": 'On', "littlefield": 'On', "lot118": 'On', "lot37": 'On', "manor garage": 'On',
  "mccombs": 'On', "moody cma": 'On', "moody (center or dmc?)": 'On', "moore-hill": 'On', "pcl": 'On', "prather": 'On', "rec center": 'On',
  "roberts": 'On', "rowling": 'On', "sag": 'On', "san jacinto": 'On', "san jacinto garage": 'On', "guadalupe garage": 'On', "scottish rite dormitory": 'On', "sunchase condos": 'Off', "whitaker fields": 'Off', "21 rio": 'West', "orange tree condos": 'West', "st thomas condos": 'West', "overlook condos": 'West', "lenox condos": 'West',
  "lonestar lofts": 'West', "miller square": 'Off', "2001comal": 'On', "2105comal": 'On', "smc": 'On',
  "texas union": 'On', "trg": 'On', "utc": 'On',

  "somerset condos": 'West',
  "avon at 22nd": 'West',
  "phi gamma delta": 'West',
  "stoneleigh condominiums": 'West',
  "robbins place": 'West',
  "1010w23rd": 'West', "1101w22nd": 'West', "1883 cameron": 'West', "1883 montgomery": 'West',
  "1883atcameron2707riogrande": 'West', "21 pearl": 'West', "2200pearl": 'West', "2204sangabriel": 'West',
  "2208pearl": 'West', "2212sangabriel": 'West', "2216sangabriel": 'West', "2222pearl": 'West', "2300nueces": 'West',
  "24 longview": 'West', "2414longview": 'West', "2502leon": 'West', "2510leon": 'West', "2510sangabriel": 'West',
  "26 west": 'West', "2606riogrande": 'West', "2616saladounitedstates": 'West', "2706salado": 'West',
  "2707riogrande": 'West', "2707riogrande1883atcameron": 'West', "2710nueces": 'West', "2801riogrande": 'West',
  "2810salado": 'West', "2815riogrande": 'West', "2819riogrande": 'West', "702w25th": 'West',
  "708grahamplace": 'West', "709w22nd": 'West', "714w22nd": 'West', "715w2212": 'West', "910w26th": 'West',
  "910west26th": 'West', "9atrio": 'West', "axis": 'West', "axis west": 'West', "blockon25th": 'West',
  "blockonrio": 'West', "callaway": 'West', "caminoflats": 'West', "castilian": 'West', "centennialcondos": 'West',
  "crest at pearl": 'West', "croixcondos": 'West', "grandmarc": 'West', "hilltop": 'West', "icon": 'West',
  "inspire": 'West', "inspire on 22nd": 'West', "ion": 'West', "lark": 'West', "legacy on rio": 'West',
  "lenoxcondos915w23": 'West', "lonestarlofts": 'West', "mark": 'West', "moontower": 'West', "noble 2500": 'West',
  "nueceshouse": 'West', "overlookcondominiums": 'West', "parkplace": 'West', "pearlcoop": 'West',
  "pointnorth": 'West', "quarters nueces": 'West', "quartersgrayson": 'West', "quartersgraysonhouse": 'West',
  "quartersongrayson": 'West', "quarterssterling": 'West', "regents west at 24th": 'West',
  "regents west at 26th": 'West', "riograndecondominiums": 'West', "rise at west campus": 'West',
  "sangabrielsquare": 'West', "sanremo": 'West', "signature 1909": 'West', "skyloft": 'West', "sparq on rio": 'West',
  "standardat": 'West', "stoneleighcondominiumson2409leon": 'West', "taoscoop": 'West', "texan west campus": 'West',
  "texanpearl": 'West', "the block on 23rd": 'West', "the block on 25th east": 'West',
  "the block on 25th west": 'West', "the block on 28th": 'West', "the block on leon": 'West', "the block on pearl": 'West', "the block on pearl north": 'West', "the block on pearl south": 'West',
  "the nine at rio": 'West', "the nine at west campus": 'West', "the rambler": 'West', "the sinclair": 'West',
  "the standard": 'West', "torre": 'West', "tradition on rio": 'West', "twenty two 15": 'West',
  "union on 24th": 'West', "union on san antonio": 'West', "venue on guadalupe": 'West',
  "villas on guadalupe": 'West', "villas on nueces": 'West', "villas on rio": 'West', "villas on 26th": 'West', "villason24th": 'West',
  "vintagewestcampus": 'West', "waterloo": 'West', "westridge": 'West', "yugo corner": 'West', "yugo nueces": 'West',
  "yugo rio": 'West', "yugo waterloo": 'West',

  "1021e45th": 'Off', "106e30th": 'Off', "201e34th": 'Off', "2900cole": 'Off', "2910medicalarts": 'Off',
  "306e30th": 'Off', "306e30th2": 'Off', "3106duval": 'Off', "3111tomgreen": 'Off', "3115helms": 'Off',
  "3121speedway": 'Off', "31stcondominiums": 'Off', "3301speedway": 'Off', "3409speedway": 'Off',
  "3506speedway": 'Off', "3606grooms": 'Off', "403w38th": 'Off', "415w39th": 'Off', "4200avenuea": 'Off',
  "4305duval": 'Off', "4310avenueb": 'Off', "4406bennettave": 'Off', "4504speedway": 'Off', "4505speedway": 'Off',
  "4527nlamarblvd": 'Off', "609e49th": 'Off', "810w32nd": 'Off', "918e37th": 'Off', "duval villas": 'Off',
  "hydeparkcourtnorthcampus": 'Off', "kingston": 'Off', "oakpark": 'Off', "sucasa": 'Off', "tanglewood": 'Off',
  "tanglewood north": 'Off', "windtreecondominiums": 'Off',
};

function zoneOf_(dropoff) {
  if (!String(dropoff || '').trim()) return '';
  const bare = dropNote_(dropoff);
  return ZONE_OF[placeKey_(dropoff)] || partZone_(dropoff) || (/\b78712\b/.test(dropoff) ? 'On' : '') ||
    zoneFromAddress_(dropoff) || (bare && zoneFromAddress_(bare)) ||
    embeddedAddressZone_(dropoff) || nameZone_(dropoff) || (mentionsCampusPlace_(dropoff) ? 'On' : 'Off');
}

function partZone_(text) {
  const parts = String(text).split(/\s*[,;()\/]\s*|\s+-\s+|\s+or\s+/i).map(x => x.trim()).filter(Boolean);
  if (parts.length < 2) return '';
  for (const p of parts) { const c = canonPlace_(p); if (c && ZONE_OF[c.toLowerCase()]) return ZONE_OF[c.toLowerCase()]; }
  return '';
}

function embeddedAddressZone_(text) {
  const re = /\b\d{2,5}\s+(?:[nsew]\.?\s+|west\s+|east\s+)?[a-z0-9]+(?:\s+[a-z]+)?/gi;
  let m;
  while ((m = re.exec(String(text)))) { const z = zoneFromAddress_(m[0]); if (z) return z; }
  return '';
}

const GREEK_ = /\b(sorority|fraternity|alpha|beta|gamma|delta|epsilon|zeta|theta|iota|kappa|lambda|sigma|tau|upsilon|phi|chi|psi|omega|omicron)\b/i;
function nameZone_(text) {
  const t = String(text);
  if (/west\s*campus/i.test(t)) return 'West';
  if (/^\s*(ut\s+)?lot\s*#?\s*\d{1,3}\b/i.test(t)) return 'On';
  if (GREEK_.test(t) || /\bco-?\s?op\b/i.test(t)) return 'West';
  return '';
}

const WEST_CAMPUS_STREETS = ['rio grande', 'nueces', 'san gabriel', 'pearl', 'leon', 'salado', 'longview', 'san antonio',
  'montgomery', 'graham', 'san pedro', 'robbins', 'wichita'];
function zoneFromAddress_(dropoff) {
  const t = String(dropoff || '').toLowerCase().replace(/[.,#]/g, ' ').replace(/\b(austin|tx|texas|787\d\d)\b/g, ' ')
    .replace(/\bdean\s*keaton\b/g, '26th').replace(/\b(mlk|martin luther king( jr)?( blvd)?)\b/g, '19th');

  if (/^\s*\d{2,5}\s+(robbins|robinson)\s+p[li]/.test(t)) return 'West';
  let m = /^\s*(\d{2,5})\s+(?:(w|e|west|east)\s+)?(\d{1,2})(?:st|nd|rd|th)?\b/.exec(t);
  if (m) {
    const n = Number(m[3]), num = Number(m[1]);
    if (n >= 29) return 'Off';
    if (n >= 19 && /^w/.test(m[2] || '')) return 'West';
    if (n >= 19 && /^e/.test(m[2] || '') && num >= 700) return 'Off';
    return '';
  }
  m = /^\s*(\d{3,5})\s+(?:[nsew]\s+)?([a-z][a-z ]*)/.exec(t);
  if (m) {
    const num = Number(m[1]), street = m[2].trim();
    if (num >= 2900) return 'Off';
    if (num >= 1800 && WEST_CAMPUS_STREETS.some(w => street.startsWith(w))) return 'West';
  }
  return '';
}

const GEO_CENTER = [30.2849, -97.7341];
const GEO_MAX_M = 8000;
function geoDist_(a, b) {
  const dy = (a[0] - b[0]) * 110540, dx = (a[1] - b[1]) * 111320 * Math.cos(a[0] * Math.PI / 180);
  return Math.sqrt(dx * dx + dy * dy);
}

function geoPlaces_(addresses, maxNew) {
  const cache = bigGet_('GEO_CACHE') || {}, out = {};
  let added = 0;
  [...new Set(addresses.filter(Boolean))].forEach(a => {
    const k = placeKey_(a);
    if (!k) return;
    if (k in cache) { out[k] = cache[k]; return; }
    if (added >= maxNew) return;
    added++;
    let pos = 0;
    try {
      const r = Maps.newGeocoder().setBounds(30.24, -97.78, 30.33, -97.70).geocode((canonPlace_(a) || a) + ', Austin, TX');
      const loc = r && r.status === 'OK' && r.results && r.results[0] && r.results[0].geometry.location;
      if (loc && geoDist_([loc.lat, loc.lng], GEO_CENTER) < GEO_MAX_M) pos = [+loc.lat.toFixed(5), +loc.lng.toFixed(5)];
    } catch (e) { added--; return; }
    cache[k] = pos; out[k] = pos;
  });
  if (added) { try { bigSet_('GEO_CACHE', cache); } catch (e) {   } }
  return out;
}

const CHAIN_M = 250;
function samePlace_(a, b) {
  if (placeKey_(a) === placeKey_(b)) return true;
  const g = geoPlaces_([a, b], 0), pa = g[placeKey_(a)], pb = g[placeKey_(b)];
  return !!(pa && pb && geoDist_(pa, pb) <= CHAIN_M);
}

const CART_NO_GO = ['ECG', 'ECGA', 'East Campus Garage', 'SAG', 'San Antonio Garage', 'TRG', 'Trinity Garage', 'MNG', 'Manor Garage',
  'BRG', 'Brazos Garage', 'HCG', 'Health Center Garage', 'CCG', 'Conference Center Garage', 'GUG', 'Guadalupe Garage', 'NUG', 'Nueces Garage'];
let CART_NO_GO_KEYS_ = null;
function cartNoGo_(place) {
  if (!CART_NO_GO_KEYS_) CART_NO_GO_KEYS_ = new Set(CART_NO_GO.map(placeKey_));
  return CART_NO_GO_KEYS_.has(placeKey_(place));
}

const DROPOFF_ONLY = ['Whitaker Fields', 'Moody Center'];

const PLACE_REWRITES = [
  { re: /^\s*(?:the\s+)?(?:ut\s+)?target(?:\s+(?:store|on\s+guad(?:alupe)?|at\s+guad(?:alupe)?|guad(?:alupe)?|dobie|at\s+dobie))?\s*[.!]*\s*$/i, to: 'Target (Dobie)' },
  { re: /^\s*(?:the\s+)?(?:alan\s+)?wright\s+(?:intramural\s+)?fields?\s*[.!]*\s*$|^\s*alan\s+wright\s*[.!]*\s*$/i, to: 'Whitaker Fields' },
  { re: /^\s*(?:durin|duran|durren|dureen)(?:\s+(?:hall|dorms?|residence hall))?\s*[.!]*\s*$/i, to: 'Duren' },
];
function placeRewrite_(text) {
  const t = String(text == null ? '' : text).trim(), w = PLACE_REWRITES.find(x => x.re.test(t));
  return w && t !== w.to ? w.to : null;
}
function dropoffOnly_(text) {
  const keys = new Set(DROPOFF_ONLY.map(p => placeKey_(p)));
  const parts = [text].concat(String(text).split(/\s*[,;()]\s*|\s+-\s+/));
  if (parts.some(x => x && keys.has(placeKey_(x)))) return true;
  const t = ' ' + String(text).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim() + ' ';
  return DROPOFF_ONLY.some(p => [p].concat(PLACES[p] || []).some(a => { const w = String(a).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); return w.length >= 5 && t.includes(' ' + w + ' '); }));
}
function offCampusPickup_(pickup) {
  if (!String(pickup || '').trim()) return false;
  if (dropoffOnly_(pickup)) return true;
  return zoneOf_(pickup) !== 'On';
}

let CAMPUS_PHRASES_ = null;
const SHORT_OK_ = new Set(['bsom', 'pcl', 'sac', 'sjc', 'j1', 'j2', 'ua9', 'fac', 'gdc', 'eer', 'wcp', 'sjh', 'rlm', 'pma', 'ecj', 'glt', 'kin', 'utc', 'mai', 'unb', 'cma', 'hrc', 'lbj',
  'brg', 'ccg', 'ecg', 'gug', 'hcg', 'mag', 'rhg', 'sag', 'sjg', 'swg', 'trg', 'tsg', 'dhg', 'dkr', 'ksc']);
function mentionsCampusPlace_(text) {
  if (!CAMPUS_PHRASES_) {
    CAMPUS_PHRASES_ = [];
    Object.keys(PLACES).filter(k => ZONE_OF[k.toLowerCase()] === 'On' && !DROPOFF_ONLY.includes(k)).forEach(k =>
      [k].concat(PLACES[k]).forEach(v => {
        const w = String(v).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
        if (w && !/^\d+$/.test(w) && (w.length >= 4 || SHORT_OK_.has(w))) CAMPUS_PHRASES_.push(' ' + w + ' ');
      }));
  }
  const t = ' ' + String(text).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, ' ').trim() + ' ';
  return CAMPUS_PHRASES_.some(w => t.includes(w));
}

const CAMPUS_MAP = [
  ['ACA', 'Academic Annex'], ['ADH', 'Almetris Duren Hall', 'Duren'], ['AFP', 'Athletics Field Pavilion'],
  ['AHG', 'Anna Hiss Gymnasium', '', 'anna hiss'], ['ANB', 'Arno Nowotny Building'], ['', 'Andrews Dormitory', 'Andrews'],
  ['ARC', 'Animal Resources Center'], ['ART', 'Art Building'], ['ASE', 'Aerospace Engineering'],
  ['ATT', 'AT&T Executive Education and Conference Center', '', 'att center', 'at&t center', 'at&t conference center'],
  ['BAT', 'Batts Hall', '', 'batts'], ['BEL', 'L. Theo Bellmont Hall', 'Belmont', 'bellmont'], ['BEN', 'Benedict Hall', '', 'benedict'],
  ['BHD', 'Brackenridge Hall Dormitory', 'Brackenridge'], ['BIO', 'Biological Laboratories', '', 'bio labs'], ['BLD', 'Blanton Dormitory', 'Blanton'],
  ['BMA', 'Blanton Museum of Art', '', 'blanton museum'], ['BMC', 'Dealey Center for New Media', '', 'belo', 'belo center', 'belo center for new media'],
  ['BME', 'Biomedical Engineering Building'], ['BOT', 'Biological Greenhouse'], ['BRB', 'Bernard and Audre Rapoport Building', '', 'rapoport'],
  ['BTL', 'Battle Hall'], ['BUR', 'Burdine Hall', '', 'burdine'], ['BWY', 'Bridgeway Building', '', '2616 wichita'],
  ['CAL', 'Calhoun Hall', '', 'calhoun'], ['CBA', 'McCombs School of Business'], ['CCF', 'Caven Lacrosse and Sports Center'],
  ['CCJ', 'John B. Connally Center for Justice'], ['CDA', 'Comal Child Development Center Annex'], ['CDL', 'Collections Deposit Library'],
  ['CEE', 'Continuing Engineering Education'], ['CLK', 'Caven Clark Field Support Building'],
  ['CMA,CMB', 'Jesse H. Jones Communication Center'], ['CML', 'Child Development Center'], ['COM', 'Computation Center'],
  ['CPE', 'Chemical and Petroleum Engineering'], ['CRB', 'Central Receiving Building'], ['CRD', 'Carothers Dormitory', 'Carothers'],
  ['CRH', 'Creekside Residence Hall', 'Creekside'], ['CS3', 'Chilling Station No. 3'], ['CS4', 'Chilling Station No. 4'],
  ['CS5', 'Chilling Station No. 5'], ['CS6', 'Chilling Station No. 6'],
  ['DCP', 'Denton A. Cooley Pavilion', '', 'cooley pavilion'], ['DEV', 'Development Office Building'],
  ['DFA', 'E. William Doty Fine Arts Building', '', 'fine arts', 'doty fine arts'], ['DFF', 'UFCU Disch-Falk Field', '', 'disch falk', 'disch-falk', 'falk disch', 'falk-disch', 'disch falk field', 'falk disch field',
    'disch', 'baseball field', 'baseball stadium', 'ut baseball field', 'ufcu disch falk'],
  ['DHG', 'Dobie Hall Garage', '', 'dobie garage'], ['DCG', 'Dobie Center Garage', '', 'dobie twenty one garage', 'dobie 21 garage', 'dobie twenty21 garage'], ['', 'Dell Seton Medical Center', '', 'dell seton']  , ['DTO', 'Dobie Twenty One', 'Dobie Twenty21'],
  ['EAS', 'Edgar A. Smith Building'], ['ECJ', 'Ernest Cockrell Jr. Hall', '', 'cockrell'],
  ['EER,EERC', 'Engineering Education and Research Center', '', 'eer lawn'], ['EPS', 'E. P. Schoch Building', '', 'schoch'],
  ['ERC', 'Erwin Special Events Center', '', 'erwin center'], ['ESS', 'Computer Science Annex'], ['ETC', 'Engineering Teaching Center II', '', 'etc ii', 'etc 2', 'etc2', 'etcii', 'etc two', 'engineering teaching center 2', 'engineering teaching center two', 'etc building'],
  ['FAC', 'Peter T. Flawn Academic Center', '', 'flawn', 'flawn academic center'],
  ['FC1,FC2,FC3,FC4,FC5,FC6,FC7,FC8', 'Facilities Complex'], ['FDF', 'Frank Denius Fields', '', 'denius fields'],
  ['FDH', 'J. Frank Dobie House'], ['FNT', 'Faulkner Nano Science and Technology Building'], ['FPC', 'OFPC Field Staff Office'],
  ['GAR', 'Garrison Hall', '', 'garrison'], ['GDC', 'Gates Dell Complex', '', 'gates dell', 'gates'], ['GEA', 'Mary E. Gearing Hall', '', 'gearing'],
  ['GEB', 'Dorothy L. Gebauer Building', '', 'gebauer'], ['GOL', 'Goldsmith Hall', '', 'goldsmith'], ['GRE', 'Gregory Gymnasium', 'Gregory Gym'],
  ['GWB', 'Gordon White Building'], ['GSB', 'Graduate School of Business Building'],
  ['HDB', 'Health Discovery Building'], ['HLB', 'Health Learning Building'], ['HMA', 'Hogg Memorial Auditorium', '', 'hogg auditorium'],
  ['HRC', 'Harry Ransom Center', '', 'ransom center'], ['HRH', 'Homer Rainey Hall'], ['HSM', 'William Randolph Hearst Building'],
  ['HTB', 'Health Transformation Building'], ['IPF', 'Indoor Practice Facility'], ['INT', 'International Office'],
  ['JCD', 'Jester Dormitory', 'Jester'], ['JES', 'Beauford H. Jester Center', 'Jester'], ['JGB', 'Jackson Geological Sciences Building', '', 'jackson geo', 'jackson school'],
  ['JHH', 'John W. Hargis Hall'], ['JON', 'Jesse H. Jones Hall'], ['KIN', 'Kinsolving Dormitory', 'Kinsolving'],
  ['LBJ', 'Lyndon B. Johnson Library', '', 'lbj library'], ['LCH', 'Littlefield Carriage House'], ['LDH', 'Longhorn Dining Facility'],
  ['LFH', 'Littlefield Home', 'Littlefield House'], ['LLA,LLB,LLC,LLD,LLE,LLF', 'Living Learning Centers', '', 'living learning center'],
  ['LTD', 'Littlefield Dormitory', 'Littlefield'], ['LTH', 'Laboratory Theatre Building'],
  ['MAI', 'Main Building', '', 'ut tower', 'the tower', 'tower'], ['MBB', 'Moffett Molecular Biology Building'], ['MEZ', 'Mezes Hall', '', 'mezes'],
  ['MFH', 'Mithoff Track and Soccer Field House', '', 'mithoff track/soccer field house', 'mithoff track soccer field house', 'mithoff field house'], ['MHD', 'Moore-Hill Dormitory', 'Moore-Hill'],
  ['MMS', 'Mike A. Myers Track and Soccer Stadium', '', 'myers stadium'], ['MNC', 'Moncrief-Neuhaus Athletic Center'],
  ['MRH', 'Music Building and Recital Hall', 'Butler School of Music', 'music building'], ['MSB', 'Mail Services Building'], ['MTC', 'Material Transfers Center'],
  ['NEZ', 'North End Zone Building'], ['NHB', 'Norman Hackerman Building', '', 'hackerman'], ['NMS', 'Neural and Molecular Science Building'],
  ['NUR', 'Nursing School'], ['PAC', 'Performing Arts Center'], ['PAI', 'T.S. Painter Hall', '', 'painter', 'painter hall'],
  ['PAR', 'Parlin Hall', '', 'parlin'], ['PAT', 'J.T. Patterson Laboratories Building', '', 'patterson'], ['PCL', 'Perry-Castaneda Library', 'PCL'],
  ['PHD', 'Prather Hall Dormitory', 'Prather'], ['PHR', 'Pharmacy Building', '', 'pharmacy'], ['POB', "Peter O'Donnell Jr. Building", '', 'odonnell'],
  ['PPA', 'Hal C. Weaver Power Plant Annex'], ['PPE', 'Hal C. Weaver Power Plant Expansion'], ['PPL', 'Hal C. Weaver Power Plant'],
  ['RHD', 'Roberts Hall Dormitory', 'Roberts'], ['RRH', 'Robert B. Rowling Hall', '', 'rowling', 'rowling hall'], ['RLM', 'Robert Lee Moore Hall'],
  ['RLP', 'Patton Hall'], ['RSC', 'Recreational Sports Center', 'Rec Center'],
  ['SAC', 'Student Activity Center', '', 'wcp', 'william c powers', 'powers student activity center', 'william c. powers jr. student activity center', 'william c powers jr student activity center'], ['SBS', 'Red and Charline McCombs Field'],
  ['SEA', 'Sarah M. and Charles E. Seay Building', '', 'seay'], ['SER', 'Service Building'], ['SJH', 'San Jacinto Residence Hall', 'San Jacinto'],
  ['SRH', 'Sid Richardson Hall'], ['SSB', 'Student Services Building', '', 'student services'], ['SSW', 'School of Social Work Building', '', 'social work'],
  ['STD', 'DKR-Texas Memorial Stadium', 'DKR Stadium'], ['SUT', 'Sutton Hall', 'Sutton'], ['SW7', '2617 Speedway'],
  ['SZB', 'George I. Sanchez Building', '', 'sanchez'], ['TCC', 'Joe C. Thompson Conference Center', '', 'thompson conference center'],
  ['TMM', 'Texas Memorial Museum'], ['TNH', 'Townes Hall', '', 'law school'], ['TSC', 'Jamail Texas Swimming Center', '', 'swim center'],
  ['TTC', 'Texas Tennis Center'], ['UA9', '2609 University Avenue', '', '2609 university ave', '2609 university', '2609 univ ave', 'ua9 building'], ['UIL', 'University Interscholastic League'],
  ['UNB', 'Union Building', 'Texas Union'], ['UPB', 'University Police Building', '', 'utpd'], ['USS', 'University Sign Shop'],
  ['UTA', 'UT Administration Building', 'UT Administration'], ['UTC', 'University Teaching Center'], ['UTX', 'Etter-Harbin Alumni Center', '', 'alumni center', 'etter harbin', 'etter harbin alumni center', 'etter', 'ehac', 'harbin'],
  ['WAG', 'Waggener Hall'], ['WAT', 'Arthur P. Watson House'], ['WCH', 'Will C. Hogg Building'], ['WEL', 'Robert A. Welch Hall', 'Welch'],
  ['WIN', 'F. Loren Winship Drama Building', 'Winship'], ['WMB', 'West Mall Office Building'], ['WRW', 'W. R. Woolrich Laboratories', '', 'woolrich'],
  ['WWH', 'Walter Webb Hall'],
  ['GLT', 'Gary L. Thomas Energy Engineering Building', '', 'gary l thomas', 'gary thomas'],
  ['', 'Dell Medical School', '', 'dell medical center', 'dell med', 'dell medical', 'dell med school', 'dell medical school campus',
    'dell med center', 'dell medical school building'],
  ['', 'Kinsolving Dormitory', 'Kinsolving', 'kins dining', 'kins dining hall', 'kinsolving dining', 'kinsolving dining hall', 'kins cafe'],

  ['', 'Littlefield Fountain', '', 'fountain', 'little field fountain', 'littlefield fountain'],
  ['', 'Littlefield Home', 'Littlefield House', 'littlefield cafe', 'littlefield patio cafe'],
  ['PMA', 'Robert Lee Moore Hall', '', 'physics math astronomy', 'physics, math & astronomy building', 'physics math and astronomy building', 'physics math & astronomy', 'pma building'], ['', 'International Office', '', 'texas global'],
  ['', 'The Bubble', '', 'bubble'], ['', 'Clark Field', '', 'caven clark field'],
  ['', 'Performing Arts Center', '', 'bass concert hall', 'bass hall', 'bass', 'across bass concert hall', 'texas performing arts', 'texas performing arts building'],
  ['', 'Nursing School', '', 'school of nursing', 'nursing building'], ['', 'Mike A. Myers Track and Soccer Stadium', '', 'mike a myers soccer field', 'mike a myers stadium', 'myers'],
  ['', 'Moody Center', '', 'moody center arena'], ['VAC', 'Visual Arts Center', '', 'visual arts building', 'visual arts', 'visual art building', 'vac building'], ['', 'Townes Hall', '', 'ut law school', 'law school'],
  ['', 'Lyndon B. Johnson Library', '', 'lbj presidential library', 'lbj public affairs', 'lbj school'],
  ['DMC', 'Dealey Center for New Media', '', 'dealy center', 'dealey center', 'dealy center of new media', 'dealey center for new media', 'kut',
    'dealey media center', 'dealy media center', 'dealey new media', 'dealey center for new media',
    'g. b. dealey center for new media', 'gb dealey center for new media', 'g.b. dealey center', 'gb dealey center', 'dealey',
    'moody auditorium', 'dmc 2.106'],
  ['', 'DKR-Texas Memorial Stadium', 'DKR Stadium', 'dkr gate 6', 'stadium'], ['', 'Etter-Harbin Alumni Center', '', 'texas exes center', 'texas exes'],
  ['', 'Sutton Hall', 'Sutton', 'sutten', 'sutten hall'], ['', 'San Jacinto Residence Hall', 'San Jacinto', 'san jacinto resident hall', 'cypress bend cafe', 'cypress bend'],
  ['', 'Harry Ransom Center', '', 'harry randsom center', 'ransom'], ['', 'Robert B. Rowling Hall', '', 'rowlling'],
  ['', 'Art Building', '', 'art and art history building', 'art building and museum'], ['', 'Aerospace Engineering', '', 'aerospace building'],
  ['', 'Facilities Complex', '', 'facilities complex building 1'],
  ['', 'Clark Field', '', 'clarke field'], ['', 'Inner Campus Drive'], ['', 'Texas Memorial Museum', '', 'natural history museum'],
  ['ECGA', 'East Campus Graduate Apartments', '', 'east campus grad apartments'],
  ['', 'Chemical and Petroleum Engineering', '', 'chemical engineering building', 'chem eng building', 'new chemical eng building', 'petroleum engineering building', 'petroleum engineering', 'petroleum building', 'pete building', 'chemical and petroleum engineering building'],
  ['', 'Gregory Gymnasium', 'Gregory Gym', 'eddie reese outdoor pool', 'outdoor pool', 'gregory pool'], ['', 'Almetris Duren Hall', 'Duren', 'durren'],
  ['', "Peter O'Donnell Jr. Building", '', 'peter odonnell', 'peter o donnell'], ['', 'San Jacinto Residence Hall', 'San Jacinto', 'san jancinto', 'san j', 'san j bus stop'],
  ['', 'Townes Hall', '', 'school of law'], ['', 'Goldsmith Hall', '', 'school of architecture'],
  ['', 'Music Building and Recital Hall', 'Butler School of Music', 'bates recital hall', 'bates hall'], ['', 'Waggener Hall', '', 'wagner hall', 'waggener'],
  ['', 'Homer Rainey Hall', '', 'rainey hall'], ['', 'DKR-Texas Memorial Stadium', 'DKR Stadium', 'drk'],
  ['', 'Creekside Residence Hall', 'Creekside', 'waller creek', 'waller creek residence halls'], ['', 'Union Building', 'Texas Union', 'texas union on 24th', 'union on 24th st'],
  ['KSC', 'Kendra Scott Center', '', 'kendra scott', 'kendra scott centre', 'the kendra scott center'],

  ['BRG', 'Brazos Garage'], ['CCG', 'Conference Center Garage'], ['ECG', 'East Campus Garage'], ['GUG', 'Guadalupe Garage'],
  ['HCG', 'Health Center Garage'], ['MAG', 'Manor Garage'], ['RHG', 'Rowling Hall Garage'], ['SAG', 'San Antonio Garage'],
  ['SJG', 'San Jacinto Garage'], ['SWG', 'Speedway Garage'], ['TRG', 'Trinity Garage'], ['TSG', '27th Street Garage'], ['NUG', 'Nueces Garage'],
];
(function addCampusMap_() {
  const norm = placeNorm_;
  const owner = s => { const n = norm(s); return n ? Object.keys(PLACES).find(k => norm(k) === n || PLACES[k].some(v => norm(v) === n)) : null; };
  CAMPUS_MAP.forEach(([codes, name, join, ...more]) => {
    const codeList = codes ? codes.split(',') : [];
    const target = join || owner(name) || codeList.map(owner).find(Boolean) || name;
    if (!PLACES[target]) PLACES[target] = [];
    codeList.concat([name], more).forEach(a => {
      const v = a.toLowerCase(), o = owner(v);
      if (!o || o === target) { if (!PLACES[target].includes(v)) PLACES[target].push(v); }
    });
    const key = target.toLowerCase();
    if (!ZONE_OF[key]) ZONE_OF[key] = codeList.concat([name], more).map(x => ZONE_OF[x.toLowerCase()] || ZONE_OF[norm(x)]).find(Boolean) || 'On';
  });
  PLACE_INDEX_ = null;
})();

;

const FILE_VERSION_CREWS = '2026-11-07a';

const crewKey_ = c => (String(c.driver).trim() + '|' + String(c.nav || '').trim()).toLowerCase();

function loadCrewEdits_() {
  const v = JSON.parse(SP_().getProperty('CREW_EDITS') || 'null');
  const edits = v && v.night === nightKey_(new Date()) ? v : { night: nightKey_(new Date()), added: [], state: {} };
  edits.gone = edits.gone || {};
  edits.since = edits.since || {};
  return edits;
}

function saveCrewEdits_(edits) {
  SP_().setProperty('CREW_EDITS', JSON.stringify(edits));
}

function applyCrewEdits_(autoCrews) {
  const edits = loadCrewEdits_();
  const byKey = {};
  autoCrews.forEach(c => { byKey[crewKey_(c)] = Object.assign({}, c); });
  edits.added.forEach(c => {
    const k = crewKey_(c), auto = byKey[k] || autoCrews.find(a => samePair_(a, c)) || {};
    const keep = auto.van && auto.van === c.van ? { vehicle: auto.vehicle, phone: auto.phone, cap: auto.cap } : {};
    byKey[k] = Object.assign({}, keep, c, c.cap ? {} : { cap: keep.cap || 0 }, { edited: true, shift: c.shift || auto.shift || '' });
  });

  const rank = c => (c.edited ? 0 : c.shift ? 1 : 2);
  const keys = Object.keys(byKey).sort((a, b) => rank(byKey[a]) - rank(byKey[b]));
  const kept = [];
  keys.forEach(k => {
    const same = kept.find(x => samePair_(byKey[x], byKey[k]));
    if (!same) { kept.push(k); return; }
    const w = byKey[same], d = byKey[k];
    w.lastTs = Math.max(w.lastTs || 0, d.lastTs || 0);
    if (!w.loc && d.loc) w.loc = d.loc;
    delete byKey[k];
  });
  const all = Object.keys(byKey)
    .filter(k => !edits.gone[k] || (byKey[k].lastTs || 0) > edits.gone[k])
    .map(k => Object.assign(byKey[k], { key: k, state: edits.state[k] || 'on',
      breakSince: edits.state[k] === 'break' ? edits.since[k] || 0 : 0,
      gasSince: edits.state[k] === 'gas' ? edits.since[k] || 0 : 0 }))
    .sort((a, b) => a.driver.localeCompare(b.driver));
  if (typeof setCrewNames_ === 'function') setCrewNames_(all);
  return { list: all.filter(c => c.state !== 'off'), removed: all.filter(c => c.state === 'off') };
}

function saveCrew(oldKey, crew) {
  const raw = crew || {};
  const lead = String(crew.lead || '').trim();
  const cap = Math.max(0, Math.min(20, Number(crew.cap) || 0));
  crew = { driver: String(crew.driver || '').trim(), nav: String(crew.nav || '').trim(), van: crew.van || '', loc: crew.loc || '' };
  if (cap) crew.cap = cap;
  if (raw.shift) crew.shift = String(raw.shift).slice(0, 30);
  if (!crew.driver) throw new Error('A crew needs at least a driver name.');
  if (lead && [crew.driver, crew.nav].some(n => n && sameName_(n, lead))) crew.lead = lead;
  withLock_(() => {
    const edits = loadCrewEdits_();
    const key = crewKey_(crew);
    edits.added = edits.added.filter(c => crewKey_(c) !== key && crewKey_(c) !== oldKey);
    edits.added.push(crew);
    const [od, on] = String(oldKey || '').split('|');
    const swapped = oldKey && oldKey !== key && samePair_({ driver: od, nav: on }, crew);
    if (swapped) { edits.gone[oldKey] = Date.now(); delete edits.state[oldKey]; }
    else if (oldKey && oldKey !== key) edits.state[oldKey] = 'off';
    delete edits.state[key];
    delete edits.gone[key];
    saveCrewEdits_(edits);
  });
}

function removeCrews(keys) {
  withLock_(() => {
    const edits = loadCrewEdits_();
    keys.forEach(key => {
      edits.added = edits.added.filter(c => crewKey_(c) !== key);
      delete edits.state[key];
      edits.gone[key] = Date.now();
    });
    saveCrewEdits_(edits);
  });
}

function setCrewState(key, state) {
  withLock_(() => {
    const edits = loadCrewEdits_();
    if (state === 'on') delete edits.state[key];
    else edits.state[key] = state;
    if (state === 'break' || state === 'gas') edits.since[key] = Date.now(); else delete edits.since[key];
    saveCrewEdits_(edits);
  });
}

;

const FILE_VERSION_GROUPME = '2026-11-07a';

let GM_OFFLINE_ = false;

const GM = { API: 'https://api.groupme.com/v3', LOG_SIZE: 30 };

const PLACE_STOPWORDS = new Set(['the', 'hall', 'dorm', 'house', 'apartments', 'apts', 'building',
  'bldg', 'and', 'near', 'front', 'outside', 'from']);

function gmRequest_(url, opts) {
  let last;
  for (let i = 0; i < 3; i++) {
    try { return UrlFetchApp.fetch(url, opts); } catch (e) { last = e; if (i < 2) Utilities.sleep(400 * (i + 1)); }
  }
  const why = String((last && last.message) || last).replace(/:\s*https?:\/\/\S+/, '');
  throw new Error(`Couldn't reach GroupMe from Google's servers (${why}). Usually temporary — it retries on the next refresh.`);
}

function gmFetch_(path, params, tokenOverride) {
  const token = tokenOverride || SP_().getProperty('GROUPME_TOKEN');
  const query = Object.keys(params || {}).map(k => `${k}=${encodeURIComponent(params[k])}`).join('&');
  const resp = gmRequest_(`${GM.API}${path}${query ? '?' + query : ''}`,
    { headers: { 'X-Access-Token': token }, muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code === 304) return null;
  if (code === 401) throw new Error('GroupMe token rejected — copy a fresh one from web.groupme.com.');
  if (code !== 200) throw new Error(`GroupMe returned ${code}`);
  return JSON.parse(resp.getContentText()).response;
}

function saveGroupMeToken(raw) {

  const token = String(raw || '').replace(/^\s*x-access-token\s*:?\s*/i, '').trim();
  if (!/^[A-Za-z0-9_-]{20,128}$/.test(token)) throw new Error("That doesn't look like a GroupMe token — copy just the value after X-Access-Token.");
  const me = gmFetch_('/users/me', {}, token);
  SP_().setProperty('GROUPME_TOKEN', token);
  SP_().setProperty('GROUPME_ME_ID', String(me.user_id || me.id || ''));
  SP_().setProperty('GROUPME_ME_NAME', String(me.name || ''));
  const ids = JSON.parse(SP_().getProperty('DISPATCHER_GM_IDS') || '[]'), mine = String(me.user_id || me.id || '');
  if (mine && !ids.includes(mine)) SP_().setProperty('DISPATCHER_GM_IDS', JSON.stringify(ids.concat(mine).slice(-30)));
  SP_().deleteProperty('GM_ERROR');
  return { name: me.name };
}

function forgetGroupMeToken() {
  const p = SP_();

  ['GROUPME_TOKEN', 'GROUPME_ME_ID', 'GROUPME_ME_NAME'].forEach(k => p.deleteProperty(k));
}

function listMyGroups() {
  return gmFetch_('/groups', { per_page: 100, omit: 'memberships' }).map(g => ({ id: g.id, name: g.name }));
}

function setGroupMeGroup(id, name) {
  const p = SP_();
  p.setProperty('GROUPME_GROUP_ID', String(id));
  p.setProperty('GROUPME_GROUP_NAME', String(name || id));
  ['SYNC_ON', 'GM_LAST_ID', 'GM_LOG', 'GM_ERROR'].forEach(k => p.deleteProperty(k));
  CacheService.getScriptCache().remove('recent');
}

function listGroupMeGroups() {
  listMyGroups().forEach(g => console.log(`${g.id}  ${g.name}`));
}

function setSyncRunning(on) {
  const p = SP_();
  p.setProperty('SYNC_ON', on ? 'true' : 'false');
  if (on) {

    if (p.getProperty('GM_LAST_NIGHT') !== nightKey_(new Date())) p.deleteProperty('GM_LAST_ID');
    try { backfillCrews_(); } catch (e) { p.setProperty('GM_ERROR', e.message); }
  }
}

const GAS_DONE_RE = /\b(ready|fill(ed)?( (up|it|gas|the tank))?|got gas|gassed( up)?|done|finished|back|full( tank)?|tank( is|'?s)? full|topped (off|up)|all set|good to go|(we'?re|we are|were) good)\b/i;

function isEmptyMsg_(text) {
  const w = msgWords_(text);
  return w.some((x, i) => isEmptyWord_(x) && w[i - 1] !== 'not') || /\bready for (rides|more)\b/i.test(text);
}

function loadPeople_() {
  const v = JSON.parse(SP_().getProperty('CREW_PEOPLE') || 'null');
  const people = v && v.night === nightKey_(new Date()) ? v.people : {};

  POSTER_NAMES_ = new Set(Object.keys(people).filter(k => people[k] && people[k].spoke));
  return people;
}

function savePeople_(people) {
  SP_()
    .setProperty('CREW_PEOPLE', JSON.stringify({ night: nightKey_(new Date()), people }));
}

function notePerson_(people, name, field, ms) {
  const k = norm_(String(name).trim().split(/\s+/)[0]);
  if (k.length < 2) return;
  const cur = people[k] || (people[k] = {});
  cur[field] = Math.max(cur[field] || 0, ms);
}

function trackCrews_(people, m) {
  if (m.system || !m.text) return;
  const ms = m.created_at * 1000;
  if (m.sender_type === 'user') {
    notePerson_(people, m.name, 'spoke', ms);
    const k = norm_(String(m.name).trim().split(/\s+/)[0]);
    if (people[k]) people[k].full = m.name;
  }
  if (m.sender_type === 'user' && isEmptyMsg_(m.text)) notePerson_(people, m.name, 'empty', ms);
  if (m.sender_type === 'user' && GAS_DONE_RE.test(m.text)) notePerson_(people, m.name, 'ready', ms);

  if (!phonesIn_(m.text).length) return;
  const mentions = (m.attachments || []).filter(a => a.type === 'mentions');
  mentions.forEach(a =>
    (a.loci || []).forEach(([start, len]) =>
      notePerson_(people, m.text.substr(start, len).replace(/^@/, ''), 'assigned', ms)));

  if (!mentions.length) {
    (m.text.match(/@[A-Za-z][\w'-]*/g) || []).forEach(t => notePerson_(people, t.slice(1), 'assigned', ms));
  }
}

function backfillCrews_() {
  const p = SP_();
  const group = p.getProperty('GROUPME_GROUP_ID');
  if (!p.getProperty('GROUPME_TOKEN') || !group) return;
  const since = Date.now() / 1000 - 2 * 3600;
  const people = loadPeople_(), meId = dispatcherIds_(), seenMsgs = [];
  let beforeId = null, reachedOld = false;
  for (let page = 0; page < 5 && !reachedOld; page++) {
    const r = gmFetch_(`/groups/${group}/messages`, beforeId ? { limit: 100, before_id: beforeId } : { limit: 100 });
    if (!r || !r.messages.length) break;
    for (const m of r.messages) {
      if (m.created_at < since) { reachedOld = true; break; }
      trackCrews_(people, m);
      if (!m.system && m.text) seenMsgs.push(chatEntry_(m, meId));
    }
    beforeId = r.messages[r.messages.length - 1].id;
  }
  savePeople_(people);
  addChat_(seenMsgs.reverse());
}

function setCrewStatus(key, status) {
  const [driver, nav] = String(key).split('|');
  withLock_(() => {
    const people = loadPeople_(), now = Date.now();
    [driver, nav].filter(Boolean).forEach(n => {
      notePerson_(people, n, status === 'available' ? 'empty' : 'assigned', now);
      notePerson_(people, n, 'manual', now);
    });
    savePeople_(people);
  });
}

function noteAssigned_(crew) {
  const people = loadPeople_();
  [crew.driver, crew.nav].filter(Boolean).forEach(n => notePerson_(people, n, 'assigned', Date.now()));
  savePeople_(people);
}

function crewBoard_(crews, rides) {
  const people = loadPeople_();
  const lookup = name => {
    const n = norm_(String(name || '').split(/\s+/)[0]), keys = Object.keys(people).filter(k => sameName_(k, name));
    const close = keys.filter(k => n.length >= 3 && (k === n || k.startsWith(n) || n.startsWith(k)));
    return (close.length ? close : keys).map(k => people[k]);
  };
  const order = { available: 0, busy: 1, unknown: 2, gas: 3, break: 4 };
  return crews.map((c, idx) => {
    const ppl = [c.driver, c.nav].filter(Boolean).flatMap(lookup);
    const empty = Math.max(0, ...ppl.map(x => x.empty || 0));
    const assigned = Math.max(0, ...ppl.map(x => x.assigned || 0));

    const ready = Math.max(0, ...ppl.map(x => x.ready || 0));

    const gasDone = c.state === 'gas' && ready > (c.gasSince || 0);
    const gas = c.state === 'gas' && !gasDone;
    const freeAt = gasDone ? ready : empty;
    const status = c.state === 'break' ? 'break' : gas ? 'gas' : freeAt > assigned ? 'available' : assigned ? 'busy' : 'unknown';
    const sinceMs = status === 'available' ? freeAt : assigned;
    const manual = !!sinceMs && ppl.some(x => x.manual === sinceMs);
    const toPick = rides.filter(r => r.driver && !r.picked && samePair_(r, c)).length;

    const spoke = n => { const p = n ? lookup(n) : []; return { ms: Math.max(0, ...p.map(x => x.spoke || 0)), full: (p.find(x => x.full) || {}).full }; };
    const d = spoke(c.driver), v = spoke(c.nav);
    const pinned = c.lead ? [c.driver, c.nav].find(n => n && sameName_(n, c.lead)) : '';
    const lead = pinned ? { name: pinned, other: pinned === c.driver ? c.nav : c.driver, full: spoke(pinned).full }
      : v.ms > d.ms ? { name: c.nav, other: c.driver, full: v.full } : d.ms ? { name: c.driver, other: c.nav, full: d.full } : null;
    return Object.assign({}, c, { idx, status, sinceMs, emptyMs: empty, since: sinceMs ? fmtTime_(new Date(sinceMs)) : '', toPick,
      lead: lead ? lead.name : '', other: lead ? lead.other : '', leadFull: lead ? (lead.full || lead.name) : '', leadPinned: !!pinned, manual,

      roleNav: c.nav ? (lead ? lead.name : c.nav) : '', roleDriver: c.nav ? (lead ? lead.other : c.driver) : c.driver });
  }).sort((a, b) => order[a.status] - order[b.status] || (a.status === 'available' ? a.sinceMs - b.sinceMs : 0));
}

function gasBack_(people, crews, m) {
  const ms = m.created_at * 1000;
  const c = crews.find(x => x.state === 'gas' && ms > (x.gasSince || 0) && [x.driver, x.nav].some(n => n && sameName_(n, m.name)));
  if (!c) return null;
  const edits = loadCrewEdits_();
  delete edits.state[c.key]; delete edits.since[c.key];
  saveCrewEdits_(edits);
  c.state = 'on'; c.gasSince = 0;
  [c.driver, c.nav].filter(Boolean).forEach(n => notePerson_(people, n, 'empty', ms));
  return `Back from gas — ${[c.driver, c.nav].filter(Boolean).join(' / ')} marked empty`;
}

const idAfter_ = (a, b) => (String(a).length !== String(b).length ? String(a).length > String(b).length : String(a) > String(b));

const CATCHUP_MIN = 60;

function syncGroupMe(pushed) {
  const p = SP_();
  const group = p.getProperty('GROUPME_GROUP_ID');
  if (!p.getProperty('GROUPME_TOKEN') || !group || p.getProperty('SYNC_ON') !== 'true') return 0;

  let lastId = p.getProperty('GM_LAST_ID');
  if (!lastId) {
    const since = Date.now() / 1000 - CATCHUP_MIN * 60;
    let beforeId = null, oldest = null;
    for (let page = 0; page < 4 && !lastId; page++) {
      const r = gmFetch_(`/groups/${group}/messages`, beforeId ? { limit: 100, before_id: beforeId } : { limit: 100 });
      const ms = (r && r.messages) || [];
      if (!ms.length) break;
      const older = ms.find(m => m.created_at < since);
      if (older) lastId = older.id;
      oldest = ms[ms.length - 1].id;
      beforeId = oldest;
    }
    if (!lastId) { if (!oldest) return 0; lastId = oldest; }
  }
  const msgs = [];
  let newLast = lastId;
  for (let page = 0; page < 5; page++) {
    const r = gmFetch_(`/groups/${group}/messages`, { after_id: newLast, limit: 100 });
    if (!r || !r.messages.length) break;
    msgs.push(...r.messages);
    newLast = r.messages[r.messages.length - 1].id;
    if (r.messages.length < 100) break;
  }
  const live = (pushed || []).filter(m => m && m.id && String(m.group_id || group) === String(group))
    .map(m => Object.assign({}, m, { id: String(m.id), sender_type: m.sender_type || (m.system ? 'system' : 'user'), attachments: m.attachments || [] }));
  live.forEach(m => { if (!msgs.some(x => String(x.id) === m.id)) msgs.push(m); });
  msgs.sort((a, b) => (idAfter_(a.id, b.id) ? 1 : idAfter_(b.id, a.id) ? -1 : 0));
  if (!msgs.length) return 0;
  const meId = dispatcherIds_();
  try { recentMessages_(); } catch (e) {   }

  const likes = [], replies = [];
  LAST_UPDATE_ = {};
  BATCH_MSGS_ = msgs.map(m => ({ id: String(m.id), who: m.name, ts: m.created_at, text: String(m.text || '') }));
  const locked = tryWithLock_(5000, () => {
    const p2 = SP_();

    const saved = p2.getProperty('GM_LAST_ID') || lastId;
    const done = JSON.parse(p2.getProperty('GM_PUSHED') || '[]');
    const fresh = msgs.filter(m => idAfter_(m.id, saved) && !done.includes(String(m.id)));
    if (!fresh.length) return 0;
    const fromPush = fresh.filter(m => live.some(x => x.id === String(m.id))).map(m => String(m.id));
    if (fromPush.length) p2.setProperty('GM_PUSHED', JSON.stringify(done.concat(fromPush).slice(-200)));

    const sh = sheet_();
    const t = tonight_(), rides = t.rides;
    rides.olderActive = t.olderActive;
    const crewsNow = applyCrewEdits_(t.crews).list;
    tagGroups_(rides, fresh);
    const log = JSON.parse(p2.getProperty('GM_LOG') || '[]');
    const people = loadPeople_(), entries = [];
    fresh.forEach(m => {
      trackCrews_(people, m);
      if (m.system || !m.text) return;
      const entry = chatEntry_(m, meId);

      const back = m.sender_type === 'user' && !entry.me && GAS_DONE_RE.test(m.text) ? gasBack_(people, crewsNow, m) : null;
      if (back) {
        Object.assign(entry, { kind: 'gasback', status: 'ok', result: back });
        log.unshift({ t: fmtTime_(new Date(m.created_at * 1000)), who: m.name, text: String(m.text).slice(0, 200), ok: true, result: back });
        entries.push(entry);
        return;
      }

      if (entry.kind === 'assign' && entry.me) {
        const res = applyMyPost_(sh, rides, crewsNow, m);
        if (res) {
          log.unshift({ t: fmtTime_(new Date(m.created_at * 1000)), who: m.name, text: String(m.text).slice(0, 200), ok: res.ok, result: res.result });
          Object.assign(entry, { status: res.ok ? 'ok' : 'ask', result: res.result, rideIds: res.ids || [] });
        }
      }
      if (m.sender_type === 'user' && !['assign', 'mine'].includes(entry.kind)) {
        const res = applyMessage_(sh, rides, m);
        if (res) {
          log.unshift({ t: fmtTime_(new Date(m.created_at * 1000)), who: m.name, text: String(m.text).slice(0, 200), ok: res.ok, result: res.result });
          Object.assign(entry, { status: res.ok ? 'ok' : 'ask', result: res.result, rows: res.rows || [], rideIds: res.ids || [] });
          if (res.cancelled) likes.push(String(m.id));
          if (res.ok && ['at', 'picked', 'atpicked'].includes(entry.kind) && (res.ids || []).length)
            LAST_UPDATE_[m.name] = { kind: entry.kind, ids: res.ids, ms: m.created_at * 1000 };
          if (res.pickAsk) {
            const asks = JSON.parse(p2.getProperty('PICK_ASKS') || '[]').filter(a => Date.now() - a.askMs < 60 * 60e3);
            p2.setProperty('PICK_ASKS', JSON.stringify(asks.concat([res.pickAsk])));
          }
          if (res.replies) replies.push(...res.replies);
          if (res.wait) {
            const waits = JSON.parse(p2.getProperty('SKIP_WAITS') || '[]').filter(w => w.askId !== res.wait.askId);
            p2.setProperty('SKIP_WAITS', JSON.stringify(waits.concat([res.wait])));
          }
        }
      }
      entries.push(entry);
    });
    p2.setProperty('GM_LOG', JSON.stringify(log.slice(0, GM.LOG_SIZE)));
    savePeople_(people);
    addChat_(entries);
    p2.setProperty('GM_LAST_ID', idAfter_(newLast, saved) ? newLast : saved);
    p2.setProperty('GM_LAST_NIGHT', nightKey_(new Date()));
    return fresh.length;
  });

  if (locked.busy) return 0;
  const applied = locked.value;
  likes.forEach(id => { try { gmPost_(`/messages/${group}/${id}/like`, {}); } catch (e) {   } });
  replies.forEach(r => { try { postMessage_(r.text, r.mentions, r.reply); } catch (e) { console.error('skip answer', e); } });
  if (likes.length) try { CacheService.getScriptCache().remove('recent'); } catch (e) {   }
  return applied;
}

function recentMessages_() {
  const p = SP_(), group = p.getProperty('GROUPME_GROUP_ID');
  if (!p.getProperty('GROUPME_TOKEN') || !group || p.getProperty('SYNC_ON') !== 'true') return [];
  const cache = CacheService.getScriptCache(), hit = cache.get('recent');
  if (hit) return JSON.parse(hit);
  if (GM_OFFLINE_) {
    const keep = cache.get('recent_keep');
    return keep ? JSON.parse(keep) : [];
  }
  const since = Date.now() / 1000 - (CFG.EDIT_WINDOW_MIN + 15) * 60, out = [];
  let beforeId = null, reachedOld = false;
  for (let page = 0; page < 4 && !reachedOld; page++) {
    const r = gmFetch_(`/groups/${group}/messages`, beforeId ? { limit: 100, before_id: beforeId } : { limit: 100 });
    if (!r || !r.messages.length) break;
    for (const m of r.messages) {
      if (m.created_at < since) { reachedOld = true; break; }
      out.push({ id: String(m.id), ts: m.created_at, who: m.name, text: String(m.text || '').slice(0, 300), rx: reactionsOf_(m) });
    }
    beforeId = r.messages[r.messages.length - 1].id;
  }
  try { cache.put('recent', JSON.stringify(out.slice(0, 250)), 20); } catch (e) {   }

  try { cache.put('recent_keep', JSON.stringify(out.filter(m => m.rx.length || phonesIn_(m.text).length).slice(0, 150)), 21600); }
  catch (e) {   }
  return out;
}

function reactionsOf_(m) {
  const out = [], add = (emoji, uids) => {
    uids = (uids || []).map(String);
    if (!uids.length) return;
    const e = /^(❤|❤️|♥|♥️)$/.test(emoji) ? '❤️' : emoji, same = out.find(x => x.emoji === e);
    if (same) same.uids = [...new Set(same.uids.concat(uids))];
    else out.push({ emoji: e, uids: [...new Set(uids)] });
  };
  add('❤️', m.favorited_by);
  (m.reactions || []).forEach(r => add(r.code && /[^\x00-\x7f]/.test(r.code) ? r.code : '❤️', r.user_ids));
  return out;
}

function recentlyPosted_() {
  const since = Date.now() / 1000 - CFG.EDIT_WINDOW_MIN * 60, posted = {};
  recentMessages_().filter(m => m.ts >= since).forEach(m => m.text.replace(/\d[\d\s().-]{8,}\d/g, s => {
    const digits = s.replace(/\D/g, '').slice(-10);
    if (digits.length === 10 && !posted[digits]) posted[digits] = { t: fmtTime_(new Date(m.ts * 1000)), by: m.who, id: m.id };
  }));
  return posted;
}

function syncGroupMeQuietly_(pushed) {
  try {
    const n = syncGroupMe(pushed) || 0;
    try { flushSkipWaits_(); } catch (e) { console.error('skip follow-ups', e); }
    if (SP_().getProperty('GM_ERROR')) SP_().deleteProperty('GM_ERROR');
    return n;
  } catch (e) {
    console.error('syncGroupMe', e);

    if (!/lock timeout|another process/i.test(String(e && e.message))) {
      try { SP_().setProperty('GM_ERROR', 'GroupMe sync: ' + e.message); } catch (e2) {   }
    }
    return 0;
  }
}

function syncTickState() {
  return syncTick() > 0 ? getState(true) : null;
}

function syncTick() {
  if (SP_().getProperty('SYNC_ON') !== 'true') return 0;
  let n = syncGroupMeQuietly_();

  const cache = CacheService.getScriptCache();
  try {
    const sig = JSON.stringify(recentMessages_().filter(m => m.rx.length).map(m => [m.id, m.rx.map(r => r.emoji + r.uids.length)]));
    if (sig !== cache.get('rx_sig')) { cache.put('rx_sig', sig, 21600); n++; }
  } catch (e) {   }
  return n;
}

function liveConfig() {
  const p = SP_(), token = p.getProperty('GROUPME_TOKEN'), group = p.getProperty('GROUPME_GROUP_ID');
  if (!token || !group) return { ok: false, why: 'GroupMe is not set up.' };
  let viewer = '', owner = '';
  try { viewer = Session.getActiveUser().getEmail(); owner = Session.getEffectiveUser().getEmail(); } catch (e) {   }

  if (!excelMode_() && (!viewer || viewer !== owner)) return { ok: false, why: 'Live updates only run for the account that owns this app (Deploy: Execute as Me, Who has access: Only myself).' };
  const me = meId_();
  if (!me) return { ok: false, why: "Couldn't read your GroupMe user id yet." };
  return { ok: true, token, userId: me, groupId: String(group) };
}

function liveTick(reactions, pushed) {
  if (SP_().getProperty('SYNC_ON') !== 'true') return null;
  syncGroupMeQuietly_(Array.isArray(pushed) ? pushed.slice(0, 20) : []);
  if (reactions) {
    CacheService.getScriptCache().remove('recent');
    try { recentMessages_(); } catch (e) {   }
  }
  return getState(true);
}

function groupMeStatus_() {
  const p = SP_();
  return {
    enabled: !!(p.getProperty('GROUPME_TOKEN') && p.getProperty('GROUPME_GROUP_ID')),
    tokenHint: (p.getProperty('GROUPME_TOKEN') || '').slice(-4),
    groupName: p.getProperty('GROUPME_GROUP_NAME') || p.getProperty('GROUPME_GROUP_ID') || '',
    running: p.getProperty('SYNC_ON') === 'true',
    autopilot: p.getProperty('AUTOPILOT') === 'true',
    error: p.getProperty('GM_ERROR') || '',
    log: JSON.parse(p.getProperty('GM_LOG') || '[]'),
  };
}

function setAutopilot(on) {
  if (on) SP_().setProperty('AUTOPILOT', 'true'); else SP_().deleteProperty('AUTOPILOT');
  return !!on;
}
const autopilotOn_ = () => SP_().getProperty('AUTOPILOT') === 'true' && SP_().getProperty('SYNC_ON') === 'true';

function clearGroupMeLog() {
  SP_().deleteProperty('GM_LOG');
}

const norm_ = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');

let CREW_NAMES_ = new Set(), POSTER_NAMES_ = new Set();
function setCrewNames_(crews) {
  CREW_NAMES_ = new Set();
  (crews || []).forEach(c => [c.driver, c.nav].forEach(n => { const k = norm_(String(n || '').split(/\s+/)[0]); if (k.length >= 3) CREW_NAMES_.add(k); }));
}

const ownedByOther_ = z => [...CREW_NAMES_, ...POSTER_NAMES_].some(h => h !== z && (h.startsWith(z) || z.startsWith(h)));

function sameName_(a, b) {
  const x = norm_(String(a).split(/\s+/)[0]), y = norm_(String(b).split(/\s+/)[0]);
  if (x.length < 3 || y.length < 3) return x === y && !!x;
  return x.startsWith(y) || y.startsWith(x);
}

function nameIssues_(board) {
  let members = [];
  try { members = groupMembers_(); } catch (e) { return []; }
  if (!members.length) return [];
  const firsts = members.map(m => norm_(String(m.name).split(/\s+/)[0])).filter(Boolean), out = [];
  board.forEach(b => [b.driver, b.nav].filter(Boolean).forEach(n => {
    if (!firsts.some(f => sameName_(f, n))) out.push({ key: b.key, name: n, crew: [b.driver, b.nav].filter(Boolean).join(' / ') });
  }));
  return out;
}

function placeTokens_(s) {
  return String(s).toLowerCase().split(/[^a-z0-9]+/).filter(w => w.length >= 3 && !PLACE_STOPWORDS.has(w));
}

function placeScore_(place, pickup) {
  const pick = placeTokens_(pickup);
  return placeTokens_(place).filter(w => pick.some(p => p.startsWith(w) || w.startsWith(p))).length;
}

function typoDist_(a, b) {
  const d = Array.from({ length: a.length + 1 }, (_, i) => [i].concat(Array(b.length).fill(0)));
  for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
    }
  }
  return d[a.length][b.length];
}
const msgWords_ = text => String(text).toLowerCase().replace(/[^a-z0-9/\s]/g, ' ').split(/\s+/).filter(Boolean);
const isPickedWord_ = w => ['pu', 'p/u', 'pickup', 'pickedup', 'picked'].includes(w) ||
  (w[0] === 'p' && w.length >= 4 && !/^picking/.test(w) && typoDist_(w, 'picked') <= 2);

const isDroppedWord_ = w => ['drop', 'drops', 'dropping', 'droppin', 'dropoff', 'droppedoff', 'dropped'].includes(w) ||
  (w[0] === 'd' && w.length >= 5 && !/^dropping/.test(w) && typoDist_(w, 'dropped') <= 2);

const isEmptyWord_ = w => {
  if (w === 'mt') return true;
  const x = String(w).replace(/(.)\1+/g, '$1');
  return x[0] === 'e' && x.length >= 4 && x.length <= 7 && typoDist_(x, 'empty') <= 1;
};
const isArrivedWord_ = w => ['at', 'here', 'outside', 'arrived'].includes(w) ||
  (w[0] === 'a' && w.length >= 5 && typoDist_(w, 'arrived') <= 2);
const GROUP_WORDS = new Set(['all', 'everyone', 'everybody', 'both', 'them', 'all of them', 'group', 'the group', 'all riders']);
const LEAD_FILLER = new Set(['just', 'ok', 'okay', 'we', 'yes', 'yep', 'and', 'also', 'have', 'was', 'were', 'been',
  'i', 'im', 're', 'm', 'already']);

const WHEN_NUM = '(\\d{1,3}|an?|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|fifteen|twenty|thirty|forty|' +
  'a\\s+couple(?:\\s+of)?|couple(?:\\s+of)?|a\\s+few|few|half\\s+an?|half)';
const WHEN_UNIT = '(m|mn|min|mins|minute|minutes|h|hr|hrs|hour|hours)';
const WHEN_WORDS = { a: 1, an: 1, one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10,
  eleven: 11, twelve: 12, fifteen: 15, twenty: 20, thirty: 30, forty: 40, couple: 2, few: 3, half: 0.5 };
const WHEN_AGO_RE = new RegExp(`(?:\\b(?:like|about|around|abt|approx|maybe)\\s+|~\\s*)?\\b${WHEN_NUM}\\s*-?\\s*${WHEN_UNIT}?\\s*(?:ago|back|earlier|before)\\b`, 'i');
const WHEN_FOR_RE = new RegExp(`\\b(?:for|since|been(?:\\s+(?:here|there))?(?:\\s+for)?|waiting(?:\\s+(?:here|there))?(?:\\s+for)?)` +
  `\\s+(?:like\\s+|about\\s+|around\\s+|~\\s*)?${WHEN_NUM}\\s*-?\\s*${WHEN_UNIT}\\b`, 'i');
const WHEN_CLOCK_RE = /(?:\b(at|since|around|by)\s+|@\s*)?\b(\d{1,2}):(\d{2})\s*(am|pm|a\.m\.|p\.m\.|a|p)?(?![\w:])/i;

function whenMinutes_(num, unit) {
  const n = String(num).toLowerCase().replace(/\s+(of|an?)$/, '').replace(/^a\s+/, '').trim();
  const v = /^\d+$/.test(n) ? Number(n) : WHEN_WORDS[n];
  if (v == null) return null;
  return v * (unit && /^h/i.test(unit) ? 60 : 1);
}

function extractWhen_(text, msgMs) {
  text = String(text || '');
  const cut = (m, keepAt) => (text.slice(0, m.index) + (keepAt ? ' at ' : ' ') + text.slice(m.index + m[0].length)).replace(/\s+/g, ' ').trim();
  let m = WHEN_AGO_RE.exec(text) || WHEN_FOR_RE.exec(text);
  if (m) {
    const min = whenMinutes_(m[1], m[2]);
    if (min != null && min > 0 && min <= 120) return { text: cut(m, false), ms: msgMs - Math.round(min) * 60e3, said: true };
  }
  m = WHEN_CLOCK_RE.exec(text);
  if (m) {
    let h = Number(m[2]) % 12, min = Number(m[3]);
    const ap = (m[4] || '').toLowerCase()[0];
    const msgD = new Date(msgMs);
    const now = Number(Utilities.formatDate(msgD, CFG.TZ, 'H')) * 60 + Number(Utilities.formatDate(msgD, CFG.TZ, 'm'));

    const cands = ap ? [(h + (ap === 'p' ? 12 : 0)) * 60 + min] : [h * 60 + min, (h + 12) * 60 + min];
    const back = Math.min(...cands.map(c => (now - c + 1440) % 1440));
    if (min < 60 && back <= 120) {

      const keepAt = /^\s*$/.test(text.slice(0, m.index)) && /^(at|@)/i.test(m[0].trim());
      return { text: cut(m, keepAt), ms: msgMs - back * 60e3, said: true };
    }
  }
  return { text, ms: msgMs, said: false };
}

const CANCEL_RE = /\b(cancel+(?:ed|led|ing|s)?|cancell?ation|cxl|cxld|skip+(?:ed|ing|pin)?|no[- ]?show(?:ed)?)\b/i;

let BATCH_MSGS_ = [];
function recentFrom_(m) {
  const seen = new Set();
  return loadChat_().concat(BATCH_MSGS_)
    .filter(e => e.who === m.name && e.ts < m.created_at && m.created_at - e.ts <= 600 && !seen.has(e.id) && seen.add(e.id))
    .sort((a, b) => b.ts - a.ts);
}
const isCancelMsg_ = text => CANCEL_RE.test(String(text || ''));
function cancelReason_(text) {
  const t = String(text).toLowerCase();
  if (/not answer|no answer|didn'?t answer|wasn'?t answer|unresponsive|no response|not picking up|no[- ]?show/.test(t)) return 'Not Answering';
  if (/not ready/.test(t)) return 'Not Ready';
  if (/uber|lyft|walk(ed|ing)?|got a ride|found a ride|another ride|friend|bus|alternative/.test(t)) return 'Chose Alternative Way Home';
  if (/2 ?min|two min|waited/.test(t)) return 'Exceeded 2min Wait';
  return 'Other';
}
function cancelByMessage_(sh, rides, m) {
  const msgMs = m.created_at * 1000;
  const open = rides.filter(r => r.driver && !r.picked && r.tsMs < msgMs && [r.driver, r.nav].some(n => sameName_(n, m.name)));
  if (SKIP_ASK_RE.test(m.text)) return skipAsk_(rides, open, m);
  if (!open.length) return { ok: true, result: 'Nothing open to cancel for this crew (already done in the sheet)' };

  const clean = t => t.replace(/\b(picked|pu|p\/u|at|here|dropped|drop|dropping)\s+(up\s+|off\s+)?[a-z]+/gi, ' ')
    .replace(new RegExp(CANCEL_RE.source, 'gi'), ' ').replace(/\b(she|he|they|rider|riders|has|have|just|ride|the|is|was)\b/gi, ' ');

  const text = mentionNames_(m).reduce((t, n) => t.split('@' + n).join(' ').split(n).join(' '), String(m.text)).replace(/@\S+/g, ' ');
  const parts = text.split(/[,;.\n!]|\bbut\b/i);
  let targets = byRiderName_(clean(parts.filter(isCancelMsg_).join(' ')), open);
  if (!targets.length) {
    targets = byRoute_(text, open);
    if (targets.length > 1) return { ok: false, result: 'Which one cancelled? ' + targets.map(r => `${r.first} (${r.pickup} → ${r.dropoff})`).join(', ') };
  }
  if (!targets.length) targets = byRiderName_(clean(parts.filter(x => !parseUpdate_(x)).join(' ')), open);
  let context = '';
  if (!targets.length) {
    const w = JSON.parse(SP_().getProperty('SKIP_WAITS') || '[]').concat(JSON.parse(SP_().getProperty('PICK_ASKS') || '[]'))
      .filter(x => x.who === m.name && msgMs - x.askMs < 20 * 60e3).sort((a, b) => a.askMs - b.askMs).pop();
    if (w) { targets = open.filter(r => w.ids.includes(r.id)); if (targets.length && w.dueMs) context = '2 min'; }
  }
  if (!targets.length) {
    for (const e of recentFrom_(m)) {
      targets = byRiderName_(clean(e.text), open);
      if (targets.length) { context = e.text; break; }
    }
  }
  if (!targets.length) {
    if (open.length > 1) return { ok: false, result: 'Which one cancelled? ' + open.map(r => `${r.first} (${r.pickup})`).join(', ') };
    targets = open;
  }
  let reason = cancelReason_(m.text + ' ' + context);
  if (reason === 'Other' && toldSkip_(m.name, msgMs)) reason = 'Exceeded 2min Wait';
  const done = [];
  try {
    targets.forEach(r => { cancelRow_(sh, r.row, reason); r.cancelled = true; done.push(r); });
  } catch (e) { return { ok: false, rows: done.map(r => r.row), ids: done.map(r => r.id), result: e.message }; }
  done.forEach(r => { const i = rides.indexOf(r); if (i >= 0) rides.splice(i, 1); });
  return { ok: true, cancelled: true, rows: done.map(r => r.row), ids: done.map(r => r.id),
           result: `Cancelled (${reason}): ${done.map(r => r.first).join(', ')}` };
}

const SKIP_ASK_RE = /\?|^\s*(can|could|should|may|shall|do)\s+(we|i)\b|\b(ok|okay|good|fine) to\b/i;
const SKIP_WAIT_MS = 2 * 60e3;
function skipAsk_(rides, open, m) {
  const msgMs = m.created_at * 1000;
  if (!open.length || Date.now() - msgMs > 10 * 60e3)
    return { ok: false, result: 'Asks about skipping — answer by hand' };
  const clean = String(m.text).replace(new RegExp(CANCEL_RE.source, 'gi'), ' ');
  let targets = byRiderName_(clean, open);
  if (!targets.length) targets = open.filter(r => r.arrived);
  if (!targets.length) targets = open;
  const arrivedMs = Math.max(0, ...targets.map(r => clockMs_(r.arrived) || 0));
  const startMs = arrivedMs && arrivedMs <= msgMs ? arrivedMs : msgMs, dueMs = startMs + SKIP_WAIT_MS;
  const tag = '@' + m.name, mention = { id: String(m.user_id), start: 0, len: tag.length }, reply = { id: String(m.id) };
  const names = targets.map(r => r.first).join(', ');
  if (dueMs - msgMs < 30e3)
    return { ok: true, result: `Asked to skip ${names} — waited 2 min, told them to skip`,
             replies: [{ text: `${tag} you can skip`, mentions: [mention], reply }] };
  const left = Math.ceil((dueMs - msgMs) / 60e3);
  return { ok: true, result: `Asked to skip ${names} — told them to wait ${left} min, then skip`,
           replies: [{ text: `${tag} wait ${left === 2 ? '2 min' : left + ' more min'}`, mentions: [mention], reply }],
           wait: { askId: String(m.id), who: m.name, uid: String(m.user_id), ids: targets.map(r => r.id), askMs: msgMs, dueMs } };
}

function flushSkipWaits_() {
  const p = SP_(), raw = p.getProperty('SKIP_WAITS');
  if (!raw || !JSON.parse(raw).some(w => !w.told && w.dueMs <= Date.now())) return;
  const locked = tryWithLock_(5000, () => {
    const list = JSON.parse(p.getProperty('SKIP_WAITS') || '[]'), now = Date.now(), out = [];
    const due = list.filter(w => !w.told && w.dueMs <= now);
    if (!due.length) return out;
    const status = tonight_().rides, mine = loadChat_().filter(e => e.me);
    due.forEach(w => {
      w.told = now;
      const stillOpen = w.ids.some(id => status.some(r => r.id === id && r.driver && !r.picked));
      const auto = e => /^@.+ (wait (\d+ more )?\d* ?min|you can skip)/.test(String(e.text));
      const answered = mine.some(e => e.ts * 1000 > w.askMs && !auto(e) &&
        (String(e.text).toLowerCase().includes('@' + String(w.who).toLowerCase().split(' ')[0]) || (e.reply && e.reply.id === w.askId)));
      if (stillOpen && !answered && now - w.dueMs < 5 * 60e3) out.push(w);
    });
    p.setProperty('SKIP_WAITS', JSON.stringify(list.filter(w => now - w.askMs < 60 * 60e3)));
    return out;
  });
  if (locked.busy) return;
  const send = locked.value;
  send.forEach(w => {
    const tag = '@' + w.who;
    try { postMessage_(`${tag} you can skip now`, [{ id: w.uid, start: 0, len: tag.length }], { id: w.askId }); }
    catch (e) { console.error('skip follow-up', e); }
  });
}

function byRoute_(text, rides) {
  const t = ' ' + String(text).toLowerCase().replace(/[^a-z0-9]+/g, ' ') + ' ';
  const named = (raw, canon) => [raw, canon].concat(PLACES[canon] || []).map(x => String(x || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim())
    .some(x => x.length >= 3 && t.includes(' ' + x + ' '));
  const at = rides.filter(r => named(r.pickup, r.pickupName));
  if (at.length <= 1) return at;
  const to = at.filter(r => named(r.dropoff, canonPlace_(r.dropoff)));
  return to.length ? to : at;
}

function toldSkip_(who, ms) {
  return JSON.parse(SP_().getProperty('SKIP_WAITS') || '[]').some(w => w.who === who && w.told && ms - w.told < 15 * 60e3 && ms >= w.told - 60e3);
}

function parseUpdate_(text) {

  text = String(text).replace(/\b(got|gotten)\s+(to|here|there)\b|\bpulled\s+up(\s+to)?\b/gi, ' at ');
  const w = msgWords_(text).flatMap(x => (x === 'p/u' ? [x] : x.split('/'))).filter(Boolean);
  while (w.length && LEAD_FILLER.has(w[0])) w.shift();
  if (!w.length) return null;
  let kind = isPickedWord_(w[0]) ? 'picked' : isDroppedWord_(w[0]) ? 'dropped' : isArrivedWord_(w[0]) ? 'at' : null;
  if (!kind) return null;
  let rest = w.slice(1), both = false;
  if (kind === 'at') {
    const i = rest[0] === 'and' || rest[0] === '&' ? 1 : 0;
    if (rest[i] && isPickedWord_(rest[i])) { kind = 'picked'; both = true; rest = rest.slice(i + 1); }
  }
  if (kind === 'picked' && rest[0] === 'up') rest.shift();
  if (kind === 'dropped' && rest[0] === 'off') rest.shift();
  if (kind === 'at' && rest[0] === 'at') rest.shift();
  while (rest.length && (isEmptyWord_(rest[rest.length - 1]) || rest[rest.length - 1] === 'now')) rest.pop();
  const place = rest.join(' ');
  return { kind, place: GROUP_WORDS.has(place) ? '' : place, group: GROUP_WORDS.has(place), both };
}

const NAME_STOP_ = new Set(['the', 'and', 'all', 'both', 'them', 'him', 'her', 'his', 'she', 'they', 'up', 'off', 'now', 'here', 'there',
  'with', 'from', 'for', 'one', 'two', 'three', 'four', 'last', 'first', 'next', 'rider', 'riders', 'girl', 'guy', 'just', 'got', 'done']);
function byRiderName_(text, rides) {
  const words = String(text).toLowerCase().split(/[^a-z]+/).filter(w => w.length >= 3 && !NAME_STOP_.has(w));
  const same = (w, name) => {
    const n = norm_(String(name || '').split(/\s+/)[0]);
    if (n.length < 3) return false;
    if (w === n) return true;
    return w.length >= 4 && n.length >= 4 && (n.startsWith(w) || w.startsWith(n));
  };
  return words.length ? rides.filter(r => words.some(w => same(w, r.first) || same(w, r.last))) : [];
}

function placeList_(place, rides) {
  const words = String(place).toLowerCase().split(/\s+/).filter(w => w && !['and', '&', '+', 'then', 'n'].includes(w));
  const at = seg => {
    const k = placeKey_(seg), toks = placeTokens_(seg), exact = rides.filter(r => r.pickupKey === k);
    return exact.length ? exact
      : rides.filter(r => toks.length && toks.every(w => placeTokens_(r.pickup + ' ' + (r.pickupName || '')).some(p => p.startsWith(w) || w.startsWith(p))));
  };
  const found = [];
  let parts = 0;
  for (let i = 0; i < words.length;) {
    let j = words.length, hit = [];
    for (; j > i; j--) { hit = at(words.slice(i, j).join(' ')); if (hit.length) break; }
    if (!hit.length) { i++; continue; }
    parts++; hit.forEach(r => { if (!found.includes(r)) found.push(r); });
    i = j;
  }
  return parts >= 2 && new Set(found.map(r => r.pickupKey)).size >= 2 ? found : null;
}

function tagGroups_(rides, newMsgs) {
  const batches = loadBatches_();
  let posted = {};
  try { posted = Object.assign({}, recentlyPosted_()); } catch (e) {   }
  (newMsgs || []).forEach(m => (m.text || '').replace(/\d[\d\s().-]{8,}\d/g, s => {
    const d = s.replace(/\D/g, '').slice(-10);
    if (d.length === 10) posted[d] = { id: m.id };
  }));
  rides.forEach(r => {
    const ph = String(r.phone).replace(/\D/g, '').slice(-10);
    r.group = batches[r.id] || (posted[ph] && posted[ph].id ? 'gm' + posted[ph].id : null);
  });
}

function mentionNames_(m) {
  const text = String(m.text || ''), out = [];
  (m.attachments || []).filter(a => a.type === 'mentions').forEach(a =>
    (a.loci || []).forEach(([s, l]) => out.push(text.substr(s, l).replace(/^@/, '').trim())));
  if (!out.length) (text.match(/@[A-Za-z][\w'-]*(?:\s+[A-Z][\w'-]*)?/g) || []).forEach(t => out.push(t.slice(1).trim()));
  return out.filter(Boolean);
}

const withoutRemoved_ = text => String(text || '').replace(/^REMOVED\b[^\n]*(\n[^\n]*\S[^\n]*)*/gim, '');
function applyMyPost_(sh, rides, crews, m) {
  const phones = new Set(phonesIn_(withoutRemoved_(m.text)));
  const targets = rides.filter(r => !r.driver && !r.picked && phones.has(String(r.phone).replace(/\D/g, '').slice(-10)) &&
    r.tsMs <= m.created_at * 1000 + 60e3);
  if (!targets.length) return null;
  const names = targets.map(r => r.first).join(', '), tagged = mentionNames_(m);
  const matches = crews.filter(c => c.state !== 'off' && tagged.some(n => [c.driver, c.nav].filter(Boolean).some(x => sameName_(x, n))));
  if (matches.length !== 1) {
    return { ok: false, ids: targets.map(r => r.id),
      result: `Your post for ${names}: ${matches.length ? `"@${tagged.join(', @')}" matches more than one crew` : `couldn't tell which crew "@${tagged.join(', @') || '?'}" is`} — assign them in the panel` };
  }
  const crew = matches[0], when = fmtTime_(new Date(m.created_at * 1000));
  const dispatcher = dispatcherName_();
  const done = [];
  targets.forEach(r => {
    try {
      assertEditable_(sh, r.row);
      const cur = sh.getRange(r.row, 1, 1, CFG.LAST_COL).getDisplayValues()[0];
      if (String(cur[COL.DRIVER - 1]).trim() || hasReason_(cur[COL.REASON - 1])) return;
      if (!String(cur[COL.DISPATCH - 1]).trim() && dispatcher) sh.getRange(r.row, COL.DISPATCH).setValue(dispatcher);
      if (!String(cur[COL.CALL - 1]).trim()) sh.getRange(r.row, COL.CALL).setValue(when);
      sh.getRange(r.row, COL.LOC, 1, 4).setValues([[r.zone || zoneOf_(r.dropoff) || 'Off', crew.van || '', crew.driver, crew.nav || '']]);
      clearInheritedColor_(sh, r.row);
      Object.assign(r, { driver: crew.driver, nav: crew.nav || '' });
      done.push(r);
    } catch (e) {   }
  });
  if (!done.length) return null;
  saveBatch_(done.map(r => r.id));
  ledgerAdd_(done.map(r => ({ id: r.id, row: r.row, first: r.first, phone: r.phone, pickup: r.pickup, dropoff: r.dropoff, time: r.time,
    driver: crew.driver, nav: crew.nav || '', van: crew.van || '', loc: r.zone || zoneOf_(r.dropoff) || 'Off' })));
  return { ok: true, ids: done.map(r => r.id),
    result: `Saved your post to the sheet: ${done.map(r => r.first).join(', ')} → ${crew.driver}${crew.nav ? ' / ' + crew.nav : ''}` };
}

function nightMinutes_(d) {
  const t = Number(Utilities.formatDate(d, CFG.TZ, 'H')) * 60 + Number(Utilities.formatDate(d, CFG.TZ, 'm'));
  return t < 12 * 60 ? t + 1440 : t;
}
function nightMinutesOf_(hhmm) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(hhmm).trim());
  if (!m) return -1;
  const t = Number(m[1]) * 60 + Number(m[2]);
  return t < 12 * 60 ? t + 1440 : t;
}

function placeTypos_(place, r) {
  const p = placeNorm_(place);
  const names = [r.pickup, r.pickupName].concat(PLACES[r.pickupName] || []);
  return Math.min(...names.map(n => placeNorm_(n)).filter(Boolean).map(q => typoDist_(p, q)));
}

function pickOnEmpty_(sh, rides, m) {
  const msgMs = m.created_at * 1000;
  const waiting = rides.filter(r => r.driver && !r.picked && r.arrived && r.tsMs < msgMs && [r.driver, r.nav].some(n => sameName_(n, m.name)));
  const forgot = forgottenOnEmpty_(rides, m);
  if (!waiting.length) return forgot;
  const done = [];
  try {
    waiting.forEach(r => {
      const atMs = clockMs_(r.arrived);
      if (!atMs) return;
      const when = Math.min(atMs + PICK_GAP_MIN * 60e3, msgMs);
      pick_(sh, r.row, when, false);
      r.picked = true;
      done.push(r);
    });
  } catch (e) { return { ok: false, rows: done.map(r => r.row), ids: done.map(r => r.id), result: e.message }; }
  if (!done.length) return forgot;
  const res = { ok: true, rows: done.map(r => r.row), ids: done.map(r => r.id),
                result: 'Said empty without "picked" — picked 2 min after arrival: ' + done.map(r => r.first).join(', ') };
  return forgot ? Object.assign(forgot, { rows: res.rows.concat(forgot.rows || []), ids: res.ids.concat(forgot.ids || []), result: res.result + ' · ' + forgot.result }) : res;
}

const FORGOT_MIN = 8;
function forgottenOnEmpty_(rides, m) {
  const msgMs = m.created_at * 1000;
  const mine = rides.filter(r => r.driver && r.tsMs < msgMs && [r.driver, r.nav].some(n => sameName_(n, m.name)));

  const sentAt = {};
  let posts = [];
  try { posts = recentMessages_().map(x => ({ ts: x.ts, text: x.text })); } catch (e) {   }
  posts.concat(BATCH_MSGS_).filter(x => x.ts * 1000 < msgMs).forEach(x => phonesIn_(x.text).forEach(ph => {
    sentAt[ph] = Math.min(sentAt[ph] || Infinity, x.ts * 1000);
  }));
  const sent = r => sentAt[String(r.phone).replace(/\D/g, '').slice(-10)] || 0;
  const forgot = mine.filter(r => !r.picked && !r.arrived && sent(r) &&
    (msgMs - sent(r) >= FORGOT_MIN * 60e3 || mine.some(x => x !== r && x.picked && sent(x) && Math.abs(sent(x) - sent(r)) < 60e3)));
  if (!forgot.length) return null;

  const and = a => (a.length > 1 ? a.slice(0, -1).join(', ') + ' and ' + a[a.length - 1] : a[0]);
  const byPlace = {};
  forgot.forEach(r => { const p = r.pickupName || r.pickup; (byPlace[p] = byPlace[p] || []).push(r.first); });
  const list = and(Object.entries(byPlace).map(([p, names]) => `${and(names)} at ${p}`));
  const was = forgot.length > 1 ? 'were' : 'was';
  if (!autopilotOn_() || Date.now() - msgMs > 10 * 60e3)
    return { ok: false, ids: forgot.map(r => r.id), result: `Said empty, but ${list} ${was} never marked at or picked — ask them?` };
  const tag = '@' + m.name;
  return { ok: true, ids: forgot.map(r => r.id), result: `Said empty, but ${list} ${was} never marked picked — asked them`,
           replies: [{ text: `${tag} did you pick up ${list}?`, mentions: [{ id: String(m.user_id), start: 0, len: tag.length }], reply: { id: String(m.id) } }],
           pickAsk: { who: m.name, ids: forgot.map(r => r.id), askMs: msgMs } };
}

const YES_RE = /^\s*(y|ya|yah|yes+|yeah|yea|yep|yup|yessir|correct|we did|did|i did|got (them|him|her|em|'em)|picked( (them|him|her) up)?|mhm)\b/i;
const NO_RE = /^\s*(n|no+|nope|nah|not yet|didn'?t|did not|haven'?t|have not|on (our|the) way|omw|heading|going)\b/i;
function pickAnswer_(sh, rides, m) {
  const p = SP_(), msgMs = m.created_at * 1000, asks = JSON.parse(p.getProperty('PICK_ASKS') || '[]');
  const a = asks.filter(x => x.who === m.name && !x.done && msgMs > x.askMs && msgMs - x.askMs < 15 * 60e3).pop();
  if (!a) return null;
  const text = String(m.text || ''), yes = YES_RE.test(text), no = !yes && NO_RE.test(text);
  if (!yes && !no) return null;
  a.done = true;
  p.setProperty('PICK_ASKS', JSON.stringify(asks));
  const open = rides.filter(r => a.ids.includes(r.id) && !r.picked);
  const names = open.map(r => r.first).join(', ');
  if (!open.length) return { ok: true, result: 'Already done in the sheet' };
  if (no) return { ok: true, ids: open.map(r => r.id), result: `Not picked yet: ${names} (still with them)` };
  const when = extractWhen_(text, msgMs);
  const done = [];
  try {
    open.forEach(r => {
      const callMs = clockMs_(r.call) || r.tsMs;

      const ms = when.said ? when.ms : Math.max(callMs + 3 * 60e3, Math.min(callMs + 6 * 60e3, a.askMs - 6 * 60e3));
      pick_(sh, r.row, ms, false);
      r.picked = true; done.push({ r, ms });
    });
  } catch (e) { return { ok: false, ids: open.map(r => r.id), result: e.message }; }
  const times = done.map(d => `${d.r.first} ${fmtTime_(new Date(d.ms - PICK_GAP_MIN * 60e3))}`).join(', ');
  return when.said
    ? { ok: true, rows: done.map(d => d.r.row), ids: done.map(d => d.r.id), result: `Picked (they said ${fmtTime_(new Date(when.ms))}): ${names}` }
    : { ok: false, rows: done.map(d => d.r.row), ids: done.map(d => d.r.id), result: `Picked ${names} — arrival estimated (${times}); fix in the sheet if you know better` };
}

function dropRiders_(rides, m, u) {
  const msgMs = m.created_at * 1000, dropped = loadDroppedAt_();
  const inCar = rides.filter(r => r.driver && r.picked && !dropped[r.id] && r.tsMs < msgMs && [r.driver, r.nav].some(n => sameName_(n, m.name)));
  if (!inCar.length) return null;
  let hit = [];
  if (u.group) hit = inCar;
  else if (u.place) {
    const k = placeKey_(u.place);
    hit = inCar.filter(r => r.dropoffKey === k);
    if (!hit.length) {
      const scored = inCar.map(r => ({ r, s: placeScore_(u.place, r.dropoff + ' ' + (canonPlace_(r.dropoff) || '')) })), best = Math.max(...scored.map(x => x.s));
      if (best > 0) hit = scored.filter(x => x.s === best).map(x => x.r);
    }
    if (!hit.length) hit = byRiderName_(u.place, inCar);
  }
  if (!hit.length) return null;
  noteDropped_(hit.map(r => r.id), msgMs);
  return { ok: true, ids: hit.map(r => r.id), result: 'Dropped: ' + hit.map(r => `${r.first} (${r.dropoff})`).join(', ') };
}

const TIME_FILLER = /^(sorry|sry|srry|mb|my bad|oops|oh|ok|okay|it was|that was|was|actually|like|about|around|approx|we|i|it|that|lol|\s)*$/i;
let LAST_UPDATE_ = {};
function timeFix_(sh, rides, m) {
  const msgMs = m.created_at * 1000, when = extractWhen_(m.text, msgMs);
  if (!when.said || !TIME_FILLER.test(when.text.replace(/[^a-z ]/gi, ' ').replace(/\s+/g, ' ').trim())) return null;
  let last = LAST_UPDATE_[m.name];
  if (!last) {
    const e = loadChat_().filter(x => x.who === m.name && ['at', 'picked', 'atpicked'].includes(x.kind) && (x.rideIds || []).length &&
      x.ts * 1000 < msgMs && msgMs - x.ts * 1000 <= 10 * 60e3).pop();
    if (e) last = { kind: e.kind, ids: e.rideIds, ms: e.ts * 1000 };
  }
  if (!last || msgMs - last.ms > 10 * 60e3) return null;
  const targets = rides.filter(r => last.ids.includes(r.id));
  if (!targets.length) return null;
  const fixed = [];
  try {
    targets.forEach(r => {

      const derived = fmtTime_(clampToCall_(sh, r.row, new Date(last.ms - PICK_GAP_MIN * 60e3)));
      const arrive = last.kind === 'picked' ? (r.arrived && r.arrived !== derived ? null : when.ms - PICK_GAP_MIN * 60e3) : when.ms;
      if (arrive != null) {
        assertEditable_(sh, r.row);
        const t = fmtTime_(clampToCall_(sh, r.row, new Date(arrive)));
        sh.getRange(r.row, COL.ARRIVE).setValue(t);
        r.arrived = t;
      }
      if (last.kind !== 'at') notePicked_([r.id], when.ms);
      fixed.push(r);
    });
  } catch (e) { return { ok: false, result: e.message }; }
  delete LAST_UPDATE_[m.name];
  return { ok: true, rows: fixed.map(r => r.row), ids: fixed.map(r => r.id),
           result: `Time fixed (said ${fmtTime_(new Date(when.ms))}): ${fixed.map(r => `${r.first}${r.arrived ? ' arrival ' + r.arrived : ''}`).join(', ')}` };
}

function applyMessage_(sh, rides, m) {
  const answer = pickAnswer_(sh, rides, m);
  if (answer) return answer;
  const fix = timeFix_(sh, rides, m);
  if (fix) return fix;
  if (!isCancelMsg_(m.text)) return applyUpdate_(sh, rides, m);

  const c = cancelByMessage_(sh, rides, m);

  const rest = [m.text].concat(String(m.text).split(/[,;.\n!]|\bbut\b/i).filter(x => !isCancelMsg_(x)).map(x => x.trim()))
    .find(x => { const u = parseUpdate_(extractWhen_(x, m.created_at * 1000).text); return u && u.kind !== 'dropped'; });
  const o = rest ? applyUpdate_(sh, rides, Object.assign({}, m, { text: rest })) : null;
  if (!o) return c;
  return { ok: c.ok && o.ok, cancelled: c.cancelled, replies: c.replies, wait: c.wait, rows: (c.rows || []).concat(o.rows || []), ids: (c.ids || []).concat(o.ids || []), result: c.result + ' · ' + o.result };
}
function applyUpdate_(sh, rides, m) {
  const when = extractWhen_(m.text, m.created_at * 1000);
  const u = parseUpdate_(when.text);
  if ((!u || u.kind === 'dropped') && isEmptyMsg_(m.text)) return pickOnEmpty_(sh, rides, m);
  if (u && u.kind === 'dropped') return dropRiders_(rides, m, u);
  if (!u) return null;
  const kind = u.kind, place = u.place;

  const msgMs = m.created_at * 1000;
  const msgMin = nightMinutes_(new Date(msgMs));
  const isCrew = r => r.driver && r.tsMs < msgMs && [r.driver, r.nav].some(n => sameName_(n, m.name)) &&
    !(r.call && nightMinutesOf_(r.call) > msgMin + 1);
  const names = list => list.map(r => r.first).join(', ');
  const mine = rides.filter(isCrew), open = mine.filter(r => !r.picked);
  const tooOld = (rides.olderActive || []).filter(isCrew);
  const oldMsg = list => `${names(list)} requested over ${CFG.EDIT_WINDOW_MIN} min ago — update by hand (the app doesn't edit older rows)`;
  if (!mine.length) return { ok: false, result: tooOld.length ? oldMsg(tooOld) : 'No ride assigned to this crew in the sheet yet' };

  const byPlace = list => {
    if (!list.length) return [];
    const hit = list.filter(r => r.pickupKey === placeKey_(place));
    if (hit.length) return hit;
    const scored = list.map(r => ({ r, s: placeScore_(place, r.pickup + ' ' + (r.pickupName || '')) }));
    const best = Math.max(...scored.map(x => x.s));
    if (best > 0) return scored.filter(x => x.s === best).map(x => x.r);

    const near = list.map(r => ({ r, d: placeTypos_(place, r) }));
    const dmin = Math.min(...near.map(x => x.d)), p = placeNorm_(place);
    const allowed = p.length <= 5 ? 1 : p.length <= 9 ? 2 : 3;
    return p.length >= 3 && dmin <= allowed ? near.filter(x => x.d === dmin).map(x => x.r) : [];
  };
  let cands, named = false;
  if (place) {
    cands = byPlace(open);
    if (!cands.length) cands = byPlace(mine);
    if (!cands.length) {
      cands = byRiderName_(place, open);
      if (!cands.length) cands = byRiderName_(place, mine);
      named = cands.length > 0;
    }
    if (!cands.length) {
      const old = byPlace(tooOld);
      return { ok: false, result: old.length ? oldMsg(old) : `No crew ride picking up at "${place}"` };
    }
  } else {
    if (!open.length) return { ok: true, result: `Already done in the sheet: ${names(mine.slice(-4))}` };

    const gid = r => r.group || 'p:' + r.pickupKey;
    const oldest = open.reduce((a, b) => (a.row < b.row ? a : b));
    cands = open.filter(r => gid(r) === gid(oldest));
  }

  const todo = cands.filter(r => !r.picked && (kind === 'picked' || u.both || !r.arrived || when.said));
  if (!todo.length) return { ok: true, rows: cands.map(r => r.row), ids: cands.map(r => r.id), result: `Already done in the sheet: ${names(cands)}` };
  cands = todo;
  if (!u.group && !named && new Set(cands.map(r => r.pickupKey)).size > 1) {

    const listed = place && placeList_(place, cands);
    if (!listed) return { ok: false, result: 'Which one? ' + cands.map(r => `${r.first} (${r.pickup})`).join(', ') };
    cands = listed;
  }

  const ms = when.ms, conflicts = [];
  try {
    cands.forEach(r => {
      if (kind === 'at' || u.both) {
        const t = fmtTime_(clampToCall_(sh, r.row, new Date(ms)));

        if (when.said && r.arrived && r.arrived !== t) conflicts.push(`${r.first} already at ${r.arrived}`);
        arrive_(sh, r.row, t);
        r.arrived = r.arrived || t;
      }
      if (kind === 'picked') {
        pick_(sh, r.row, ms, u.both);
        r.picked = true;
      }
    });
  } catch (e) {
    return { ok: false, result: e.message };
  }
  const said = when.said ? ` (said ${fmtTime_(new Date(ms))})` : '';
  const result = `${kind === 'at' ? 'Arrived' : u.both ? 'At + picked' : 'Picked'}${said}: ${cands.map(r => r.first).join(', ')}`;
  return conflicts.length
    ? { ok: false, rows: cands.map(r => r.row), ids: cands.map(r => r.id), result: `${result} — ${conflicts.join(', ')}; message says ${fmtTime_(new Date(ms))}. Fix Arrival by hand if needed.` }
    : { ok: true, rows: cands.map(r => r.row), ids: cands.map(r => r.id), result };
}

const ACK_RE = /^(ok(ay)?|k+|kk|thanks?|thank you|thx|ty|bet|got it|copy( that)?|10-?4|on it|omw|yes|yeah|yep|no|nope|np|sure|lol|haha|done|sounds good|will do)[\s.!]*$/i;

function meId_() {
  const p = SP_();
  if ((!p.getProperty('GROUPME_ME_ID') || !p.getProperty('GROUPME_ME_NAME')) && p.getProperty('GROUPME_TOKEN') && !GM_OFFLINE_) {
    try { const me = gmFetch_('/users/me', {}); p.setProperty('GROUPME_ME_ID', String(me.user_id || me.id || '')); p.setProperty('GROUPME_ME_NAME', String(me.name || '')); }
    catch (e) {   }
  }
  return p.getProperty('GROUPME_ME_ID') || '';
}

function chatEntry_(m, meId) {
  const text = String(m.text || '');
  const ids = [].concat(meId || []).map(String);
  const e = { id: String(m.id), ts: m.created_at, who: m.name, text: text.slice(0, 200), me: ids.includes(String(m.user_id)) };
  const phones = (text.match(/\d[\d\s().-]{8,}\d/g) || []).map(s => s.replace(/\D/g, '').slice(-10)).filter(d => d.length === 10);
  const u = parseUpdate_(extractWhen_(text, (m.created_at || 0) * 1000).text);
  if (phones.length) Object.assign(e, { kind: 'assign', phones });
  else if (e.me) e.kind = 'mine';
  else if (m.sender_type !== 'user') e.kind = 'ack';
  else if (isCancelMsg_(text)) e.kind = 'cancel';
  else if (u) e.kind = u.kind === 'picked' && u.both ? 'atpicked' : u.kind;
  else if (isEmptyMsg_(text)) e.kind = 'empty';
  else if (!/[a-z0-9]/i.test(text) || ACK_RE.test(text.trim())) e.kind = 'ack';
  else e.kind = 'chat';
  if (e.kind === 'dropped' && isEmptyMsg_(text)) e.empty = true;
  return e;
}

function loadChat_() {
  const v = bigGet_('GM_MSGS');
  return v && v.night === nightKey_(new Date()) ? v.list : [];
}

function addChat_(entries) {
  if (!entries.length) return;
  const list = loadChat_(), have = new Set(list.map(e => e.id));
  entries.forEach(e => { if (!have.has(e.id)) list.push(e); });
  list.sort((a, b) => a.ts - b.ts);
  bigSet_('GM_MSGS', { night: nightKey_(new Date()), list: list.slice(-120) });
}

function dismissChat(ids) {
  const p = SP_(), d = JSON.parse(p.getProperty('GM_DISMISSED') || '[]');
  p.setProperty('GM_DISMISSED', JSON.stringify(d.concat([].concat(ids)).slice(-150)));
}

function chatView_(active, board) {
  const chat = loadChat_(), use = new Set(), byId = {};
  const dismissed = new Set(JSON.parse(SP_().getProperty('GM_DISMISSED') || '[]'));
  const rideChat = {}, crewChat = {};
  active.forEach(r => {
    const ph = String(r.phone).replace(/\D/g, '').slice(-10);

    const ids = chat.filter(c => (c.rideIds ? c.rideIds.includes(r.id) : (c.rows || []).includes(r.row)) ||
      (c.kind === 'assign' && c.phones.includes(ph) && c.ts * 1000 >= r.tsMs)).map(c => c.id);
    if (ids.length) { rideChat[r.id] = ids; ids.forEach(i => use.add(i)); }
  });
  board.forEach(c => {
    const ids = chat.filter(m => !m.me && m.kind !== 'assign' && [c.driver, c.nav].filter(Boolean).some(n => sameName_(n, m.who)))
      .slice(-6).map(m => m.id);
    if (ids.length) { crewChat[c.key] = ids; ids.forEach(i => use.add(i)); }
  });
  const now = Date.now() / 1000;
  const inbox = chat.filter(c => !dismissed.has(c.id) && now - c.ts < 3600 && (c.kind === 'chat' || c.status === 'ask'))
    .map(c => c.id).reverse();
  inbox.forEach(i => use.add(i));
  chat.forEach(c => { if (use.has(c.id)) byId[c.id] = Object.assign({}, c); });

  let recent = [], names = {};
  try { recent = recentMessages_(); } catch (e) { console.error('recentMessages_', e); }
  try { groupMembers_().forEach(x => { names[x.id] = x.name; }); } catch (e) {   }
  const rxById = {};
  recent.forEach(m => { rxById[m.id] = m.rx; });
  const named = rx => rx.map(r => ({ emoji: r.emoji, names: r.uids.map(u => names[u] || 'someone') }));
  Object.values(byId).forEach(e => { if (rxById[e.id]) e.rx = named(rxById[e.id]); });

  const rideAck = {};
  active.forEach(r => {
    const posts = (rideChat[r.id] || []).map(i => byId[i]).filter(e => e && e.kind === 'assign');
    if (!posts.length) return;
    const post = posts[posts.length - 1];
    const crewRx = (rxById[post.id] || []).map(x => ({ emoji: x.emoji, by: x.uids.map(u => names[u] || '')
      .filter(n => n && [r.driver, r.nav].filter(Boolean).some(c => sameName_(c, n))) })).filter(x => x.by.length);
    rideAck[r.id] = { postTs: post.ts, ack: crewRx.length > 0, emoji: crewRx.length ? crewRx[0].emoji : '',
                       by: [...new Set(crewRx.flatMap(x => x.by))] };
  });
  return { byId, rideChat, crewChat, inbox, rideAck };
}

function groupId_() {
  const p = SP_(), g = p.getProperty('GROUPME_GROUP_ID');
  if (!p.getProperty('GROUPME_TOKEN') || !g) throw new Error('Set up GroupMe first (⚙ GroupMe).');
  return g;
}

function gmPost_(path, body) {
  const resp = gmRequest_(`${GM.API}${path}`, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true,
    headers: { 'X-Access-Token': SP_().getProperty('GROUPME_TOKEN') },
  });
  const code = resp.getResponseCode();
  if (code === 401) throw new Error('GroupMe token rejected — copy a fresh one from web.groupme.com.');
  if (code >= 300) throw new Error(`GroupMe returned ${code}`);
  return JSON.parse(resp.getContentText() || '{}').response;
}

function groupMembers_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('gm_members');
  if (hit) return JSON.parse(hit);
  if (GM_OFFLINE_) { const keep = cache.get('gm_members_keep'); return keep ? JSON.parse(keep) : []; }
  const g = gmFetch_(`/groups/${groupId_()}`, {});
  const members = ((g && g.members) || []).map(x => ({ id: String(x.user_id), name: x.nickname || x.name, real: x.name || '' }))
    .sort((a, b) => a.name.localeCompare(b.name));
  cache.put('gm_members', JSON.stringify(members), 120);
  try { cache.put('gm_members_keep', JSON.stringify(members), 21600); } catch (e) {   }
  return members;
}

function getChatFeed() {
  const meId = meId_();
  const r = gmFetch_(`/groups/${groupId_()}/messages`, { limit: 40 });
  const msgs = ((r && r.messages) || []).slice().reverse().map(m => ({
    id: String(m.id), ts: m.created_at, who: m.name, me: String(m.user_id) === meId, system: !!m.system,
    text: m.text || '',
    mentions: (m.attachments || []).filter(a => a.type === 'mentions').flatMap(a => a.loci || []),
    mentionIds: (m.attachments || []).filter(a => a.type === 'mentions').flatMap(a => (a.user_ids || []).map(String)),
    images: (m.attachments || []).filter(a => a.type === 'image' && a.url).map(a => a.url),
    reply: replyOf_(m),
    rx: reactionsOf_(m),
  }));
  const members = groupMembers_(), names = {};
  members.forEach(x => { names[x.id] = x.name; });
  msgs.forEach(m => { m.rx = m.rx.map(r => ({ emoji: r.emoji, names: r.uids.map(u => names[u] || 'someone'), mine: r.uids.includes(meId) })); });

  const shown = new Set(msgs.map(m => m.id)), quotes = {};
  const need = msgs.filter(m => m.reply && !shown.has(m.reply.id)).map(m => m.reply.id);
  if (need.length) loadChat_().filter(e => need.includes(e.id)).forEach(e => { quotes[e.id] = { who: e.who, text: e.text }; });
  return { msgs, members, quotes };
}

function replyOf_(m) {
  const a = (m.attachments || []).find(x => x.type === 'reply' && x.reply_id);
  return a ? { id: String(a.reply_id), base: String(a.base_reply_id || a.reply_id) } : null;
}

function likeMessage(messageId, on) {
  gmPost_(`/messages/${groupId_()}/${messageId}/${on ? 'like' : 'unlike'}`, {});
  CacheService.getScriptCache().remove('recent');
}

function sendGroupMe(text, mentions, reply) {
  text = String(text || '').trim();
  const { messageId, replyDropped } = postMessage_(text, mentions, reply);

  const replaced = /^updated ride!/i.test(text) ? deleteSuperseded_(text, messageId) : null;
  return { sent: true, replaced, replyDropped, messageId };
}

function editMyMessage(oldId, text, mentions, reply) {
  text = String(text || '').trim();
  const r = gmFetch_(`/groups/${groupId_()}/messages`, { limit: 100 });
  const old = ((r && r.messages) || []).find(x => String(x.id) === String(oldId));
  if (!old) throw new Error("Couldn't find that message anymore — send a new one instead.");
  if (String(old.user_id) !== meId_()) throw new Error("That message isn't yours — you can only edit your own.");
  const { messageId, replyDropped } = postMessage_(text, mentions, reply);
  const del = deleteMyMessage_(oldId);
  return { sent: true, messageId, replyDropped, deleted: del.ok, note: del.ok ? '' : del.note };
}

function autoMentions_(text, ms) {
  let members = [];
  try { members = groupMembers_(); } catch (e) { return ms; }
  const out = ms.slice(), lower = String(text).toLowerCase();
  const free = (at, len) => !out.some(x => at < x.start + x.len && x.start < at + len);
  const add = (id, tag) => {
    for (let at = lower.indexOf(tag); at >= 0; at = lower.indexOf(tag, at + 1)) {
      const next = lower[at + tag.length], prev = at ? lower[at - 1] : ' ';
      if (!/[a-z0-9]/.test(prev) && (!next || !/[a-z0-9]/.test(next)) && free(at, tag.length)) out.push({ id: String(id), start: at, len: tag.length });
    }
  };
  members.slice().sort((a, b) => b.name.length - a.name.length).forEach(m => add(m.id, '@' + String(m.name).toLowerCase()));
  const firsts = {};
  members.forEach(m => { const f = String(m.name).toLowerCase().split(/\s+/)[0]; (firsts[f] = firsts[f] || []).push(m); });
  Object.keys(firsts).filter(f => f.length >= 2 && firsts[f].length === 1).forEach(f => add(firsts[f][0].id, '@' + f));

  for (const m of String(text).matchAll(/(^|[^a-z0-9])@([a-z][a-z'-]{2,})/gi)) {
    const at = m.index + m[1].length, word = m[2].toLowerCase(), len = word.length + 1;
    if (!free(at, len)) continue;
    const who = Object.keys(firsts).filter(f => f.length >= 3 && firsts[f].length === 1 && word.startsWith(f));
    if (who.length === 1) out.push({ id: String(firsts[who[0]][0].id), start: at, len });
  }
  return out.sort((a, b) => a.start - b.start);
}

function postMessage_(text, mentions, reply) {
  text = String(text || '').trim();
  if (!text) throw new Error('Type a message first.');
  if (text.length > 1000) throw new Error('GroupMe messages are limited to 1,000 characters.');
  const message = { source_guid: Utilities.getUuid(), text };
  const ms = autoMentions_(text, (mentions || []).filter(x => x.id && x.start >= 0 && x.len > 0 && x.start + x.len <= text.length));
  const atts = [];
  if (ms.length) atts.push({ type: 'mentions', user_ids: ms.map(x => String(x.id)), loci: ms.map(x => [x.start, x.len]) });
  const rep = reply && reply.id ? { type: 'reply', reply_id: String(reply.id), base_reply_id: String(reply.base || reply.id) } : null;
  if (rep) atts.push(rep);
  if (atts.length) message.attachments = atts;
  let sent, replyDropped = false;
  try {
    sent = gmPost_(`/groups/${groupId_()}/messages`, { message });
  } catch (e) {

    if (!rep || !/returned 4\d\d/.test(e.message) || /token/i.test(e.message)) throw e;
    message.source_guid = Utilities.getUuid();
    message.attachments = atts.filter(a => a !== rep);
    if (!message.attachments.length) delete message.attachments;
    sent = gmPost_(`/groups/${groupId_()}/messages`, { message });
    replyDropped = true;
  }
  CacheService.getScriptCache().remove('recent');
  return { messageId: sent && sent.message && String(sent.message.id), replyDropped };
}

function postToGroup_(text, mentions) {
  const message = { source_guid: Utilities.getUuid(), text: String(text) };
  const ms = (mentions || []).filter(x => x.id && x.start >= 0 && x.len > 0);
  if (ms.length) message.attachments = [{ type: 'mentions', user_ids: ms.map(x => String(x.id)), loci: ms.map(x => [x.start, x.len]) }];
  gmPost_(`/groups/${groupId_()}/messages`, { message });
  CacheService.getScriptCache().remove('recent');
}

function deleteMyMessage_(id) {
  const group = groupId_(), me = meId_();
  const r = gmFetch_(`/groups/${group}/messages`, { limit: 100 });
  const m = ((r && r.messages) || []).find(x => String(x.id) === String(id));
  if (!m) return { ok: false, note: "Couldn't find that GroupMe post anymore — delete it by hand if it's still there." };
  if (String(m.user_id) !== me) return { ok: false, note: 'That GroupMe post isn\'t yours — not deleting it.' };
  const resp = gmRequest_(`${GM.API}/conversations/${group}/messages/${id}`, {
    method: 'delete', muteHttpExceptions: true, headers: { 'X-Access-Token': SP_().getProperty('GROUPME_TOKEN') } });
  const code = resp.getResponseCode();
  CacheService.getScriptCache().remove('recent');
  if (code >= 200 && code < 300) {
    bigSet_('GM_MSGS', { night: nightKey_(new Date()), list: loadChat_().filter(e => e.id !== String(id)) });
    return { ok: true };
  }
  return { ok: false, note: `GroupMe refused to delete the post (HTTP ${code}) — delete it by hand.` };
}

const phonesIn_ = text => (String(text || '').match(/\d[\d\s().-]{8,}\d/g) || [])
  .map(s => s.replace(/\D/g, '').slice(-10)).filter(d => d.length === 10);

function deleteSuperseded_(text, newId) {
  const phones = new Set(phonesIn_(text)), me = meId_(), group = groupId_();
  if (!phones.size || !me) return { deleted: 0, failed: 0, note: me ? '' : 'Unknown GroupMe user — nothing deleted.' };
  const r = gmFetch_(`/groups/${group}/messages`, { limit: 100 });
  const old = ((r && r.messages) || []).filter(m => String(m.user_id) === me && String(m.id) !== String(newId || '') &&
    Date.now() / 1000 - m.created_at < 3 * 3600 && (() => { const p = phonesIn_(m.text); return p.length && p.every(x => phones.has(x)); })());
  let deleted = 0, failed = 0, code = 0;
  old.forEach(m => {
    try {
      const resp = gmRequest_(`${GM.API}/conversations/${group}/messages/${m.id}`, {
        method: 'delete', muteHttpExceptions: true, headers: { 'X-Access-Token': SP_().getProperty('GROUPME_TOKEN') } });
      code = resp.getResponseCode();
      if (code >= 200 && code < 300) deleted++; else failed++;
    } catch (e) { failed++; }
  });
  const reposts = old.map(m => ({ text: m.text || '', mentions: (m.attachments || []).filter(a => a.type === 'mentions')
    .flatMap(a => (a.loci || []).map((l, i) => ({ id: String((a.user_ids || [])[i] || ''), start: l[0], len: l[1] }))) }));
  if (deleted) {
    const gone = new Set(old.map(m => String(m.id)));
    bigSet_('GM_MSGS', { night: nightKey_(new Date()), list: loadChat_().filter(e => !gone.has(e.id)) });
  }
  return { deleted, failed, reposts: deleted && !failed ? reposts : [], note: failed ? `GroupMe refused to delete ${failed} old message(s) (HTTP ${code || '?'}) — delete it by hand.` : '' };
}

;

const FILE_VERSION_CODE = '2026-11-07a';

const CFG = {
  SPREADSHEET_ID: '',
  SHEET_INDEX: 0,
  TZ: 'America/Chicago',
  LOOKBACK_HOURS: 14,
  EDIT_WINDOW_MIN: 60,
  DUP_WINDOW_MIN: 40,
  SCAN_ROWS: 600,
  LAST_COL: 23,
};

const COL = {
  TS: 1, DISPATCH: 2, CALL: 3, ARRIVE: 4, EID: 6, LAST: 7, FIRST: 8, PHONE: 9,
  PICKUP: 10, DROPOFF: 11, ADA: 12, LOC: 13, VAN: 14, DRIVER: 15, NAV: 16, REASON: 22,
};

const isCart_ = v => /cart/i.test(String(v || ''));
const hasReason_ = v => { const s = String(v == null ? '' : v).trim(); return !!s && s !== '0'; };

const LOCATIONS = ['West', 'On', 'Off', 'North'];
const VANS = ['Van', 'ADA', 'Sienna', 'Cart'];
const CANCEL_REASONS = [
  'Chose Alternative Way Home', 'Not Answering', 'Long Wait Time Expected', 'Before Start',
  'Not Ready', 'Not Within Pickup Drop-off Parameters', 'Exceeded 2min Wait', 'After Deadline',
  'Wrong Number', 'Not Grouped with Friends', 'Luggage etc.', 'Other',
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('SURE Walk Dispatch')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function sheet_() {

  if (typeof XL_ !== 'undefined' && XL_) {
    if (!XL_.snap) throw new Error('The Excel add-in did not send the sheet — reload the add-in.');
    return XL_.sheet || (XL_.sheet = new XlSheet_(XL_.snap));
  }
  return SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheets()[CFG.SHEET_INDEX];
}

let SP_CACHE_ = null;

const PER_USER_KEYS_ = new Set(['GROUPME_TOKEN', 'GROUPME_ME_ID', 'GROUPME_ME_NAME']);
let UP_CACHE_ = null;
function SP_() {
  const real = PropertiesService.getScriptProperties();

  const cache = () => SP_CACHE_ || (SP_CACHE_ = real.getProperties());
  const user = () => userProps_();
  const ucache = () => UP_CACHE_ || (UP_CACHE_ = user().getProperties());
  return {
    getProperty: k => { const m = PER_USER_KEYS_.has(k) ? ucache() : cache(); return k in m ? m[k] : null; },
    setProperty: (k, v) => {
      if (PER_USER_KEYS_.has(k)) { ucache()[k] = String(v); user().setProperty(k, String(v)); return; }
      cache()[k] = String(v); real.setProperty(k, String(v));
    },
    setProperties: obj => { Object.keys(obj).forEach(k => SP_().setProperty(k, obj[k])); },
    deleteProperty: k => {
      if (PER_USER_KEYS_.has(k)) { delete ucache()[k]; user().deleteProperty(k); return; }
      delete cache()[k]; real.deleteProperty(k);
    },
  };
}

function dispatcherIds_() {
  const ids = JSON.parse(SP_().getProperty('DISPATCHER_GM_IDS') || '[]'), me = meId_();
  return me && !ids.includes(me) ? ids.concat(me) : ids;
}

const BIG_CHUNK = 2000;
function bigGet_(key) {
  const p = SP_(), n = Number(p.getProperty(key + '#n') || 0);
  let s = '';
  for (let i = 0; i < n; i++) s += p.getProperty(key + '#' + i) || '';
  try { return s ? JSON.parse(s) : null; } catch (e) { return null; }
}
function bigSet_(key, obj) {
  const p = SP_(), s = JSON.stringify(obj), oldN = Number(p.getProperty(key + '#n') || 0), out = {};
  const n = Math.ceil(s.length / BIG_CHUNK);
  for (let i = 0; i < n; i++) out[key + '#' + i] = s.slice(i * BIG_CHUNK, (i + 1) * BIG_CHUNK);
  out[key + '#n'] = String(n);
  p.setProperties(out);
  for (let i = n; i < oldN; i++) p.deleteProperty(key + '#' + i);
}

function fmtTime_(date) {
  return Utilities.formatDate(date || new Date(), CFG.TZ, 'H:mm');
}

let LAST_RIDE_CACHE_ = null;
function lastRideRow_(sh) {
  if (LAST_RIDE_CACHE_) return LAST_RIDE_CACHE_;
  return (LAST_RIDE_CACHE_ = lastRideRowUncached_(sh));
}
function lastRideRowUncached_(sh) {

  const lastRow = sh.getLastRow();
  for (const start of [Math.max(1, lastRow - 999), 1]) {
    const a = sh.getRange(start, COL.TS, lastRow - start + 1, 1).getDisplayValues();
    for (let i = a.length - 1; i >= 0; i--) {
      if (start + i > 1 && /\d{1,2}:\d{2}/.test(a[i][0])) return start + i;
    }
    if (start === 1) break;
  }
  return 1;
}

function peek(lastKnown) {
  const sh = sheet_(), last = lastRideRow_(sh), first = Math.max(2, last - 59);
  const rng = sh.getRange(first, 1, last - first + 1, CFG.LAST_COL);
  const disp = rng.getDisplayValues(), bgs = rng.getBackgrounds();
  const out = { sig: sigOf_(last, disp, bgs), last, fresh: [] };
  if (lastKnown && last > lastKnown && last - lastKnown <= 60) {
    const vals = sh.getRange(lastKnown + 1, 1, last - lastKnown, CFG.LAST_COL).getValues();
    const d = disp.slice(-(last - lastKnown));
    out.fresh = vals.map((v, i) => newRide_(lastKnown + 1 + i, v, d[i])).filter(Boolean);
  }
  return out;
}

function newRide_(row, vals, disp) {
  const ts = vals[COL.TS - 1], get = col => String(disp[col - 1]).trim();
  if (!(ts instanceof Date) || hasReason_(get(COL.REASON)) || get(COL.DRIVER)) return null;
  return { row, id: rideId_(ts, get(COL.PHONE), get(COL.FIRST)), tsMs: ts.getTime(), time: fmtTime_(ts),
    first: get(COL.FIRST), last: get(COL.LAST), phone: get(COL.PHONE), pickup: get(COL.PICKUP), dropoff: get(COL.DROPOFF),
    ada: /yes/i.test(get(COL.ADA)), zone: zoneOf_(get(COL.DROPOFF)), pickupKey: placeKey_(get(COL.PICKUP)),
    dropoffKey: placeKey_(get(COL.DROPOFF)), pickupName: canonPlace_(get(COL.PICKUP)) || get(COL.PICKUP),
    arrived: '', call: '', picked: false, dupOf: null, posted: null };
}
function peekSig_(sh) {
  const last = lastRideRow_(sh), first = Math.max(2, last - 59);
  const rng = sh.getRange(first, 1, last - first + 1, CFG.LAST_COL);
  return sigOf_(last, rng.getDisplayValues(), rng.getBackgrounds());
}
function sigOf_(last, disp60, bgs60) {
  const raw = JSON.stringify([last, disp60, bgs60]);
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8));
}

function rgb_(hex) {
  const m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex || '');
  return m ? m.slice(1).map(h => parseInt(h, 16)) : null;
}

function isBlack_(hex) {
  const c = rgb_(hex);
  return !!c && Math.max(...c) <= 0x50 && Math.max(...c) - Math.min(...c) <= 40;
}

function isRed_(hex) {
  const c = rgb_(hex);
  if (!c || isBlack_(hex)) return false;
  const [r, g, b] = c, max = Math.max(r, g, b), min = Math.min(r, g, b);
  if (max !== r || max - min < 25 || (max - min) / max < 0.15) return false;
  const hue = ((g - b) / (max - min)) * 60;
  return hue >= -15 && hue <= 15;
}

function isColored_(hex) {
  const c = rgb_(hex);
  return !!c && !isBlack_(hex) && !isRed_(hex) && Math.max(...c) - Math.min(...c) >= 12;
}

function rowColor_(bgRow) {
  let black = 0, red = 0;
  const colors = {};
  bgRow.forEach(h => {
    h = String(h).toLowerCase();
    if (isBlack_(h)) black++;
    else if (isRed_(h)) red++;
    else if (isColored_(h)) colors[h] = (colors[h] || 0) + 1;
  });
  if (black >= 3) return { kind: 'black' };
  if (red >= 3) return { kind: 'red' };
  const top = Object.keys(colors).sort((a, b) => colors[b] - colors[a])[0];
  return top && colors[top] >= 3 ? { kind: 'color', color: top } : { kind: 'none' };
}

const NIGHT_ROLLOVER_MIN = 2 * 60 + 30;
function nightKey_(date) {
  return Utilities.formatDate(new Date(date.getTime() - NIGHT_ROLLOVER_MIN * 60e3), CFG.TZ, 'yyyy-MM-dd');
}

const PASTEL_OF = { '#ff9900': '#fce5cd', '#ffff00': '#fff2cc', '#00ff00': '#d9ead3', '#00ffff': '#b3f0ea',
                    '#4a86e8': '#c9daf8', '#0000ff': '#cfe2f3', '#9900ff': '#d9d2e9', '#ff00ff': '#ead1dc' };

function isPastel_(hex) {
  const c = rgb_(hex);
  return !!c && Math.min(...c) >= 140;
}
function pastelOf_(hex) {
  hex = String(hex).toLowerCase();
  if (PASTEL_OF[hex]) return PASTEL_OF[hex];
  const c = rgb_(hex);
  return c ? '#' + c.map(v => Math.round(v + (255 - v) * 0.75).toString(16).padStart(2, '0')).join('') : null;
}
function brightOf_(hex) {
  return Object.keys(PASTEL_OF).find(b => PASTEL_OF[b] === String(hex).toLowerCase()) || null;
}

function getColors_(key) {
  const v = JSON.parse(SP_().getProperty(key) || 'null');
  return v && v.night === nightKey_(new Date()) ? v : {};
}

function tonightColors_() {
  const o = getColors_('COLOR_OVERRIDE'), d = getColors_('COLOR_DETECTED');
  let A = o.A || d.A || null, B = o.B || d.B || null;
  if (!B && A) B = pastelOf_(A);
  if (!A && B) A = brightOf_(B);
  return { A, B, chosenA: o.A || null, chosenB: o.B || null, detectedA: d.A || null, detectedB: d.B || null };
}

function setTonightColor(hex, which) {
  const p = SP_();
  if (!hex) return p.deleteProperty('COLOR_OVERRIDE');
  if (!isColored_(hex)) throw new Error('Pick a real color (not white, gray, black = cancelled, or red = duplicate).');
  const o = getColors_('COLOR_OVERRIDE'), next = { night: nightKey_(new Date()), A: o.A || null, B: o.B || null };
  hex = hex.toLowerCase();
  if (which === 'B') next.B = hex;
  else { next.A = hex; next.B = pastelOf_(hex); }
  p.setProperty('COLOR_OVERRIDE', JSON.stringify(next));
}

let ROSTER_CACHE_ = null;
function roster_(sh) {
  if (!ROSTER_CACHE_) ROSTER_CACHE_ = { r: parseRoster_(sh, lastRideRow_(sh)) };
  return ROSTER_CACHE_.r;
}

const pairKey_ = (a, b) => [a, b].map(s => String(s || '').trim().toLowerCase()).sort().join('|');

function shadeForRow_(sh, row, whenMs) {
  const name = String(sh.getRange(row, COL.DISPATCH).getDisplayValue() || '').trim() || dispatcherName_();
  const letter = dispatcherShift_(roster_(sh), name);
  return letter ? (letter === 'B' ? 'B' : 'A') : shadeFor_(whenMs);
}
function dispatcherShift_(roster, name) {
  const words = String(name || '').toLowerCase().replace(/[^a-z .]/g, '').split(/[\s.]+/).filter(Boolean);
  if (!roster || !words.length) return '';
  const first = words[0], last = words.length > 1 ? words[words.length - 1] : '';
  const all = [];
  roster.shifts.forEach(s => (s.dispatchers || []).forEach(d => all.push({ d: String(d).toLowerCase().trim(), letter: shiftLetter_(s.name) })));

  let hits = all.filter(x => first.length > 1 ? x.d.split(/\s+/)[0] === first : x.d[0] === first);
  if (hits.length > 1 && last) {
    let members = [];
    try { members = JSON.parse(CacheService.getScriptCache().get('gm_members') || CacheService.getScriptCache().get('gm_members_keep') || '[]'); } catch (e) {   }
    const full = members.map(m => String(m.name || '').toLowerCase().split(/\s+/)).filter(w => w.length > 1 && w[w.length - 1].startsWith(last));
    hits = hits.filter(x => full.some(w => w[0] === x.d.split(/\s+/)[0]));
  }
  const letters = [...new Set(hits.map(x => x.letter))];
  return letters.length === 1 ? letters[0] : '';
}

const B_SHADE_FROM = 22 * 60 + 30;
function shadeFor_(whenMs) {
  const d = new Date(whenMs || Date.now());
  const t = Number(Utilities.formatDate(d, CFG.TZ, 'H')) * 60 + Number(Utilities.formatDate(d, CFG.TZ, 'm'));
  return (t < 12 * 60 ? t + 1440 : t) >= B_SHADE_FROM ? 'B' : 'A';
}

function samePair_(a, b) {
  const an = String(a.nav || '').trim(), bn = String(b.nav || '').trim();
  if (!an || !bn) return !an && !bn && sameName_(a.driver, b.driver);
  return (sameName_(a.driver, b.driver) && sameName_(an, bn)) || (sameName_(a.driver, bn) && sameName_(an, b.driver));
}

const WAIT_SAMPLE = 5;
function waitMinutes_(call, arrive) {
  const p = s => { const m = /^(\d{1,2}):(\d{2})$/.exec(String(s).trim()); return m ? Number(m[1]) * 60 + Number(m[2]) : null; };
  const c = p(call), a = p(arrive);
  if (c == null || a == null) return null;
  const w = (a - c + 1440) % 1440;
  return w <= 120 ? w : null;
}
function waitSummary_(waits) {
  const byNewest = waits.slice().sort((x, y) => y.row - x.row);
  const recent = byNewest.slice(0, WAIT_SAMPLE), before = byNewest.slice(WAIT_SAMPLE, WAIT_SAMPLE * 2);
  const avg = list => (list.length ? list.reduce((s, x) => s + x.wait, 0) / list.length : null);
  return { avg: avg(recent), prevAvg: before.length >= 3 ? avg(before) : null, n: recent.length, recent };
}

function vanFromLabel_(label) {
  const s = label.toLowerCase();
  if (s.includes('cart')) return 'Cart';
  if (s.includes('ada') || s.includes('voyager')) return 'ADA';
  if (s.includes('sienna')) return 'Sienna';
  return 'Van';
}

function capOfLabel_(label) {
  const m = /\bx\s*(\d{1,2})\s*$/i.exec(String(label).trim());
  return m ? Number(m[1]) : 0;
}

function vehicleName_(label) {
  return label.replace(/\s*x\d+\s*$/i, '').replace(/\s*:\s*/, ' ').replace(/\s+/g, ' ').trim();
}

function rosterDateKey_(s) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(String(s).trim());
  if (!m) return null;
  const y = m[3].length === 2 ? '20' + m[3] : m[3];
  return `${y}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
}

function parseRoster_(sh, lastRide) {
  const lastRow = sh.getLastRow();
  if (lastRow <= lastRide) return null;
  const width = Math.min(sh.getLastColumn(), 26);
  const v = sh.getRange(lastRide + 1, 1, lastRow - lastRide, width).getDisplayValues()
    .map(r => r.map(c => String(c).trim()));

  const starts = v.map((r, i) => (rosterDateKey_(r[0]) ? i : -1)).filter(i => i >= 0);
  if (!starts.length) return null;
  const night = nightKey_(new Date());

  const dated = starts.map(i => ({ i, d: rosterDateKey_(v[i][0]) }));
  const pastOrToday = dated.filter(x => x.d <= night).sort((a, b) => (a.d < b.d ? -1 : a.d > b.d ? 1 : a.i - b.i));
  const si = (dated.find(x => x.d === night) || pastOrToday[pastOrToday.length - 1] || dated[dated.length - 1]).i;
  const block = v.slice(si, starts.find(i => i > si) ?? v.length);

  const hdr = block.findIndex(r => r.some(c => /^driver$/i.test(c)));
  if (hdr < 0) return null;
  const driverCols = block[hdr].map((c, j) => (/^driver$/i.test(c) ? j : -1)).filter(j => j >= 0);

  const labels = block.slice(0, hdr).reverse().find(r => r.some(c => /\bphone\b/i.test(c))) || [];
  const shifts = driverCols.map((dc, n) => ({

    name: (/shift/i.test(labels[dc] || '') ? String(labels[dc]).replace(/:\s*$/, '').trim() : '') ||
      (driverCols.length === 3 ? ['A Shift', 'M Shift', 'B Shift'][n] : n === 0 ? 'A Shift' : `Shift ${n + 1}`),
    crews: [],
    people: [],
    dispatchers: [],
  }));

  block.forEach((r, i) => {
    if (i !== hdr && /dispatcher/i.test(r[0] || ''))
      driverCols.forEach((dc, n) => [r[dc], r[dc + 1]].filter(x => /^[a-z][a-z .'-]*$/i.test(x || '')).forEach(x => shifts[n].dispatchers.push(x)));
  });
  block.slice(hdr + 1).forEach(r => {
    const label = r[0];
    if (/dispatcher/i.test(label || '')) {
      return;
    }
    if (!label) return;
    driverCols.forEach((dc, n) => [r[dc], r[dc + 1]].filter(x => /^[a-z][a-z .'-]*$/i.test(x || '') && !/^repair$/i.test(x)).forEach(x => shifts[n].people.push(x)));
    if (/floater/i.test(label)) return;
    driverCols.forEach((dc, n) => {

      const shifted = /^@\s*\d/.test(r[dc] || '');
      const driver = (shifted ? r[dc + 1] : r[dc]) || '', nav = (shifted ? r[dc + 2] : r[dc + 1]) || '';
      if (!driver && !nav) return;
      if (/repair/i.test(driver + ' ' + nav)) return;
      shifts[n].crews.push({
        driver: driver || nav, nav: driver ? nav : '',
        van: vanFromLabel_(label), vehicle: vehicleName_(label), cap: capOfLabel_(label),
        note: (shifted ? r[dc] : '') || r.slice(1).find(x => /^@\s*\d/.test(x)) || '',
        phone: (/^\s*-*\s*(\d{3,4})\s*$/.exec(r[dc + 2] || '') || [])[1] || '',
      });
    });
  });
  markMidMoves_(shifts);
  return { date: rosterDateKey_(block[0][0]), shifts };
}

function markMidMoves_(shifts) {

  shifts.forEach(sh => {
    sh.crews = sh.crews.reduce((out, c) => {
      const prev = out.find(x => samePair_(x, c));
      if (!prev) { out.push(c); return out; }
      if (!noteMinutes_(c.note) && !noteMinutes_(prev.note)) return out;
      const [base, later] = noteMinutes_(c.note) ? [prev, c] : [c, prev];
      out[out.indexOf(prev)] = Object.assign({}, base, { note: '', move: { at: noteMinutes_(later.note), van: later.van, vehicle: later.vehicle, cap: later.cap, phone: later.phone } });
      return out;
    }, []);
  });
  const m = shifts.find(s => shiftLetter_(s.name) === 'M'), b = shifts.find(s => shiftLetter_(s.name) === 'B');
  if (!m || !b) return;
  b.crews = b.crews.filter(bc => {
    const mc = m.crews.find(x => samePair_(x, bc));
    if (!mc) return true;
    mc.move = { at: noteMinutes_(bc.note) || 22 * 60 + 30, van: bc.van, vehicle: bc.vehicle, cap: bc.cap, phone: bc.phone };
    return false;
  });
}

function noteMinutes_(note) {
  const x = /^@\s*(\d{1,2})(?::(\d{2}))?\s*(am|pm|a|p)?/i.exec(String(note || '').trim());
  if (!x) return 0;
  let h = Number(x[1]) % 12;
  const ap = (x[3] || '').toLowerCase()[0];
  if (ap !== 'a' && h >= 6) h += 12;
  if (ap === 'a' || h < 6) h += 24;
  return h * 60 + Number(x[2] || 0);
}

const SHIFT_CHANGES = [{ at: 21 * 60, on: ['A', 'M'] }, { at: 22 * 60 + 30, on: ['M', 'B'] }, { at: 25 * 60, on: ['B'] }];
const SHIFT_DEFAULT = ['A'];

const SHIFT_LAST_ENDS = { B: 26 * 60 + 30 };

function lettersOnAt_(t) {
  let on = SHIFT_DEFAULT;
  SHIFT_CHANGES.forEach(c => { if (t >= c.at) on = c.on; });
  return on;
}

function offAt_(crew, rosterShifts) {
  const letter = shiftLetter_(crew.shift || ''), end = shiftEnds_()[letter];
  if (!end) return 0;
  const later = lettersOnAt_(end).filter(l => l !== letter);
  const continues = rosterShifts.some(s => later.includes(shiftLetter_(s.name)) && s.crews.some(x => samePair_(x, crew)));
  return continues ? 0 : end;
}

function shiftStarts_() {
  const starts = Object.fromEntries(SHIFT_DEFAULT.map(l => [l, 0]));
  SHIFT_CHANGES.forEach(c => c.on.forEach(l => { if (!(l in starts)) starts[l] = c.at; }));
  return starts;
}
function personOffAt_(name, rosterShifts) {
  const ends = shiftEnds_(), starts = shiftStarts_();
  const letters = [...new Set(rosterShifts.filter(s => (s.people || []).concat(s.crews.flatMap(c => [c.driver, c.nav]))
    .some(x => x && sameName_(x, name))).map(s => shiftLetter_(s.name)))].filter(l => l in ends && l in starts)
    .sort((a, b) => starts[a] - starts[b]);
  if (!letters.length) return 0;
  let end = ends[letters[0]];
  letters.slice(1).forEach(l => { if (starts[l] >= end) end = Math.max(end, ends[l]); });
  return end;
}

function offInfo_(crew, rosterShifts) {
  const names = [crew.driver, crew.nav].filter(Boolean);
  const per = names.map(n => ({ n, at: personOffAt_(n, rosterShifts) })).filter(x => x.at);
  if (!per.length) return { offAt: offAt_(crew, rosterShifts), offWho: '', offStays: '' };
  const at = Math.min(...per.map(x => x.at));
  const stays = per.filter(x => x.at > at).map(x => x.n);
  return { offAt: at, offWho: stays.length ? per.filter(x => x.at === at).map(x => x.n).join(' & ') : '', offStays: stays.join(' & ') };
}

function shiftEnds_() {
  const ends = {};
  let on = SHIFT_DEFAULT;
  SHIFT_CHANGES.forEach(c => { on.filter(l => !c.on.includes(l)).forEach(l => { ends[l] = c.at; }); on = c.on; });
  return Object.assign(ends, SHIFT_LAST_ENDS);
}

function shiftLetter_(name) {
  const m = /^([a-z])\b/i.exec(String(name).trim());
  return m ? m[1].toUpperCase() : '';
}

function autoShiftLetters_(now) {
  let t = Number(Utilities.formatDate(now, CFG.TZ, 'H')) * 60 + Number(Utilities.formatDate(now, CFG.TZ, 'm'));
  if (t < 12 * 60) t += 24 * 60;
  let on = SHIFT_DEFAULT;
  SHIFT_CHANGES.forEach(c => { if (t >= c.at) on = c.on; });
  return on;
}

function activeShiftNames_(roster) {
  if (!roster || !roster.shifts.length) return [];
  const auto = autoShiftLetters_(new Date());
  const v = JSON.parse(SP_().getProperty('ACTIVE_SHIFTS') || 'null');
  if (v && v.night === nightKey_(new Date()) && v.period === auto.join('')) {
    return roster.shifts.map(s => s.name).filter(n => v.names.includes(n));
  }
  const names = roster.shifts.filter(s => auto.includes(shiftLetter_(s.name))).map(s => s.name);
  return names.length ? names : [roster.shifts[0].name];
}

function toggleShift(name) {
  const sh = sheet_();
  const roster = parseRoster_(sh, lastRideRow_(sh));
  const on = new Set(activeShiftNames_(roster));
  on.has(name) ? on.delete(name) : on.add(name);
  SP_().setProperty('ACTIVE_SHIFTS', JSON.stringify({
    night: nightKey_(new Date()), period: autoShiftLetters_(new Date()).join(''), names: [...on],
  }));
}

function tonight_() {
  const sh = sheet_();
  const last = lastRideRow_(sh);
  const first = Math.max(2, last - CFG.SCAN_ROWS + 1);
  const n = last - first + 1;
  const range = sh.getRange(first, 1, n, CFG.LAST_COL);
  const vals = range.getValues(), disp = range.getDisplayValues(), bgs = range.getBackgrounds();
  const cutoff = Date.now() - CFG.LOOKBACK_HOURS * 3600e3;
  const editCutoff = Date.now() - CFG.EDIT_WINDOW_MIN * 60e3;
  const night = nightKey_(new Date());

  const rides = [], pickedTonight = [], olderActive = [], colorCounts = { A: {}, B: {} }, prior = [], rowInfo = {}, waits = [], statusById = {};

  vals.forEach((r, i) => {
    const ts = r[COL.TS - 1];
    if (!(ts instanceof Date) || ts.getTime() < cutoff) return;
    const get = c => String(disp[i][c - 1]).trim();
    const rid = rideId_(ts, get(COL.PHONE), get(COL.FIRST)), color = rowColor_(bgs[i]);
    statusById[rid] = { row: first + i, driver: get(COL.DRIVER), arrived: get(COL.ARRIVE), color: color.color || '',
      tonight: nightKey_(ts) === night,
      state: hasReason_(get(COL.REASON)) || color.kind === 'black' ? 'cancelled' : color.kind === 'red' ? 'dup' : isPickedRow_(color, get(COL.ARRIVE)) ? 'picked' : 'open' };
    if (hasReason_(get(COL.REASON))) return;
    if (color.kind === 'black' || color.kind === 'red') return;
    if (isPickedRow_(color, get(COL.ARRIVE)) && nightKey_(ts) === night) {
      const shade = colorCounts[isPastel_(color.color) ? 'B' : 'A'];
      shade[color.color] = (shade[color.color] || 0) + 1;
    }
    const ride = {
      row: first + i,
      id: rideId_(ts, get(COL.PHONE), get(COL.FIRST)),
      tsMs: ts.getTime(),
      time: fmtTime_(ts),
      first: get(COL.FIRST), last: get(COL.LAST), phone: get(COL.PHONE),
      pickup: get(COL.PICKUP), dropoff: get(COL.DROPOFF), ada: /yes/i.test(get(COL.ADA)),
      loc: get(COL.LOC), van: get(COL.VAN), driver: get(COL.DRIVER), nav: get(COL.NAV),
      arrived: get(COL.ARRIVE), call: get(COL.CALL),
      pickupKey: placeKey_(get(COL.PICKUP)), dropoffKey: placeKey_(get(COL.DROPOFF)),
      pickupName: canonPlace_(get(COL.PICKUP)) || get(COL.PICKUP),
      zone: zoneOf_(get(COL.DROPOFF)),
      picked: isPickedRow_(color, get(COL.ARRIVE)),
      dupOf: null,
    };

    const same = s => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
    const me = { route: same(ride.pickup) + '>' + same(ride.dropoff),
                 phone: ride.phone.replace(/\D/g, '').slice(-10), name: norm_(ride.first + ride.last) };
    const dup = me.phone.length === 10 && me.name.length > 3 ? prior.filter(q => q.ts >= ride.tsMs - CFG.DUP_WINDOW_MIN * 60e3 &&
      q.route === me.route && q.phone === me.phone && q.name === me.name).pop() : null;
    if (dup) ride.dupOf = { row: dup.row, first: dup.first, time: dup.time };
    prior.push(Object.assign({ ts: ride.tsMs, row: ride.row, first: ride.first, time: ride.time }, me));

    const w = waitMinutes_(get(COL.CALL), get(COL.ARRIVE));
    if (w != null && nightKey_(ts) === night) waits.push({ row: ride.row, first: ride.first, pickup: ride.pickup, wait: w, at: get(COL.ARRIVE) });
    if (ride.driver) {

      rowInfo[pairKey_(ride.driver, ride.nav)] =
        { driver: ride.driver, nav: ride.nav, van: ride.van, loc: ride.loc, lastTs: ride.tsMs };
    }

    if (ts.getTime() >= editCutoff || (!ride.picked && nightKey_(ts) === night)) rides.push(ride);
    if (ride.picked && ride.driver && nightKey_(ts) === night) pickedTonight.push(ride);
  });

  const roster = parseRoster_(sh, last);
  const shiftsOn = activeShiftNames_(roster);
  const crews = {};
  const nowMin = (d => { const t = Number(Utilities.formatDate(d, CFG.TZ, 'H')) * 60 + Number(Utilities.formatDate(d, CFG.TZ, 'm')); return t < 12 * 60 ? t + 1440 : t; })(new Date());
  (roster ? roster.shifts.filter(s => shiftsOn.includes(s.name)).flatMap(s => s.crews.map(c => Object.assign({ shift: s.name }, c,
    c.move && nowMin >= c.move.at ? { van: c.move.van, vehicle: c.move.vehicle, cap: c.move.cap, phone: c.move.phone || c.phone } : {}))) : []).forEach(c => {
    const k = (c.driver + '|' + c.nav).toLowerCase(), info = rowInfo[pairKey_(c.driver, c.nav)] || {};
    crews[k] = Object.assign({}, c, { loc: info.loc || '', lastTs: info.lastTs || 0 });
  });
  rides.filter(r => r.driver && !r.picked).forEach(r => {
    const k = (r.driver + '|' + r.nav).toLowerCase();
    if (!crews[k] && !Object.values(crews).some(c => samePair_(c, r))) crews[k] = Object.assign({}, rowInfo[pairKey_(r.driver, r.nav)]);
  });

  const top = m => Object.keys(m).sort((a, b) => m[b] - m[a])[0] || null;
  const det = { A: top(colorCounts.A), B: top(colorCounts.B) }, was = getColors_('COLOR_DETECTED');
  if ((det.A || det.B) && (det.A !== (was.A || null) || det.B !== (was.B || null))) {
    SP_().setProperty('COLOR_DETECTED', JSON.stringify({ night, A: det.A, B: det.B }));
  }

  const k60 = Math.min(60, last - Math.max(2, last - 59) + 1);
  return {
    rides,
    pickedTonight,
    olderActive,
    statusById,
    sig: sigOf_(last, disp.slice(-k60), bgs.slice(-k60)),
    lastRide: last,
    crews: Object.values(crews).sort((a, b) => a.driver.localeCompare(b.driver)),
    shifts: roster ? roster.shifts.map(s => s.name) : [],
    rosterShifts: roster ? roster.shifts.map(s => ({ name: s.name, people: s.people || [], crews: s.crews.map(c => ({ driver: c.driver, nav: c.nav })) })) : [],
    shiftsOn,
    wait: waitSummary_(waits),

    rosterInfo: { sheet: sh.getParent().getName() + ' › ' + sh.getName(), found: !!roster, date: roster ? roster.date : '',
                  counts: roster ? roster.shifts.map(s => ({ name: s.name, n: s.crews.length })) : [] },
    rosterDate: roster ? roster.date : '',
  };
}

function fileVersions_() {
  return {
    'Code.gs': FILE_VERSION_CODE,
    'GroupMe.gs': typeof FILE_VERSION_GROUPME !== 'undefined' ? FILE_VERSION_GROUPME : 'missing',
    'Crews.gs': typeof FILE_VERSION_CREWS !== 'undefined' ? FILE_VERSION_CREWS : 'missing',
    'Places.gs': typeof FILE_VERSION_PLACES !== 'undefined' ? FILE_VERSION_PLACES : 'missing',
  };
}

function getState(fast) {
  return Object.assign(getStateCore_(fast), { build: typeof BUILD_ID !== 'undefined' ? BUILD_ID : '' });
}
function getStateCore_(fast) {
  GM_OFFLINE_ = !!fast;

  const legacy = PropertiesService.getScriptProperties();
  if (legacy.getProperty('GROUPME_TOKEN')) ['GROUPME_TOKEN', 'GROUPME_ME_ID', 'GROUPME_ME_NAME'].forEach(k => legacy.deleteProperty(k));
  if (!fast) syncGroupMeQuietly_();
  const tonight = tonight_();
  const rides = tonight.rides;
  const { list: crews, removed: removedCrews } = applyCrewEdits_(tonight.crews);
  crews.forEach(c => Object.assign(c, offInfo_(c, tonight.rosterShifts || [])));
  const board = crewBoard_(crews, rides);
  let chat = { byId: {}, rideChat: {}, crewChat: {}, inbox: [], rideAck: {} };
  try { chat = chatView_(rides.filter(r => r.driver && !r.picked), board); } catch (e) { console.error('chatView_', e); }
  let posted = {};
  try { posted = recentlyPosted_(); } catch (e) {   }
  const pending = rides.filter(r => !r.driver);
  const active = rides.filter(r => r.driver && !r.picked).concat(ledgerExtras_(tonight.statusById), carrying_(tonight.pickedTonight || [], crews, board));
  let geo = {};
  try { geo = geoPlaces_(pending.flatMap(r => [r.pickup, r.dropoff]).concat(active.map(r => r.dropoff)), fast ? 3 : 8); }
  catch (e) {   }
  return {

    pending: pending.map(r => Object.assign(r, { posted: posted[String(r.phone).replace(/\D/g, '').slice(-10)] || null,
      pPos: geo[placeKey_(r.pickup)] || 0, dPos: geo[placeKey_(r.dropoff)] || 0,
      offPickup: offCampusPickup_(r.pickup),
      cartNoGo: cartNoGo_(r.pickup) || cartNoGo_(r.dropoff) })),

    active: active.map(r => Object.assign(r, { crewKey: (crews.find(c => samePair_(c, r)) || {}).key || (r.driver + '|' + r.nav).toLowerCase(),
      dPos: geo[placeKey_(r.dropoff)] || 0 })),
    crews,
    board,
    nameIssues: nameIssues_(board),
    chat,
    versions: fileVersions_(),
    diag: inProgressDiag_(tonight.statusById),
    sig: tonight.sig,
    lastRide: tonight.lastRide,
    wait: tonight.wait, shifts: tonight.shifts, shiftsOn: tonight.shiftsOn, rosterDate: tonight.rosterDate, rosterInfo: tonight.rosterInfo,
    removedCrews,
    locations: LOCATIONS, vans: VANS, reasons: CANCEL_REASONS,
    dispatcher: userProps_().getProperty('dispatcher') || '',
    dispatcherUsed: dispatcherName_(),
    prefs: JSON.parse(userProps_().getProperty('prefs') || 'null'),
    groupme: groupMeStatus_(),
    color: tonightColors_(),
    shiftEnds: shiftEnds_(),
  };
}

function act(name, args) {

  const actions = { assignRides, unassignRides, moveRides, markArrived, markPicked, cancelRide, markDuplicate, markNotAda, editRide, setCrewStatus, setCrewState, restoreRide,
    saveCrew, removeCrews, setTonightColor, toggleShift, dismissChat, setSyncRunning, clearGroupMeLog,
    saveGroupMeToken, setGroupMeGroup, forgetGroupMeToken, undoAction };
  if (!Object.prototype.hasOwnProperty.call(actions, name)) throw new Error(`Unknown action "${name}" — re-paste all files.`);
  const how = UNDOABLE_[name];
  let before = null;
  if (how) { try { before = undoSnap_(how, args || []); } catch (e) { before = null; } }
  const result = actions[name].apply(null, args || []);
  SP_CACHE_ = null; LAST_RIDE_CACHE_ = null; ROSTER_CACHE_ = null;
  let undo = null;
  if (before) { try { undo = undoTicket_(how, args || [], before); } catch (e) { undo = null; } }
  return { result, state: getState(true), undo };
}

function savePrefs(prefs) {
  const p = prefs || {};
  const clean = { theme: ['system', 'light', 'dark'].includes(p.theme) ? p.theme : 'system', mascot: p.mascot === 'bevo' ? 'bevo' : 'cat',
    accent: p.accent && /^#[0-9a-f]{6}$/i.test(p.accent.l) && /^#[0-9a-f]{6}$/i.test(p.accent.d)
      ? { name: String(p.accent.name || '').slice(0, 30), l: p.accent.l, d: p.accent.d } : null, t: Number(p.t) || Date.now(),
    alerts: { sound: !(p.alerts && p.alerts.sound === false), desktop: !!(p.alerts && p.alerts.desktop) } };
  if (!clean.accent) delete clean.accent;
  userProps_().setProperty('prefs', JSON.stringify(clean));
}

function dispatcherName_() {
  const saved = (userProps_().getProperty('dispatcher') || '').trim();
  if (saved) return saved;
  if (!SP_().getProperty('GROUPME_ME_NAME')) { try { meId_(); } catch (e) {   } }
  const gm = String(SP_().getProperty('GROUPME_ME_NAME') || '').trim().split(/\s+/).filter(Boolean);
  return gm.length > 1 ? `${gm[0][0].toUpperCase()}. ${gm[gm.length - 1]}` : (gm[0] || '');
}

function setDispatcher(name) {
  userProps_().setProperty('dispatcher', name.trim());
}

function assertEditable_(sh, row) {
  const ts = sh.getRange(row, COL.TS).getValue();
  if (!(ts instanceof Date)) throw new Error(`Row ${row} isn't a ride — not editing it.`);
  const kind = rowColor_(sh.getRange(row, 1, 1, CFG.LAST_COL).getBackgrounds()[0]).kind;
  if (kind === 'black' || kind === 'red') {
    throw new Error(`Row ${row} is marked ${kind === 'red' ? 'duplicate (red)' : 'cancelled (black)'} — not editing it.`);
  }
  const open = !(kind === 'color' && String(sh.getRange(row, COL.ARRIVE).getDisplayValue()).trim()) &&
    !hasReason_(sh.getRange(row, COL.REASON).getDisplayValue()) &&
    nightKey_(ts) === nightKey_(new Date());
  if (!open && ts.getTime() < Date.now() - CFG.EDIT_WINDOW_MIN * 60e3) {
    throw new Error(`Row ${row} is finished and older than ${CFG.EDIT_WINDOW_MIN} min — not editing it; fix it by hand if needed.`);
  }
}

function tryWithLock_(ms, fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(ms)) return { busy: true };
  SP_CACHE_ = null; LAST_RIDE_CACHE_ = null;
  try { return { value: fn() }; } finally { lock.releaseLock(); }
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  SP_CACHE_ = null; LAST_RIDE_CACHE_ = null;
  try { return fn(); } finally { lock.releaseLock(); }
}

function rideId_(ts, phone, first) {
  const p = String(phone || '').replace(/\D/g, '').slice(-10);
  return (ts instanceof Date ? ts.getTime() : '') + '|' + (p || norm_(first));
}
function rowIdAt_(sh, row) {
  if (!row || row < 2 || row > sh.getLastRow()) return '';
  const rng = sh.getRange(row, 1, 1, COL.PHONE), v = rng.getValues()[0], d = rng.getDisplayValues()[0];
  return rideId_(v[COL.TS - 1], d[COL.PHONE - 1], d[COL.FIRST - 1]);
}
function rideIndex_(sh) {
  const last = lastRideRow_(sh), first = Math.max(2, last - CFG.SCAN_ROWS + 1);
  const rng = sh.getRange(first, 1, last - first + 1, COL.PHONE), v = rng.getValues(), d = rng.getDisplayValues(), idx = {};
  v.forEach((r, i) => { if (r[0] instanceof Date) idx[rideId_(r[0], d[i][COL.PHONE - 1], d[i][COL.FIRST - 1])] = first + i; });
  return idx;
}

function resolveRows_(sh, refs) {
  let index = null;
  return [].concat(refs).map(ref => {
    if (typeof ref === 'number' || /^\d+$/.test(String(ref))) return Number(ref);
    if (rowIdAt_(sh, ref.row) === ref.id) return ref.row;
    if (!index) index = rideIndex_(sh);
    if (index[ref.id]) return index[ref.id];
    throw new Error("A ride you clicked isn't in the sheet anymore (deleted or archived) — nothing was changed. Refreshing.");
  });
}

function loadLedger_() {
  const v = bigGet_('RIDE_LEDGER');
  return v && v.night === nightKey_(new Date()) ? v.rides : {};
}
function saveLedger_(rides) {
  bigSet_('RIDE_LEDGER', { night: nightKey_(new Date()), rides });
}
function ledgerAdd_(entries) {
  const L = loadLedger_();
  entries.forEach(e => {
    if (!e.id) return;
    L[e.id] = Object.assign({ pickupKey: placeKey_(e.pickup), pickupName: canonPlace_(e.pickup) || e.pickup, sentAt: Date.now() },
      L[e.id] && { sentAt: L[e.id].sentAt }, e);
  });
  saveLedger_(L);
}

function ledgerExtras_(statusById) {
  const L = loadLedger_(), extras = [];
  let changed = false;
  Object.keys(L).forEach(id => {
    const e = L[id], s = statusById[id];
    if (s && s.state !== 'open') { delete L[id]; changed = true; return; }
    if (s && s.driver) return;
    if (!s && Date.now() - e.sentAt > 3 * 3600e3) { delete L[id]; changed = true; return; }
    extras.push(Object.assign({}, e, { row: s ? s.row : e.row, arrived: s ? s.arrived : (e.arrived || ''), picked: false, dupOf: null,
      lost: s ? 'crew missing in sheet' : 'row not found in sheet' }));
  });
  if (changed) { try { saveLedger_(L); } catch (e) {   } }
  return extras;
}

function restoreRide(ref) {
  withLock_(() => {
    const sh = sheet_(), row = resolveRows_(sh, [ref])[0], e = loadLedger_()[ref.id];
    if (!e) throw new Error('No record of that ride — assign it again.');
    assertEditable_(sh, row);
    const cur = sh.getRange(row, 1, 1, CFG.LAST_COL).getDisplayValues()[0];
    if (String(cur[COL.DRIVER - 1]).trim()) throw new Error(`${e.first} already has a crew in the sheet (${cur[COL.DRIVER - 1]}) — refresh.`);
    sh.getRange(row, COL.LOC, 1, 4).setValues([[e.loc || zoneOf_(e.dropoff) || '', e.van || '', e.driver, e.nav || '']]);
  });
}

function unassignRides(refs) {
  return withLock_(() => unassign_(sheet_(), refs));
}
function unassign_(sh, refs) {
  const rows = resolveRows_(sh, refs);
  const plan = rows.map(row => {
    assertEditable_(sh, row);
    const r = sh.getRange(row, 1, 1, CFG.LAST_COL).getDisplayValues()[0], get = c => String(r[c - 1]).trim();
    if (!get(COL.DRIVER)) throw new Error(`${get(COL.FIRST)} (row ${row}) isn't assigned to a crew anymore — refresh.`);
    if (hasReason_(get(COL.REASON))) throw new Error(`${get(COL.FIRST)} (row ${row}) was cancelled — refresh.`);
    if (isPickedRow_(rowColor_(sh.getRange(row, 1, 1, CFG.LAST_COL).getBackgrounds()[0]), get(COL.ARRIVE)))
      throw new Error(`${get(COL.FIRST)} (row ${row}) is already picked up — they can't be taken off the ride.`);
    return { row, id: rowIdAt_(sh, row), first: get(COL.FIRST), before: { arrive: sh.getRange(row, COL.ARRIVE).getValue(), mp: sh.getRange(row, COL.LOC, 1, 4).getValues()[0] } };
  });
  plan.forEach(p => {
    sh.getRange(p.row, COL.ARRIVE).clearContent();
    sh.getRange(p.row, COL.LOC, 1, 4).clearContent();
  });
  const L = loadLedger_(); let changed = false;
  plan.forEach(p => { if (L[p.id]) { delete L[p.id]; changed = true; } });
  if (changed) saveLedger_(L);
  LAST_RIDE_CACHE_ = null;
  return plan;
}
function moveRides(refs, crew, loc, van) {
  const plan = withLock_(() => unassign_(sheet_(), refs));
  try {
    return assignRides(refs, crew, loc, van);
  } catch (e) {
    withLock_(() => {
      const sh = sheet_();
      plan.forEach(p => {
        sh.getRange(p.row, COL.LOC, 1, 4).setValues([p.before.mp]);
        if (p.before.arrive !== '' && p.before.arrive != null) sh.getRange(p.row, COL.ARRIVE).setValue(p.before.arrive);
      });
    });
    throw new Error(e.message.replace(/nothing was assigned\.?/i, 'nothing was changed.'));
  }
}

function inProgressDiag_(statusById) {
  const d = { night: nightKey_(new Date()), withCrew: 0, open: 0, picked: 0, cancelled: 0, dup: 0, colors: {} };
  Object.values(statusById).forEach(s => {
    if (!s.tonight || !s.driver) return;
    d.withCrew++; d[s.state]++;
    if (s.state === 'picked') d.colors[s.color] = (d.colors[s.color] || 0) + 1;
  });
  return d;
}

function sortedLines_(items) {

  return items.map(x => Object.assign({ k: String(canonPlace_(x.pickup) || x.pickup || '').toLowerCase(), k2: String(x.pickup || '').trim().toLowerCase() }, x))
    .sort((a, b) => a.k.localeCompare(b.k) || a.k2.localeCompare(b.k2) || a.row - b.row).map(x => x.line);
}
function rideLine_(first, phone, pickup, dropoff) {
  const t = s => String(s == null ? '' : s).trim();
  return `${t(first)}  ${t(phone)}  ${t(pickup)} → ${t(dropoff)}`;
}

function assignRides(refs, crew, loc, van) {
  return withLock_(() => {
    const sh = sheet_();
    const rows = resolveRows_(sh, refs);
    const dispatcher = dispatcherName_();
    const lines = [];

    const t = tonight_(), crewsNow = applyCrewEdits_(t.crews).list;
    const c = crewsNow.find(x => samePair_(x, crew)) || crew, cap = Number(c.cap) || 0;
    if (cap) {
      const withCrew = t.rides.filter(r => r.driver && !r.picked && samePair_(r, c))
        .concat(carrying_(t.pickedTonight || [], crewsNow, crewBoard_(crewsNow, t.rides)).filter(r => samePair_(r, c)));

      const pickups = rows.map(row => sh.getRange(row, COL.PICKUP).getDisplayValue());
      const freed = withCrew.filter(r => pickups.every(p => samePlace_(r.dropoff, p))).length;
      const load = withCrew.length - freed;
      if (load + rows.length > cap) {
        throw new Error(`${c.vehicle || 'Their vehicle'} fits ${cap}: ${[c.driver, c.nav].filter(Boolean).join(' / ')} already ` +
          `${load === 1 ? 'has 1 rider' : `has ${load} riders`}, so only ${Math.max(0, cap - load)} more fit — nothing was assigned.`);
      }
    }

    const roles = crewBoard_([crew], [])[0] || {};
    const driver = roles.roleDriver || crew.driver, nav = roles.roleNav || crew.nav || '';
    const plan = rows.map(row => {
      assertEditable_(sh, row);
      const r = sh.getRange(row, 1, 1, CFG.LAST_COL).getDisplayValues()[0];
      const get = c => String(r[c - 1]).trim();
      if (get(COL.DRIVER) || hasReason_(get(COL.REASON))) {
        throw new Error(`${get(COL.FIRST)} (row ${row}) was already assigned or cancelled — refresh.`);
      }

      const fixes = [COL.PICKUP, COL.DROPOFF].map(c => [c, placeRewrite_(r[c - 1])]).filter(([, f]) => f);
      fixes.forEach(([c, f]) => { r[c - 1] = f; });

      const rowLoc = zoneOf_(get(COL.DROPOFF)) || (loc && loc !== 'Auto' ? loc : '') || 'Off';
      return { row, get, rowLoc, fixes };
    });

    if (isCart_(van)) {
      const off = plan.filter(p => p.rowLoc !== 'On');
      if (off.length) throw new Error(`Golf carts only take on-campus rides: ${off.map(p => `${p.get(COL.FIRST)} (${p.rowLoc})`).join(', ')} — nothing was assigned.`);
      const far = plan.filter(p => cartNoGo_(p.get(COL.PICKUP)) || cartNoGo_(p.get(COL.DROPOFF)));
      if (far.length) throw new Error(`Golf carts can't go to ECG, ECGA or the detached garages: ${far.map(p => `${p.get(COL.FIRST)} (${p.get(COL.PICKUP)} → ${p.get(COL.DROPOFF)})`).join(', ')} — nothing was assigned.`);
    }
    plan.forEach(({ row, get, rowLoc, fixes }) => {
      fixes.forEach(([c, f]) => sh.getRange(row, c).setValue(f));
      if (!get(COL.DISPATCH) && dispatcher) sh.getRange(row, COL.DISPATCH).setValue(dispatcher);
      if (!get(COL.CALL)) sh.getRange(row, COL.CALL).setValue(fmtTime_());
      sh.getRange(row, COL.LOC, 1, 4).setValues([[rowLoc, van, driver, nav]]);
      clearInheritedColor_(sh, row);
      lines.push({ pickup: get(COL.PICKUP), row, line: rideLine_(get(COL.FIRST), get(COL.PHONE), get(COL.PICKUP), get(COL.DROPOFF)) });
    });

    LAST_RIDE_CACHE_ = null;
    const others = tonight_().rides.filter(r => r.driver && !r.picked && !rows.includes(r.row) && samePair_(r, crew))
      .sort((a, b) => a.row - b.row);
    noteAssigned_(crew);
    saveBatch_(others.map(r => r.id).concat(rows.map(row => rowIdAt_(sh, row))));
    ledgerAdd_(plan.map(({ row, get, rowLoc }) => ({ id: rowIdAt_(sh, row), row, first: get(COL.FIRST), phone: get(COL.PHONE),
      pickup: get(COL.PICKUP), dropoff: get(COL.DROPOFF), time: get(COL.TS), driver, nav, van, loc: rowLoc })));
    const newLines = sortedLines_(lines);
    if (!others.length) return newLines.join('\n');

    return ['Updated ride!'].concat(sortedLines_(others.map(r => ({ pickup: r.pickup, row: r.row, line: rideLine_(r.first, r.phone, r.pickup, r.dropoff) }))), ['NEW:'], newLines).join('\n');
  });
}

function loadBatches_() {
  const v = JSON.parse(SP_().getProperty('ASSIGN_BATCHES') || 'null');
  return v && v.night === nightKey_(new Date()) && v.ids ? v.ids : {};
}

function saveBatch_(rideIds) {
  const b = loadBatches_(), group = 'a' + Date.now();
  rideIds.filter(Boolean).forEach(id => { b[id] = group; });
  const keep = Object.keys(b).sort((x, y) => Number(y.split('|')[0]) - Number(x.split('|')[0])).slice(0, 200);
  SP_().setProperty('ASSIGN_BATCHES', JSON.stringify({ night: nightKey_(new Date()), ids: Object.fromEntries(keep.map(k => [k, b[k]])) }));
}

function arrive_(sh, row, time) {
  assertEditable_(sh, row);
  const cell = sh.getRange(row, COL.ARRIVE);
  if (!cell.getDisplayValue()) {
    clearInheritedColor_(sh, row);
    cell.setValue(time);
  }
}

function isPickedRow_(color, arrival) {
  return color.kind === 'color' && !!String(arrival || '').trim();
}
function clearInheritedColor_(sh, row) {
  const kind = rowColor_(sh.getRange(row, 1, 1, CFG.LAST_COL).getBackgrounds()[0]).kind;
  if (kind === 'color') sh.getRange(row, 1, 1, sh.getMaxColumns()).setBackground(null);
}

const PICK_GAP_MIN = 2;

function clampToCall_(sh, row, arrive) {
  const call = /^(\d{1,2}):(\d{2})$/.exec(sh.getRange(row, COL.CALL).getDisplayValue().trim());
  if (!call) return arrive;
  const mins = d => Number(Utilities.formatDate(d, CFG.TZ, 'H')) * 60 + Number(Utilities.formatDate(d, CFG.TZ, 'm'));
  const before = (Number(call[1]) * 60 + Number(call[2]) - mins(arrive) + 1440) % 1440;
  return before > 0 && before <= 120 ? new Date(arrive.getTime() + before * 60e3) : arrive;
}

const CARRY_MAX_MIN = 45;
const CARRY_GUESS_MIN = 30;

function notePicked_(ids, ms) {
  const night = nightKey_(new Date()), v = JSON.parse(SP_().getProperty('PICKED_AT') || 'null');
  const at = v && v.night === night ? v.at : {};
  ids.forEach(id => { at[id] = ms; });
  const keep = Object.keys(at).sort((a, b) => at[b] - at[a]).slice(0, 300);
  SP_().setProperty('PICKED_AT', JSON.stringify({ night, at: Object.fromEntries(keep.map(k => [k, at[k]])) }));
}

function noteDropped_(ids, ms) {
  const night = nightKey_(new Date()), v = JSON.parse(SP_().getProperty('DROPPED_AT') || 'null');
  const at = v && v.night === night ? v.at : {};
  ids.forEach(id => { at[id] = ms; });
  const keep = Object.keys(at).sort((a, b) => at[b] - at[a]).slice(0, 300);
  SP_().setProperty('DROPPED_AT', JSON.stringify({ night, at: Object.fromEntries(keep.map(k => [k, at[k]])) }));
}
function loadDroppedAt_() {
  const v = JSON.parse(SP_().getProperty('DROPPED_AT') || 'null');
  return v && v.night === nightKey_(new Date()) ? v.at : {};
}
function loadPickedAt_() {
  const v = JSON.parse(SP_().getProperty('PICKED_AT') || 'null');
  return v && v.night === nightKey_(new Date()) ? v.at : {};
}

function clockMs_(s) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(s || '').trim());
  if (!m) return 0;
  const [h, mi] = fmtTime_(new Date()).split(':').map(Number);
  return Date.now() - (((h * 60 + mi) - (Number(m[1]) * 60 + Number(m[2])) + 1440) % 1440) * 60e3;
}
function carrying_(picked, crews, board) {
  const at = loadPickedAt_(), dropped = loadDroppedAt_(), now = Date.now();
  const kept = picked.map(r => {
    if (dropped[r.id]) return null;
    const c = crews.find(x => samePair_(x, r)), b = c && board.find(x => x.key === c.key);
    if (!b) return null;
    const ms = at[r.id] || clockMs_(r.arrived), max = (at[r.id] ? CARRY_MAX_MIN : CARRY_GUESS_MIN) * 60e3;
    if (!ms || now - ms > max || (b.emptyMs || 0) >= ms) return null;
    return Object.assign({}, r, { pickedAt: fmtTime_(new Date(ms)), pickedMs: ms, carryKey: c.key, carryCap: Number(c.cap) || 0 });
  }).filter(Boolean);

  const byCrew = {};
  kept.sort((a, b) => b.pickedMs - a.pickedMs).forEach(r => { (byCrew[r.carryKey] = byCrew[r.carryKey] || []).push(r); });
  return Object.values(byCrew).flatMap(list => (list[0].carryCap ? list.slice(0, list[0].carryCap) : list));
}

function pick_(sh, row, whenMs, sameTime) {
  const shade = shadeForRow_(sh, row, whenMs);
  const color = tonightColors_()[shade];
  if (!color) throw new Error("Set tonight's picked color at the top of the panel first.");
  arrive_(sh, row, fmtTime_(clampToCall_(sh, row, new Date(whenMs - (sameTime ? 0 : PICK_GAP_MIN * 60e3)))));
  sh.getRange(row, 1, 1, sh.getMaxColumns()).setBackground(color);
  notePicked_([rowIdAt_(sh, row)], whenMs);
}

function crewRows_(sh, refs) {
  const rows = resolveRows_(sh, refs);
  const crews = new Set(rows.map(row => {
    assertEditable_(sh, row);
    const v = sh.getRange(row, COL.DRIVER, 1, 2).getDisplayValues()[0];
    return pairKey_(v[0], v[1]);
  }));
  if (crews.size > 1) throw new Error('Those rides belong to different crews — refresh and try again.');
  return rows;
}

function markArrived(rows) {
  withLock_(() => {
    const sh = sheet_(), t = fmtTime_();
    crewRows_(sh, rows).forEach(row => arrive_(sh, row, t));
  });
}

function markPicked(rows) {
  withLock_(() => {
    const sh = sheet_(), now = Date.now();
    crewRows_(sh, rows).forEach(row => pick_(sh, row, now, false));
  });
}

const DUPLICATE_COLOR = '#ff0000';
function markDuplicate(ref) {
  withLock_(() => {
    const sh = sheet_(), row = resolveRows_(sh, [ref])[0];
    assertEditable_(sh, row);
    sh.getRange(row, 1, 1, sh.getMaxColumns()).setBackground(DUPLICATE_COLOR);
  });
}

function markNotAda(ref) {
  withLock_(() => {
    const sh = sheet_(), row = resolveRows_(sh, [ref])[0];
    assertEditable_(sh, row);
    sh.getRange(row, COL.ADA).setValue('No');
  });
}

const EDIT_FIELDS_ = { first: COL.FIRST, last: COL.LAST, pickup: COL.PICKUP, dropoff: COL.DROPOFF };
function editRide(ref, fields) {
  withLock_(() => {
    const sh = sheet_(), row = resolveRows_(sh, [ref])[0], f = fields || {};
    assertEditable_(sh, row);
    Object.keys(EDIT_FIELDS_).forEach(k => {
      if (!(k in f)) return;
      const v = String(f[k] == null ? '' : f[k]).replace(/^[=+@]+/, '').trim().slice(0, 200);
      if (!v && k !== 'last') throw new Error(`The ${k === 'first' ? 'first name' : k} can't be empty.`);
      const cell = sh.getRange(row, EDIT_FIELDS_[k]);
      if (String(cell.getDisplayValue()).trim() !== v) cell.setValue(v);
    });
  });
}

function cancelRide(ref, reason) {
  withLock_(() => {
    const sh = sheet_();
    cancelRow_(sh, resolveRows_(sh, [ref])[0], reason);
  });
}
function cancelRow_(sh, row, reason) {
  assertEditable_(sh, row);
  sh.getRange(row, COL.DISPATCH, 1, 3).clearContent();
  sh.getRange(row, COL.LOC, 1, 4).clearContent();
  sh.getRange(row, COL.REASON).setValue(reason);
  sh.getRange(row, 1, 1, sh.getMaxColumns()).setBackground('#000000').setFontColor('#ffffff');
}

function rideMessage(refs) {
  const sh = sheet_();
  return sortedLines_(resolveRows_(sh, refs).map(row => {
    const r = sh.getRange(row, 1, 1, CFG.LAST_COL).getDisplayValues()[0];
    return { pickup: r[COL.PICKUP - 1], row, line: rideLine_(r[COL.FIRST - 1], r[COL.PHONE - 1], r[COL.PICKUP - 1], r[COL.DROPOFF - 1]) };
  })).join('\n');
}

const UNDO_MAX_MIN = 30;
const UNDOABLE_ = {

  assignRides:   { rows: a => a[0], people: a => [a[1] && a[1].driver, a[1] && a[1].nav], ledger: true, keepCols: [COL.CALL] },
  unassignRides: { rows: a => a[0] },
  moveRides:     { rows: a => a[0], people: a => [a[1] && a[1].driver, a[1] && a[1].nav], ledger: true, keepCols: [COL.CALL] },
  markArrived:   { rows: a => a[0] },
  markPicked:    { rows: a => a[0] },
  cancelRide:    { rows: a => [a[0]] },
  markDuplicate: { rows: a => [a[0]] },
  markNotAda:    { rows: a => [a[0]] },
  editRide:      { rows: a => [a[0]] },
  restoreRide:   { rows: a => [a[0]] },
  setCrewStatus: { people: a => String(a[0]).split('|') },
  saveCrew:      { crewEdits: true },
  removeCrews:   { crewEdits: true },
  setCrewState:  { crewEdits: true },
  dismissChat:   { dismissed: a => [].concat(a[0]) },
  toggleShift:   { inverse: a => a[0] },
};
const PEOPLE_FIELDS_ = ['assigned', 'empty', 'manual'];

const undoEnc_ = v => (v instanceof Date ? { d: v.getTime() } : v);
const undoDec_ = v => (v && typeof v === 'object' && 'd' in v ? new Date(v.d) : v);

function snapRows_(sh, rows) {
  const w = sh.getMaxColumns();
  return rows.map(row => {
    const r = sh.getRange(row, 1, 1, CFG.LAST_COL);
    return { row, id: rowIdAt_(sh, row), v: r.getValues()[0], f: r.getFormulas()[0], d: r.getDisplayValues()[0],
             bg: sh.getRange(row, 1, 1, w).getBackgrounds()[0], fc: sh.getRange(row, 1, 1, w).getFontColors()[0] };
  });
}
function snapPeople_(names) {
  const P = loadPeople_(), out = {};
  names.filter(Boolean).forEach(n => {
    const k = norm_(String(n).trim().split(/\s+/)[0]);
    if (k.length >= 2) out[k] = Object.fromEntries(PEOPLE_FIELDS_.map(f => [f, (P[k] || {})[f] || 0]));
  });
  return out;
}
function undoSnap_(how, args) {
  const s = { at: Date.now() };
  if (how.rows) { const sh = sheet_(); s.sh = sh; s.rows = snapRows_(sh, resolveRows_(sh, how.rows(args))); }
  if (how.people) s.people = snapPeople_(how.people(args));
  if (how.crewEdits) s.crewEdits = SP_().getProperty('CREW_EDITS') || '';
  return s;
}
function undoTicket_(how, args, b) {
  const t = { at: b.at };
  if (b.rows) {
    const after = snapRows_(b.sh, b.rows.map(x => x.row));
    t.rows = b.rows.map((x, i) => {
      const y = after[i], cells = [];
      for (let c = 0; c < x.v.length; c++) {
        if (x.f[c] || y.f[c]) continue;
        if (how.keepCols && how.keepCols.includes(c + 1)) continue;
        if (x.d[c] !== y.d[c]) cells.push({ c: c + 1, before: undoEnc_(x.v[c]), after: y.d[c] });
      }
      const bgChanged = x.bg.join() !== y.bg.join(), fcChanged = x.fc.join() !== y.fc.join();
      return { row: y.row, id: y.id, cells, bgBefore: bgChanged ? x.bg : null, bgAfter: bgChanged ? y.bg : null,
               fcBefore: fcChanged ? x.fc : null, fcAfter: fcChanged ? y.fc : null };
    }).filter(r => r.cells.length || r.bgBefore || r.fcBefore);
    if (how.ledger) t.ledgerIds = after.map(y => y.id);
  }
  if (b.people) t.people = { before: b.people, after: snapPeople_(Object.keys(b.people)) };
  if (b.crewEdits != null) t.crewEdits = { before: b.crewEdits, after: SP_().getProperty('CREW_EDITS') || '' };
  if (how.dismissed) t.dismissed = how.dismissed(args);
  if (how.inverse) t.toggleShift = how.inverse(args);
  const empty = !(t.rows && t.rows.length) && !t.people && !t.crewEdits && !t.dismissed && !t.toggleShift;
  return empty ? null : t;
}

function undoAction(t) {
  if (!t || !t.at) throw new Error('Nothing to undo.');
  if (Date.now() - t.at > UNDO_MAX_MIN * 60e3) throw new Error(`That was over ${UNDO_MAX_MIN} minutes ago — too late to undo; fix it by hand.`);
  const out = { skipped: 0, notes: [] };
  if (t.msgId) {
    const r = deleteMyMessage_(t.msgId);
    if (!r.ok) out.notes.push(r.note);
  }
  withLock_(() => {
    if (t.rows && t.rows.length) {
      const sh = sheet_();
      t.rows.forEach(x => {
        let row;
        try { row = resolveRows_(sh, [{ row: x.row, id: x.id }])[0]; } catch (e) { out.skipped++; return; }
        x.cells.forEach(c => {
          const cell = sh.getRange(row, c.c);
          if (String(cell.getDisplayValue()) === String(c.after)) cell.setValue(undoDec_(c.before)); else out.skipped++;
        });
        if (x.bgBefore) {
          const rng = sh.getRange(row, 1, 1, x.bgBefore.length);
          if (rng.getBackgrounds()[0].join() === x.bgAfter.join()) rng.setBackgrounds([x.bgBefore]); else out.skipped++;
        }
        if (x.fcBefore) {
          const rng = sh.getRange(row, 1, 1, x.fcBefore.length);
          if (rng.getFontColors()[0].join() === x.fcAfter.join()) rng.setFontColors([x.fcBefore]); else out.skipped++;
        }
      });
    }
    if (t.ledgerIds) { const L = loadLedger_(); t.ledgerIds.forEach(id => delete L[id]); saveLedger_(L); }
    if (t.people) {
      const P = loadPeople_();
      Object.keys(t.people.after).forEach(k => PEOPLE_FIELDS_.forEach(f => {
        const now = (P[k] || {})[f] || 0, was = (t.people.before[k] || {})[f] || 0;
        if (now !== t.people.after[k][f]) { if (now !== was) out.skipped++; return; }
        if (was) (P[k] = P[k] || {})[f] = was; else if (P[k]) delete P[k][f];
      }));
      savePeople_(P);
    }
    if (t.crewEdits) {
      if ((SP_().getProperty('CREW_EDITS') || '') === t.crewEdits.after) SP_().setProperty('CREW_EDITS', t.crewEdits.before);
      else out.skipped++;
    }
    if (t.dismissed) {
      const p = SP_(), gone = new Set([].concat(t.dismissed).map(String));
      p.setProperty('GM_DISMISSED', JSON.stringify(JSON.parse(p.getProperty('GM_DISMISSED') || '[]').filter(i => !gone.has(String(i)))));
    }
  });
  if (t.toggleShift) toggleShift(t.toggleShift);
  (t.reposts || []).forEach(m => { try { postToGroup_(m.text, m.mentions); } catch (e) { out.notes.push("Couldn't re-post your earlier message: " + e.message); } });
  return out;
}

;

const FILE_VERSION_EXCEL = '2026-11-10a';

let XL_ = null;
function excelMode_() { return !!XL_; }

function setupExcelAddin() {
  const sp = PropertiesService.getScriptProperties();
  let key = sp.getProperty('EXCEL_KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, '').slice(0, 24); sp.setProperty('EXCEL_KEY', key); }
  Logger.log('Excel add-in team key: ' + key);
  return key;
}

function xlJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function xlSameKey_(a, b) {
  a = String(a || ''); b = String(b || '');
  if (!a || a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

function xlCallable_(fn) {
  return typeof fn === 'string' && /^[A-Za-z][A-Za-z0-9]*$/.test(fn) && !/^(doGet|doPost|setup)/.test(fn)
    && typeof globalThis[fn] === 'function';
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return xlJson_({ ok: false, error: 'Bad request.' }); }
  const key = PropertiesService.getScriptProperties().getProperty('EXCEL_KEY');
  if (!key) return xlJson_({ ok: false, badKey: true, error: 'The Excel add-in is not set up yet: open the Apps Script editor and run setupExcelAddin().' });
  if (!xlSameKey_(req.key, key)) return xlJson_({ ok: false, badKey: true, error: 'Wrong team key.' });
  if (!/^[A-Za-z0-9_-]{16,64}$/.test(String(req.uid || ''))) return xlJson_({ ok: false, error: 'Missing install id.' });

  if (req.fn === '__page') {
    return xlJson_({ ok: true, result: HtmlService.createHtmlOutputFromFile('Sidebar').getContent(), build: typeof BUILD_ID !== 'undefined' ? BUILD_ID : '' });
  }
  if (!xlCallable_(req.fn)) return xlJson_({ ok: false, error: `Unknown function "${req.fn}".` });

  XL_ = { uid: String(req.uid), snap: req.snap ? xlDecodeSnap_(req.snap) : null, ops: [], sheet: null };
  let result = null, error = null;
  try { result = globalThis[req.fn].apply(null, Array.isArray(req.args) ? req.args : []); } catch (err) { error = (err && err.message) || String(err); }
  const ops = XL_.ops;
  XL_ = null;

  return xlJson_(error ? { ok: false, error, ops } : { ok: true, result: result === undefined ? null : result, ops });
}

function xlDecodeSnap_(s) {
  if (typeof s === 'object') return s;
  const blob = Utilities.newBlob(Utilities.base64Decode(s), 'application/x-gzip');
  return JSON.parse(Utilities.ungzip(blob).getDataAsString('UTF-8'));
}

function userProps_() {
  if (!XL_) return PropertiesService.getUserProperties();
  const sp = PropertiesService.getScriptProperties(), pre = 'U:' + XL_.uid + ':';
  const P = {
    getProperty: k => sp.getProperty(pre + k),
    setProperty: (k, v) => { sp.setProperty(pre + k, String(v)); return P; },
    deleteProperty: k => { sp.deleteProperty(pre + k); return P; },
    getProperties: () => {
      const all = sp.getProperties(), out = {};
      Object.keys(all).forEach(k => { if (k.indexOf(pre) === 0) out[k.slice(pre.length)] = all[k]; });
      return out;
    },
  };
  return P;
}

const XL_FMT_ = 'yyyy-MM-dd HH:mm:ss';
function XlSheet_(snap) {
  this.s = snap;
  this.rows = snap.rows || {};
  this.width = snap.width || 26;
}
XlSheet_.prototype.getLastRow = function () { return this.s.lastRow || 1; };
XlSheet_.prototype.getLastColumn = function () { return this.s.lastCol || this.width; };
XlSheet_.prototype.getMaxColumns = function () { return this.width; };
XlSheet_.prototype.getName = function () { return this.s.name || 'Sheet1'; };
XlSheet_.prototype.getParent = function () { const n = this.s.book || 'Workbook'; return { getName: () => n }; };
XlSheet_.prototype.getRange = function (row, col, nr, nc) {
  if (typeof row === 'string') throw new Error('A1 ranges are not supported in Excel mode.');
  return new XlRange_(this, row, col, nr || 1, nc || 1);
};
XlSheet_.prototype.row_ = function (r, make) {
  let x = this.rows[r];
  if (!x && make) x = this.rows[r] = { d: [], v: {}, f: {} };
  return x || null;
};
XlSheet_.prototype.cell_ = function (r, c) {
  const x = this.rows[r];
  if (!x) return { d: '', v: undefined, f: '', bg: '#ffffff', fc: '#000000' };
  const pick = (a, dflt) => (a == null ? dflt : typeof a === 'string' ? a : (a[c - 1] || dflt));
  const d = x.d && x.d[c - 1] != null ? String(x.d[c - 1]) : '';
  return { d, v: x.v ? x.v[c] : undefined, f: (x.f && x.f[c]) || '', bg: pick(x.bg, '#ffffff'), fc: pick(x.fc, '#000000') };
};

XlSheet_.prototype.colors_ = function (x, k, dflt) {
  if (Array.isArray(x[k])) return x[k];
  const one = x[k] || dflt, a = [];
  for (let i = 0; i < this.width; i++) a.push(one);
  return (x[k] = a);
};

function xlValueOf_(cell) {
  const v = cell.v;
  if (v && typeof v === 'object') {
    if ('t' in v) return Utilities.parseDate(v.t, CFG.TZ, XL_FMT_);
    if ('n' in v) return v.n;
    if ('b' in v) return v.b;
  }
  return cell.d;
}

function xlEncode_(v) {
  if (v instanceof Date) {
    const t = Utilities.formatDate(v, CFG.TZ, XL_FMT_);
    const d = t.slice(0, 4) === '1899' ? Utilities.formatDate(v, CFG.TZ, 'H:mm') : Utilities.formatDate(v, CFG.TZ, 'M/d/yyyy H:mm:ss');
    return { out: { t }, d, v: { t } };
  }
  if (v == null) return { out: '', d: '', v: undefined };
  if (typeof v === 'number') return { out: v, d: String(v), v: { n: v } };
  if (typeof v === 'boolean') return { out: v, d: v ? 'TRUE' : 'FALSE', v: { b: v } };
  return { out: String(v), d: String(v), v: undefined };
}

function XlRange_(sh, row, col, nr, nc) { this.sh = sh; this.r = row; this.c = col; this.h = nr; this.w = nc; }
XlRange_.prototype.grid_ = function (fn) {
  const out = [];
  for (let i = 0; i < this.h; i++) {
    const line = [];
    for (let j = 0; j < this.w; j++) line.push(fn(this.sh.cell_(this.r + i, this.c + j)));
    out.push(line);
  }
  return out;
};
XlRange_.prototype.getValues = function () { return this.grid_(xlValueOf_); };
XlRange_.prototype.getDisplayValues = function () { return this.grid_(x => x.d); };
XlRange_.prototype.getFormulas = function () { return this.grid_(x => x.f); };
XlRange_.prototype.getBackgrounds = function () { return this.grid_(x => x.bg); };
XlRange_.prototype.getFontColors = function () { return this.grid_(x => x.fc); };
XlRange_.prototype.getValue = function () { return this.getValues()[0][0]; };
XlRange_.prototype.getDisplayValue = function () { return this.getDisplayValues()[0][0]; };
XlRange_.prototype.getRow = function () { return this.r; };
XlRange_.prototype.getColumn = function () { return this.c; };
XlRange_.prototype.getNumRows = function () { return this.h; };
XlRange_.prototype.getNumColumns = function () { return this.w; };
XlRange_.prototype.op_ = function (op) {
  op.r = this.r; op.c = this.c; op.h = this.h; op.w = this.w;
  if (this.c === 1 && this.w >= this.sh.width) op.row = true;
  if (XL_) XL_.ops.push(op);
  return this;
};

XlRange_.prototype.setValues = function (vals) {
  if (vals.length !== this.h || vals.some(r => r.length !== this.w)) throw new Error('setValues: size does not match the range.');
  const out = vals.map((line, i) => line.map((v, j) => {
    const e = xlEncode_(v), x = this.sh.row_(this.r + i, true), c = this.c + j;
    x.d[c - 1] = e.d;
    if (!x.v) x.v = {};
    if (!x.f) x.f = {};
    if (e.v) x.v[c] = e.v; else delete x.v[c];
    if (typeof v === 'string' && v.charAt(0) === '=') x.f[c] = v; else delete x.f[c];
    return e.out;
  }));
  return this.op_({ t: 'v', vals: out });
};
XlRange_.prototype.setValue = function (v) {
  const vals = [];
  for (let i = 0; i < this.h; i++) { const line = []; for (let j = 0; j < this.w; j++) line.push(v); vals.push(line); }
  return this.setValues(vals);
};
XlRange_.prototype.clearContent = function () {
  for (let i = 0; i < this.h; i++) {
    const x = this.sh.row_(this.r + i, false);
    if (!x) continue;
    for (let j = 0; j < this.w; j++) {
      const c = this.c + j;
      if (x.d) x.d[c - 1] = '';
      if (x.v) delete x.v[c];
      if (x.f) delete x.f[c];
    }
  }
  return this.op_({ t: 'clr' });
};
XlRange_.prototype.colorGrid_ = function (k, dflt, grid) {
  for (let i = 0; i < this.h; i++) {
    const x = this.sh.row_(this.r + i, true);
    if (this.c === 1 && this.w >= this.sh.width && grid[i].every(v => v === grid[i][0])) { x[k] = grid[i][0]; continue; }
    const a = this.sh.colors_(x, k, dflt);
    for (let j = 0; j < this.w; j++) if (this.c + j - 1 < this.sh.width) a[this.c + j - 1] = grid[i][j];
  }
};
const xlColor_ = (c, dflt) => (c == null || c === '' ? dflt : String(c).toLowerCase());
XlRange_.prototype.fill_ = function (k, dflt, t, color) {
  const c = xlColor_(color, dflt), grid = [];
  for (let i = 0; i < this.h; i++) { const line = []; for (let j = 0; j < this.w; j++) line.push(c); grid.push(line); }
  this.colorGrid_(k, dflt, grid);
  return this.op_({ t, color: color == null ? null : c });
};
XlRange_.prototype.fills_ = function (k, dflt, t, colors) {
  const grid = colors.map(line => line.map(c => xlColor_(c, dflt)));
  this.colorGrid_(k, dflt, grid);
  return this.op_({ t, colors: grid });
};
XlRange_.prototype.setBackground = function (c) { return this.fill_('bg', '#ffffff', 'bg', c); };
XlRange_.prototype.setBackgrounds = function (g) { return this.fills_('bg', '#ffffff', 'bg', g); };
XlRange_.prototype.setFontColor = function (c) { return this.fill_('fc', '#000000', 'fc', c); };
XlRange_.prototype.setFontColors = function (g) { return this.fills_('fc', '#000000', 'fc', g); };

['setFontWeight', 'setNumberFormat', 'setHorizontalAlignment', 'setNote'].forEach(m => {
  XlRange_.prototype[m] = function () { return this; };
});

function __swReset() {
  SP_CACHE_ = null; UP_CACHE_ = null; LAST_RIDE_CACHE_ = null; ROSTER_CACHE_ = null;
  GM_OFFLINE_ = false; CREW_NAMES_ = new Set(); POSTER_NAMES_ = new Set(); BATCH_MSGS_ = []; LAST_UPDATE_ = {};
}
function __swCall(msg) {
  if (msg.props) __SW.script = msg.props;
  __SW.changed = false;
  __swReset();
  const out = { id: msg.id };
  if (!xlCallable_(msg.fn)) { out.ok = false; out.error = `Unknown function "${msg.fn}".`; out.ops = []; return out; }
  XL_ = { uid: String(msg.uid), snap: msg.snap, ops: [], sheet: null };
  try {
    const r = globalThis[msg.fn].apply(null, msg.args || []);
    out.ok = true;
    out.result = r === undefined ? null : JSON.parse(JSON.stringify(r));
  } catch (e) { out.ok = false; out.error = (e && e.message) || String(e); }
  out.ops = XL_.ops;
  XL_ = null;
  if (__SW.changed) out.props = __SW.script;
  return out;
}
if (typeof WorkerGlobalScope !== 'undefined' && self instanceof WorkerGlobalScope) {
  self.onmessage = e => { self.postMessage(__swCall(e.data)); };
} else {
  globalThis.__swCall = __swCall;
}
