/*
 * SURE Walk dispatch rules — the same rules as the Apps Script panel (apps-script/Code.gs), without any sheet access.
 * Everything here works on plain text grids read from the workbook, so it can be tested on its own.
 * Needs lib/places.js (canonPlace_, placeKey_, zoneOf_, samePlace_, cartNoGo_, placeRewrite_, offCampusPickup_).
 */
/* exported R */
const R = (() => {
  const TZ = 'America/Chicago';
  // Column numbers (1 = A). Set from the sheet's header row by setLayout(); these defaults are the Excel
  // "Sure Walk Fall 2026" layout (Microsoft Forms: Id, Start time, Form Submitted, Time Submitted, Dispatcher, …).
  const COL = { TS: 3, DISPATCH: 5, CALL: 6, ARRIVE: 7, WAIT: 8, EID: 9, LAST: 10, FIRST: 11, PHONE: 12, PICKUP: 13, DROPOFF: 14,
    ADA: 15, LOC: 16, VAN: 17, DRIVER: 18, NAV: 19, EMAIL: 23, REASON: 24, NOTES: 25 };
  let LAST_COL = 25;
  const READ_COLS = 30;         // columns read from the sheet (A:AD); the layout is found inside them
  const colLetter = n => { let s = ''; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const colNumber = s => String(s).toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  // Header text -> column. First match wins, so "Form Submitted" beats "Start time" for the timestamp.
  const HEADERS = [
    ['TS', /^(form submitted|timestamp|completion time)/], ['DISPATCH', /^dispatch/], ['CALL', /^call time/], ['ARRIVE', /^arrival/],
    ['WAIT', /^wait/], ['EID', /\beid\b/], ['LAST', /^last name/], ['FIRST', /^first name/], ['PHONE', /^phone/],
    ['PICKUP', /^pick ?-?up/], ['DROPOFF', /^drop ?-?off/], ['ADA', /disabilit|^ada\b/], ['LOC', /^location$/],
    ['VAN', /^(van|vehicle)/], ['DRIVER', /^driver$/], ['NAV', /^navigator$/], ['EMAIL', /^email/], ['REASON', /cancel/], ['NOTES', /^notes?$/],
  ];
  /** Reads the header row (row 1 text) and points COL at the right columns. Throws if a needed column is missing. */
  function setLayout(header) {
    const h = (header || []).map(x => String(x || '').toLowerCase().replace(/[^a-z0-9 -]/g, ' ').replace(/\s+/g, ' ').trim());
    const found = {};
    HEADERS.forEach(([k, re]) => { const i = h.findIndex((x, j) => re.test(x) && !Object.values(found).includes(j + 1)); if (i >= 0) found[k] = i + 1; });
    const need = ['TS', 'FIRST', 'PHONE', 'PICKUP', 'DROPOFF', 'DRIVER', 'NAV', 'ARRIVE', 'CALL'];
    const missing = need.filter(k => !found[k]);
    if (missing.length) throw new Error(`Can't find these columns in row 1: ${missing.join(', ')} — is the rides tab selected?`);
    Object.assign(COL, found);
    LAST_COL = Math.max(...Object.values(COL));
    return Object.assign({}, COL);
  }
  const EDIT_WINDOW_MIN = 60;   // never write to finished rides submitted longer ago than this
  const DUP_WINDOW_MIN = 40;    // possible-duplicate check looks back this far
  const LOOKBACK_HOURS = 14;    // tonight's rows
  const PICK_GAP_MIN = 2;       // arrival assumed this long before a pick when "At" was skipped
  const LOCATIONS = ['West', 'On', 'Off', 'North'];
  const VANS = ['Van', 'ADA', 'Sienna', 'Cart'];
  const CANCEL_REASONS = ['Chose Alternative Way Home', 'Not Answering', 'Long Wait Time Expected', 'Before Start',
    'Not Ready', 'Not Within Pickup Drop-off Parameters', 'Exceeded 2min Wait', 'After Deadline',
    'Wrong Number', 'Not Grouped with Friends', 'Luggage etc.', 'Other'];
  const DUPLICATE_COLOR = '#ff0000';

  /* ---------- clock (always Austin time, whatever the computer is set to) ---------- */
  const pad = n => String(n).padStart(2, '0');
  const FMT = new Intl.DateTimeFormat('en-US', { timeZone: TZ, year: 'numeric', month: 'numeric', day: 'numeric',
    hour: 'numeric', minute: 'numeric', second: 'numeric', hourCycle: 'h23' });
  function parts(ms) {
    const o = {};
    FMT.formatToParts(new Date(ms)).forEach(p => { o[p.type] = Number(p.value); });
    return { y: o.year, mo: o.month, d: o.day, h: o.hour % 24, mi: o.minute, s: o.second };
  }
  // Austin wall-clock time -> ms
  function wallToMs(y, mo, d, h, mi, s) {
    const guess = Date.UTC(y, mo - 1, d, h, mi, s || 0);
    let ms = guess;
    for (let i = 0; i < 2; i++) { const p = parts(ms); ms = guess - (Date.UTC(p.y, p.mo - 1, p.d, p.h, p.mi, p.s) - ms); }
    return ms;
  }
  const fmtTime = ms => { const p = parts(ms == null ? Date.now() : ms); return `${p.h}:${pad(p.mi)}`; };
  const dateKey = ms => { const p = parts(ms); return `${p.y}-${pad(p.mo)}-${pad(p.d)}`; };
  // A dispatch night rolls over at 2:30 AM: 1:30 AM on 10/18 is still the 10/17 night.
  const nightKey = ms => dateKey((ms == null ? Date.now() : ms) - 150 * 60e3);
  // minutes into the night; past midnight = +24h
  const nightMin = ms => { const p = parts(ms == null ? Date.now() : ms), t = p.h * 60 + p.mi; return t < 12 * 60 ? t + 1440 : t; };

  // "11/4/2026 0:29:13", "11/4/26 12:29 AM", "2026-11-04 00:29:13" -> ms (Austin time); null if it isn't a timestamp.
  function parseTs(text) {
    const s = String(text || '').trim();
    let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?$/i.exec(s);
    let y, mo, d, h, mi, sec, ap;
    if (m) { [, mo, d, y, h, mi, sec, ap] = m; if (y.length === 2) y = '20' + y; }
    else if ((m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s))) [, y, mo, d, h, mi, sec] = m;
    else return null;
    h = Number(h) % (ap ? 12 : 24) + (ap && /pm/i.test(ap) ? 12 : 0);
    return wallToMs(+y, +mo, +d, h, +mi, +(sec || 0));
  }
  const hasTime = text => /\d{1,2}:\d{2}/.test(String(text || ''));
  // A ride's id: its timestamp + phone (or first name) — stable even when rows move.
  const rideIdOf = text => { const ts = parseTs(text[COL.TS - 1]); return ts == null ? '' : ts + '|' + (String(text[COL.PHONE - 1] || '').replace(/\D/g, '').slice(-10) || String(text[COL.FIRST - 1] || '').toLowerCase().replace(/[^a-z0-9]/g, '')); };

  /* ---------- names ---------- */
  const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  function sameName(a, b) {
    const x = norm(String(a).split(/\s+/)[0]), y = norm(String(b).split(/\s+/)[0]);
    if (x.length < 3 || y.length < 3) return x === y && !!x;
    return x.startsWith(y) || y.startsWith(x);
  }
  // the same two people in either order ("Esther/Cam" = "Camden/Esther")
  function samePair(a, b) {
    const an = String(a.nav || '').trim(), bn = String(b.nav || '').trim();
    if (!an || !bn) return !an && !bn && sameName(a.driver, b.driver);
    return (sameName(a.driver, b.driver) && sameName(an, bn)) || (sameName(a.driver, bn) && sameName(an, b.driver));
  }
  const crewKey = c => (String(c.driver || '') + '|' + String(c.nav || '')).toLowerCase();
  const crewLabel = c => [c.driver, c.nav].filter(Boolean).join(' / ');

  /* ---------- row colors: black = cancelled, red = duplicate, any other real color = picked ---------- */
  function rgb(hex) {
    const m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex || '');
    return m ? m.slice(1).map(h => parseInt(h, 16)) : null;
  }
  function isBlack(hex) { const c = rgb(hex); return !!c && Math.max(...c) <= 0x50 && Math.max(...c) - Math.min(...c) <= 40; }
  function isRed(hex) {
    const c = rgb(hex);
    if (!c || isBlack(hex)) return false;
    const [r, g, b] = c, max = Math.max(r, g, b), min = Math.min(r, g, b);
    if (max !== r || max - min < 25 || (max - min) / max < 0.15) return false;
    const hue = ((g - b) / (max - min)) * 60;
    return hue >= -15 && hue <= 15;
  }
  function isColored(hex) { const c = rgb(hex); return !!c && !isBlack(hex) && !isRed(hex) && Math.max(...c) - Math.min(...c) >= 12; }
  // one fill color for the whole row (what the extension reads) -> kind
  function rowKind(fill) {
    const h = String(fill || '').toLowerCase();
    if (isBlack(h)) return { kind: 'black' };
    if (isRed(h)) return { kind: 'red' };
    if (isColored(h)) return { kind: 'color', color: h };
    return { kind: 'none' };
  }
  const PASTEL_OF = { '#ff9900': '#fce5cd', '#ffff00': '#fff2cc', '#00ff00': '#d9ead3', '#00ffff': '#b3f0ea',
    '#4a86e8': '#c9daf8', '#0000ff': '#cfe2f3', '#9900ff': '#d9d2e9', '#ff00ff': '#ead1dc' };
  const isPastel = hex => { const c = rgb(hex); return !!c && Math.min(...c) >= 140; };
  function pastelOf(hex) {
    hex = String(hex).toLowerCase();
    if (PASTEL_OF[hex]) return PASTEL_OF[hex];
    const c = rgb(hex);
    return c ? '#' + c.map(v => Math.round(v + (255 - v) * 0.75).toString(16).padStart(2, '0')).join('') : null;
  }
  const brightOf = hex => Object.keys(PASTEL_OF).find(b => PASTEL_OF[b] === String(hex).toLowerCase()) || null;
  // A colored row only counts as picked when Arrival (D) is filled too (Forms copies the row above's color).
  const isPickedRow = (kind, arrival) => kind.kind === 'color' && !!String(arrival || '').trim();
  const PICKED_GUESS_MIN = 30;
  // "19:07" -> minutes since that clock time (within the last 24 h); null if it isn't a time.
  function minutesAgo(hhmm, nowMs) {
    const m = /^(\d{1,2}):(\d{2})/.exec(String(hhmm || '').trim());
    if (!m) return null;
    const p = parts(nowMs == null ? Date.now() : nowMs);
    return ((p.h * 60 + p.mi) - (Number(m[1]) * 60 + Number(m[2])) + 1440) % 1440;
  }
  const hasReason = v => { const s = String(v == null ? '' : v).trim(); return !!s && s !== '0'; };

  /* ---------- roster block below the rides ---------- */
  function rosterDateKey(s) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(String(s).trim());
    if (!m) return null;
    const y = m[3].length === 2 ? '20' + m[3] : m[3];
    return `${y}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
  }
  function vanFromLabel(label) {
    const s = label.toLowerCase();
    if (s.includes('cart')) return 'Cart';
    if (s.includes('ada') || s.includes('voyager')) return 'ADA';
    if (s.includes('sienna')) return 'Sienna';
    return 'Van';
  }
  const capOfLabel = label => { const m = /\bx\s*(\d{1,2})\s*$/i.exec(String(label).trim()); return m ? Number(m[1]) : 0; };
  const vehicleName = label => label.replace(/\s*x\d+\s*$/i, '').replace(/\s*:\s*/, ' ').replace(/\s+/g, ' ').trim();
  function noteMinutes(note) {
    const x = /^@\s*(\d{1,2})(?::(\d{2}))?\s*(am|pm|a|p)?/i.exec(String(note || '').trim());
    if (!x) return 0;
    let h = Number(x[1]) % 12;
    const ap = (x[3] || '').toLowerCase()[0];
    if (ap !== 'a' && h >= 6) h += 12;
    if (ap === 'a' || h < 6) h += 24;
    return h * 60 + Number(x[2] || 0);
  }
  const shiftLetter = name => { const m = /^([a-z])\b/i.exec(String(name).trim()); return m ? m[1].toUpperCase() : ''; };
  function markMidMoves(shifts) {
    shifts.forEach(sh => {
      sh.crews = sh.crews.reduce((out, c) => {
        const prev = out.find(x => samePair(x, c));
        if (!prev) { out.push(c); return out; }
        if (!noteMinutes(c.note) && !noteMinutes(prev.note)) return out;
        const [base, later] = noteMinutes(c.note) ? [prev, c] : [c, prev];
        out[out.indexOf(prev)] = Object.assign({}, base, { note: '', move: { at: noteMinutes(later.note), van: later.van, vehicle: later.vehicle, cap: later.cap, phone: later.phone } });
        return out;
      }, []);
    });
    const m = shifts.find(s => shiftLetter(s.name) === 'M'), b = shifts.find(s => shiftLetter(s.name) === 'B');
    if (!m || !b) return;
    b.crews = b.crews.filter(bc => {
      const mc = m.crews.find(x => samePair(x, bc));
      if (!mc) return true;
      mc.move = { at: noteMinutes(bc.note) || 22 * 60 + 30, van: bc.van, vehicle: bc.vehicle, cap: bc.cap, phone: bc.phone };
      return false;
    });
  }
  // v: text rows below the last ride (each row = columns A, B, …). Tonight's block, else the newest dated before tonight.
  function parseRoster(v, nowMs) {
    v = v.map(r => r.map(c => String(c == null ? '' : c).trim()));
    const starts = v.map((r, i) => (rosterDateKey(r[0]) ? i : -1)).filter(i => i >= 0);
    if (!starts.length) return null;
    const night = nightKey(nowMs);
    const dated = starts.map(i => ({ i, d: rosterDateKey(v[i][0]) }));
    const past = dated.filter(x => x.d <= night).sort((a, b) => (a.d < b.d ? -1 : a.d > b.d ? 1 : a.i - b.i));
    const si = (dated.find(x => x.d === night) || past[past.length - 1] || dated[dated.length - 1]).i;
    const nextStart = starts.find(i => i > si);
    const block = v.slice(si, nextStart == null ? v.length : nextStart);
    const hdr = block.findIndex(r => r.some(c => /^driver$/i.test(c)));
    if (hdr < 0) return null;
    const driverCols = block[hdr].map((c, j) => (/^driver$/i.test(c) ? j : -1)).filter(j => j >= 0);
    const labels = block.slice(0, hdr).reverse().find(r => r.some(c => /\bphone\b/i.test(c))) || [];
    const shifts = driverCols.map((dc, n) => ({
      name: (/shift/i.test(labels[dc] || '') ? String(labels[dc]).replace(/:\s*$/, '').trim() : '') ||
        (driverCols.length === 3 ? ['A Shift', 'M Shift', 'B Shift'][n] : n === 0 ? 'A Shift' : `Shift ${n + 1}`),
      crews: [], dispatchers: [],
    }));
    const isName = x => /^[a-z][a-z .'-]*$/i.test(x || '');
    block.forEach((r, i) => {
      if (i !== hdr && /dispatcher/i.test(r[0] || '')) driverCols.forEach((dc, n) => [r[dc], r[dc + 1]].filter(isName).forEach(x => shifts[n].dispatchers.push(x)));
    });
    block.slice(hdr + 1).forEach(r => {
      const label = r[0];
      if (!label || /dispatcher|floater/i.test(label)) return;
      driverCols.forEach((dc, n) => {
        const shifted = /^@\s*\d/.test(r[dc] || '');
        const driver = (shifted ? r[dc + 1] : r[dc]) || '', nav = (shifted ? r[dc + 2] : r[dc + 1]) || '';
        if (!driver && !nav) return;
        if (/repair/i.test(driver + ' ' + nav)) return;
        shifts[n].crews.push({ driver: driver || nav, nav: driver ? nav : '', van: vanFromLabel(label), vehicle: vehicleName(label),
          cap: capOfLabel(label), note: (shifted ? r[dc] : '') || r.slice(1).find(x => /^@\s*\d/.test(x)) || '',
          phone: (/^\s*-*\s*(\d{3,4})\s*$/.exec(r[dc + 2] || '') || [])[1] || '' });
      });
    });
    markMidMoves(shifts);
    return { date: rosterDateKey(block[0][0]), shifts };
  }

  /* ---------- which shifts are on (A 6:30–10:30 PM, M 9 PM–1 AM, B 10:30 PM–2:30 AM) ---------- */
  const SHIFT_CHANGES = [{ at: 21 * 60, on: ['A', 'M'] }, { at: 22 * 60 + 30, on: ['M', 'B'] }, { at: 25 * 60, on: ['B'] }];
  function autoShiftLetters(nowMs) {
    const t = nightMin(nowMs);
    let on = ['A'];
    SHIFT_CHANGES.forEach(c => { if (t >= c.at) on = c.on; });
    return on;
  }
  // toggles = { night, period, names } saved when the dispatcher flips a shift; it holds until the next shift change.
  function activeShifts(roster, toggles, nowMs) {
    if (!roster || !roster.shifts.length) return [];
    const auto = autoShiftLetters(nowMs);
    if (toggles && toggles.night === nightKey(nowMs) && toggles.period === auto.join('')) return roster.shifts.map(s => s.name).filter(n => toggles.names.includes(n));
    const names = roster.shifts.filter(s => auto.includes(shiftLetter(s.name))).map(s => s.name);
    return names.length ? names : [roster.shifts[0].name];
  }
  // Picked shade: the dispatcher's roster shift (B = pastel), else the clock (pastel from 10:30 PM).
  function dispatcherShift(roster, name) {
    const words = String(name || '').toLowerCase().replace(/[^a-z .]/g, '').split(/[\s.]+/).filter(Boolean);
    if (!roster || !words.length) return '';
    const first = words[0], all = [];
    roster.shifts.forEach(s => (s.dispatchers || []).forEach(d => all.push({ d: String(d).toLowerCase().trim(), letter: shiftLetter(s.name) })));
    const hits = all.filter(x => (first.length > 1 ? x.d.split(/\s+/)[0] === first : x.d[0] === first));
    const letters = [...new Set(hits.map(x => x.letter))];
    return letters.length === 1 ? letters[0] : '';
  }
  function shadeFor(roster, dispatcher, nowMs) {
    const l = dispatcherShift(roster, dispatcher);
    if (l) return l === 'B' ? 'B' : 'A';
    return nightMin(nowMs) >= 22 * 60 + 30 ? 'B' : 'A';
  }

  /* ---------- wait time (column E = Arrival − Call), last 5 rides ---------- */
  function waitMinutes(call, arrive) {
    const p = s => { const m = /^(\d{1,2}):(\d{2})$/.exec(String(s).trim()); return m ? Number(m[1]) * 60 + Number(m[2]) : null; };
    const c = p(call), a = p(arrive);
    if (c == null || a == null) return null;
    const w = (a - c + 1440) % 1440;
    return w <= 120 ? w : null;
  }

  /*
   * The night from the sheet.
   * rows: [{ row, text: [A..W], fill: '#rrggbb' | '' }] (fill '' = not read yet / mixed), newest last.
   * Returns { rides, pending, active, crews, roster, wait, colors }.
   */
  function buildNight({ rows, rosterRows, toggles, extraCrews, chosenColors, nowMs }) {
    nowMs = nowMs == null ? Date.now() : nowMs;
    const night = nightKey(nowMs), cutoff = nowMs - LOOKBACK_HOURS * 3600e3, editCutoff = nowMs - EDIT_WINDOW_MIN * 60e3;
    const rides = [], prior = [], waits = [], seen = {}, colorCounts = { A: {}, B: {} };
    rows.forEach(({ row, text, fill }) => {
      const ts = parseTs(text[COL.TS - 1]);
      if (ts == null || ts < cutoff) return;
      const get = c => String(text[c - 1] == null ? '' : text[c - 1]).trim();
      const kind = rowKind(fill), tonight = nightKey(ts) === night;
      const r = {
        row, id: rideIdOf(text),
        tsMs: ts, time: fmtTime(ts), tonight,
        first: get(COL.FIRST), last: get(COL.LAST), phone: get(COL.PHONE), pickup: get(COL.PICKUP), dropoff: get(COL.DROPOFF),
        ada: /yes/i.test(get(COL.ADA)), dispatcher: get(COL.DISPATCH), call: get(COL.CALL), arrived: get(COL.ARRIVE),
        loc: get(COL.LOC), van: get(COL.VAN), driver: get(COL.DRIVER), nav: get(COL.NAV), reason: get(COL.REASON), notes: get(COL.NOTES),
        fill: String(fill || '').toLowerCase(), kind: kind.kind,
        pickupName: (typeof canonPlace_ === 'function' && canonPlace_(get(COL.PICKUP))) || get(COL.PICKUP),
        zone: typeof zoneOf_ === 'function' ? zoneOf_(get(COL.DROPOFF)) : '',
        offCampus: typeof offCampusPickup_ === 'function' ? !!offCampusPickup_(get(COL.PICKUP)) : false,
      };
      // Picked = a colored row with an Arrival time. A row nobody colored counts as picked once its Arrival is
      // PICKED_GUESS_MIN old (the Excel sheet isn't always colored), so finished rides don't fill up the crews.
      const arrAgo = minutesAgo(r.arrived, nowMs);
      const picked = isPickedRow(kind, r.arrived) || (kind.kind === 'none' && r.driver && arrAgo != null && arrAgo >= PICKED_GUESS_MIN);
      r.state = hasReason(r.reason) || kind.kind === 'black' ? 'cancelled' : kind.kind === 'red' ? 'dup'
        : picked ? 'picked' : r.driver ? 'assigned' : 'pending';
      if (r.state === 'picked' && tonight && kind.kind === 'color') { const s = colorCounts[isPastel(kind.color) ? 'B' : 'A']; s[kind.color] = (s[kind.color] || 0) + 1; }
      if (r.state === 'cancelled' || r.state === 'dup') { if (ts >= editCutoff) rides.push(r); return; }
      // possible duplicate: same route, same name and same phone within 40 min (a suggestion only)
      const same = s => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
      const me = { route: same(r.pickup) + '>' + same(r.dropoff), phone: r.phone.replace(/\D/g, '').slice(-10), name: norm(r.first + r.last) };
      const dup = me.phone.length === 10 && me.name.length > 3
        ? prior.filter(q => q.ts >= ts - DUP_WINDOW_MIN * 60e3 && q.route === me.route && q.phone === me.phone && q.name === me.name).pop() : null;
      if (dup) r.dupOf = { row: dup.row, first: dup.first, time: dup.time };
      prior.push(Object.assign({ ts, row, first: r.first, time: r.time }, me));
      const w = waitMinutes(r.call, r.arrived);
      if (w != null && tonight) waits.push({ row, wait: w });
      if (r.driver) seen[crewKey(r)] = { driver: r.driver, nav: r.nav, van: r.van, loc: r.loc };
      if (ts >= editCutoff || (r.state !== 'picked' && tonight)) rides.push(r);
    });

    const roster = rosterRows ? parseRoster(rosterRows, nowMs) : null;
    const shiftsOn = activeShifts(roster, toggles, nowMs);
    const t = nightMin(nowMs), crews = [];
    const add = c => { if (!crews.some(x => samePair(x, c))) crews.push(Object.assign({ key: crewKey(c) }, c)); };
    if (roster) roster.shifts.filter(s => shiftsOn.includes(s.name)).forEach(s => s.crews.forEach(c =>
      add(Object.assign({ shift: s.name }, c, c.move && t >= c.move.at ? { van: c.move.van, vehicle: c.move.vehicle, cap: c.move.cap, phone: c.move.phone || c.phone } : {}))));
    (extraCrews || []).forEach(c => add(Object.assign({ shift: 'Added' }, c)));
    rides.filter(r => r.state === 'assigned').forEach(r => add(Object.assign({ shift: 'On a ride' }, seen[crewKey(r)] || { driver: r.driver, nav: r.nav, van: r.van })));
    crews.forEach(c => { c.riders = rides.filter(r => r.state === 'assigned' && samePair(r, c)); });

    const top = m => Object.keys(m).sort((a, b) => m[b] - m[a])[0] || null;
    const det = { A: top(colorCounts.A), B: top(colorCounts.B) };
    const ch = chosenColors && chosenColors.night === night ? chosenColors : {};
    let A = ch.A || det.A || null, B = ch.B || det.B || null;
    if (!B && A) B = pastelOf(A);
    if (!A && B) A = brightOf(B);

    const recent = waits.sort((a, b) => b.row - a.row).slice(0, 5);
    return {
      night, rides, roster, shiftsOn, crews,
      pending: rides.filter(r => r.state === 'pending'),
      active: rides.filter(r => r.state === 'assigned'),
      wait: recent.length ? Math.round(recent.reduce((s, x) => s + x.wait, 0) / recent.length) : null,
      colors: { A, B, detected: det },
    };
  }

  /* ---------- checks before writing (the fresh row is re-read first) ---------- */
  // fresh: { text: [A..W], fill }
  function assertEditable(fresh, row, nowMs) {
    nowMs = nowMs == null ? Date.now() : nowMs;
    const ts = parseTs(fresh.text[COL.TS - 1]);
    if (ts == null) throw new Error(`Row ${row} isn't a ride — not editing it.`);
    const kind = rowKind(fresh.fill);
    if (kind.kind === 'black' || kind.kind === 'red') throw new Error(`Row ${row} is marked ${kind.kind === 'red' ? 'duplicate (red)' : 'cancelled (black)'} — not editing it.`);
    const get = c => String(fresh.text[c - 1] || '').trim();
    const open = !(kind.kind === 'color' && get(COL.ARRIVE)) && !hasReason(get(COL.REASON)) && nightKey(ts) === nightKey(nowMs);
    if (!open && ts < nowMs - EDIT_WINDOW_MIN * 60e3) throw new Error(`Row ${row} is finished and older than ${EDIT_WINDOW_MIN} min — fix it by hand if needed.`);
  }
  // Seats + golf-cart rules for an assignment. riders: the rides being added; crew: from buildNight().crews.
  function checkAssign(crew, riders, van) {
    const cap = Number(crew.cap) || 0, have = (crew.riders || []).filter(r => !riders.some(x => x.row === r.row));
    if (cap) {
      const freed = have.filter(r => riders.every(p => typeof samePlace_ === 'function' ? samePlace_(r.dropoff, p.pickup) : false)).length;
      const load = have.length - freed;
      if (load + riders.length > cap) throw new Error(`${crew.vehicle || 'Their vehicle'} fits ${cap}: ${crewLabel(crew)} already ${load === 1 ? 'has 1 rider' : `has ${load} riders`}, so only ${Math.max(0, cap - load)} more fit.`);
    }
    if (/cart/i.test(van || '')) {
      const off = riders.filter(r => (r.zone || 'Off') !== 'On');
      if (off.length) throw new Error(`Golf carts only take on-campus rides: ${off.map(r => `${r.first} (${r.zone || 'Off'})`).join(', ')}.`);
      const far = riders.filter(r => typeof cartNoGo_ === 'function' && (cartNoGo_(r.pickup) || cartNoGo_(r.dropoff)));
      if (far.length) throw new Error(`Golf carts can't go to ECG, ECGA or the detached garages: ${far.map(r => r.first).join(', ')}.`);
    }
  }

  /* ---------- GroupMe posts (same wording as the panel) ---------- */
  const rideLine = r => { const t = s => String(s == null ? '' : s).trim(); return `${t(r.first)}  ${t(r.phone)}  ${t(r.pickup)} → ${t(r.dropoff)}`; };
  const pk = r => String(r.pickupName || r.pickup || '').toLowerCase(), pk2 = r => String(r.pickup || '').trim().toLowerCase();
  const byPickup = rs => rs.slice().sort((a, b) => pk(a).localeCompare(pk(b)) || pk2(a).localeCompare(pk2(b)) || a.row - b.row);
  function assignText(others, added) {
    const lines = byPickup(added).map(rideLine);
    return others.length ? ['Updated ride!'].concat(byPickup(others).map(rideLine), ['NEW:'], lines).join('\n') : lines.join('\n');
  }
  const removedText = (rest, gone, word) => (rest.length ? ['Updated ride!'].concat(byPickup(rest).map(rideLine)) : [])
    .concat([`${word}: ` + gone.map(r => String(r.first || '').trim()).join(', ')]).join('\n');

  return { TZ, COL, get LAST_COL() { return LAST_COL; }, READ_COLS, colLetter, colNumber, setLayout, EDIT_WINDOW_MIN, PICK_GAP_MIN, LOCATIONS, VANS, CANCEL_REASONS, DUPLICATE_COLOR, PASTEL_OF,
    parts, wallToMs, fmtTime, nightKey, nightMin, parseTs, hasTime, rideIdOf, norm, sameName, samePair, crewKey, crewLabel,
    rowKind, isColored, isPastel, pastelOf, brightOf, hasReason, parseRoster, autoShiftLetters, activeShifts, shadeFor,
    waitMinutes, buildNight, assertEditable, checkAssign, rideLine, byPickup, assignText, removedText };
})();
if (typeof module !== 'undefined') module.exports = R;
