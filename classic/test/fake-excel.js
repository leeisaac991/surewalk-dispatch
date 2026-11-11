// Fake Office.js / Excel over an in-memory workbook built from a CSV-style export (test only).
// Dates/times become real Excel serials with number formats, like the live workbook.
(() => {
  const W = 40;
  const get = f => { const q = new XMLHttpRequest(); q.open('GET', f, false); q.send(); return q.responseText; };
  const raw = JSON.parse(get('/sheet.json')), CONF = JSON.parse(get('/config.json'));
  const serial = (y, mo, d, h = 0, mi = 0, s = 0) => Date.UTC(y, mo - 1, d, h, mi, s) / 864e5 + 25569;
  const pad = n => String(n).padStart(2, '0');
  function typed(s) { // what Excel stores for typed text
    let m;
    if (s === '' || s == null) return { v: '', fmt: 'General' };
    if (typeof s === 'number') return { v: s, fmt: 'General' };
    if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{2})$/.exec(s))) return { v: serial(+m[3], +m[1], +m[2], +m[4], +m[5]), fmt: 'm/d/yyyy h:mm' };
    if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s))) return { v: serial(+m[3], +m[1], +m[2]), fmt: 'm/d/yyyy' };
    if ((m = /^(\d{1,2}):(\d{2})$/.exec(s))) return { v: (+m[1] * 60 + +m[2]) / 1440, fmt: 'h:mm' };
    if (/^\d{1,15}$/.test(s)) return { v: Number(s), fmt: 'General' };
    return { v: s, fmt: 'General' };
  }
  function text(c) {
    if (typeof c.v === 'number' && c.fmt !== 'General') {
      const d = new Date(Math.round((c.v - 25569) * 864e5));
      const date = `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`, time = `${d.getUTCHours()}:${pad(d.getUTCMinutes())}`;
      return c.fmt === 'h:mm' ? time : c.fmt === 'm/d/yyyy' ? date : `${date} ${time}`;
    }
    return c.v === '' || c.v == null ? '' : String(c.v);
  }
  const newCell = s => Object.assign(typed(s), { bg: '#FFFFFF', fc: '#000000' });
  const grid = raw.map(r => { const line = []; for (let j = 0; j < W; j++) line.push(newCell(r[j] == null ? '' : r[j])); return line; });
  const cell = (r, c) => { while (grid.length <= r) { const l = []; for (let j = 0; j < W; j++) l.push(newCell('')); grid.push(l); } return grid[r][c]; };
  (CONF.paint || []).forEach(([r, color, font]) => { for (let j = 0; j < W; j++) { cell(r - 1, j).bg = color; if (font) cell(r - 1, j).fc = font; } });
  const writes = [];
  const colN = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const colL = n => { let s = ''; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  class Rng {
    constructor(r, c, h, w, a) { this.r = r; this.c = c; this.h = h; this.w = w; this.a = a || `${colL(c + 1)}${r + 1}`; const self = this;
      this.format = {
        fill: { set color(v) { self.each(x => { x.bg = v.toUpperCase(); }); writes.push(`${self.a} fill ${v}`); }, clear() { self.each(x => { x.bg = '#FFFFFF'; }); writes.push(`${self.a} fill cleared`); } },
        font: { set color(v) { self.each(x => { x.fc = v.toUpperCase(); }); writes.push(`${self.a} font ${v}`); } },
      }; }
    each(fn) { for (let i = 0; i < this.h; i++) for (let j = 0; j < this.w; j++) fn(cell(this.r + i, this.c + j), i, j); }
    grid(fn) { const o = []; for (let i = 0; i < this.h; i++) { const l = []; for (let j = 0; j < this.w; j++) l.push(fn(cell(this.r + i, this.c + j))); o.push(l); } return o; }
    load() { return this; }
    get values() { return this.grid(x => x.v); }
    set values(g) { this.each((x, i, j) => { const t = typeof g[i][j] === 'string' ? typed(g[i][j]) : { v: g[i][j], fmt: x.fmt }; x.v = t.v; if (x.fmt === 'General' || typeof g[i][j] === 'string') x.fmt = t.fmt === 'General' && typeof x.v === 'number' && x.fmt !== 'General' ? x.fmt : t.fmt; }); writes.push(`${this.a}=${JSON.stringify(g)}`); }
    get text() { return this.grid(text); }
    get formulas() { return this.grid(x => x.v); }
    get numberFormat() { return this.grid(x => x.fmt); }
    set numberFormat(g) { this.each((x, i, j) => { x.fmt = g[i][j]; }); }
    clear() { this.each(x => { x.v = ''; }); writes.push(`${this.a} cleared`); }
    getCellProperties() { return { value: this.grid(x => ({ format: { fill: { color: x.bg }, font: { color: x.fc } } })) }; }
  }
  const lastRow = () => { for (let r = grid.length; r > 0; r--) if (grid[r - 1].some(x => x.v !== '')) return r; return 0; };
  const ws = { name: 'Sheet1', load() {},
    getRangeByIndexes: (r, c, h, w) => new Rng(r, c, h, w),
    getRange(a) { const m = /^(\d+):(\d+)$/.exec(a); return new Rng(+m[1] - 1, 0, m[2] - m[1] + 1, W, a); },
    getUsedRangeOrNullObject() { return { isNullObject: false, rowIndex: 0, rowCount: lastRow(), load() {} }; } };
  const workbook = { name: 'Sure Walk Fall 2026.xlsx', load() {}, worksheets: { getFirst: () => ws, getItem: () => ws } };
  window.Excel = { run: async fn => fn({ workbook, sync: async () => {} }), ClearApplyTo: { contents: 'Contents' } };
  // Office, including a stand-in for Excel's dialog window: a test window whose messages go through the same calls
  const EventType = { DialogMessageReceived: 'dm', DialogEventReceived: 'de', DialogParentMessageReceived: 'dp' };
  const AsyncResultStatus = { Succeeded: 'succeeded', Failed: 'failed' };
  window.Office = { onReady: cb => setTimeout(() => cb({ host: 'Excel' }), 0), EventType, AsyncResultStatus,
    context: { ui: { displayDialogAsync(url, opts, cb) {
      const w = window.open(url, 'swc_dialog', 'popup=yes,width=1200,height=850');
      if (!w) { cb({ status: 'failed', error: { message: 'blocked', code: 12009 } }); return; }
      const on = {}, queue = [];
      w.__swcQueue = queue;
      window.__fakeDlg = { toParent: s => on.dm && on.dm({ message: s }) };
      const dialog = { addEventHandler(t, fn) { on[t] = fn; }, close() { w.close(); },
        messageChild(s) { if (w.__fromParent) w.__fromParent(s); else queue.push(s); } };
      const watch = setInterval(() => { if (w.closed) { clearInterval(watch); if (on.de) on.de({ error: 12006 }); } }, 500);
      dialog.close = () => { clearInterval(watch); w.close(); };
      cb({ status: 'succeeded', value: dialog });
    } } } };
  window.__xl = { writes, text: (r, c) => text(cell(r - 1, colN(c) - 1)), bg: (r, c) => cell(r - 1, colN(c) - 1).bg, fc: (r, c) => cell(r - 1, colN(c) - 1).fc,
    type(r, vals) { vals.forEach((s, j) => Object.assign(cell(r - 1, j), typed(s))); }, paint(r, color) { for (let j = 0; j < W; j++) cell(r - 1, j).bg = color; } };
  // the test clock, here and in the worker
  const RealDate = Date, T0 = CONF.now - (CONF.skew || 0), R0 = RealDate.now(); // skew: this computer's clock runs behind GroupMe's
  window.Date = class extends RealDate { constructor(...a) { super(...(a.length ? a : [T0 + RealDate.now() - R0])); } static now() { return T0 + RealDate.now() - R0; } };
  const clock = `(() => { const R = Date, T0 = ${T0}, R0 = R.now() - ${RealDate.now() - R0}; self.Date = class extends R { constructor(...a) { super(...(a.length ? a : [T0 + R.now() - R0])); } static now() { return T0 + R.now() - R0; } }; })();\n`;
  window.__SWC = { backend: clock + get('/dist/backend.js'), sidebar: get('/dist/Sidebar.html') };
})();
