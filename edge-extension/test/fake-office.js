// Fake Office.js / Excel API over an in-memory workbook (test only). Loaded in place of office.js.
(() => {
  const W = 30;
  const xhr = new XMLHttpRequest(); xhr.open('GET', '/sheet.json', false); xhr.send();
  const grid = JSON.parse(xhr.responseText).map(r => { const t = r.map(x => (x == null ? '' : String(x))); while (t.length < W) t.push(''); return t; });
  const fill = grid.map(() => Array(W).fill('#FFFFFF')), font = grid.map(() => Array(W).fill('#000000'));
  grid.forEach((r, i) => { if (i > 0 && r[14] && r[3]) fill[i].fill('#00FF00'); }); // picked tonight = colored
  fill[1560].fill('#000000');
  const writes = [];
  const colN = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  function parse(a) {
    let m = /^(\d+):(\d+)$/.exec(a);
    if (m) return { r1: +m[1], r2: +m[2], c1: 1, c2: W };
    m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a);
    return { r1: +m[2], r2: +(m[4] || m[2]), c1: colN(m[1]), c2: colN(m[3] || m[1]) };
  }
  const ensure = r => { while (grid.length < r) { grid.push(Array(W).fill('')); fill.push(Array(W).fill('#FFFFFF')); font.push(Array(W).fill('#000000')); } };
  class Rng {
    constructor(a) { this.a = a; this.x = parse(a); ensure(this.x.r2); const self = this;
      this.format = {
        fill: { set color(v) { self.each((r, c) => { fill[r][c] = v.toUpperCase(); }); writes.push(`${a} fill ${v}`); }, clear() { self.each((r, c) => { fill[r][c] = '#FFFFFF'; }); writes.push(`${a} fill cleared`); } },
        font: { set color(v) { self.each((r, c) => { font[r][c] = v.toUpperCase(); }); writes.push(`${a} font ${v}`); } },
      }; }
    each(fn) { for (let r = this.x.r1 - 1; r < this.x.r2; r++) for (let c = this.x.c1 - 1; c < this.x.c2; c++) fn(r, c); }
    grid(src) { const o = []; for (let r = this.x.r1 - 1; r < this.x.r2; r++) o.push(src[r].slice(this.x.c1 - 1, this.x.c2)); return o; }
    load() { return this; }
    get text() { return this.grid(grid); }
    set values(v) { v.forEach((row, i) => row.forEach((x, j) => { grid[this.x.r1 - 1 + i][this.x.c1 - 1 + j] = String(x); })); writes.push(`${this.a}=${JSON.stringify(v)}`); }
    clear() { this.each((r, c) => { grid[r][c] = ''; }); writes.push(`${this.a} cleared`); }
    getCellProperties() { const f = this.grid(fill), t = this.grid(font); return { value: f.map((row, i) => row.map((c, j) => ({ format: { fill: { color: c }, font: { color: t[i][j] } } }))) }; }
  }
  const lastRow = () => { for (let r = grid.length; r > 0; r--) if (grid[r - 1].some(x => x !== '')) return r; return 0; };
  const ws = { name: 'MASTER 2026-27', load() {}, getRange: a => new Rng(a),
    getUsedRangeOrNullObject() { return { isNullObject: false, rowIndex: 0, rowCount: lastRow(), load() {} }; } };
  const workbook = { name: 'Master Form 2026-27.xlsx', load() {}, worksheets: { getFirst: () => ws, getItem: n => ws } };
  window.Excel = { run: async fn => fn({ workbook, sync: async () => {} }), ClearApplyTo: { contents: 'Contents' } };
  window.Office = { HostType: { Excel: 'Excel' }, onReady: cb => setTimeout(() => cb({ host: 'Excel' }), 0) };
  window.__xlWrites = writes;
  // test clock (11/4/2026 0:33 AM Austin) + settings: GroupMe goes to the fake server
  const RealDate = Date, T0 = RealDate.UTC(2026, 10, 4, 6, 33, 0), R0 = RealDate.now();
  window.Date = class extends RealDate { constructor(...a) { super(...(a.length ? a : [T0 + RealDate.now() - R0])); } static now() { return T0 + RealDate.now() - R0; } };
  localStorage.setItem('swd.settings', JSON.stringify({ dispatcher: 'T. Tester', gmToken: 'testtokentesttokentest1234', gmGroupId: 'G1', gmGroupName: 'SURE Walk Dispatch', autoPost: true, gmBase: location.origin + '/gm' }));
})();
