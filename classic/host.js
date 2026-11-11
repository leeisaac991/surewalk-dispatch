/*
 * SURE Walk Dispatch (classic) — runs the ORIGINAL dispatch panel inside Excel on the web.
 *
 *   Sidebar.html  : the original panel page, unchanged.
 *   backend.js    : the original Code.gs / GroupMe.gs / Crews.gs / Places.gs + Excel.gs (a stand-in sheet) +
 *                   services.js (stand-ins for Google's services), run in a Web Worker in this browser.
 *   host.js       : this file. Stands in for google.script.run: before each panel call it reads the rides sheet
 *                   with Office.js, presents it in the OLD Google-sheet column layout, runs the call in the worker,
 *                   and writes the call's changes back into the right Excel cells.
 *
 * Expects window.__SWC = { backend: <backend.js text>, sidebar: <Sidebar.html text> } (set by the Script Lab loader).
 */
(function () {
  'use strict';
  const SWC = window.__SWC || {};
  const READ_COLS = 30;     // Excel columns read (A:AD)
  const VW = 26;            // columns the original code sees (A:Z of the old layout)
  const SCAN_ROWS = 600, FIND_ROWS = 1000, MAX_AGE_MS = 8000;

  /* ---------------- this browser's settings and the panel's saved state ---------------- */
  const ls = {
    get(k, d) { try { const v = localStorage.getItem('swc.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem('swc.' + k, JSON.stringify(v)); } catch (e) { /* this session only */ } },
  };
  let uid = ls.get('uid', '');
  if (!uid) { const a = new Uint8Array(16); crypto.getRandomValues(a); uid = Array.from(a, b => b.toString(16).padStart(2, '0')).join(''); ls.set('uid', uid); }
  let props = ls.get('props', {}); // the panel's script properties (crews, colors, GroupMe login, …)
  let propsSent = false;

  /* ---------------- dates: Excel serial <-> Austin wall clock ---------------- */
  const pad = n => String(n).padStart(2, '0');
  function serialToWall(serial) {
    const d = new Date(Math.round((serial - 25569) * 864e5));
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
  }
  function wallToSerial(t) {
    const m = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/.exec(t);
    return m ? Date.UTC(+m[1], m[2] - 1, +m[3], +m[4], +m[5], +m[6]) / 864e5 + 25569 : t;
  }
  // "11/10/2026 23:14", "11/10/26 11:14 PM" typed as text -> wall clock (so the code sees a real timestamp)
  function textToWall(s) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?$/i.exec(String(s || '').trim());
    if (!m) return null;
    let h = +m[4];
    if (m[7]) h = (h % 12) + (/pm/i.test(m[7]) ? 12 : 0);
    const y = m[3].length === 2 ? '20' + m[3] : m[3];
    return `${y}-${pad(m[1])}-${pad(m[2])} ${pad(h)}:${m[5]}:${m[6] || '00'}`;
  }
  const bareFmt = f => String(f || '').replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  const isDateFmt = f => !/^general$/i.test(String(f || '')) && /[dmyhs]/i.test(bareFmt(f)) && !/^[0#.,%E+-]*$/i.test(bareFmt(f));
  function showDate(serial, f) {
    const t = serialToWall(serial), Y = t.slice(0, 4), M = +t.slice(5, 7), D = +t.slice(8, 10);
    const h = +t.slice(11, 13), mi = t.slice(14, 16), s = t.slice(17, 19), b = bareFmt(f).toLowerCase();
    const ampm = /am\/pm|a\/p/.test(b), hh = ampm ? ((h + 11) % 12) + 1 : h;
    const date = /[dy]/.test(b) ? `${M}/${D}/${/yyyy/.test(b) ? Y : Y.slice(2)}` : '';
    const time = /h/.test(b) ? `${hh}:${mi}${/s/.test(b) ? ':' + s : ''}${ampm ? (h < 12 ? ' AM' : ' PM') : ''}` : '';
    return [date, time].filter(Boolean).join(' ');
  }

  /* ---------------- the old layout <- the Excel layout ---------------- */
  // Old Google-sheet columns (what the original code's COL uses) and how to find each one in Excel's header row.
  const OLD = [
    [1, /^(form submitted|timestamp|completion time)/], [2, /^dispatch/], [3, /^call time/], [4, /^arrival/], [5, /^wait/],
    [6, /\beid\b/], [7, /^last name/], [8, /^first name/], [9, /^phone/], [10, /^pick ?-?up/], [11, /^drop ?-?off/],
    [12, /disabilit|^ada\b/], [13, /^location$/], [14, /^(van|vehicle)/], [15, /^driver$/], [16, /^navigator$/],
    [22, /cancel/], [23, /^notes?$/],
  ];
  function layoutOf(header) {
    const h = header.map(x => String(x || '').toLowerCase().replace(/[^a-z0-9 -]/g, ' ').replace(/\s+/g, ' ').trim());
    const map = {}, used = new Set();
    OLD.forEach(([k, re]) => { const i = h.findIndex((x, j) => re.test(x) && !used.has(j)); if (i >= 0) { map[k] = i + 1; used.add(i); } });
    const missing = [1, 8, 9, 10, 11, 15, 16, 3, 4].filter(k => !map[k]);
    if (missing.length) throw new Error("Can't find the rides columns in row 1 of this tab (Form Submitted / Timestamp, First Name, Phone, Pick-up, Drop-off, Driver, Navigator, Call Time, Arrival Time). Pick the rides tab in ⚙ below.");
    // the rest of the old columns (agreements, email…) take the remaining Excel columns in order
    let next = 1;
    for (let k = 1; k <= VW; k++) {
      if (map[k]) continue;
      while (next <= READ_COLS && used.has(next - 1)) next++;
      if (next <= READ_COLS) { map[k] = next; used.add(next - 1); next++; }
    }
    return { map, ts: map[1] };
  }
  // Excel column for an old-layout column on a given row: ride rows use the header map; the roster block below
  // the rides is the old block shifted to start in the timestamp column.
  const xcolOf = (L, row, k) => (row <= L.lastRide ? L.map[k] : k + L.ts - 1);

  /* ---------------- reading the workbook ---------------- */
  let conf = { sheet: ls.get('sheet', '') };
  const sheetOf = ctx => (conf.sheet ? ctx.workbook.worksheets.getItem(conf.sheet) : ctx.workbook.worksheets.getFirst());
  let cache = null, cacheAt = 0, dirty = true;

  async function readSheet() {
    if (cache && !dirty && Date.now() - cacheAt < MAX_AGE_MS) return cache;
    dirty = false;
    const got = await Excel.run(async ctx => {
      const ws = sheetOf(ctx);
      ws.load('name'); ctx.workbook.load('name');
      const used = ws.getUsedRangeOrNullObject(true);
      used.load('rowIndex,rowCount');
      const hdr = ws.getRangeByIndexes(0, 0, 1, READ_COLS); hdr.load('text');
      await ctx.sync();
      const lastRow = used.isNullObject ? 1 : used.rowIndex + used.rowCount;
      const L = layoutOf(hdr.text[0]);
      const aTop = Math.max(1, lastRow - FIND_ROWS + 1);
      const tsCol = ws.getRangeByIndexes(aTop - 1, L.ts - 1, lastRow - aTop + 1, 1); tsCol.load('text');
      await ctx.sync();
      L.lastRide = 1;
      for (let i = tsCol.text.length - 1; i >= 0; i--) if (aTop + i > 1 && /\d{1,2}:\d{2}/.test(tsCol.text[i][0])) { L.lastRide = aTop + i; break; }
      const top = Math.max(2, L.lastRide - SCAN_ROWS + 1);
      const parts = [{ r: 1, n: 1 }];
      if (lastRow >= top) parts.push({ r: top, n: lastRow - top + 1 });
      const loaded = parts.map(p => {
        const rg = ws.getRangeByIndexes(p.r - 1, 0, p.n, READ_COLS);
        rg.load('values,text,formulas,numberFormat');
        return { p, rg, props: rg.getCellProperties({ format: { fill: { color: true }, font: { color: true } } }) };
      });
      await ctx.sync();
      const cells = {}; // row -> array of Excel cells {t, v, f, fmt, bg, fc}
      loaded.forEach(({ p, rg, props: pr }) => {
        for (let i = 0; i < p.n; i++) {
          const line = [];
          for (let j = 0; j < READ_COLS; j++) {
            const c = pr.value[i][j].format;
            line.push({ t: rg.text[i][j] == null ? '' : String(rg.text[i][j]), v: rg.values[i][j], f: rg.formulas[i][j], fmt: rg.numberFormat[i][j],
              bg: String((c.fill && c.fill.color) || '#ffffff').toLowerCase(), fc: String((c.font && c.font.color) || '#000000').toLowerCase() });
          }
          cells[p.r + i] = line;
        }
      });
      return { book: ctx.workbook.name, name: ws.name, lastRow, top, L, cells };
    });
    await clearCopiedColors(got);
    cache = got; cacheAt = Date.now();
    return got;
  }

  /*
   * Excel gives each new form response the formatting of the row above, so after a picked (colored) or cancelled
   * (black) ride, new rides arrive already highlighted. A ride nobody has touched yet (no dispatcher, call time,
   * arrival, crew or cancel reason) whose color is exactly the row above's had it copied: clear it. Only rides from
   * the last hour are touched (the panel never edits older rows).
   */
  const UNTOUCHED = [2, 3, 4, 13, 14, 15, 16, 22]; // old-layout columns: dispatch, call, arrival, location…navigator, reason
  async function clearCopiedColors(x) {
    const L = x.L, rideCols = [1, 2, 3, 4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16].map(k => L.map[k]).filter(Boolean);
    const rowFill = r => { // the row's color when every rides column has it, else ''
      const c = x.cells[r];
      if (!c) return '';
      const first = c[rideCols[0] - 1].bg;
      return rideCols.every(xc => c[xc - 1].bg === first) ? first : '';
    };
    const now = Date.now(), fixes = [];
    for (let r = Math.max(3, x.top + 1); r <= L.lastRide; r++) {
      const c = x.cells[r], above = rowFill(r - 1), mine = rowFill(r);
      if (!c || !mine || mine === '#ffffff' || mine !== above) continue;
      if (UNTOUCHED.some(k => { const xc = L.map[k]; return xc && String(c[xc - 1].t).trim() && String(c[xc - 1].t).trim() !== '0'; })) continue;
      const w = textToWall(c[L.ts - 1].t) || (typeof c[L.ts - 1].v === 'number' ? serialToWall(c[L.ts - 1].v) : null);
      if (!w) continue;
      const [d, t] = w.split(' '), [y, mo, dd] = d.split('-').map(Number), [h, mi] = t.split(':').map(Number);
      const age = (now - (Date.UTC(y, mo - 1, dd, h, mi) - austinOffsetMs(Date.UTC(y, mo - 1, dd, h, mi)))) / 60e3;
      if (age < 0 || age > 60) continue;
      fixes.push(r);
    }
    if (!fixes.length) return;
    await Excel.run(async ctx => {
      const ws = sheetOf(ctx);
      fixes.forEach(r => { const rg = ws.getRange(`${r}:${r}`); rg.format.fill.clear(); rg.format.font.color = '#000000'; });
      await ctx.sync();
    });
    fixes.forEach(r => x.cells[r].forEach(z => { z.bg = '#ffffff'; z.fc = '#000000'; }));
  }
  // Austin's offset from UTC at a moment (ms), e.g. -6 h in winter
  function austinOffsetMs(ms) {
    const p = {};
    new Intl.DateTimeFormat('en-US', { timeZone: 'America/Chicago', year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', hourCycle: 'h23' })
      .formatToParts(new Date(ms)).forEach(z => { p[z.type] = +z.value; });
    return Date.UTC(p.year, p.month - 1, p.day, p.hour % 24, p.minute) - Math.floor(ms / 60e3) * 60e3;
  }

  // The snapshot Excel.gs expects, in the OLD column layout (see apps-script/Excel.gs XlSheet_).
  function virtualSnap(x) {
    const rows = {};
    Object.keys(x.cells).forEach(rk => {
      const r = +rk, src = x.cells[r], d = [], v = {}, f = {}, bg = [], fc = [];
      let any = false;
      for (let k = 1; k <= VW; k++) {
        const xc = xcolOf(x.L, r, k), c = xc ? src[xc - 1] : null;
        if (!c) { d.push(''); bg.push('#ffffff'); fc.push('#000000'); continue; }
        let t = c.t;
        if (typeof c.v === 'number' && isDateFmt(c.fmt)) { v[k] = { t: serialToWall(c.v) }; if (/^#+$/.test(t)) t = showDate(c.v, c.fmt); }
        else if (typeof c.v === 'number') { v[k] = { n: c.v }; if (/^#+$/.test(t)) t = String(c.v); }
        else if (typeof c.v === 'boolean') v[k] = { b: c.v };
        else if (k === 1 && textToWall(t)) v[k] = { t: textToWall(t) }; // a timestamp stored as text
        if (typeof c.f === 'string' && c.f.charAt(0) === '=') f[k] = c.f;
        d.push(t); if (t) any = true;
        bg.push(c.bg); fc.push(c.fc);
      }
      while (d.length && d[d.length - 1] === '') d.pop();
      const one = a => (a.every(c => c === a[0]) ? a[0] : a);
      const row = { d, bg: one(bg), fc: one(fc) };
      if (Object.keys(v).length) row.v = v;
      if (Object.keys(f).length) row.f = f;
      if (any || row.bg !== '#ffffff' || row.v) rows[r] = row;
    });
    return { book: x.book, name: x.name, lastRow: x.lastRow, lastCol: VW, width: VW, top: x.top, rows };
  }

  /* ---------------- writing: old-layout ops -> Excel ---------------- */
  function toExcelOps(ops, L) {
    const out = [];
    ops.forEach(op => {
      if (op.t === 'v' || op.t === 'clr') {
        for (let i = 0; i < op.h; i++) for (let j = 0; j < op.w; j++) {
          const row = op.r + i, xc = xcolOf(L, row, op.c + j);
          if (!xc) continue;
          out.push(op.t === 'v' ? { t: 'v', r: row, c: xc, vals: [[op.vals[i][j]]] } : { t: 'clr', r: row, c: xc });
        }
      } else if (op.t === 'bg' || op.t === 'fc') {
        for (let i = 0; i < op.h; i++) {
          const row = op.r + i;
          if (op.colors) {
            const line = op.colors[i];
            if (op.row) out.push({ t: op.t, row, color: line[line.length - 1] }); // columns past the read ones, as in Google
            line.forEach((col, j) => { const xc = xcolOf(L, row, op.c + j); if (xc) out.push({ t: op.t, r: row, c: xc, color: col }); });
          } else if (op.row) out.push({ t: op.t, row, color: op.color });
          else for (let j = 0; j < op.w; j++) { const xc = xcolOf(L, row, op.c + j); if (xc) out.push({ t: op.t, r: row, c: xc, color: op.color }); }
        }
      }
    });
    return out;
  }
  const toCell = v => (v && typeof v === 'object' && 't' in v ? wallToSerial(v.t) : v);
  async function applyOps(ops) {
    if (!ops.length) return;
    await Excel.run(async ctx => {
      const ws = sheetOf(ctx);
      ops.forEach(op => {
        const rg = op.row != null && op.c == null ? ws.getRange(`${op.row}:${op.row}`) : ws.getRangeByIndexes((op.r || op.row) - 1, op.c - 1, 1, 1);
        if (op.t === 'v') {
          const v = op.vals[0][0];
          rg.values = [[toCell(v)]];
          if (v && typeof v === 'object' && 't' in v) rg.numberFormat = [[String(v.t).slice(0, 4) === '1899' ? 'h:mm' : 'm/d/yyyy h:mm']];
        } else if (op.t === 'clr') rg.clear(Excel.ClearApplyTo.contents);
        else if (op.t === 'bg') { if (!op.color || op.color === '#ffffff') rg.format.fill.clear(); else rg.format.fill.color = op.color; }
        else if (op.t === 'fc') rg.format.font.color = op.color || '#000000';
      });
      await ctx.sync();
    });
    dirty = true;
  }

  /* ---------------- the original code, in a worker ---------------- */
  let worker = null, direct = null, seq = 0;
  const waiting = {};
  function startBackend() {
    try {
      worker = new Worker(URL.createObjectURL(new Blob([SWC.backend], { type: 'text/javascript' })));
      worker.onmessage = e => { const w = waiting[e.data.id]; if (w) { delete waiting[e.data.id]; w(e.data); } };
      worker.onerror = e => { console.error('backend', e.message); };
    } catch (e) { // workers not allowed here: run it on this page (GroupMe calls then pause the page briefly)
      worker = null;
      (0, eval)(SWC.backend); // eslint-disable-line no-eval
      direct = globalThis.__swCall;
    }
  }
  function backend(msg) {
    msg.id = ++seq;
    if (direct) return Promise.resolve(direct(msg));
    return new Promise(res => { waiting[msg.id] = res; worker.postMessage(msg); });
  }

  // Calls run one at a time, so each one sees the previous call's writes.
  let chain = Promise.resolve();
  const lastPeek = { key: '', result: null };
  function rpc(fn, args) {
    const job = chain.then(async () => {
      const x = await readSheet();
      const snap = virtualSnap(x);
      const peekKey = fn === 'peek' ? JSON.stringify(snap).length + '|' + hash(JSON.stringify(snap)) + '|' + JSON.stringify(args) : '';
      if (peekKey && peekKey === lastPeek.key) return lastPeek.result;
      const msg = { fn, args: args || [], snap, uid };
      if (!propsSent) { msg.props = props; propsSent = true; }
      const out = await backend(msg);
      if (out.props) { props = out.props; ls.set('props', props); }
      if (out.ops && out.ops.length) await applyOps(toExcelOps(out.ops, x.L));
      if (!out.ok) { lastPeek.key = ''; throw new Error(out.error || 'Error'); }
      if (peekKey) { lastPeek.key = peekKey; lastPeek.result = out.result; } else lastPeek.key = '';
      return out.result;
    });
    chain = job.catch(() => {});
    return job;
  }
  function hash(s) { let h = 0x811c9dc5; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193); } return (h >>> 0).toString(36); }

  /* ---------------- the original page: in the side panel, or popped out into its own window ---------------- */
  // Each copy of the page (a frame in the side panel, or the pop-out window) gets this small stand-in for
  // google.script.run that hands every call to this host — the only place that talks to Excel.
  const SHIM = `<script>(function () {
  var H = null;
  try { H = (window.parent !== window ? window.parent : window.opener).SWC_HOST; } catch (e) { /* no host */ }
  function gone() {
    document.body.innerHTML = '<div style="font:15px Segoe UI,system-ui,sans-serif;padding:24px;max-width:520px">' +
      '<h3>SURE Walk Dispatch</h3><p>The panel in Excel was closed or reloaded, so this window lost its connection to the sheet.</p>' +
      '<p>In Excel, open <b>Script Lab \\u25B8 Run</b> again, then press <b>\\u29C9 Pop out</b>.</p></div>';
  }
  function runner(ok, fail) {
    return new Proxy({}, { get: function (_, name) {
      if (name === 'withSuccessHandler') return function (h) { return runner(h, fail); };
      if (name === 'withFailureHandler') return function (h) { return runner(ok, h); };
      if (name === 'withUserObject') return function () { return runner(ok, fail); };
      if (typeof name !== 'string' || name === 'then') return undefined;
      return function () {
        var args = JSON.parse(JSON.stringify(Array.prototype.slice.call(arguments))), p;
        try { p = H.call(name, args); } catch (e) { gone(); return; }
        p.then(function (r) { if (ok) ok(r == null ? r : JSON.parse(JSON.stringify(r))); },
               function (e) { if (fail) fail(new Error((e && e.message) || String(e))); });
      };
    } });
  }
  window.google = { script: { run: runner(null, null), host: { close: function () {}, setHeight: function () {}, setWidth: function () {} },
    url: { getLocation: function (cb) { cb({ hash: '', parameter: {}, parameters: {} }); } } } };
  document.addEventListener('click', function (e) { // "New version available — Reload": reload the panel in Excel
    var a = e.target && e.target.closest && e.target.closest('a[href*="script.google.com/macros"]');
    if (a) { e.preventDefault(); try { H.reload(); } catch (x) { gone(); } }
  }, true);
  if (!H) document.addEventListener('DOMContentLoaded', gone);
})();<\/script>`;
  const pageHtml = () => SWC.sidebar.replace(/<base[^>]*>/i, '').replace(/<head>/i, '<head>' + SHIM);

  const BAR = 30; // the strip above the panel in Excel, with the Pop out button
  const barStyle = `position:fixed;top:0;left:0;right:0;height:${BAR}px;display:flex;align-items:center;gap:8px;padding:0 10px;` +
    'background:#bf5700;color:#fff;font:600 13px Segoe UI,system-ui,sans-serif;z-index:5;box-sizing:border-box';
  const btnStyle = 'border:0;border-radius:99px;padding:3px 12px;background:#fff;color:#bf5700;font:600 12px Segoe UI,system-ui,sans-serif;cursor:pointer';

  function showInPane() {
    document.body.style.margin = '0';
    document.body.innerHTML = `<div style="${barStyle}"><span>SURE Walk</span><span id="swcMsg" style="flex:1;font-weight:400;font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"></span>
      <button id="swcPop" style="${btnStyle}" title="Open the panel in its own window: resize it, maximize it, or move it to another screen">⧉ Pop out</button></div>
      <div id="swcNote" hidden style="position:fixed;top:${BAR}px;left:0;right:0;z-index:6;background:#fff4e5;color:#5a3000;border-bottom:1px solid #f0c48a;
        padding:8px 12px;font:13px Segoe UI,system-ui,sans-serif;line-height:1.4"></div>
      <iframe id="swcFrame" title="SURE Walk Dispatch" style="position:fixed;top:${BAR}px;left:0;width:100%;height:calc(100% - ${BAR}px);border:0"></iframe>`;
    document.getElementById('swcFrame').srcdoc = pageHtml();
    document.getElementById('swcPop').onclick = popOut;
  }
  // Excel add-ins can't show alert boxes, so problems are explained in a note under the bar.
  function note(html) {
    const n = document.getElementById('swcNote');
    if (!n) return;
    n.innerHTML = html + ' <a href="#" id="swcNoteX" style="color:#bf5700;font-weight:600;margin-left:6px">OK</a>';
    n.hidden = false;
    document.getElementById('swcNoteX').onclick = e => { e.preventDefault(); n.hidden = true; };
  }
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  // One copy of the panel at a time: while it's popped out, the side panel only keeps the connection to the sheet.
  function poppedOut(back) {
    document.body.innerHTML = `<div style="${barStyle}"><span>SURE Walk</span></div>
      <div style="padding:${BAR + 18}px 18px 18px;font:14px Segoe UI,system-ui,sans-serif;color:#333">
        <p style="margin:0 0 12px"><b>The panel is open in its own window.</b><br>Keep this side panel open: it's the window's connection to the sheet.</p>
        <button id="swcBack" style="${btnStyle};border:1px solid #bf5700">Bring it back here</button></div>`;
    document.getElementById('swcBack').onclick = back;
  }

  /*
   * Pop out. Excel on the web doesn't let add-ins open ordinary browser windows (window.open is blocked in the
   * add-in frame), so the panel uses Excel's own dialog window. Excel only opens a page from the add-in's own site
   * there, so this works in the SURE Walk add-in, not in Script Lab. The dialog page (dialog.html) shows the same panel and sends every call here.
   */
  const CHUNK = 24000;
  let dlg = null, dlgSeq = 0;
  const dlgParts = {};
  function dlgSend(obj) {
    if (!dlg) return;
    const s = JSON.stringify(obj), id = 'p' + (++dlgSeq), n = Math.max(1, Math.ceil(s.length / CHUNK));
    for (let i = 0; i < n; i++) dlg.messageChild('SWC|' + id + '|' + i + '|' + n + '|' + s.slice(i * CHUNK, (i + 1) * CHUNK));
  }
  function dlgReceive(str) {
    const m = /^SWC\|([^|]+)\|(\d+)\|(\d+)\|([\s\S]*)$/.exec(str);
    if (!m) return null;
    const p = dlgParts[m[1]] || (dlgParts[m[1]] = []);
    p[+m[2]] = m[4];
    if (p.filter(x => x != null).length < +m[3]) return null;
    delete dlgParts[m[1]];
    return JSON.parse(p.join(''));
  }
  function onDialogMessage(arg) {
    const m = dlgReceive(arg.message);
    if (!m) return;
    if (m.t === 'ready') dlgSend({ t: 'page', html: pageHtml() });
    else if (m.t === 'call') rpc(m.fn, m.args).then(r => dlgSend({ t: 'ret', id: m.id, ok: true, result: r }), e => dlgSend({ t: 'ret', id: m.id, ok: false, error: e.message || String(e) }));
    else if (m.t === 'reload') location.reload();
  }
  function bringBack() { try { if (dlg) dlg.close(); } catch (e) { /* already closed */ } dlg = null; showInPane(); }
  function popOut() {
    const ui = window.Office && Office.context && Office.context.ui;
    if (!window.__SWC_HOSTED || !ui || !ui.displayDialogAsync) {
      note('<b>Pop out only works in the SURE Walk add-in, not in Script Lab.</b> Excel blocks new windows from Script Lab. ' +
        'Install the add-in once (see <b>github.com/leeisaac991/surewalk-dispatch</b> ▸ classic), or widen this side panel by dragging its left edge.');
      return;
    }
    ui.displayDialogAsync(new URL('dialog.html', location.href).href, { height: 85, width: 75, displayInIframe: false }, res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        note(`<b>Excel couldn't open the window</b> (${esc(res.error && res.error.message)}, code ${esc(res.error && res.error.code)}).`);
        return;
      }
      dlg = res.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => { dlg = null; showInPane(); }); // closed with its ✕
      poppedOut(bringBack);
    });
  }
  // reloading the panel in Excel ends the window's connection, so the window closes with it
  window.addEventListener('beforeunload', () => { try { if (dlg) dlg.close(); } catch (e) { /* gone */ } });

  function problem(msg) {
    document.body.innerHTML = `<div style="font:14px Segoe UI,system-ui,sans-serif;padding:16px;max-width:460px">
      <h3 style="margin:0 0 8px">SURE Walk Dispatch</h3><p style="color:#c62828">${String(msg).replace(/</g, '&lt;')}</p>
      <label style="display:grid;gap:4px;font-size:13px">Rides tab (blank = first tab)<input id="swcTab" value="${String(conf.sheet).replace(/"/g, '&quot;')}" style="padding:6px"></label>
      <button id="swcGo" style="margin-top:10px;padding:8px 14px">Try again</button></div>`;
    document.getElementById('swcGo').onclick = () => { conf.sheet = document.getElementById('swcTab').value.trim(); ls.set('sheet', conf.sheet); start(); };
  }
  async function start() {
    try {
      if (!SWC.backend || !SWC.sidebar) throw new Error('The panel files did not load — check the internet connection and run again.');
      cache = null; dirty = true;
      await readSheet(); // checks the tab and its columns before showing the panel
      if (!worker && !direct) startBackend();
      if (dlg) return; // already showing in its own window
      showInPane();
    } catch (e) {
      problem(e && e.code === 'ItemNotFound' ? `There's no tab named "${conf.sheet}".` : (e.message || String(e)));
    }
  }
  window.SWC_HOST = { call: (fn, args) => rpc(fn, args), reload: () => location.reload(), popOut, // used by the panel page(s)
    readSheet, virtualSnap, toExcelOps, rpc, get props() { return props; } }; // console / tests
  Office.onReady(() => start());
})();
