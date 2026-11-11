// Test server: fake Microsoft Graph (an in-memory workbook from a real sheet export) + fake GroupMe,
// and the extension's panel served with a chrome.* stub. node fake-server.js <export.json> [port]
const fs = require('fs'), path = require('path'), http = require('http');
const EXT = path.join(__dirname, '..');
const EXPORT = process.argv[2], PORT = +(process.argv[3] || 8777);
const W = 30;
const grid = JSON.parse(fs.readFileSync(EXPORT, 'utf8')).map(r => { const t = r.map(x => (x == null ? '' : String(x))); while (t.length < W) t.push(''); return t; });
const fill = grid.map(() => Array(W).fill('#FFFFFF')), font = grid.map(() => Array(W).fill('#000000'));
const paint = (r, c) => { for (let j = 0; j < W; j++) fill[r - 1][j] = c; };
// rides already picked tonight (driver + arrival) are colored, as in the real sheet
grid.forEach((r, i) => { if (i > 0 && r[14] && r[3]) paint(i + 1, '#00FF00'); });
paint(1561, '#000000');
const log = { writes: [], posts: [], deletes: [], calls: 0 };

const colN = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
function parseAddr(a) {
  let m = /^(\d+):(\d+)$/.exec(a);
  if (m) return { r1: +m[1], r2: +m[2], c1: 1, c2: W };
  m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a);
  return { r1: +m[2], r2: +(m[4] || m[2]), c1: colN(m[1]), c2: colN(m[3] || m[1]) };
}
const lastRow = () => { for (let r = grid.length; r > 0; r--) if (grid[r - 1].some(x => x !== '')) return r; return 0; };
const ensure = r => { while (grid.length < r) { grid.push(Array(W).fill('')); fill.push(Array(W).fill('#FFFFFF')); font.push(Array(W).fill('#000000')); } };

function graph(method, url, body) {
  log.calls++;
  url = decodeURIComponent(url);
  let m;
  if (/^\/me\b/.test(url)) return [200, { displayName: 'Test Dispatcher' }];
  if (/^\/shares\/[^/]+\/driveItem/.test(url)) return [200, { id: 'ITEM', name: 'Master Form 2026-27.xlsx', parentReference: { driveId: 'DRV' } }];
  if (/\/workbook\/worksheets\?/.test(url)) return [200, { value: [{ name: 'MASTER 2026-27', position: 0 }] }];
  if (/\/workbook\/createSession$/.test(url)) return [201, { id: 'SESSION1' }];
  if (/\/usedRange\(valuesOnly=true\)/.test(url)) return [200, { rowIndex: 0, rowCount: lastRow() }];
  if ((m = /\/range\(address='([^']+)'\)(\/[^?]*)?/.exec(url))) {
    const a = parseAddr(m[1]), sub = m[2] || '';
    ensure(a.r2);
    const each = fn => { for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) fn(r, c); };
    if (method === 'GET' && !sub) { const t = []; for (let r = a.r1; r <= a.r2; r++) t.push(grid[r - 1].slice(a.c1 - 1, a.c2)); return [200, { text: t }]; }
    if (method === 'PATCH' && !sub) {
      body.values.forEach((row, i) => row.forEach((v, j) => { grid[a.r1 - 1 + i][a.c1 - 1 + j] = String(v); }));
      log.writes.push(`${m[1]}=${JSON.stringify(body.values)}`); return [200, {}];
    }
    if (method === 'POST' && sub === '/clear') { each((r, c) => { grid[r - 1][c - 1] = ''; }); log.writes.push(`${m[1]} cleared`); return [200, {}]; }
    if (sub === '/format/fill' && method === 'GET') { const s = new Set(); each((r, c) => s.add(fill[r - 1][c - 1])); return [200, { color: s.size === 1 ? [...s][0] : null }]; }
    if (sub === '/format/fill' && method === 'PATCH') { each((r, c) => { fill[r - 1][c - 1] = body.color.toUpperCase(); }); log.writes.push(`${m[1]} fill ${body.color}`); return [200, {}]; }
    if (sub === '/format/fill/clear') { each((r, c) => { fill[r - 1][c - 1] = '#FFFFFF'; }); log.writes.push(`${m[1]} fill cleared`); return [200, {}]; }
    if (sub === '/format/font' && method === 'GET') return [200, { color: font[a.r1 - 1][a.c1 - 1] }];
    if (sub === '/format/font' && method === 'PATCH') { each((r, c) => { font[r - 1][c - 1] = body.color.toUpperCase(); }); log.writes.push(`${m[1]} font ${body.color}`); return [200, {}]; }
  }
  return [404, { error: { code: 'NotFound', message: 'fake graph: ' + method + ' ' + url } }];
}
function groupme(method, url, body) {
  if (url === '/users/me') return [200, { response: { name: 'Test Dispatcher', user_id: '1' } }];
  if (url.startsWith('/groups?')) return [200, { response: [{ id: 'G1', name: 'SURE Walk Dispatch' }] }];
  if (url === '/groups/G1') return [200, { response: { members: [{ user_id: '11', nickname: 'Alex Test' }, { user_id: '12', nickname: 'Sam Test' }, { user_id: '13', nickname: 'Jordan' }] } }];
  if (url === '/groups/G1/messages' && method === 'POST') { const id = 'M' + (log.posts.length + 1); log.posts.push(Object.assign({ id }, body.message)); return [201, { response: { message: { id } } }]; }
  if (/^\/conversations\/G1\/messages\//.test(url) && method === 'DELETE') { log.deletes.push(url.split('/').pop()); return [204, null]; }
  return [404, { meta: { code: 404 } }];
}

const STUB = `<script>
(() => {
  const RealDate = Date, T0 = RealDate.UTC(2026, 10, 4, 6, 33, 0), R0 = RealDate.now();
  window.Date = class extends RealDate { constructor(...a) { super(...(a.length ? a : [T0 + RealDate.now() - R0])); } static now() { return T0 + RealDate.now() - R0; } };
  const mem = { settings: { dispatcher: 'T. Tester', clientId: 'test-client', tenant: 'organizations', fileLink: 'https://utexas-my.sharepoint.com/:x:/g/personal/x/abc',
      sheet: '', gmToken: 'testtokentesttokentest1234', gmGroupId: 'G1', gmGroupName: 'SURE Walk Dispatch', autoPost: true,
      graphBase: location.origin + '/graph', loginBase: location.origin + '/login', gmBase: location.origin + '/gm' },
    msToken: { access: 'test', refresh: 'r', exp: RealDate.now() + 36e5 } };
  window.chrome = { storage: { local: {
      get: async k => { const o = {}; [].concat(k).forEach(x => { if (x in mem) o[x] = JSON.parse(JSON.stringify(mem[x])); }); return o; },
      set: async o => { Object.assign(mem, JSON.parse(JSON.stringify(o))); }, remove: async k => { [].concat(k).forEach(x => delete mem[x]); } } },
    identity: { getRedirectURL: () => 'https://cegmolcbfonbjcbgapchnpffnmcknebg.chromiumapp.org/', launchWebAuthFlow: async () => { throw new Error('no'); } },
    runtime: { getURL: p => '/' + p }, tabs: { create() {} } };
})();
</script>`;

http.createServer((req, res) => {
  let body = ''; req.on('data', c => { body += c; }); req.on('end', () => {
    const send = (code, obj, type) => { res.writeHead(code, { 'Content-Type': type || 'application/json' }); res.end(obj == null ? '' : typeof obj === 'string' || Buffer.isBuffer(obj) ? obj : JSON.stringify(obj)); };
    const u = req.url;
    if (u === '/log') return send(200, log);
    if (u.startsWith('/graph/')) {
      if ((req.headers.authorization || '') !== 'Bearer test') return send(401, { error: { code: 'InvalidAuthenticationToken', message: 'no token' } });
      const p = u.slice(6), b = body ? JSON.parse(body) : undefined;
      if (p === '/$batch') return send(200, { responses: b.requests.map(q => { const [s, o] = graph(q.method, q.url, q.body); return { id: q.id, status: s, body: o }; }) });
      const [s, o] = graph(req.method, p, b); return send(s, o);
    }
    if (u.startsWith('/gm/')) { const [s, o] = groupme(req.method, u.slice(3), body ? JSON.parse(body) : undefined); return send(s, o); }
    if (u === '/sheet.json') return send(200, fs.readFileSync(EXPORT));
    if (u === '/fake-office.js') return send(200, fs.readFileSync(path.join(__dirname, 'fake-office.js')), 'text/javascript');
    if (u.startsWith('/addin/') && process.env.ADDIN_DIST) { // the built Excel add-in, with the fake Office.js
      const f = u.split('?')[0].slice(7) || 'taskpane.html';
      let data; try { data = fs.readFileSync(path.join(process.env.ADDIN_DIST, f)); } catch (e) { return send(404, 'not found', 'text/plain'); }
      if (f.endsWith('.html')) data = String(data).replace('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', '/fake-office.js');
      return send(200, data, f.endsWith('.js') ? 'text/javascript' : f.endsWith('.css') ? 'text/css' : 'text/html');
    }
    const f = u === '/' ? 'panel.html' : u.split('?')[0].slice(1);
    let data; try { data = fs.readFileSync(path.join(EXT, f)); } catch (e) { return send(404, 'not found', 'text/plain'); }
    if (f === 'panel.html') data = String(data).replace('<script src="lib/places.js">', STUB + '<script src="lib/places.js">');
    send(200, data, f.endsWith('.js') ? 'text/javascript' : f.endsWith('.css') ? 'text/css' : f.endsWith('.png') ? 'image/png' : 'text/html');
  });
}).listen(PORT, () => console.log('fake server on ' + PORT));
module.exports = { grid, fill };
