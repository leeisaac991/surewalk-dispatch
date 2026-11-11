// Test server for the classic panel: serves dist/ + a page with the fake Excel, and a fake GroupMe API at /gm.
//   NOW=<ISO> PAINT='[[row,"#00FF00"],…]' node server.js <sheet.json> [port]
const fs = require('fs'), path = require('path'), http = require('http');
const HERE = __dirname, DIST = path.join(HERE, '..', 'dist'), EXPORT = process.argv[2], PORT = +(process.argv[3] || 8778);
const NOW = Date.parse(process.env.NOW || '2026-11-11T05:20:00Z'), started = Date.now();
const nowSec = () => Math.floor((NOW + (Date.now() - started)) / 1000);
const gm = { posts: [], likes: [], msgs: [] }; // msgs: the chat, oldest first
// Chat members = everyone named in the export's Driver / Navigator columns (ids 11, 12, …), plus the dispatcher.
const MEMBERS = (() => {
  const d = JSON.parse(fs.readFileSync(EXPORT, 'utf8')), h = d[0].map(x => String(x).trim().toLowerCase());
  const cols = [h.indexOf('driver'), h.indexOf('navigator')], names = new Set();
  d.slice(1).forEach(r => cols.forEach(c => { const n = String(r[c] || '').trim(); if (/^[A-Za-z][A-Za-z' -]+$/.test(n)) names.add(n); }));
  return [...names].sort().map((n, i) => ({ user_id: String(11 + i), nickname: n, name: n + ' T' }))
    .concat({ user_id: '1', nickname: 'Test Dispatcher', name: 'Test Dispatcher' });
})();
let nextId = 1000;
function groupme(method, url, body) {
  const u = new URL(url, 'http://x');
  if (u.pathname === '/users/me') return [200, { response: { name: 'Test Dispatcher', user_id: '1', id: '1' } }];
  if (u.pathname === '/groups') return [200, { response: [{ id: 'G1', name: 'SURE Walk Dispatch' }] }];
  if (u.pathname === '/groups/G1') return [200, { response: { id: 'G1', name: 'SURE Walk Dispatch', members: MEMBERS } }];
  if (u.pathname === '/groups/G1/messages' && method === 'GET') {
    const after = u.searchParams.get('after_id'), before = u.searchParams.get('before_id'), limit = +(u.searchParams.get('limit') || 20);
    let list;
    if (after) list = gm.msgs.filter(m => +m.id > +after).slice(0, limit); // oldest first
    else { list = gm.msgs.slice().reverse(); if (before) list = list.filter(m => +m.id < +before); list = list.slice(0, limit); } // newest first
    return list.length ? [200, { response: { count: gm.msgs.length, messages: list } }] : [304, null];
  }
  if (u.pathname === '/groups/G1/messages' && method === 'POST') {
    const m = Object.assign({ id: String(++nextId), group_id: 'G1', user_id: '1', sender_id: '1', name: 'Test Dispatcher', sender_type: 'user',
      created_at: nowSec(), favorited_by: [], attachments: [] }, body.message);
    gm.posts.push(m); gm.msgs.push(m);
    return [201, { response: { message: m } }];
  }
  let m;
  if ((m = /^\/messages\/G1\/(\d+)\/like$/.exec(u.pathname))) { gm.likes.push(m[1]); return [200, {}]; }
  if ((m = /^\/conversations\/G1\/messages\/(\d+)$/.exec(u.pathname)) && method === 'DELETE') { gm.msgs = gm.msgs.filter(x => x.id !== m[1]); return [204, null]; }
  return [404, { meta: { code: 404 } }];
}
http.createServer((req, res) => {
  let body = ''; req.on('data', c => { body += c; }); req.on('end', () => {
    const send = (code, obj, type) => { res.writeHead(code, { 'Content-Type': type || 'application/json', 'Access-Control-Allow-Origin': '*' }); res.end(obj == null ? '' : typeof obj === 'string' || Buffer.isBuffer(obj) ? obj : JSON.stringify(obj)); };
    const u = req.url;
    if (u.startsWith('/gm/')) { const [s, o] = groupme(req.method, u.slice(3), body ? JSON.parse(body) : null); return send(s, o); }
    if (u === '/gm-state') return send(200, gm);
    if (req.method === 'POST' && u === '/gm-crew') { // a crew member posts in the chat (test helper)
      const b = JSON.parse(body), m = { id: String(++nextId), group_id: 'G1', user_id: b.user_id, sender_id: b.user_id, name: b.name, sender_type: 'user',
        text: b.text, created_at: nowSec(), favorited_by: [], attachments: [] };
      gm.msgs.push(m); return send(200, m);
    }
    if (req.method === 'POST' && u === '/gm-like') { const b = JSON.parse(body), m = gm.msgs.find(x => x.id === b.id); if (m) m.favorited_by.push(b.user_id); return send(200, {}); }
    if (u === '/sheet.json') return send(200, fs.readFileSync(EXPORT));
    if (u === '/config.json') return send(200, { now: NOW, skew: +(process.env.SKEW || 0), paint: JSON.parse(process.env.PAINT || '[]') });
    // /?snippet: only the fake Excel — the Script Lab snippet's loader is run on it (it downloads the panel itself)
    if (u === '/?snippet') return send(200, '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body><script src="/fake-excel.js"></script></body></html>', 'text/html');
    if (u === '/') return send(200, '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head><body><script src="/fake-excel.js"></script><script src="/dist/host.js"></script></body></html>', 'text/html');
    if (u === '/fake-excel.js') return send(200, fs.readFileSync(path.join(HERE, 'fake-excel.js')), 'text/javascript');
    if (u.startsWith('/dist/')) {
      let data = fs.readFileSync(path.join(DIST, u.slice(6)), 'utf8');
      if (u === '/dist/backend.js') data = data.replace(/https:\/\/api\.groupme\.com\/v3/g, `http://localhost:${PORT}/gm`);
      return send(200, data, u.endsWith('.js') ? 'text/javascript' : 'text/html');
    }
    send(404, 'not found', 'text/plain');
  });
}).listen(PORT, () => console.log('classic test server on ' + PORT));
