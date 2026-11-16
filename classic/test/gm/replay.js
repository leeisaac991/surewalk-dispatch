// Replays driver messages ("at/picked both greg", "picked them up"…) through the built backend's applyMessage_
// against made-up crew rides, and checks each result and sheet write against cases.json.
//   node classic/test/gm/replay.js [classic/dist/backend.js] [classic/test/gm/cases.json]
// A case: { name, text, from?, rides: [{first, pickup, picked?, arrived?, group?, dropoff?}], expect, writes }
const fs = require('fs'), path = require('path'), vm = require('vm');
const BACKEND = process.argv[2] || path.join(__dirname, '..', '..', 'dist', 'backend.js');
const CASES = process.argv[3] || path.join(__dirname, 'cases.json');
const writes = [];
const ctx = { console, Intl, Date, Math, JSON, Set, Map, Array, Object, String, Number, RegExp, Error, Promise, URL, __w: writes,
  XMLHttpRequest: function () { throw new Error('no network in this test'); } };
ctx.globalThis = ctx; ctx.self = ctx;
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(BACKEND, 'utf8'), ctx);
vm.runInContext(`
  arrive_ = (sh, row, t) => __w.push('arrive row ' + row + ' ' + t);
  pick_ = (sh, row, ms, both) => __w.push('pick row ' + row + (both ? ' (same minute)' : ''));
  clampToCall_ = (sh, row, d) => d;
  loadBatches_ = () => [];
  recentlyPosted_ = () => ({});
`, ctx);
const msgMs = Date.UTC(2026, 10, 15, 4, 30); // 10:30 PM in Austin
let bad = 0;
for (const c of JSON.parse(fs.readFileSync(CASES, 'utf8'))) {
  writes.length = 0;
  const rides = c.rides.map((r, i) => Object.assign({ row: 100 + i, id: 'r' + i, driver: 'Greg Smith', nav: 'Ana Lopez',
    tsMs: msgMs - 20 * 60e3, call: '22:15', picked: false, arrived: '', last: 'X' }, r));
  rides.forEach(r => { r.pickupName = ctx.canonPlace_(r.pickup) || r.pickup; r.pickupKey = ctx.placeKey_(r.pickup); });
  rides.olderActive = [];
  let out;
  try { out = ctx.applyMessage_(null, rides, { text: c.text, name: c.from || 'Greg Smith', created_at: msgMs / 1000, id: 'm1' }); }
  catch (e) { out = { result: 'THREW ' + e.message }; }
  const got = out ? out.result : null, w = writes.join('; ');
  const pass = got === c.expect && w === c.writes;
  if (!pass) bad++;
  console.log(`${pass ? 'ok  ' : 'FAIL'} ${c.name}: "${c.text}" -> ${JSON.stringify(got)}${pass ? '' : `\n     expected ${JSON.stringify(c.expect)} | writes "${w}" vs "${c.writes}"`}`);
}
console.log(bad ? `${bad} failed` : 'all passed');
process.exit(bad ? 1 : 0);
