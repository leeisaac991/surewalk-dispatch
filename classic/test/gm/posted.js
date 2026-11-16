// "Already posted" flags: a rider's earlier ride post only counts for the SAME pickup → drop-off.
// Also checks the never-a-pickup places (The Otis Hotel).
//   node classic/test/gm/posted.js [classic/dist/backend.js]
const fs = require('fs'), path = require('path'), vm = require('vm');
const BACKEND = process.argv[2] || path.join(__dirname, '..', '..', 'dist', 'backend.js');
const ctx = { console, Intl, Date, Math, JSON, Set, Map, Array, Object, String, Number, RegExp, Error, Promise, URL,
  XMLHttpRequest: function () { throw new Error('no network in this test'); } };
ctx.globalThis = ctx; ctx.self = ctx;
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(BACKEND, 'utf8'), ctx);
let bad = 0;
const check = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) bad++;
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${name}${ok ? '' : `: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`}`);
};
const ts = Date.now() / 1000 - 600;
const msgs = [ // newest first, as recentMessages_ returns them
  { id: '3', ts: ts + 300, who: 'Disp', text: 'Updated ride!\nMaya  512-555-0101  PCL → Castilian\n@Luke @Elijah' },
  { id: '2', ts: ts + 100, who: 'Disp', text: 'New ride!\nMaya  (512) 555-0101  Jester West → Duren\nLeo  5125550102  Kinsolving → Target' },
  { id: '1', ts, who: 'Disp', text: 'zoe 512 555 0103 at greg' }, // typed by hand: no "→"
];
const posted = ctx.postsIn_(msgs, {});
const ride = (phone, pickup, dropoff) => ({ phone, pickup, dropoff });
const id = r => { const p = ctx.postedFor_(posted, r); return p ? p.id : null; };
check('same ride as the newest post', id(ride('512-555-0101', 'PCL', 'Castilian')), '3');
check('same ride as the earlier post (Jester = Jester West)', id(ride('5125550101', 'Jester', 'Duren')), '2');
check('same rider, new pickup + drop-off: not posted', id(ride('5125550101', 'Gregory Gym', 'Jester')), null);
check('same rider, same pickup, new drop-off: not posted', id(ride('5125550101', 'PCL', 'Duren')), null);
check('pickup typed a bit differently still matches', id(ride('5125550102', 'kinsolving dorm', 'Target (Dobie)')), '2');
check('hand-typed post (no route): still counts', id(ride('5125550103', 'Gregory Gym', 'Jester')), '1');
check('never posted', id(ride('5125550199', 'PCL', 'Jester')), null);
// Group tags use the same matching: the rider's second ride isn't put in their first ride's group.
const rides = [{ id: 'a', phone: '5125550101', pickup: 'Gregory Gym', dropoff: 'Jester' }, { id: 'b', phone: '5125550102', pickup: 'Kinsolving', dropoff: 'Target' }];
vm.runInContext('loadBatches_ = () => ({}); recentMessages_ = () => __msgs;', Object.assign(ctx, { __msgs: msgs }));
vm.runInContext('CFG.EDIT_WINDOW_MIN = 60;', ctx);
ctx.tagGroups_(rides, []);
check('group tag: new ride of a posted rider has no post group', rides[0].group, null);
check('group tag: posted ride gets its post', rides[1].group, 'gm2');
// The Otis Hotel is never a pickup, also next to a campus landmark; as a drop-off it's West Campus.
for (const t of ['Otis Hotel', 'the otis hotel near jester', 'Otis hotel by dobie', 'The Otis']) check(`not a pickup: ${t}`, ctx.offCampusPickup_(t), true);
check('Otis as a drop-off: West Campus', ctx.zoneOf_('Otis Hotel'), 'West');
for (const t of ['Jester', 'PCL', 'Dobie', 'greg gym next to jester']) check(`still a pickup: ${t}`, ctx.offCampusPickup_(t), false);
console.log(bad ? `${bad} failed` : 'all passed');
process.exit(bad ? 1 : 0);
