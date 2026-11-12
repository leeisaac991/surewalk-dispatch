// Drives the classic panel (CDP on port 9337) against server.js. Prints what happened; screenshots to $SHOTS.
import fs from 'fs';
const sleep = ms => new Promise(r => setTimeout(r, ms));
const BASE = process.env.BASE || 'http://localhost:8778', SHOTS = process.env.SHOTS || '';
let tabs; for (let i = 0; i < 40; i++) { try { tabs = await (await fetch('http://localhost:9337/json/list')).json(); if (tabs.some(t => t.type === 'page')) break; } catch (e) {} await sleep(300); }
const ws = new WebSocket(tabs.find(t => t.type === 'page').webSocketDebuggerUrl);
await new Promise(r => { ws.onopen = r; });
let id = 0; const pend = {}, logs = [];
ws.onmessage = e => { const m = JSON.parse(e.data);
  if (m.method === 'Runtime.exceptionThrown') logs.push('EXC: ' + JSON.stringify(m.params.exceptionDetails).slice(0, 600));
  if (m.method === 'Runtime.consoleAPICalled' && m.params.type === 'error') logs.push('ERR: ' + m.params.args.map(a => a.value || a.description).join(' ').slice(0, 300));
  if (pend[m.id]) { pend[m.id](m.result); delete pend[m.id]; } };
const send = (method, params = {}) => new Promise(r => { pend[++id] = r; ws.send(JSON.stringify({ id, method, params })); });
await send('Runtime.enable'); await send('Page.enable');
const ev = async x => { const r = await send('Runtime.evaluate', { expression: x, awaitPromise: true, returnByValue: true }); return r.exceptionDetails ? 'EVAL-ERR ' + JSON.stringify(r.exceptionDetails).slice(0, 300) : r.result.value; };
const F = `document.getElementById('swcFrame').contentWindow`; // the panel page, in its frame in the side panel
const call = (fn, args) => ev(`new Promise(ok => ${F}.google.script.run.withSuccessHandler(ok).withFailureHandler(e => ok('FAIL ' + e.message))[${JSON.stringify(fn)}](...${JSON.stringify(args || [])}))`);
const act = (name, args) => call('act', [name, args]);
const shot = async f => { if (!SHOTS) return; const s = await send('Page.captureScreenshot', { format: 'png' }); fs.writeFileSync(`${SHOTS}/${f}`, Buffer.from(s.data, 'base64')); };
const row = r => ev(`JSON.stringify({E: __xl.text(${r},'E'), F: __xl.text(${r},'F'), G: __xl.text(${r},'G'), PQRS: ['P','Q','R','S'].map(c => __xl.text(${r},c)).join('|'), X: __xl.text(${r},'X'), fill: __xl.bg(${r},'A'), fillZ: __xl.bg(${r},'AE'), font: __xl.fc(${r},'C')})`);
const ok = r => (typeof r === 'string' ? r : 'ok');
const post = (path, b) => fetch(BASE + path, { method: 'POST', body: JSON.stringify(b) }).then(r => r.json());

await sleep(9000);
console.log('panel up:', await ev(`!!${F}.document.querySelector('#newBuild') && !${F}.document.body.innerText.includes("Can't find")`), '| pop-out button:', await ev(`!!document.getElementById('swcPop')`));
// copied highlight: rows 225/226 were painted like 224 (see PAINT) and are untouched -> cleared; 224 has a dispatcher -> kept
console.log('copied colors: 224', await ev(`__xl.bg(224,'A')`), '225', await ev(`__xl.bg(225,'A')`), '226', await ev(`__xl.bg(226,'A')`));
let st = await call('getState', [true]);
if (typeof st === 'string') { console.log(st); process.exit(1); }
console.log('pending', st.pending.length, '| active', st.active.length, '| crews', st.crews.length, '| roster', JSON.stringify(st.rosterInfo));
await shot('c1.png');

await call('setDispatcher', ['T. Tester']);
console.log('color:', ok(await act('setTonightColor', ['#00ff00', 'A'])));
const p = st.pending.find(x => x.row === 226), ref = { row: p.row, id: p.id };
const free = st.crews.filter(c => c.cap).sort((a, b) => (st.active.filter(r => r.driver === a.driver).length - st.active.filter(r => r.driver === b.driver).length));
let crew = null;
for (const c of free) { const r = await act('assignRides', [[ref], c, 'Auto', c.van]); if (typeof r !== 'string') { crew = c; break; } }
console.log('assign 226 ->', crew && crew.driver + '/' + crew.nav, await row(226));
console.log('arrived ->', ok(await act('markArrived', [[ref]])), await row(226));
console.log('picked ->', ok(await act('markPicked', [[ref]])), await row(226));
const p2 = st.pending.find(x => x.row === 225), ref2 = { row: p2.row, id: p2.id };
const c2 = await act('cancelRide', [ref2, 'Not Answering']);
console.log('cancel 225 ->', ok(c2), await row(225));
console.log('undo ->', ok(await act('undoAction', [c2.undo])), await row(225));
// ✎ Edit ride: name / pickup / drop-off go into J, K, M, N of the Excel sheet; undo puts them back
const names = r => ev(`JSON.stringify(['J','K','M','N'].map(c => __xl.text(${r}, c)))`);
console.log('before edit 225:', await names(225));
const ed = await act('editRide', [ref2, { first: 'Kat', last: 'Fixed', pickup: 'Kendra Scott Center', dropoff: 'Jester' }]);
const edited = ed.state && ed.state.pending.find(x => x.row === 225);
console.log('edit 225 ->', ok(ed), await names(225), '| panel shows', edited && `${edited.first} ${edited.pickup} → ${edited.dropoff} (zone ${edited.zone})`);
console.log('undo edit ->', ok(await act('undoAction', [ed.undo])), await names(225));

// GroupMe: sign in, pick the chat, sync on
const members = (await (await fetch(BASE + '/gm/groups/G1')).json()).response.members, idOf = n => (members.find(m => m.nickname === n) || {}).user_id;
await post('/gm-crew', { user_id: idOf(crew.nav) || '12', name: crew.nav + ' T', text: 'good evening' }); // the chat has history before tonight's posts
console.log('token:', JSON.stringify(await call('saveGroupMeToken', ['testtokentesttokentest1234'])), '| group:', ok(await act('setGroupMeGroup', ['G1', 'SURE Walk Dispatch'])), '| sync on:', ok(await act('setSyncRunning', [true])));
// the crew says "empty" in GroupMe, then gets a ride: they must show busy right away
await post('/gm-crew', { user_id: idOf(crew.driver), name: crew.driver + ' T', text: 'empty' });
await sleep(1500);
st = await call('getState', [false]);
const statusOf = s => ((s.board || []).find(b => b.key === crew.key) || {}).status;
console.log('after "empty": crew is', statusOf(st));
const p3 = st.pending.find(x => x.row === 223), ref3 = { row: p3.row, id: p3.id };
const a3 = await act('assignRides', [[ref3], crew, 'Auto', crew.van]);
console.log('assign 223 ->', typeof a3 === 'string' ? a3 : 'ok', await row(223), '| crew is now', typeof a3 === 'string' ? '?' : statusOf(a3.state));
const msg = typeof a3 === 'string' ? '' : a3.result;
const sent = await call('sendGroupMe', [msg + '\n@' + crew.driver, [], null]);
console.log('ride post sent:', typeof sent === 'string' ? sent : 'ok');
const gs = await (await fetch(BASE + '/gm-state')).json(), postId = gs.posts[gs.posts.length - 1].id;
await post('/gm-like', { id: postId, user_id: idOf(crew.driver) }); // the crew hearts it
await post('/gm-crew', { user_id: idOf(crew.driver), name: crew.driver + ' T', text: 'at jester' }); // and arrives
await sleep(1500);
st = await call('getState', [false]); // a full refresh syncs GroupMe
console.log('after "at jester":', await row(223));
await sleep(21000); st = await call('getState', [false]); // reactions are re-read every 20 s
console.log('heart on 223:', JSON.stringify(st.chat && st.chat.rideAck && st.chat.rideAck[p3.id]));
console.log('heart chip in panel:', await ev(`(${F}.document.body.innerText.match(/❤️[^\\n]{0,30}/) || [''])[0]`));
await post('/gm-crew', { user_id: idOf(crew.driver), name: crew.driver + ' T', text: 'picked jester' });
await sleep(1500);
st = await call('getState', [false]);
console.log('after "picked jester":', await row(223));
await ev(`${F}.document.querySelectorAll('button').forEach(b => { if (b.textContent.trim() === '✕') b.click(); }), 1`);
await sleep(1500); await shot('c2.png');

// Pop out: the panel moves into its own window and still works; closing the window brings it back
await send('Runtime.evaluate', { expression: `document.getElementById('swcPop').click()`, userGesture: true });
await sleep(6000);
const pages = (await (await fetch('http://localhost:9337/json/list')).json()).filter(t => t.type === 'page');
const popT = pages.find(t => /dialog\.html/.test(t.url));
console.log('pop-out window opened:', !!popT, '| side panel says:', JSON.stringify(await ev(`document.body.innerText.slice(0, 60)`)));
if (popT) {
  const pw = new WebSocket(popT.webSocketDebuggerUrl); await new Promise(r => { pw.onopen = r; });
  let pid = 0; const pp = {};
  pw.onmessage = e => { const m = JSON.parse(e.data); if (pp[m.id]) { pp[m.id](m.result); delete pp[m.id]; } };
  const pev = x => new Promise(r => { pp[++pid] = r; pw.send(JSON.stringify({ id: pid, method: 'Runtime.evaluate', params: { expression: x, awaitPromise: true, returnByValue: true } })); }).then(r => r.result && r.result.value);
  const PF = `document.getElementById('swcFrame').contentWindow`;
  console.log('window shows the panel:', await pev(`!!(document.getElementById('swcFrame') && ${PF}.document.querySelector('#newBuild'))`), '| call from the window:',
    await pev(`new Promise(ok => ${PF}.google.script.run.withSuccessHandler(s => ok('pending ' + s.pending.length + ', crews ' + s.crews.length + ', board ' + (s.board || []).length)).withFailureHandler(e => ok('FAIL ' + e.message)).getState(true))`));
  console.log('assign from the window:', await pev(`new Promise(ok => ${PF}.google.script.run.withSuccessHandler(r => ok(typeof r === 'object' && r.state ? 'ok' : JSON.stringify(r))).withFailureHandler(e => ok('FAIL ' + e.message)).act('setTonightColor', ['#ffff00', 'A']))`));
  await pev('window.close(), 1');
  await sleep(2000);
  console.log('after closing the window, back in the side panel:', await ev(`!!document.getElementById('swcFrame')`));
}
console.log(logs.join('\n') || 'no page errors');
process.exit(0);
