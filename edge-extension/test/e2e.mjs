// Drives the panel in headless Chromium (CDP on port 9335) against fake-server.js. Prints what happened.
import fs from 'fs';
const sleep = ms => new Promise(r => setTimeout(r, ms));
const BASE = process.env.BASE || 'http://localhost:8777';
let tabs; for (let i = 0; i < 40; i++) { try { tabs = await (await fetch('http://localhost:9335/json/list')).json(); if (tabs.some(t => t.type === 'page')) break; } catch (e) {} await sleep(300); }
const ws = new WebSocket(tabs.find(t => t.type === 'page').webSocketDebuggerUrl);
await new Promise(r => { ws.onopen = r; });
let id = 0; const pend = {}, logs = [];
ws.onmessage = e => { const m = JSON.parse(e.data);
  if (m.method === 'Runtime.exceptionThrown') logs.push('EXC: ' + JSON.stringify(m.params.exceptionDetails).slice(0, 500));
  if (m.method === 'Runtime.consoleAPICalled' && m.params.type === 'error') logs.push('ERR: ' + m.params.args.map(a => a.value || a.description).join(' ').slice(0, 300));
  if (pend[m.id]) { pend[m.id](m.result); delete pend[m.id]; } };
const send = (method, params = {}) => new Promise(r => { pend[++id] = r; ws.send(JSON.stringify({ id, method, params })); });
await send('Runtime.enable'); await send('Page.enable');
const ev = async x => { const r = await send('Runtime.evaluate', { expression: x, awaitPromise: true, returnByValue: true }); return r.exceptionDetails ? 'EVAL-ERR ' + JSON.stringify(r.exceptionDetails).slice(0, 300) : r.result.value; };
const shot = async f => { const s = await send('Page.captureScreenshot', { format: 'png' }); fs.writeFileSync(f, Buffer.from(s.data, 'base64')); };
const srv = async () => { const l = await (await fetch(BASE + '/log')).json(); if (process.env.ADDIN) l.writes = await ev('window.__xlWrites'); return l; };
const click = sel => ev(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'missing ' + ${JSON.stringify(sel)}; b.click(); return 'ok'; })()`);
const toast = () => ev(`document.getElementById('toast').hidden ? '' : document.getElementById('toastText').textContent`);
const waitIdle = async () => { await sleep(400); for (let i = 0; i < 40; i++) { if (await ev('document.body.style.cursor') !== 'progress') break; await sleep(250); } await sleep(600); };
const n = () => ev('JSON.stringify({pending: __swd.night.pending.length, active: __swd.night.active.length, crews: __swd.night.crews.length, colors: __swd.night.colors, wait: __swd.night.wait, shifts: __swd.night.shiftsOn})');

await sleep(6000);
console.log('status:', await ev(`document.getElementById('statusLine').textContent`));
console.log('night:', await n());
console.log('pending rows:', await ev(`__swd.night.pending.map(r => r.row + ':' + r.first).join(' ')`));
await shot('e2e-1.png');

// 1. tonight's color
await click('#colorBtn'); await sleep(200); await click('#swatches button[data-color="#00ff00"]'); await sleep(1500);
console.log('colors:', JSON.stringify(await ev('__swd.night.colors')));

// 2. assign Eleanor (row 1574) to the first crew with room
const target = await ev(`__swd.night.pending.find(r => r.row === 1574).id`);
await ev(`__swd.selected.add(${JSON.stringify(target)}), 1`); await ev('__swd.load()'); await sleep(500);
const crewKey = await ev(`__swd.night.crews.find(c => !c.cap || c.riders.length < c.cap).key`);
await ev(`document.getElementById('crewSel').value = ${JSON.stringify(crewKey)}; document.getElementById('crewSel').dispatchEvent(new Event('change')); 1`);
await click('#assignBtn'); await waitIdle();
console.log('assign toast:', await toast());
let L = await srv();
console.log('writes:', L.writes.slice(-6).join(' | '));
console.log('post:', JSON.stringify(L.posts.slice(-1)[0]));
await shot('e2e-2.png');

// 3. At, then Picked, on that ride's group
const grp = `[...document.querySelectorAll('[data-act="at"]')].find(b => b.dataset.ids.includes(${JSON.stringify(target)}))`;
await ev(`${grp}.click(), 1`); await waitIdle();
console.log('at toast:', await toast());
await ev(`[...document.querySelectorAll('[data-act="picked"]')].find(b => b.dataset.ids.includes(${JSON.stringify(target)})).click(), 1`); await waitIdle();
console.log('picked toast:', await toast());
L = await srv(); console.log('writes:', L.writes.slice(-3).join(' | '));
console.log('still active?', await ev(`__swd.night.active.some(r => r.id === ${JSON.stringify(target)})`));

// 4. Cancel a pending ride, then undo
const p2 = await ev(`__swd.night.pending.find(r => r.row === 1567).id`);
await ev(`document.querySelector('[data-menu="${p2}"]').click(), 1`); await sleep(200);
await ev(`document.querySelector('[data-act="cancel"][data-id="${p2}"]').click(), 1`); await sleep(300);
await ev(`document.getElementById('dReason').value = 'Not Answering'; document.getElementById('dlgOk').click(), 1`); await waitIdle();
console.log('cancel toast:', await toast());
L = await srv(); console.log('writes:', L.writes.slice(-5).join(' | '));
await click('#undoBtn'); await waitIdle();
console.log('undo toast:', await toast());
L = await srv(); console.log('writes:', L.writes.slice(-5).join(' | '));
console.log('1567 pending again?', await ev(`__swd.night.pending.some(r => r.row === 1567)`));

// 5. assign two riders, take one off (GroupMe REMOVED post), undo deletes the post
const two = await ev(`JSON.stringify(__swd.night.pending.filter(r => [1571, 1572].includes(r.row)).map(r => r.id))`);
await ev(`JSON.parse(${JSON.stringify(two)}).forEach(i => __swd.selected.add(i)); __swd.load()`); await sleep(600);
await ev(`document.getElementById('crewSel').value = ${JSON.stringify(crewKey)}; 1`);
await click('#assignBtn'); await waitIdle();
console.log('assign2 toast:', await toast());
L = await srv(); console.log('post:', JSON.stringify(L.posts.slice(-1)[0].text));
const one = JSON.parse(two)[0];
await ev(`document.querySelector('[data-menu="${one}"]').click(), 1`); await sleep(200);
await ev(`document.querySelector('[data-act="takeoff1"][data-id="${one}"]').click(), 1`); await waitIdle();
console.log('takeoff toast:', await toast());
L = await srv(); console.log('post:', JSON.stringify(L.posts.slice(-1)[0].text));
await click('#undoBtn'); await waitIdle();
L = await srv(); console.log('undo toast:', await toast(), '| deleted posts:', L.deletes.join(','));

// 6. a refused action: cart can't take an off-campus ride
console.log('graph calls:', L.calls);
await shot('e2e-3.png');
await ev(`__swd.selected.clear(); __swd.load()`); await sleep(800);
await ev(`document.getElementById('crews').scrollIntoView(), 1`); await sleep(400); await shot('e2e-4.png');
console.log(logs.join('\n') || 'no page errors');
process.exit(0);
