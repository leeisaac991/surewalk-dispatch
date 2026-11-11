/*
 * SURE Walk Dispatch — Edge side panel.
 * Reads the rides workbook (Excel on the web, via Microsoft Graph), shows Pending rides and Crews,
 * and writes dispatch changes back with the same rules as the Apps Script panel. Ride messages post to GroupMe.
 */
(() => {
  'use strict';
  const { COL } = R;
  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const colL = n => { let s = ''; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const addr = (row, c1, c2) => `${colL(c1)}${row}` + (c2 && c2 !== c1 ? `:${colL(c2)}${row}` : '');

  /* ---------------- settings & per-night state (this browser) ---------------- */
  const DEFAULTS = { dispatcher: '', clientId: '', tenant: 'organizations', fileLink: '', sheet: '', gmToken: '', gmGroupId: '', gmGroupName: '', autoPost: true,
    graphBase: '', loginBase: '', gmBase: '' };
  let S = Object.assign({}, DEFAULTS);
  let nightState = { colors: null, toggles: null, extraCrews: null };
  const save = (k, v) => chrome.storage.local.set({ [k]: v });

  let night = null, busy = false, lastUndo = null, connected = false, loading = false;
  const selected = new Set();          // pending ride ids
  const fillCache = new Map();         // row -> { key, fill, at }
  let openMenu = null;
  const touched = new Set();           // rows an action read or wrote: their colors are re-read right after

  /* ---------------- status / toast ---------------- */
  function status(msg, err) { $('statusLine').textContent = msg || ''; $('statusLine').classList.toggle('err', !!err); }
  let toastTimer = null;
  function toast(msg, opts) {
    opts = opts || {};
    $('toastText').textContent = msg;
    $('toast').classList.toggle('err', !!opts.err);
    $('undoBtn').hidden = !opts.undo;
    $('toast').hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { $('toast').hidden = true; }, opts.err ? 9000 : opts.undo ? 12000 : 5000);
  }

  /* ---------------- dialogs ---------------- */
  function dialog(title, bodyHtml, okLabel, danger) {
    return new Promise(resolve => {
      $('dlgTitle').textContent = title;
      $('dlgBody').innerHTML = bodyHtml || '';
      $('dlgOk').textContent = okLabel || 'OK';
      $('dlgOk').classList.toggle('danger', !!danger);
      $('dlg').hidden = false;
      const done = ok => {
        $('dlg').hidden = true;
        $('dlgOk').onclick = $('dlgCancel').onclick = null;
        resolve(ok ? $('dlgBody') : null);
      };
      $('dlgOk').onclick = () => done(true);
      $('dlgCancel').onclick = () => done(false);
      const first = $('dlgBody').querySelector('input,select,textarea');
      (first || $('dlgOk')).focus();
    });
  }

  /* ---------------- reading the sheet ---------------- */
  const SCAN = 600;
  async function load(force) {
    if (!connected || loading) return;
    loading = true;
    try {
      const t = await G.readText(last => last - SCAN + 1);
      let lastIdx = -1;
      t.text.forEach((r, i) => { if (R.hasTime(r[0])) lastIdx = i; });
      const rows = t.text.slice(0, lastIdx + 1).map((text, i) => ({ row: t.top + i, text }));
      const rosterRows = t.text.slice(lastIdx + 1);
      // Row colors: tonight's rows only. Open rides are re-checked every refresh, finished ones every 2 min.
      const now = Date.now(), need = [], keyOf = r => r.text[0] + '|' + r.text[7] + '|' + r.text[3] + '|' + r.text[21];
      rows.forEach(r => {
        const ts = R.parseTs(r.text[0]);
        if (ts == null || ts < now - 14 * 3600e3) return;
        const c = fillCache.get(r.row), k = R.rowKind(c && c.fill).kind;
        const done = k === 'black' || k === 'red' || (k === 'color' && String(r.text[3]).trim());
        if (force || !c || c.key !== keyOf(r) || now - c.at > (done ? 120e3 : 8e3)) need.push(r.row);
      });
      if (need.length) {
        const f = await G.readFills(need);
        const byRow = new Map(rows.map(r => [r.row, r]));
        need.forEach(row => fillCache.set(row, { key: keyOf(byRow.get(row)), fill: f[row] || '#ffffff', at: now }));
      }
      rows.forEach(r => { r.fill = (fillCache.get(r.row) || {}).fill || ''; });
      const n = R.nightKey();
      night = R.buildNight({ rows, rosterRows, toggles: nightState.toggles, chosenColors: nightState.colors,
        extraCrews: nightState.extraCrews && nightState.extraCrews.night === n ? nightState.extraCrews.list : [] });
      // forget selections of rides that are no longer pending
      [...selected].forEach(id => { if (!night.pending.some(r => r.id === id)) selected.delete(id); });
      render();
      status(`${G.workbook.name} › ${G.workbook.sheet} · updated ${R.fmtTime()}`);
    } catch (e) {
      status(e.message, true);
      if (e instanceof G.NeedSignIn) showSetup(e.message);
    } finally { loading = false; }
  }

  /* ---------------- rendering ---------------- */
  function rideTags(r) {
    return (r.ada ? '<span class="tag ada">ADA</span>' : '') +
      (r.dupOf ? `<span class="tag dup">possible duplicate of ${esc(r.dupOf.first)} ${esc(r.dupOf.time)} (row ${r.dupOf.row})</span>` : '') +
      (r.offCampus ? '<span class="tag off">off-campus pickup</span>' : '');
  }
  function menuHtml(id, items) {
    return `<div class="menu"><button data-menu="${esc(id)}" title="More">⋯</button>${openMenu === id ? `<div class="items">${items.map(([a, label, cls]) =>
      `<button data-act="${a}" data-id="${esc(id)}" class="${cls || ''}">${label}</button>`).join('')}</div>` : ''}</div>`;
  }
  function render() {
    if (!night) return;
    // wait + colors + shifts
    $('wait').hidden = night.wait == null;
    $('wait').textContent = night.wait == null ? '' : `~${night.wait} min wait`;
    $('tonightBar').hidden = false;
    $('swA').style.background = night.colors.A || 'transparent';
    $('swB').style.background = night.colors.B || 'transparent';
    $('colorTxt').textContent = night.colors.A ? '' : 'Set';
    $('shiftChips').innerHTML = (night.roster ? night.roster.shifts : []).map(s =>
      `<button class="shift ${night.shiftsOn.includes(s.name) ? 'on' : ''}" data-shift="${esc(s.name)}">${esc(s.name.replace(/\s*shift/i, ''))}</button>`).join('');

    // pending
    $('pendingCount').textContent = night.pending.length || '';
    $('pending').innerHTML = night.pending.length ? night.pending.map(r => `
      <div class="ride ${selected.has(r.id) ? 'sel' : ''}" data-ride="${esc(r.id)}">
        <input type="checkbox" data-pick="${esc(r.id)}" ${selected.has(r.id) ? 'checked' : ''}>
        <div><div class="route">${esc(r.pickup)} → ${esc(r.dropoff)}</div>
          <div class="meta">${esc(r.first)} ${esc(r.last)} · ${esc(r.time)} · ${esc(r.phone)}${r.zone ? ' · ' + esc(r.zone) : ''}</div>${rideTags(r)}</div>
        ${menuHtml(r.id, [['edit', 'Edit…'], ['dup', 'Mark duplicate (red)'], ['cancel', 'Cancel…', 'danger']])}
      </div>`).join('') : '<div class="empty">No pending rides.</div>';

    // crews
    $('crews').innerHTML = night.crews.length ? night.crews.map(c => {
      const groups = [];
      R.byPickup(c.riders).forEach(r => {
        const g = groups.find(x => x.key === (r.pickupName || r.pickup).toLowerCase());
        if (g) g.rides.push(r); else groups.push({ key: (r.pickupName || r.pickup).toLowerCase(), name: r.pickupName || r.pickup, rides: [r] });
      });
      const cap = Number(c.cap) || 0;
      return `<div class="crew">
        <div class="crewhead"><span class="dot ${c.riders.length ? 'busy' : ''}"></span>
          <b>${esc(R.crewLabel(c))}</b>
          <span class="meta">${esc(c.vehicle || c.van || '')}${cap ? ` · ${c.riders.length}/${cap}` : c.riders.length ? ` · ${c.riders.length}` : ''}${c.shift ? ' · ' + esc(c.shift) : ''}</span></div>
        ${groups.map(g => {
          const ids = g.rides.map(r => r.id).join(',');
          const arrived = g.rides.every(r => r.arrived);
          return `<div class="group">
            ${g.rides.map(r => `<div class="rider"><span class="grow"><b>${esc(r.first)}</b> ${esc(r.pickup)} → ${esc(r.dropoff)}
              <span class="meta">${esc(r.time)}${r.arrived ? ' · at ' + esc(r.arrived) : ''}</span>${rideTags(r)}</span>
              ${menuHtml(r.id, [['edit', 'Edit…'], ['takeoff1', 'Take off ride'], ['cancel', 'Cancel…', 'danger']])}</div>`).join('')}
            <div class="row btns">
              <button data-act="at" data-ids="${esc(ids)}" ${arrived ? 'disabled' : ''}>${arrived ? 'At ✓' : 'At'}</button>
              <button data-act="picked" data-ids="${esc(ids)}" class="primary">Picked</button>
              <button data-act="takeoff" data-ids="${esc(ids)}">Take off</button>
            </div></div>`;
        }).join('')}
      </div>`;
    }).join('') : `<div class="empty">${night.roster ? 'No crews on the active shifts.' : 'No roster found below the rides — add crews with “+ Add crew”.'}</div>`;
    renderAssignBar();
  }
  function renderAssignBar() {
    const rs = night ? night.pending.filter(r => selected.has(r.id)) : [];
    $('assignBar').hidden = !rs.length;
    document.body.classList.toggle('assigning', !!rs.length);
    if (!rs.length) return;
    $('selCount').textContent = `${rs.length} selected`;
    const prev = $('crewSel').value;
    $('crewSel').innerHTML = '<option value="">Choose a crew…</option>' + night.crews.map(c =>
      `<option value="${esc(c.key)}">${esc(R.crewLabel(c))} — ${c.riders.length}${c.cap ? '/' + c.cap : ''} · ${esc(c.van || '')}</option>`).join('');
    if (night.crews.some(c => c.key === prev)) $('crewSel').value = prev;
    if (!$('locSel').options.length) $('locSel').innerHTML = ['Auto'].concat(R.LOCATIONS).map(x => `<option>${x}</option>`).join('');
    const crew = night.crews.find(c => c.key === $('crewSel').value);
    const vanNow = $('vanSel').value;
    $('vanSel').innerHTML = R.VANS.map(x => `<option>${x}</option>`).join('');
    $('vanSel').value = crew && crew.van ? crew.van : (vanNow || 'Van');
  }

  /* ---------------- actions ---------------- */
  const findRide = id => night && night.rides.find(r => r.id === id);
  const ridesOf = ids => String(ids).split(',').map(findRide).filter(Boolean);
  const crewOfRide = r => night.crews.find(c => R.samePair(c, r)) || { driver: r.driver, nav: r.nav, van: r.van, riders: [] };

  async function act(label, fn) {
    if (busy) return toast('Still saving the last change…');
    busy = true; document.body.style.cursor = 'progress';
    status(label + '…');
    try {
      const res = await fn();
      if (res && res.undo) lastUndo = Object.assign({ at: Date.now(), label }, res.undo);
      toast(res && res.msg ? res.msg : label + ' — done.', { undo: !!(res && res.undo) });
    } catch (e) {
      toast(e.message, { err: true });
    } finally {
      busy = false; document.body.style.cursor = '';
      touched.forEach(r => fillCache.delete(r)); touched.clear();
      await load();
    }
  }

  // Re-read the rows just before writing and make sure they are still the same rides.
  async function fresh(rides) {
    rides.forEach(r => touched.add(r.row));
    const f = await G.readRows(rides.map(r => r.row));
    return rides.map(r => {
      const x = f[r.row];
      if (R.rideIdOf(x.text) !== r.id) throw new Error('The sheet changed under that ride (rows moved) — refreshed; try again.');
      return Object.assign({ ride: r }, x);
    });
  }
  const getOf = x => c => String(x.text[c - 1] == null ? '' : x.text[c - 1]).trim();
  // Undo = the cells and colors as they were before (written back only on request).
  function undoFor(frs, ops) {
    const back = [];
    ops.forEach(op => {
      const x = frs.find(f => f.ride.row === (op.row || Number(/\d+/.exec(op.addr)[0])));
      if (!x) return;
      if (op.t === 'v' || op.t === 'clr') {
        const m = /^([A-Z]+)(\d+)(?::([A-Z]+)\d+)?$/.exec(op.addr);
        const c1 = m[1].charCodeAt(0) - 64, c2 = (m[3] || m[1]).charCodeAt(0) - 64;
        back.push({ t: 'v', addr: op.addr, values: [x.text.slice(c1 - 1, c2).map(v => (v == null ? '' : String(v)))] });
      } else if (op.t === 'fill') back.push({ t: 'fill', row: op.row, color: x.fill });
      else if (op.t === 'font') back.push({ t: 'font', row: op.row, color: x.font });
    });
    return back.reverse();
  }
  async function announce(text, crew) {
    if (!text) return { posted: false };
    if (S.autoPost && GM.ready()) {
      try { return { posted: true, id: await GM.post(text, [crew.driver, crew.nav]) }; }
      catch (e) { toast('Saved in the sheet, but GroupMe failed: ' + e.message, { err: true }); return { posted: false }; }
    }
    try { await navigator.clipboard.writeText(text + '\n@' + [crew.driver, crew.nav].filter(Boolean).join(' @')); } catch (e) { /* no clipboard */ }
    return { posted: false, copied: true };
  }
  const postNote = p => (p.posted ? ' Posted in GroupMe.' : p.copied ? ' Message copied — paste it in GroupMe.' : '');

  function assign() {
    const rs = night.pending.filter(r => selected.has(r.id));
    const crew = night.crews.find(c => c.key === $('crewSel').value);
    if (!crew) return toast('Choose a crew first.', { err: true });
    const loc = $('locSel').value, van = $('vanSel').value;
    act(`Assigning ${rs.length} to ${R.crewLabel(crew)}`, async () => {
      const frs = await fresh(rs);
      frs.forEach(x => {
        R.assertEditable(x, x.ride.row);
        const g = getOf(x);
        if (g(COL.DRIVER) || R.hasReason(g(COL.REASON))) throw new Error(`${g(COL.FIRST)} (row ${x.ride.row}) was already assigned or cancelled — refreshed.`);
      });
      const riders = frs.map(x => {
        const g = getOf(x), fix = c => (typeof placeRewrite_ === 'function' && placeRewrite_(g(c))) || '';
        const pickup = fix(COL.PICKUP) || g(COL.PICKUP), dropoff = fix(COL.DROPOFF) || g(COL.DROPOFF);
        const zone = zoneOf_(dropoff);
        return Object.assign({}, x.ride, { pickup, dropoff, zone, rowLoc: zone || (loc !== 'Auto' ? loc : '') || 'Off', fixP: fix(COL.PICKUP), fixD: fix(COL.DROPOFF), x });
      });
      R.checkAssign(crew, riders, van);
      const ops = [];
      riders.forEach(r => {
        const g = getOf(r.x), row = r.row;
        if (r.fixP) ops.push({ t: 'v', addr: addr(row, COL.PICKUP), values: [[r.fixP]] });
        if (r.fixD) ops.push({ t: 'v', addr: addr(row, COL.DROPOFF), values: [[r.fixD]] });
        if (!g(COL.DISPATCH) && S.dispatcher) ops.push({ t: 'v', addr: addr(row, COL.DISPATCH), values: [[S.dispatcher]] });
        if (!g(COL.CALL)) ops.push({ t: 'v', addr: addr(row, COL.CALL), values: [[R.fmtTime()]] });
        ops.push({ t: 'v', addr: addr(row, COL.LOC, COL.NAV), values: [[r.rowLoc, van, crew.driver, crew.nav || '']] });
        if (R.rowKind(r.x.fill).kind === 'color') ops.push({ t: 'fill', row, color: null }); // a color copied from the row above isn't "picked"
      });
      const undo = undoFor(frs, ops);
      await G.write(ops);
      selected.clear();
      const others = (crew.riders || []).filter(o => !riders.some(r => r.row === o.row));
      const p = await announce(R.assignText(others, riders), crew);
      return { msg: `Assigned ${riders.map(r => r.first).join(', ')} to ${R.crewLabel(crew)}.${postNote(p)}`, undo: { ops: undo, post: p.id } };
    });
  }
  // the rides of one crew only
  function oneCrew(frs) {
    const keys = new Set(frs.map(x => { const g = getOf(x); return R.crewKey({ driver: g(COL.DRIVER), nav: g(COL.NAV) }); }));
    if (keys.size > 1) throw new Error('Those rides belong to different crews — refreshed; try again.');
  }
  function markAt(ids) {
    const rs = ridesOf(ids);
    act(`At ${rs.map(r => r.first).join(', ')}`, async () => {
      const frs = await fresh(rs); oneCrew(frs);
      const ops = [], t = R.fmtTime();
      frs.forEach(x => {
        R.assertEditable(x, x.ride.row);
        if (getOf(x)(COL.ARRIVE)) return;
        if (R.rowKind(x.fill).kind === 'color') ops.push({ t: 'fill', row: x.ride.row, color: null });
        ops.push({ t: 'v', addr: addr(x.ride.row, COL.ARRIVE), values: [[t]] });
      });
      const undo = undoFor(frs, ops);
      await G.write(ops);
      return { msg: `Arrival ${t} for ${rs.map(r => r.first).join(', ')}.`, undo: { ops: undo } };
    });
  }
  function markPicked(ids) {
    const rs = ridesOf(ids);
    act(`Picked ${rs.map(r => r.first).join(', ')}`, async () => {
      const frs = await fresh(rs); oneCrew(frs);
      const now = Date.now(), ops = [];
      frs.forEach(x => {
        R.assertEditable(x, x.ride.row);
        const g = getOf(x), shade = R.shadeFor(night.roster, g(COL.DISPATCH) || S.dispatcher, now), color = night.colors[shade];
        if (!color) throw new Error("Set tonight's picked color first (top of the panel).");
        if (!g(COL.ARRIVE)) { // no "at" yet: arrival = 2 min before the pick, never before the call time
          let at = now - R.PICK_GAP_MIN * 60e3;
          const call = /^(\d{1,2}):(\d{2})$/.exec(g(COL.CALL));
          if (call) {
            const p = R.parts(at), before = (Number(call[1]) * 60 + Number(call[2]) - (p.h * 60 + p.mi) + 1440) % 1440;
            if (before > 0 && before <= 120) at += before * 60e3;
          }
          ops.push({ t: 'v', addr: addr(x.ride.row, COL.ARRIVE), values: [[R.fmtTime(at)]] });
        }
        ops.push({ t: 'fill', row: x.ride.row, color });
      });
      const undo = undoFor(frs, ops);
      await G.write(ops);
      return { msg: `Picked up ${rs.map(r => r.first).join(', ')}.`, undo: { ops: undo } };
    });
  }
  function takeOff(ids) {
    const rs = ridesOf(ids);
    act(`Taking ${rs.map(r => r.first).join(', ')} off`, async () => {
      const frs = await fresh(rs); oneCrew(frs);
      const ops = [];
      frs.forEach(x => {
        R.assertEditable(x, x.ride.row);
        const g = getOf(x);
        if (!g(COL.DRIVER)) throw new Error(`${g(COL.FIRST)} isn't assigned to a crew anymore — refreshed.`);
        if (R.rowKind(x.fill).kind === 'color' && g(COL.ARRIVE)) throw new Error(`${g(COL.FIRST)} is already picked up.`);
        ops.push({ t: 'clr', addr: addr(x.ride.row, COL.ARRIVE) });
        ops.push({ t: 'clr', addr: addr(x.ride.row, COL.LOC, COL.NAV) });
      });
      const undo = undoFor(frs, ops);
      await G.write(ops);
      const crew = crewOfRide(rs[0]), rest = (crew.riders || []).filter(o => !rs.some(r => r.row === o.row));
      const p = await announce(R.removedText(rest, rs, 'REMOVED'), crew);
      return { msg: `${rs.map(r => r.first).join(', ')} back in Pending.${postNote(p)}`, undo: { ops: undo, post: p.id } };
    });
  }
  async function cancelRide(id) {
    const r = findRide(id);
    const body = await dialog(`Cancel ${r.first}'s ride?`, `<label>Reason <select id="dReason">${R.CANCEL_REASONS.map(x => `<option>${esc(x)}</option>`).join('')}</select></label>
      <p class="hint">${esc(r.pickup)} → ${esc(r.dropoff)} · row ${r.row}. The row turns black.</p>`, 'Cancel ride', true);
    if (!body) return;
    const reason = body.querySelector('#dReason').value;
    act(`Cancelling ${r.first}`, async () => {
      const [x] = await fresh([r]);
      R.assertEditable(x, r.row);
      const g = getOf(x), hadCrew = g(COL.DRIVER) && !(R.rowKind(x.fill).kind === 'color' && g(COL.ARRIVE));
      const ops = [{ t: 'clr', addr: addr(r.row, COL.DISPATCH, COL.ARRIVE) }, { t: 'clr', addr: addr(r.row, COL.LOC, COL.NAV) },
        { t: 'v', addr: addr(r.row, COL.REASON), values: [[reason]] }, { t: 'fill', row: r.row, color: '#000000' }, { t: 'font', row: r.row, color: '#ffffff' }];
      const undo = undoFor([x], ops);
      await G.write(ops);
      let p = {};
      if (hadCrew) {
        const crew = crewOfRide(r), rest = (crew.riders || []).filter(o => o.row !== r.row);
        p = await announce(R.removedText(rest, [r], 'CANCELLED'), crew);
      }
      return { msg: `Cancelled ${r.first} (${reason}).${postNote(p)}`, undo: { ops: undo, post: p.id } };
    });
  }
  async function markDup(id) {
    const r = findRide(id);
    if (!(await dialog(`Mark ${r.first}'s ride as a duplicate?`, `<p class="hint">${esc(r.pickup)} → ${esc(r.dropoff)} · ${esc(r.time)} · row ${r.row}. The row turns red and is ignored from then on.</p>`, 'Mark duplicate', true))) return;
    act(`Marking ${r.first} duplicate`, async () => {
      const [x] = await fresh([r]);
      R.assertEditable(x, r.row);
      const ops = [{ t: 'fill', row: r.row, color: R.DUPLICATE_COLOR }];
      const undo = undoFor([x], ops);
      await G.write(ops);
      return { msg: `Row ${r.row} marked duplicate.`, undo: { ops: undo } };
    });
  }
  const EDIT_FIELDS = [[COL.DISPATCH, 'Dispatch'], [COL.CALL, 'Call time'], [COL.ARRIVE, 'Arrival time'], [COL.PICKUP, 'Pick-up'], [COL.DROPOFF, 'Drop-off'],
    [COL.ADA, 'ADA (Yes/No)'], [COL.LOC, 'Location'], [COL.VAN, 'Van'], [COL.DRIVER, 'Driver'], [COL.NAV, 'Navigator'], [COL.REASON, 'Cancel reason'], [COL.NOTES, 'Notes']];
  async function editRide(id) {
    const r = findRide(id);
    const [x0] = await fresh([r]).catch(e => { toast(e.message, { err: true }); return []; });
    if (!x0) return;
    const g = getOf(x0);
    const body = await dialog(`Edit ${r.first} ${r.last} (row ${r.row})`, EDIT_FIELDS.map(([c, label]) =>
      `<label>${label} <input data-col="${c}" value="${esc(g(c))}"></label>`).join(''), 'Save');
    if (!body) return;
    const changes = [...body.querySelectorAll('input[data-col]')].map(i => [Number(i.dataset.col), i.value.trim()]).filter(([c, v]) => v !== g(c));
    if (!changes.length) return;
    act(`Saving row ${r.row}`, async () => {
      const [x] = await fresh([r]);
      const ops = changes.map(([c, v]) => ({ t: 'v', addr: addr(r.row, c), values: [[v]] }));
      const undo = undoFor([x], ops);
      await G.write(ops);
      return { msg: `Row ${r.row} saved.`, undo: { ops: undo } };
    });
  }
  function undo() {
    const u = lastUndo;
    if (!u || Date.now() - u.at > 30 * 60e3) return toast('Nothing to undo.');
    lastUndo = null; $('toast').hidden = true;
    act(`Undoing “${u.label}”`, async () => {
      await G.write(u.ops);
      if (u.post) { try { await GM.remove(u.post); } catch (e) { /* post stays */ } }
      return { msg: `Undid “${u.label}”${u.post ? ' and deleted its GroupMe post' : ''}.` };
    });
  }

  /* ---------------- tonight's color, shifts, extra crews ---------------- */
  function renderSwatches() {
    $('swatches').innerHTML = Object.keys(R.PASTEL_OF).map(h => `<button data-color="${h}" style="background:${h}" title="${h}"></button>`).join('') +
      '<button data-color="" title="Auto (from the sheet)">auto</button>';
  }
  async function setColor(hex) {
    nightState.colors = hex ? { night: R.nightKey(), A: hex, B: R.pastelOf(hex) } : null;
    await save('colors', nightState.colors);
    $('colorPick').hidden = true;
    await load();
  }
  async function toggleShift(name) {
    const on = new Set(night.shiftsOn);
    on.has(name) ? on.delete(name) : on.add(name);
    nightState.toggles = { night: R.nightKey(), period: R.autoShiftLetters().join(''), names: [...on] };
    await save('toggles', nightState.toggles);
    await load();
  }
  async function addCrew() {
    const body = await dialog('Add a crew for tonight', `<label>Driver <input id="dDrv"></label><label>Navigator <input id="dNav"></label>
      <label>Van <select id="dVan">${R.VANS.map(v => `<option>${v}</option>`).join('')}</select></label><label>Seats (0 = no limit) <input id="dCap" type="number" min="0" value="0"></label>`, 'Add');
    if (!body) return;
    const c = { driver: body.querySelector('#dDrv').value.trim(), nav: body.querySelector('#dNav').value.trim(), van: body.querySelector('#dVan').value, cap: Number(body.querySelector('#dCap').value) || 0 };
    if (!c.driver) return toast('A crew needs at least a driver.', { err: true });
    const n = R.nightKey(), cur = nightState.extraCrews && nightState.extraCrews.night === n ? nightState.extraCrews.list : [];
    nightState.extraCrews = { night: n, list: cur.concat(c) };
    await save('extraCrews', nightState.extraCrews);
    await load();
  }

  /* ---------------- settings & connecting ---------------- */
  function showSetup(msg) {
    $('setupNote').hidden = false;
    $('setupNote').innerHTML = `${esc(msg || 'Finish setup to start dispatching.')} <button class="link" id="openSettingsInline">Open Settings</button>`;
  }
  function fillSettings() {
    $('sDispatcher').value = S.dispatcher; $('sClientId').value = S.clientId; $('sTenant').value = S.tenant;
    $('sFile').value = S.fileLink; $('sSheet').value = S.sheet; $('sGmToken').value = S.gmToken; $('sAutoPost').checked = S.autoPost !== false;
    $('sGmGroup').innerHTML = `<option value="">—</option>` + (S.gmGroupId ? `<option value="${esc(S.gmGroupId)}" selected>${esc(S.gmGroupName || S.gmGroupId)}</option>` : '');
    try { $('redirectUri').textContent = G.redirectUrl(); } catch (e) { $('redirectUri').textContent = '(open from Edge)'; }
  }
  function readSettings() {
    S.dispatcher = $('sDispatcher').value.trim(); S.clientId = $('sClientId').value.trim(); S.tenant = $('sTenant').value.trim() || 'organizations';
    S.fileLink = $('sFile').value.trim(); S.sheet = $('sSheet').value.trim(); S.gmToken = $('sGmToken').value.trim(); S.autoPost = $('sAutoPost').checked;
    const g = $('sGmGroup'); S.gmGroupId = g.value; S.gmGroupName = g.value ? g.options[g.selectedIndex].text : '';
  }
  function applyConfig() {
    G.configure(Object.assign({ clientId: S.clientId, tenant: S.tenant || 'organizations' }, S.graphBase ? { graph: S.graphBase } : {}, S.loginBase ? { login: S.loginBase } : {}));
    GM.configure(Object.assign({ token: S.gmToken, groupId: S.gmGroupId }, S.gmBase ? { base: S.gmBase } : {}));
  }
  async function saveSettings() { readSettings(); await save('settings', S); applyConfig(); }
  const IN_EXCEL = window.SW_HOST === 'excel'; // running as an Excel add-in: the open workbook, no Microsoft sign-in
  async function connect() {
    connected = false;
    if (!IN_EXCEL && !S.clientId) return showSetup('Add the Microsoft app (client) ID in Settings, then sign in.');
    if (!IN_EXCEL && !S.fileLink) return showSetup('Paste the rides workbook link in Settings.');
    try {
      status('Opening the workbook…');
      const w = await G.openWorkbook(S.fileLink, S.sheet);
      $('fileWho').textContent = `${w.name} › ${w.sheet}`;
      $('setupNote').hidden = true;
      connected = true;
      fillCache.clear();
      await load(true);
    } catch (e) {
      status(e.message, true);
      showSetup(e instanceof G.NeedSignIn ? e.message : 'Could not open the workbook: ' + e.message);
    }
  }

  /* ---------------- events ---------------- */
  document.addEventListener('click', e => {
    const t = e.target.closest('button, input[type=checkbox], .ride');
    if (!t) { if (openMenu) { openMenu = null; render(); } return; }
    if (t.dataset.menu) { openMenu = openMenu === t.dataset.menu ? null : t.dataset.menu; render(); return; }
    if (t.dataset.act) {
      const a = t.dataset.act, id = t.dataset.id, ids = t.dataset.ids;
      openMenu = null;
      if (a === 'at') return markAt(ids);
      if (a === 'picked') return markPicked(ids);
      if (a === 'takeoff') return takeOff(ids);
      if (a === 'takeoff1') return takeOff(id);
      if (a === 'cancel') return cancelRide(id);
      if (a === 'dup') return markDup(id);
      if (a === 'edit') return editRide(id);
    }
    if (t.dataset.pick || t.classList.contains('ride')) {
      const id = t.dataset.pick || t.dataset.ride;
      if (!id || e.target.closest('.menu')) return;
      if (t.tagName !== 'INPUT') { selected.has(id) ? selected.delete(id) : selected.add(id); }
      else { t.checked ? selected.add(id) : selected.delete(id); }
      render();
      return;
    }
    if (t.dataset.shift) return toggleShift(t.dataset.shift);
    if (t.dataset.color !== undefined) return setColor(t.dataset.color);
    if (t.id === 'openSettingsInline') { fillSettings(); $('settings').hidden = false; }
  });
  $('assignBtn').addEventListener('click', assign);
  $('clearSel').addEventListener('click', () => { selected.clear(); render(); });
  $('crewSel').addEventListener('change', renderAssignBar);
  $('refreshBtn').addEventListener('click', () => { fillCache.clear(); connected ? load(true) : connect(); });
  $('undoBtn').addEventListener('click', undo);
  $('colorBtn').addEventListener('click', () => { renderSwatches(); $('colorPick').hidden = !$('colorPick').hidden; });
  $('addCrewBtn').addEventListener('click', addCrew);
  $('settingsBtn').addEventListener('click', () => { fillSettings(); $('settings').hidden = false; });
  $('closeSettings').addEventListener('click', () => { $('settings').hidden = true; });
  $('saveSettings').addEventListener('click', async () => { await saveSettings(); $('settings').hidden = true; connect(); });
  $('openTab').addEventListener('click', () => chrome.tabs.create({ url: chrome.runtime.getURL('panel.html') }));
  $('signInBtn').addEventListener('click', async () => {
    await saveSettings();
    try { await G.signIn(true); const me = await G.me(); $('msWho').textContent = 'Signed in as ' + (me.displayName || me.mail || me.userPrincipalName); }
    catch (e) { $('msWho').textContent = e.message; }
  });
  $('signOutBtn').addEventListener('click', async () => { await G.signOut(); $('msWho').textContent = 'Signed out.'; connected = false; });
  $('openFileBtn').addEventListener('click', async () => { await saveSettings(); await connect(); });
  $('gmCheckBtn').addEventListener('click', async () => {
    try {
      const r = await GM.check($('sGmToken').value);
      $('sGmToken').value = r.token;
      $('gmWho').textContent = 'Token OK — ' + r.name;
      const gs = await GM.groups(), cur = $('sGmGroup').value || S.gmGroupId;
      $('sGmGroup').innerHTML = '<option value="">Choose the dispatch chat…</option>' + gs.map(g => `<option value="${esc(g.id)}">${esc(g.name)}</option>`).join('');
      if (gs.some(g => g.id === cur)) $('sGmGroup').value = cur;
    } catch (e) { $('gmWho').textContent = e.message; }
  });

  // Refresh every 10 s while the panel is showing.
  setInterval(() => { if (!document.hidden && !busy) load(); }, 10000);
  document.addEventListener('visibilitychange', () => { if (!document.hidden) load(); });

  (async () => {
    if (window.SW_READY) await window.SW_READY;
    if (IN_EXCEL) { $('msCard').hidden = true; $('fileLinkRow').hidden = true; $('openTab').hidden = true; }
    const st = await chrome.storage.local.get(['settings', 'colors', 'toggles', 'extraCrews']);
    S = Object.assign({}, DEFAULTS, st.settings || {});
    nightState = { colors: st.colors || null, toggles: st.toggles || null, extraCrews: st.extraCrews || null };
    applyConfig();
    connect();
  })();
  window.__swd = { get night() { return night; }, load, assign, markAt, markPicked, takeOff, undo, selected, setColor, get S() { return S; } }; // console / tests
})();
