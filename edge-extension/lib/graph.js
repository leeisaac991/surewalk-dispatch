/*
 * Microsoft sign-in + the Excel workbook through Microsoft Graph (the official Microsoft 365 API).
 * Sign-in: OAuth 2.0 authorization code + PKCE in an Edge pop-up (chrome.identity.launchWebAuthFlow).
 * Workbook: one persistent session; reads and writes go out in $batch requests (20 per request).
 */
/* exported G */
const G = (() => {
  const SCOPES = 'offline_access User.Read Files.ReadWrite.All';
  let cfg = { graph: 'https://graph.microsoft.com/v1.0', login: 'https://login.microsoftonline.com', tenant: 'organizations', clientId: '' };
  let tok = null;        // { access, refresh, exp }
  let wb = null;         // { drive, item, name, sheet, session }
  const store = {
    get: async k => (await chrome.storage.local.get(k))[k],
    set: (k, v) => chrome.storage.local.set({ [k]: v }),
  };

  function configure(c) { cfg = Object.assign(cfg, c); }
  class NeedSignIn extends Error {}

  /* ---------- sign-in ---------- */
  const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const rand = n => b64url(crypto.getRandomValues(new Uint8Array(n)));
  async function tokenRequest(body) {
    const res = await fetch(`${cfg.login}/${cfg.tenant}/oauth2/v2.0/token`, {
      method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(Object.assign({ client_id: cfg.clientId, scope: SCOPES }, body)),
    });
    const j = await res.json().catch(() => ({}));
    if (!res.ok || !j.access_token) throw new Error('Microsoft sign-in failed: ' + (j.error_description || j.error || res.status).toString().split('\n')[0]);
    tok = { access: j.access_token, refresh: j.refresh_token || (tok && tok.refresh), exp: Date.now() + (j.expires_in - 120) * 1000 };
    await store.set('msToken', tok);
    return tok.access;
  }
  async function signIn(interactive) {
    if (!cfg.clientId) throw new NeedSignIn('Add the Microsoft app (client) ID in Settings first.');
    const verifier = rand(48), state = rand(12);
    const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
    const redirect = chrome.identity.getRedirectURL();
    const url = `${cfg.login}/${cfg.tenant}/oauth2/v2.0/authorize?` + new URLSearchParams({
      client_id: cfg.clientId, response_type: 'code', redirect_uri: redirect, scope: SCOPES, state,
      code_challenge: challenge, code_challenge_method: 'S256', prompt: interactive ? 'select_account' : 'none',
    });
    let back;
    try { back = await chrome.identity.launchWebAuthFlow({ url, interactive: !!interactive }); }
    catch (e) { throw new NeedSignIn(interactive ? 'Sign-in was closed.' : 'Sign in with Microsoft (Settings).'); }
    const q = new URL(back).searchParams;
    if (q.get('error')) throw new NeedSignIn('Microsoft sign-in: ' + (q.get('error_description') || q.get('error')));
    if (q.get('state') !== state) throw new Error('Sign-in answer did not match — try again.');
    return tokenRequest({ grant_type: 'authorization_code', code: q.get('code'), redirect_uri: redirect, code_verifier: verifier });
  }
  async function accessToken() {
    if (!tok) tok = (await store.get('msToken')) || null;
    if (tok && tok.access && Date.now() < tok.exp) return tok.access;
    if (tok && tok.refresh) { try { return await tokenRequest({ grant_type: 'refresh_token', refresh_token: tok.refresh }); } catch (e) { /* expired: try the browser session */ } }
    return signIn(false);
  }
  async function signOut() { tok = null; wb = null; await chrome.storage.local.remove(['msToken']); }

  /* ---------- requests ---------- */
  function errOf(j, status) {
    const e = (j && j.error) || {};
    const err = new Error(e.message || `Microsoft Graph error ${status}`);
    err.code = e.code || (e.innerError && e.innerError.code) || ''; err.status = status;
    return err;
  }
  async function call(method, path, body, extraHeaders) {
    for (let attempt = 0; attempt < 3; attempt++) {
      const headers = Object.assign({ Authorization: 'Bearer ' + await accessToken() }, extraHeaders || {});
      if (body !== undefined) headers['Content-Type'] = 'application/json';
      const res = await fetch(cfg.graph + path, { method, headers, body: body === undefined ? undefined : JSON.stringify(body) });
      if (res.status === 401 && attempt === 0) { tok.exp = 0; continue; }
      if (res.status === 429 || res.status === 503) { await new Promise(r => setTimeout(r, 1000 * (Number(res.headers.get('Retry-After')) || 2))); continue; }
      const j = res.status === 204 ? {} : await res.json().catch(() => ({}));
      if (!res.ok) throw errOf(j, res.status);
      return j;
    }
    throw new Error('Microsoft is busy right now — try again in a moment.');
  }
  const quote = s => "'" + String(s).replace(/'/g, "''") + "'";
  const wsPath = () => `/drives/${wb.drive}/items/${wb.item}/workbook/worksheets(${quote(wb.sheet)})`;
  const sessionHeaders = () => (wb.session ? { 'workbook-session-id': wb.session } : {});

  // Up to 20 requests per $batch. ordered = run one after another (writes); otherwise in parallel (reads).
  async function batch(reqs, ordered) {
    if (!ordered && reqs.length > 20) { // reads: up to 4 batches at once
      const parts = [];
      for (let i = 0; i < reqs.length; i += 20) parts.push(reqs.slice(i, i + 20));
      const res = new Array(parts.length);
      let next = 0;
      await Promise.all([0, 1, 2, 3].map(async () => { while (next < parts.length) { const k = next++; res[k] = await batch(parts[k], false); } }));
      return res.flat();
    }
    const out = [];
    for (let i = 0; i < reqs.length; i += 20) {
      const chunk = reqs.slice(i, i + 20).map((r, j) => {
        const q = { id: String(j + 1), method: r.method || 'GET', url: r.url, headers: Object.assign({}, sessionHeaders(), r.body !== undefined ? { 'Content-Type': 'application/json' } : {}) };
        if (r.body !== undefined) q.body = r.body;
        if (ordered && j > 0) q.dependsOn = [String(j)];
        return q;
      });
      const j = await call('POST', '/$batch', { requests: chunk });
      const byId = {};
      (j.responses || []).forEach(x => { byId[x.id] = x; });
      chunk.forEach(q => {
        const x = byId[q.id] || { status: 500, body: {} };
        if (x.status >= 300) { const e = errOf(x.body, x.status); e.request = q.url; throw e; }
        out.push(x.body);
      });
    }
    return out;
  }

  /* ---------- the workbook ---------- */
  // Any sharing link to the file (Share ▸ Copy link) or its address bar URL.
  async function openWorkbook(link, sheetName) {
    const enc = 'u!' + btoa(unescape(encodeURIComponent(String(link).trim()))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
    const item = await call('GET', `/shares/${enc}/driveItem?$select=id,name,parentReference`);
    wb = { drive: item.parentReference.driveId, item: item.id, name: item.name, sheet: sheetName || '', session: null };
    if (!wb.sheet) {
      const list = await call('GET', `/drives/${wb.drive}/items/${wb.item}/workbook/worksheets?$select=name,position`);
      const first = (list.value || []).sort((a, b) => a.position - b.position)[0];
      if (!first) throw new Error('That workbook has no tabs.');
      wb.sheet = first.name;
    }
    await newSession();
    return { name: wb.name, sheet: wb.sheet };
  }
  async function newSession() {
    try { wb.session = (await call('POST', `/drives/${wb.drive}/items/${wb.item}/workbook/createSession`, { persistChanges: true })).id; }
    catch (e) { wb.session = null; } // works without one too, just slower
  }
  // Sessions expire after a few idle minutes: make a new one and try once more.
  async function withSession(fn) {
    if (!wb) throw new Error('Open the workbook first (Settings).');
    try { return await fn(); } catch (e) {
      if (!/session/i.test(e.code + ' ' + e.message) && e.status !== 404) throw e;
      await newSession();
      return fn();
    }
  }

  /** The used rows: { lastRow, text: rows of A:W starting at `top` }. */
  function readText(top) {
    return withSession(async () => {
      const used = await call('GET', wsPath() + '/usedRange(valuesOnly=true)?$select=rowIndex,rowCount', undefined, sessionHeaders());
      const lastRow = used.rowIndex + used.rowCount;
      const start = Math.max(2, typeof top === 'function' ? top(lastRow) : top);
      if (lastRow < start) return { lastRow, top: start, text: [] };
      const rg = await call('GET', wsPath() + `/range(address='A${start}:W${lastRow}')?$select=text`, undefined, sessionHeaders());
      return { lastRow, top: start, text: rg.text || [] };
    });
  }
  /** Fill color of whole rows A:W ('' when the row's cells differ; then the first-name cell H decides). */
  function readFills(rows) {
    return withSession(async () => {
      const res = await batch(rows.map(r => ({ url: wsPath() + `/range(address='A${r}:W${r}')/format/fill?$select=color` })));
      const out = {};
      const mixed = [];
      rows.forEach((r, i) => { const c = res[i] && res[i].color; if (c) out[r] = c.toLowerCase(); else mixed.push(r); });
      if (mixed.length) {
        const res2 = await batch(mixed.map(r => ({ url: wsPath() + `/range(address='H${r}')/format/fill?$select=color` })));
        mixed.forEach((r, i) => { out[r] = String((res2[i] && res2[i].color) || '#ffffff').toLowerCase(); });
      }
      return out;
    });
  }
  /** Fresh copies of some rows just before writing: { row: { text: [A..W], fill, font } }. */
  function readRows(rows) {
    return withSession(async () => {
      const reqs = [];
      rows.forEach(r => {
        reqs.push({ url: wsPath() + `/range(address='A${r}:W${r}')?$select=text` });
        reqs.push({ url: wsPath() + `/range(address='A${r}:W${r}')/format/fill?$select=color` });
        reqs.push({ url: wsPath() + `/range(address='A${r}:W${r}')/format/font?$select=color` });
      });
      const res = await batch(reqs);
      const out = {};
      rows.forEach((r, i) => {
        out[r] = { text: (res[i * 3].text || [[]])[0], fill: String(res[i * 3 + 1].color || '').toLowerCase(), font: String(res[i * 3 + 2].color || '#000000').toLowerCase() };
      });
      if (rows.some(r => !out[r].fill)) {
        const f = await readFills(rows.filter(r => !out[r].fill));
        Object.keys(f).forEach(r => { out[r].fill = f[r]; });
      }
      return out;
    });
  }
  /*
   * Writes, in order. ops:
   *   { t: 'v', addr: 'B5:D5', values: [[…]] }   values (strings are read like typed input: "0:33" becomes a time)
   *   { t: 'clr', addr: 'M5:P5' }                clear contents
   *   { t: 'fill', row: 5, color: '#000000' | null }   whole-row fill (null = no fill)
   *   { t: 'font', row: 5, color: '#ffffff' }    whole-row text color
   */
  function write(ops) {
    if (!ops.length) return Promise.resolve();
    return withSession(() => batch(ops.map(op => {
      if (op.t === 'v') return { method: 'PATCH', url: wsPath() + `/range(address='${op.addr}')`, body: { values: op.values } };
      if (op.t === 'clr') return { method: 'POST', url: wsPath() + `/range(address='${op.addr}')/clear`, body: { applyTo: 'Contents' } };
      if (op.t === 'fill') return op.color && op.color !== '#ffffff'
        ? { method: 'PATCH', url: wsPath() + `/range(address='${op.row}:${op.row}')/format/fill`, body: { color: op.color } }
        : { method: 'POST', url: wsPath() + `/range(address='${op.row}:${op.row}')/format/fill/clear`, body: {} };
      if (op.t === 'font') return { method: 'PATCH', url: wsPath() + `/range(address='${op.row}:${op.row}')/format/font`, body: { color: op.color || '#000000' } };
      throw new Error('Unknown write ' + op.t);
    }), true));
  }
  async function me() { return call('GET', '/me?$select=displayName,mail,userPrincipalName'); }

  return { configure, signIn, signOut, accessToken, openWorkbook, readText, readFills, readRows, write, me, NeedSignIn,
    redirectUrl: () => chrome.identity.getRedirectURL(), get workbook() { return wb; } };
})();
