/*
 * GroupMe: post ride messages to the dispatch chat with your own access token (no chat view).
 * Token: web.groupme.com ▸ F12 ▸ Network ▸ any api.groupme.com request ▸ the X-Access-Token header.
 */
/* exported GM */
const GM = (() => {
  let base = 'https://api.groupme.com/v3', token = '', groupId = '', members = null, membersAt = 0;
  function configure(c) {
    if (c.base) base = c.base;
    if ('token' in c && c.token !== token) { token = c.token || ''; members = null; }
    if ('groupId' in c && c.groupId !== groupId) { groupId = c.groupId || ''; members = null; }
  }
  const ready = () => !!(token && groupId);
  async function req(method, path, body) {
    if (!token) throw new Error('Add your GroupMe token in Settings.');
    const res = await fetch(base + path, { method, headers: Object.assign({ 'X-Access-Token': token }, body ? { 'Content-Type': 'application/json' } : {}),
      body: body ? JSON.stringify(body) : undefined });
    if (res.status === 401) throw new Error('GroupMe rejected the token — copy a fresh one from web.groupme.com.');
    if (res.status >= 300) throw new Error(`GroupMe returned ${res.status}`);
    const t = await res.text();
    return t ? (JSON.parse(t).response || {}) : {};
  }
  const cleanToken = raw => String(raw || '').replace(/^\s*x-access-token\s*:?\s*/i, '').trim();
  async function check(raw) {
    const t = cleanToken(raw);
    if (!/^[A-Za-z0-9_-]{20,128}$/.test(t)) throw new Error("That doesn't look like a GroupMe token — copy just the value after X-Access-Token.");
    const old = token; token = t;
    try { const me = await req('GET', '/users/me'); return { token: t, name: me.name }; } catch (e) { token = old; throw e; }
  }
  async function groups() { return (await req('GET', '/groups?per_page=100&omit=memberships')).map(g => ({ id: String(g.id), name: g.name })); }
  async function loadMembers() {
    if (members && Date.now() - membersAt < 5 * 60e3) return members;
    const g = await req('GET', `/groups/${groupId}`);
    members = (g.members || []).map(m => ({ id: String(m.user_id), nick: m.nickname || m.name || '', real: m.name || '' }));
    membersAt = Date.now();
    return members;
  }
  // "@Name" for each crew member that matches exactly one chat member by first name (else a plain "@Name").
  async function tagsFor(names) {
    let list = [];
    try { list = await loadMembers(); } catch (e) { /* plain tags */ }
    const first = s => R.norm(String(s).split(/\s+/)[0]);
    return names.filter(Boolean).map(n => {
      const hits = list.filter(m => R.sameName(m.nick, n) || R.sameName(m.real, n));
      const exact = hits.filter(m => first(m.nick) === first(n) || first(m.real) === first(n));
      const m = (exact.length === 1 ? exact : hits.length === 1 ? hits : [])[0];
      return { tag: '@' + (m ? m.nick : n), id: m ? m.id : null };
    });
  }
  /** Posts text with the crew tagged at the end. Returns the message id (for undo). */
  async function post(text, names) {
    if (!ready()) throw new Error('GroupMe is not set up (Settings).');
    const tags = await tagsFor(names || []);
    let full = String(text).trim();
    const loci = [], ids = [];
    if (tags.length) {
      full += '\n';
      tags.forEach((t, i) => {
        if (i) full += ' ';
        if (t.id) { loci.push([full.length, t.tag.length]); ids.push(t.id); }
        full += t.tag;
      });
    }
    if (full.length > 1000) throw new Error('GroupMe messages are limited to 1,000 characters.');
    const message = { source_guid: crypto.randomUUID(), text: full };
    if (ids.length) message.attachments = [{ type: 'mentions', user_ids: ids, loci }];
    const r = await req('POST', `/groups/${groupId}/messages`, { message });
    return r.message ? String(r.message.id) : '';
  }
  async function remove(messageId) {
    if (messageId) await req('DELETE', `/conversations/${groupId}/messages/${messageId}`);
  }
  return { configure, ready, check, groups, post, remove, tagsFor };
})();
