/*
 * Runs one panel call the way Apps Script would: fresh per-call state, the workbook snapshot as the sheet,
 * then hands back the result, the sheet writes (ops) and the properties if they changed.
 * Message in:  { id, fn, args, snap, uid, props? }    out: { id, ok, result, error, ops, props? }
 */
function __swReset() {
  SP_CACHE_ = null; UP_CACHE_ = null; LAST_RIDE_CACHE_ = null; ROSTER_CACHE_ = null;
  GM_OFFLINE_ = false; CREW_NAMES_ = new Set(); POSTER_NAMES_ = new Set(); BATCH_MSGS_ = []; LAST_UPDATE_ = {};
}
function __swCall(msg) {
  if (msg.props) __SW.script = msg.props;
  __SW.changed = false;
  __swReset();
  const out = { id: msg.id };
  if (!xlCallable_(msg.fn)) { out.ok = false; out.error = `Unknown function "${msg.fn}".`; out.ops = []; return out; }
  XL_ = { uid: String(msg.uid), snap: msg.snap, ops: [], sheet: null };
  try {
    const r = globalThis[msg.fn].apply(null, msg.args || []);
    out.ok = true;
    out.result = r === undefined ? null : JSON.parse(JSON.stringify(r)); // what google.script.run would hand the page
  } catch (e) { out.ok = false; out.error = (e && e.message) || String(e); }
  out.ops = XL_.ops;
  XL_ = null;
  if (__SW.changed) out.props = __SW.script;
  return out;
}
if (typeof WorkerGlobalScope !== 'undefined' && self instanceof WorkerGlobalScope) {
  self.onmessage = e => { self.postMessage(__swCall(e.data)); };
} else {
  globalThis.__swCall = __swCall; // no workers allowed: the host calls this directly
}
