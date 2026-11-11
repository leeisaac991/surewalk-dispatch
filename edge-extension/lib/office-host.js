/*
 * Running as an Excel add-in instead of an Edge extension: the bits of the chrome.* API the panel uses,
 * backed by this browser's localStorage. Loaded before panel.js; panel.js waits for Office to be ready.
 */
window.SW_HOST = 'excel';
window.chrome = {
  storage: { local: {
    async get(keys) {
      const out = {};
      [].concat(keys).forEach(k => { try { const v = localStorage.getItem('swd.' + k); if (v != null) out[k] = JSON.parse(v); } catch (e) { /* blocked */ } });
      return out;
    },
    async set(obj) { Object.keys(obj).forEach(k => { try { localStorage.setItem('swd.' + k, JSON.stringify(obj[k])); } catch (e) { /* this session only */ } }); },
    async remove(keys) { [].concat(keys).forEach(k => { try { localStorage.removeItem('swd.' + k); } catch (e) { /* blocked */ } }); },
  } },
  runtime: { getURL: p => p },
  tabs: { create: o => window.open(o.url, '_blank') },
};
window.SW_READY = new Promise(resolve => Office.onReady(resolve));
