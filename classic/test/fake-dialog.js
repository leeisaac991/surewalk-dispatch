// Office.js inside the test dialog window (stands in for Excel's dialog runtime): messageParent / DialogParentMessageReceived.
window.Office = {
  EventType: { DialogParentMessageReceived: 'dp' }, AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
  onReady: cb => setTimeout(() => cb({ host: 'Excel' }), 0),
  context: { ui: {
    messageParent(s) { window.opener.__fakeDlg.toParent(s); },
    addHandlerAsync(type, fn, done) {
      window.__fromParent = s => setTimeout(() => fn({ message: s, origin: location.origin }), 0);
      (window.__swcQueue || []).splice(0).forEach(s => window.__fromParent(s));
      if (done) done({ status: 'succeeded' });
    },
  } },
};
