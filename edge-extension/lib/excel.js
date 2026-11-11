/*
 * The workbook through Office.js — used when the panel runs as an Excel add-in (inside Excel on the web).
 * Same interface as lib/graph.js, so panel.js works unchanged. No Microsoft app registration needed:
 * the add-in uses the workbook that's open, with the signed-in Excel user's access.
 */
/* exported G */
const G = (() => {
  class NeedSignIn extends Error {}
  let wb = null; // { name, sheet }
  const sheetOf = ctx => (wb && wb.sheet ? ctx.workbook.worksheets.getItem(wb.sheet) : ctx.workbook.worksheets.getFirst());

  async function openWorkbook(_link, sheetName) {
    wb = { name: '', sheet: sheetName || '' };
    return Excel.run(async ctx => {
      const ws = sheetOf(ctx);
      ws.load('name'); ctx.workbook.load('name');
      await ctx.sync().catch(e => { throw new Error(e.code === 'ItemNotFound' ? `There's no tab named "${sheetName}" in this workbook.` : e.message); });
      wb = { name: ctx.workbook.name || 'This workbook', sheet: ws.name };
      return { name: wb.name, sheet: wb.sheet };
    });
  }
  function readText(top) {
    return Excel.run(async ctx => {
      const ws = sheetOf(ctx), used = ws.getUsedRangeOrNullObject(true);
      used.load('rowIndex,rowCount');
      await ctx.sync();
      const lastRow = used.isNullObject ? 1 : used.rowIndex + used.rowCount;
      const start = Math.max(2, typeof top === 'function' ? top(lastRow) : top);
      const L = R.colLetter(R.READ_COLS), hd = ws.getRange(`A1:${L}1`);
      hd.load('text');
      if (lastRow < start) { await ctx.sync(); return { lastRow, top: start, header: hd.text[0], text: [] }; }
      const rg = ws.getRange(`A${start}:${L}${lastRow}`);
      rg.load('text');
      await ctx.sync();
      return { lastRow, top: start, header: hd.text[0], text: rg.text };
    });
  }
  // One color per row, judged like the old panel: 3+ black cells = black, 3+ red = red, else the most common real color.
  function rowColor(cells) {
    const n = {};
    cells.forEach(c => { c = String(c || '#ffffff').toLowerCase(); n[c] = (n[c] || 0) + 1; });
    const of = kind => Object.keys(n).filter(c => R.rowKind(c).kind === kind && n[c] >= 3).sort((a, b) => n[b] - n[a])[0];
    return of('black') || of('red') || of('color') || '#ffffff';
  }
  function readFills(rows) {
    return Excel.run(async ctx => {
      const ws = sheetOf(ctx);
      const props = rows.map(r => ws.getRange(`A${r}:${R.colLetter(R.LAST_COL)}${r}`).getCellProperties({ format: { fill: { color: true } } }));
      await ctx.sync();
      const out = {};
      rows.forEach((r, i) => { out[r] = rowColor(props[i].value[0].map(p => p.format.fill.color)); });
      return out;
    });
  }
  function readRows(rows) {
    return Excel.run(async ctx => {
      const ws = sheetOf(ctx);
      const got = rows.map(r => {
        const rg = ws.getRange(`A${r}:${R.colLetter(R.LAST_COL)}${r}`);
        rg.load('text');
        return { rg, props: rg.getCellProperties({ format: { fill: { color: true }, font: { color: true } } }) };
      });
      await ctx.sync();
      const out = {};
      rows.forEach((r, i) => {
        const cells = got[i].props.value[0];
        out[r] = { text: got[i].rg.text[0], fill: rowColor(cells.map(p => p.format.fill.color)),
          font: String(cells[0].format.font.color || '#000000').toLowerCase() };
      });
      return out;
    });
  }
  // Same ops as lib/graph.js write(): values, clear, whole-row fill, whole-row text color — applied in order.
  function write(ops) {
    if (!ops.length) return Promise.resolve();
    return Excel.run(async ctx => {
      const ws = sheetOf(ctx);
      ops.forEach(op => {
        if (op.t === 'v') ws.getRange(op.addr).values = op.values;
        else if (op.t === 'clr') ws.getRange(op.addr).clear(Excel.ClearApplyTo.contents);
        else if (op.t === 'fill') {
          const rg = ws.getRange(`${op.row}:${op.row}`);
          if (op.color && op.color !== '#ffffff') rg.format.fill.color = op.color; else rg.format.fill.clear();
        } else if (op.t === 'font') ws.getRange(`${op.row}:${op.row}`).format.font.color = op.color || '#000000';
      });
      await ctx.sync();
    });
  }
  const none = async () => { throw new Error('Not needed in the Excel add-in.'); };
  return { configure() {}, openWorkbook, readText, readFills, readRows, write, NeedSignIn,
    signIn: none, signOut: async () => {}, me: async () => ({}), redirectUrl: () => '', get workbook() { return wb; } };
})();
