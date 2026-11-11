# SURE Walk Dispatch — the original panel, inside Excel

This is the original dispatch panel (the Apps Script panel's own page and code), running in **Script Lab** on the Excel
rides workbook. It has the same look and features: crew board and shifts, pending rides and the ride plan,
assign / at / picked / take off / move / cancel / undo, picked colors, GroupMe sync (drivers' "at / picked / empty"),
❤️ markers on ride posts, the chat, auto-assign and alerts. Nothing runs on Google's servers.

**Snippet (public gist):** https://gist.github.com/leeisaac991/797e865d8e0bc875907aca9f2c645a62

## Use it
1. In Excel on the web, open the rides workbook ▸ **Home ▸ Add-ins** ▸ add **Script Lab**.
2. **Script Lab ▸ Code ▸ ☰ ▸ Import** ▸ paste the gist link ▸ **Import**.
3. **Script Lab ▸ Run**. The panel opens beside the sheet.
4. **⚙ Settings** in the panel: your name, the picked colors, and GroupMe (token, then pick the chat). Then turn on
   **GroupMe sync** at the top.

Each dispatcher sets this up once in their own browser. Their settings, crew edits and GroupMe login stay there.

**More room:** press **⧉ Pop out** (the orange strip at the top of the side panel) to open the panel in its own
window. Resize it, maximize it, or move it to another screen. Keep the Excel side panel open, because it's the
window's connection to the sheet. Close the window (or press **Bring it back here**) to return to the side panel. If
the browser blocks the window, allow pop-ups for Excel (the icon at the right of the address bar).

## How it works
- `host.js` stands in for `google.script.run`. Before each panel call it reads the rides tab with Office.js and
  presents it in the old Google-sheet column layout, so the original code runs unchanged:
  - ride columns are found by their header in row 1 (Form Submitted, Dispatcher, Call Time, …);
  - the roster block starts in the timestamp column.
  It runs the call and writes the changes back into the right Excel cells.
- `backend.js` = `services.js` (stand-ins for Google's services: properties in this browser, GroupMe through the
  browser) + the original `Places/Crews/GroupMe/Code/Excel.gs` + `worker-entry.js`. It runs in a Web Worker, so
  GroupMe calls never freeze the panel.
- **Clock:** Apps Script ran on Google's servers. Here "now" comes from the dispatcher's computer, corrected forward
  if GroupMe's timestamps show the computer is behind. Otherwise a crew that just said "empty" would stay "available"
  after being assigned.
- **Copied highlights:** Excel gives a new form response the formatting of the row above. A new ride that nobody has
  touched yet (no dispatcher, call time, arrival, crew or reason) and that has exactly the row above's color gets that
  color cleared, for rides from the last hour only.

## Update
`./build.sh` (rebuilds `dist/` from `../../apps-script`) ▸ commit and push ▸ `./snippet.sh <that commit>` ▸ update
the gist with `scriptlab/SURE-Walk-Dispatch-Classic.yaml`. The snippet loads `dist/` at that exact commit (through
jsDelivr), so open panels never change mid-shift.

Test: `./test/run.sh`. It runs the real built files in headless Chromium against a fake Excel and a fake GroupMe.
It needs `test/sheet.json`, a rides export (not in the repo).
