# SURE Walk Dispatch inside Excel (no Microsoft app registration)

The same simple dispatch panel as the Edge extension (`../edge-extension`), running inside Excel on the web. It
reads and writes the open workbook through Excel's own add-in API, so there's no sign-in setup, no Apps Script and
no Microsoft app registration. Ride messages post to GroupMe with each dispatcher's own token.

Two ways to run it:

## Option A — Script Lab (easiest: nothing to host)
Script Lab is Microsoft's free add-in for running your own code in Excel.

1. Build the snippet: `./scriptlab.sh` → `scriptlab/SURE-Walk-Dispatch.yaml`.
2. In Excel on the web, open the rides workbook ▸ **Home ▸ Add-ins** ▸ search **Script Lab** ▸ **Add**.
   (If UT blocks store add-ins, it won't show up; use the Edge extension instead.)
3. **Script Lab tab ▸ Code** ▸ the ☰ menu ▸ **Import** ▸ paste the whole contents of `SURE-Walk-Dispatch.yaml`
   ▸ **Import**.
4. **Script Lab tab ▸ Run** opens the panel beside the sheet.
5. **⚙ Settings**: your name, your GroupMe token (**Check & load chats**, then pick the dispatch chat), and the rides
   tab if it isn't the first tab. **Save**.

Each dispatcher imports the snippet once, in their own browser. Settings are saved in that browser. After an update,
re-import the new YAML.

## Option B — a real add-in (a "SURE Walk ▸ Dispatch" button on the Home tab)
The files must be served from a public HTTPS address. `./publish.sh` creates a public GitHub repo with Pages
turned on. The files are code only: tokens stay in each dispatcher's browser.
Then: **Home ▸ Add-ins ▸ More Add-ins ▸ My Add-ins ▸ Upload My Add-in** ▸ `dist/manifest.xml`.

## Notes
- The panel code lives in `../edge-extension` (`panel.*`, `lib/`). `build.sh` (add-in) and `scriptlab.sh` (snippet)
  assemble it with `lib/excel.js` (the workbook through Office.js) in place of `lib/graph.js` (Microsoft Graph).
- Test: `ADDIN=1 ../edge-extension/test/run.sh` runs the built add-in against a fake Excel and a fake GroupMe.
