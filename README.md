# SURE Walk Dispatch panel

The SURE Walk dispatch panel for the rides workbook in Excel on the web. It's the original dispatch panel: crew board
and shifts, pending rides and the ride plan, assign / at / picked / take off / move / cancel / undo, picked colors,
GroupMe sync (drivers' "at / picked / empty"), ❤️ markers on ride posts, the chat and auto-assign.

## Install the SURE Walk add-in (recommended)
1. Save the manifest: open **https://leeisaac991.github.io/surewalk-dispatch/manifest.xml** and press **Ctrl+S**
   (save it as `manifest.xml`).
2. In Excel on the web, open the rides workbook ▸ **Home ▸ Add-ins ▸ More Add-ins ▸ My Add-ins ▸ Upload My Add-in** ▸
   choose `manifest.xml` ▸ **Upload**.
3. **Home ▸ SURE Walk ▸ Dispatch** opens the panel at the side. Press **⧉ Pop out** to open it in its own window that
   you can resize, maximize or move to another screen. Keep the side panel open, because it connects the window to
   the sheet.
4. **⚙ Settings** in the panel: your name, the picked colors, and GroupMe (paste your token, then pick the dispatch
   chat). Then turn on **GroupMe sync** at the top.

Each dispatcher installs it once. The add-in loads the newest panel each time it opens, so updates reach everyone
without reinstalling. Settings, crew edits and the GroupMe login are saved in each dispatcher's browser.

## Or run it through Script Lab (no install, but no Pop out)
Snippet (public gist): https://gist.github.com/leeisaac991/797e865d8e0bc875907aca9f2c645a62. In Excel: **Home ▸
Add-ins** ▸ add **Script Lab** ▸ **Script Lab ▸ Code ▸ ☰ ▸ Import** ▸ paste the gist link ▸ **Import** ▸ **Script Lab
▸ Run**. Excel blocks new windows from Script Lab, so use the add-in when you want Pop out.

## How it works
See [classic/README.md](classic/README.md). In short:
- The original page (`Sidebar.html`) and code (`Code.gs`, `GroupMe.gs`, …) run in the browser.
- `classic/host.js` reads and writes the rides tab with Office.js. Columns are found by their header in row 1
  ("Form Submitted", "Dispatcher", …), and the roster block below the rides starts in the timestamp column.
- New rides that Excel colored like the row above (copied highlight) get that color cleared.
- No tokens or rider data are stored in this repo.

| Folder | What |
|---|---|
| `classic/` | **The panel**: build (`build.sh`), host, Google-service stand-ins, tests (`test/run.sh`) |
| `docs/` | The add-in's website (built by `classic/build.sh`, served by GitHub Pages) |
| `excel-addin/`, `edge-extension/` | Earlier, simpler panel (Script Lab / Edge side panel). Not in use. |

## Update
`classic/build.sh` (rebuilds `classic/dist/` and `docs/` from `../apps-script`) ▸ commit and push. The add-in picks up
the new `docs/` within a few minutes. For the Script Lab snippet, also run `classic/snippet.sh <commit>` and update
the gist.
