# SURE Walk Dispatch panel

A simple dispatch panel for the SURE Walk rides workbook in Excel on the web: assign crews, at, picked, take off,
cancel, duplicate, edit and undo, with ride messages posted to GroupMe. The same code runs three ways:

| Way | Folder | Needs |
|---|---|---|
| **Script Lab snippet** (easiest) | `excel-addin/scriptlab/SURE-Walk-Dispatch.yaml` | the free Script Lab add-in in Excel |
| Excel add-in (own Home-tab button) | `excel-addin/` (`build.sh`, `publish.sh`) | the files hosted on HTTPS |
| Microsoft Edge side panel | `edge-extension/` | a Microsoft (Entra) app registration |

## Script Lab
In Excel on the web, open the rides workbook ▸ **Home ▸ Add-ins** ▸ add **Script Lab** ▸ **Script Lab ▸ Code ▸ ☰ ▸
Import** ▸ paste the gist link (or the YAML) ▸ **Import** ▸ **Script Lab ▸ Run**. Then **⚙ Settings**: your name,
your GroupMe token (**Check & load chats** ▸ pick the dispatch chat) and the rides tab if it isn't the first tab.

No tokens are stored in this repo: each dispatcher's GroupMe token stays in their own browser.

## Development
- Panel: `edge-extension/panel.*` and `edge-extension/lib/` (`rules.js` = dispatch rules, `places.js` = place names,
  `excel.js` = workbook via Office.js, `graph.js` = workbook via Microsoft Graph, `groupme.js` = posting).
- After a change: `excel-addin/scriptlab.sh` rebuilds the snippet.
- Tests: `edge-extension/test/run.sh` (Edge/Graph) and `ADDIN=1 edge-extension/test/run.sh` (Excel/Office.js). They
  need a sheet export: `EXPORT=/path/to/export.json`.
