#!/usr/bin/env bash
# Write the Script Lab snippet for the classic panel: a small loader that fetches dist/ from GitHub (through the
# jsDelivr CDN) at one exact commit, so a dispatcher's panel never changes mid-shift.
#   ./snippet.sh <commit-sha>      -> scriptlab/SURE-Walk-Dispatch-Classic.yaml
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
SHA="${1:?usage: ./snippet.sh <commit sha that contains classic/dist>}"
REPO="${REPO:-leeisaac991/surewalk-dispatch}"
mkdir -p "$HERE/scriptlab"
cat > "$HERE/scriptlab/SURE-Walk-Dispatch-Classic.yaml" <<YAML
name: SURE Walk Dispatch
description: The SURE Walk dispatch panel (the original one), running on this Excel workbook. Build ${SHA:0:7}.
host: EXCEL
api_set: {}
script:
  content: |
    // @ts-nocheck
    // Loads the panel from github.com/${REPO} (folder classic/dist, commit ${SHA:0:7}). Nothing to edit here.
    const BASE = "https://cdn.jsdelivr.net/gh/${REPO}@${SHA}/classic/dist/";
    document.body.innerHTML = '<div style="font:14px Segoe UI,system-ui,sans-serif;padding:24px;color:#6e6e73">Loading SURE Walk Dispatch…</div>';
    Promise.all(["backend.js", "Sidebar.html", "host.js"].map(f =>
      fetch(BASE + f).then(r => { if (!r.ok) throw new Error(f + " (" + r.status + ")"); return r.text(); })))
      .then(([backend, sidebar, host]) => { window.__SWC = { backend, sidebar }; (0, eval)(host); })
      .catch(e => { document.body.innerHTML = '<p style="font:14px Segoe UI,system-ui,sans-serif;padding:24px;color:#c62828">Could not load the panel: ' + e.message + '. Check the internet connection and press Run again.</p>'; });
  language: typescript
template:
  content: |
    <div></div>
  language: html
style:
  content: |
    body { margin: 0; }
  language: css
libraries: |
  https://appsforoffice.microsoft.com/lib/1/hosted/office.js
  @types/office-js
YAML
echo "wrote $HERE/scriptlab/SURE-Walk-Dispatch-Classic.yaml (commit ${SHA:0:7})"
