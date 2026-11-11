#!/usr/bin/env bash
# Build the CLASSIC panel (the original dispatch panel, running inside Excel) from ../../apps-script:
#   dist/backend.js   services.js + Places/Crews/GroupMe/Code/Excel/Build .gs + worker-entry.js (comments removed)
#   dist/Sidebar.html the original page (comments removed)
#   dist/host.js      runs it all in Excel (see host.js)
# Then ./snippet.sh <commit> writes the Script Lab snippet that loads these files from GitHub (via jsDelivr).
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"; OUT="$HERE/dist"
# the Apps Script project: ../../apps-script from this repo (searched upward, so a worktree works too), or GS=...
if [ -z "${GS:-}" ]; then d="$HERE"; while [ "$d" != / ] && [ ! -f "$d/apps-script/Code.gs" ]; do d=$(dirname "$d"); done; GS="$d/apps-script"; fi
[ -f "$GS/Code.gs" ] || { echo "apps-script/ not found; set GS=/path/to/apps-script"; exit 1; }
mkdir -p "$OUT"
TMP=$(mktemp -d); trap 'rm -rf "$TMP"' EXIT
{
  cat "$HERE/services.js"; echo
  for f in Build Places Crews GroupMe Code Excel; do echo ";/* ---- $f.gs ---- */"; cat "$GS/$f.gs"; echo; done
  cat "$HERE/worker-entry.js"
} > "$TMP/backend.js"
# the old Google sheet's id and web-app address don't belong in a public file
sed -i -E "s/SPREADSHEET_ID: '[^']*'/SPREADSHEET_ID: ''/" "$TMP/backend.js"
node "$HERE/strip-comments.js" "$TMP/backend.js" > "$OUT/backend.js"
sed -E 's#https://script\.google\.com/macros/s/[A-Za-z0-9_-]+/exec#https://script.google.com/macros/s/RELOAD/exec#g' "$GS/Sidebar.html" > "$TMP/Sidebar.html"
node "$HERE/strip-comments.js" --html "$TMP/Sidebar.html" > "$OUT/Sidebar.html"
cp "$HERE/host.js" "$OUT/host.js"
node --check "$OUT/backend.js"
node --check "$OUT/host.js"
python3 -c 'import re,sys;print("\n".join(re.findall(r"<script>(.*?)</script>",open(sys.argv[1]).read(),re.S)))' "$OUT/Sidebar.html" > "$TMP/side.js" && node --check "$TMP/side.js"
echo "built $OUT: $(du -h "$OUT/backend.js" | cut -f1) backend, $(du -h "$OUT/Sidebar.html" | cut -f1) page"
