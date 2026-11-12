#!/usr/bin/env bash
# End-to-end test of the classic panel: the real built files in headless Chromium, a fake Excel workbook made from a
# sheet export (rides with a crew and an arrival are colored, like a real night), and a fake GroupMe.
#   EXPORT=rides.json [PAIRINGS=pairings.json] NOW=2026-11-11T05:20:00Z ROWS=226,225,223,224 ./classic/test/run.sh
# ROWS = the rides used: assign/at/picked, cancel/undo/edit, GroupMe sync (pickup Jester), and a touched row
# (the first two and the last are painted like the touched one: copied highlights).
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
EXPORT="${EXPORT:-$HERE/sheet.json}"; export NOW="${NOW:-2026-11-11T05:20:00Z}"; export ROWS="${ROWS:-226,225,223,224}"
[ -n "${PAIRINGS:-}" ] && export PAIRINGS
"$HERE/../build.sh" >/dev/null
PAINT=$(python3 - "$EXPORT" "$ROWS" <<'PY'
import json, sys
d = json.load(open(sys.argv[1]))
ra, rc, rs, rt = [int(x) for x in sys.argv[2].split(',')]
h = [str(x).strip().lower() for x in d[0]]
drv = h.index('driver'); arr = [i for i, x in enumerate(h) if x.startswith('arrival')][0]
p = [[i + 1, '#00FF00'] for i, r in enumerate(d) if i > 0 and len(r) > drv and r[drv].strip() and r[arr].strip()]
p += [[rt, '#00FF00'], [rc, '#00FF00'], [ra, '#00FF00']]  # rc/ra untouched: their copied color gets cleared
print(json.dumps(p))
PY
)
export PAINT
cd "$HERE"; node server.js "$EXPORT" 8778 >/dev/null & SRV=$!
PROF=$(mktemp -d)
# the add-in's side panel page (docs/taskpane.html); SCRIPTLAB=1 tests the Script Lab way instead
URL=http://localhost:8778/site/taskpane.html; [ -n "${SCRIPTLAB:-}" ] && URL=http://localhost:8778/
chromium-browser --headless=new --remote-debugging-port=9337 --password-store=basic --user-data-dir="$PROF" --window-size=430,950 "$URL" >/dev/null 2>&1 & CHR=$!
trap 'kill $SRV $CHR 2>/dev/null; sleep 1; rm -rf "$PROF" 2>/dev/null || true' EXIT
sleep 1; timeout 230 node e2e.mjs
