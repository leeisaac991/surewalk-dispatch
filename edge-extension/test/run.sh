#!/usr/bin/env bash
# End-to-end test: the real panel in headless Chromium against a fake Microsoft Graph workbook + fake GroupMe.
#   EXPORT=/path/to/sheet-export.json ./test/run.sh      (a JSON array of rows of display text, like the live sheet)
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
EXPORT="${EXPORT:-$HOME/.claude/jobs/c7e493ef/tmp/tx/live_1104.json}"
cd "$HERE"
# ADDIN=1 ./test/run.sh tests the built Excel add-in (../excel-addin/build.sh) instead of the Edge extension
URL=http://localhost:8777/
if [ -n "${ADDIN:-}" ]; then
  export ADDIN_DIST=$(mktemp -d); OUT="$ADDIN_DIST" ../../excel-addin/build.sh http://localhost:8777/addin/ >/dev/null; URL=http://localhost:8777/addin/
fi
node fake-server.js "$EXPORT" 8777 > /dev/null & SRV=$!
PROF=$(mktemp -d)
chromium-browser --headless=new --remote-debugging-port=9335 --password-store=basic --user-data-dir="$PROF" --window-size=420,950 "$URL" >/dev/null 2>&1 & CHR=$!
trap 'kill $SRV $CHR 2>/dev/null; sleep 1; rm -rf "$PROF" "$HERE"/e2e-*.png ${ADDIN_DIST:-} 2>/dev/null || true' EXIT
sleep 1; timeout 170 node e2e.mjs
