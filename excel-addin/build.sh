#!/usr/bin/env bash
# Build the Excel add-in from the same panel as the Edge extension (../edge-extension) into dist/.
#   ./build.sh https://<you>.github.io/surewalk-excel-addin/      (the address the files will be served from)
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"; SRC="$HERE/../edge-extension"; OUT="${OUT:-$HERE/dist}"
BASE="${1:?usage: ./build.sh <https base url ending in />}"
mkdir -p "$OUT/lib" "$OUT/assets"
python3 - "$SRC/panel.html" "$OUT/taskpane.html" <<'PY'
import sys
s = open(sys.argv[1]).read()
office = '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>\n<script src="lib/office-host.js"></script>\n'
assert '<script src="lib/places.js">' in s and '<script src="lib/graph.js"></script>' in s
s = s.replace('<script src="lib/places.js">', office + '<script src="lib/places.js">', 1)
s = s.replace('<script src="lib/graph.js"></script>', '<script src="lib/excel.js"></script>')
open(sys.argv[2], 'w').write(s)
PY
cp "$SRC/panel.css" "$SRC/panel.js" "$OUT/"
cp "$SRC"/lib/{places,rules,excel,groupme,office-host}.js "$OUT/lib/"
cp "$HERE"/assets/*.png "$OUT/assets/"
sed "s#{{BASE}}#$BASE#g" "$HERE/manifest.template.xml" > "$OUT/manifest.xml"
touch "$OUT/.nojekyll"
echo "built $OUT (manifest points at $BASE)"
