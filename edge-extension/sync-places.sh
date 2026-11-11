#!/usr/bin/env bash
# Copy the place list from the Apps Script project into the extension (run after editing apps-script/Places.gs).
cd "$(dirname "$0")"
{ printf '/* Copied unchanged from apps-script/Places.gs (place names, zones, cart rules). Refresh with: ./sync-places.sh */\n/* Browser shims: no geocoder here, so only exact place-name matching is used. */\nfunction bigGet_() { return null; }\nfunction bigSet_() {}\n\n'; cat ../../apps-script/Places.gs; } > lib/places.js
echo "lib/places.js updated"
