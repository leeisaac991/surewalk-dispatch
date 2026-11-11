#!/usr/bin/env bash
# Publish the add-in (built from ../edge-extension by build.sh) to GitHub Pages.
# The files hold no secrets: each dispatcher types their GroupMe token into the add-in on their own computer.
#
#   ./excel-addin/publish.sh                 first time: creates the public repo <you>/surewalk-excel-addin + Pages
#   ./excel-addin/publish.sh                 later: pushes the changed files (Pages updates within ~1 minute)
#
# Then upload dist/manifest.xml in Excel: Home ▸ Add-ins ▸ More Add-ins ▸ My Add-ins ▸ Upload My Add-in.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="${REPO:-surewalk-excel-addin}"
OWNER=$(gh api user -q .login)
BASE="https://$OWNER.github.io/$REPO/"
DIST="$HERE/dist"

rm -rf "$DIST.new"
OUT="$DIST.new" "$HERE/build.sh" "$BASE"
if [ -d "$DIST/.git" ]; then mv "$DIST/.git" "$DIST.new/.git"; fi
rm -rf "$DIST"; mv "$DIST.new" "$DIST"

cd "$DIST"
if [ ! -d .git ]; then
  git init -q -b main
  gh repo view "$OWNER/$REPO" >/dev/null 2>&1 || gh repo create "$OWNER/$REPO" --public -d "SURE Walk dispatch — Excel add-in files" >/dev/null
  git remote add origin "https://github.com/$OWNER/$REPO.git"
fi
git add -A
git commit -qm "Publish add-in $(date '+%Y-%m-%d %H:%M')" || { echo "Nothing changed."; exit 0; }
git push -q -u origin main
gh api "repos/$OWNER/$REPO/pages" >/dev/null 2>&1 || gh api -X POST "repos/$OWNER/$REPO/pages" -f 'source[branch]=main' -f 'source[path]=/' >/dev/null
echo "✓ Published to $BASE"
echo "  Manifest for Excel: $DIST/manifest.xml"
