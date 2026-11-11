# SURE Walk Dispatch — Microsoft Edge extension

A side panel in Edge that dispatches straight from the rides workbook in Excel on the web. It reads and edits the
sheet itself through Microsoft Graph (Microsoft's official API for Microsoft 365 files). Ride messages post to
GroupMe with your token. Nothing runs inside the workbook, and no Apps Script is involved.

## What it does
- **Pending**: new form rides, flagged as possible duplicates (same route, name and phone within 40 min) or
  off-campus pickups. Tick rides, choose a crew, then **Assign**. This writes Dispatch (B), Call time (C) if blank,
  and Location / Van / Driver / Navigator (M–P), then posts the ride lines to GroupMe and tags the crew. A crew that
  already has riders gets the "Updated ride! … NEW:" message.
- **Crews**: tonight's roster (the block under the rides) for the shifts that are on. A/M/B switch by the clock and
  can also be toggled. You can add crews by hand. Riders are grouped by pickup, each group with:
  **At** (Arrival D), **Picked** (row turns tonight's color: regular for A shift, pastel for B; Arrival is filled in
  if it was skipped) and **Take off** (back to Pending; posts "REMOVED: name").
- **⋯ on a ride**: Edit (any dispatch field), Mark duplicate (red row), Cancel (reason in V, row turns black, B–D and
  M–P cleared; posts "CANCELLED: name" to the crew).
- **Undo** (in the message bar after each action, for 30 min): puts the cells and colors back and deletes the
  GroupMe post the action made.
- Uses the same rules as the old panel: seats per vehicle, golf carts only on campus and never to the garages, and
  finished rows older than 60 minutes are left alone. Rows are re-read just before every write.

## One-time setup
1. **Register the app with Microsoft** (≈5 min): [portal.azure.com](https://portal.azure.com) ▸ *App registrations*
   ▸ *New registration*.
   - Name: `SURE Walk Dispatch`. Account types: *Accounts in any organizational directory*.
   - Redirect URI: platform **Single-page application (SPA)**, value
     `https://cegmolcbfonbjcbgapchnpffnmcknebg.chromiumapp.org/`
   - *API permissions* ▸ Add ▸ Microsoft Graph ▸ Delegated: `Files.ReadWrite.All`, `User.Read`, `offline_access`.
   - Copy the **Application (client) ID**.
   - If UT doesn't let students register apps, or asks for admin approval when you sign in, register it from a
     personal Azure account instead (same settings), or ask UT IT to approve it.
2. **Install in Edge**: `edge://extensions` ▸ turn on *Developer mode* ▸ *Load unpacked* ▸ choose this
   `edge-extension` folder. The ID stays `cegmolcbfonbjcbgapchnpffnmcknebg` on every computer, so the redirect URI
   above always matches. Pin the toolbar button; clicking it opens the side panel.
3. **Settings (⚙)**:
   - *You*: your name for the Dispatch column (e.g. `I. Lee`).
   - *Microsoft sign-in*: paste the client ID, then **Sign in**.
   - *Rides workbook*: paste the workbook's link (Share ▸ Copy link). Leave *Tab* blank for the first tab.
     **Connect**.
   - *GroupMe*: paste your token (web.groupme.com ▸ F12 ▸ Network ▸ any api.groupme.com request ▸
     X-Access-Token), **Check & load chats**, then pick the dispatch chat. If you turn off "Post automatically", the
     message is copied to your clipboard instead.
   - **Save**.

The workbook must keep the old sheet's columns (A timestamp … W notes), with the roster block below the rides.

## Files
- `panel.html` / `panel.css` / `panel.js`: the side panel. `background.js`: makes the toolbar button open it.
- `lib/rules.js`: dispatch rules (colors, roster, shifts, duplicates, seat and cart checks, GroupMe text).
- `lib/places.js`: a copy of `apps-script/Places.gs`. Run `./sync-places.sh` after editing that file.
- `lib/graph.js`: Microsoft sign-in and workbook reads/writes. `lib/groupme.js`: posting.
- `test/run.sh`: end-to-end test (the real panel against a fake Graph workbook and fake GroupMe).
