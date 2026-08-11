repo: danaleeskunik/kunikHome
branch: main

## Last sync
date: 2026-08-11T08:08:00Z
### Updated in this project
- Matched the Google Sheets sync pattern documented in danaleeskunik/DanaTech (two-way sync + status indicator)
- Rewrote google-apps-script.gs with doGet load-back and a _meta tab for custom categories
- Added sidebar sync status dot (idle / syncing / synced / error), click to reload from the sheet

## Reference repos
- danaleeskunik/DanaTech@main — read google-sheets-setup.txt for the prior Sheets integration pattern (read-only reference, not the source of this project's UI)

## Screen map
| Project screen | Repo files |
|---|---|
| מרכז כלכלי.dc.html (all 6 domains) | built in this project; Sheets sync pattern informed by DanaTech/google-sheets-setup.txt |
| google-apps-script.gs | written here; mirrors DanaTech's google-sheets-sync.gs approach |
