# Attachment Librarian

Outlook add-in that indexes every file attachment in your mailbox, groups
versions of the same document together (Budget_v2 / Budget FINAL / Budget (1)
are one lineage, newest first), and opens the carrying email in one click.

Outlook's attachment search can't answer "which version is the latest?" —
the Librarian can, with a single **read-only** permission and no backend.

- `manifest.xml` — add-in manifest (desktop + web; mobile section dormant)
- `src/graph.js` — MSAL nested-app-auth + Graph metadata reader (`Mail.Read`)
- `src/librarian.js` — pure lineage grouper (offline unit tests in `test/`)
- `src/taskpane/` — the Librarian pane (index, search, version timeline)

No data collection: attachment *metadata* only (names, sizes, senders,
dates — never file contents), read in the user's own Outlook session via
Microsoft Graph. Nothing leaves the Microsoft 365 boundary.

`npm run validate` checks the manifest; `npm test` runs the grouper tests.
