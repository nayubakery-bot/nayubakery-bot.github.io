# Rise & Whisk — bakery book (v1)

Mobile-first PWA for a home bakery in India (₹ INR). The price calculator replicates the Excel workbook exactly (verified against the sample 0.5 kg chocolate birthday cake quote of ₹780).

## Default PIN
`2468` — change it in Settings after signing in.

## Google Sheet
https://docs.google.com/spreadsheets/d/1LgfA7HGWy-KA5PcA6qQ7d1omzuJwiWKXIgpjEIcvyQI/edit

The original Excel file was not modified. This sheet holds Raw Materials, Recipes, Settings, Inventory, Orders and Recipe Library.

## Data (v2: live Google Sheet sync)
- No sync configured: data stays on the phone (localStorage), the same as v1.
- With sync on (Settings → Google Sheet sync): the Google Sheet is the source of truth. All tabs are read fresh on every app load and when ↻ is tapped. Every edit is written back cell by cell, and a write queue keeps edits made offline.
- There are two transports: **Apps Script web app** (`apps-script/Code.gs`, recommended) or **Sign in with Google** (direct Sheets API, needs an OAuth client ID). Setup: `apps-script/SETUP.md`.
- Code: `js/sheet-model.js` (column mapping by header name plus write planning), `js/transports.js`, `js/sync.js`, and `js/store.js` (the only API the UI uses).
- Tests: `npm test` runs the calculator test plus the sync test against a local mock of both back ends (`tests/mock-sheets.mjs`), which loads a read-only snapshot of the real workbook.

## Recipe library
Adapted (not verbatim) versions of sources listed in the workbook, credited to PastrybySaloni where applicable. Estimated raw-material prices are marked and must be confirmed.
