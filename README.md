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

## v3: edit everything from the app
All of these write back to the Google Sheet (or queue offline / stay on the phone in local mode):
- **Menu** (More → Menu & sizes): add / rename / remove sizes, layers, sponges, fillings, frostings, toppings, boxes, boards, add-ons. Rows go only into empty rows inside the ranges the Calculator tab's dropdowns/formulas use; if a list is full, "Make room" (after a confirm) inserts 5 empty rows *inside* the block so Sheets widens those ranges.
- **Selling prices** → new `Price List` tab; per-order "Your price for this order" override (saved in Orders → `Price override (₹)`).
- **Price went up** (#/price-up): shows affected recipes/cakes, then updates Raw Materials.
- **Recipes**: "+ New recipe", ✕ to clear a line, method notes (`Recipe Notes` tab).
- **Inventory**: Log a purchase (`Purchases` tab, adds stock + updates price), low-stock alert = Inventory "Reorder level".
- **Pricing rules & delivery** (`Pricing Rules`, `Delivery Zones` tabs): rush charge / festival discount toggles; flat-rate zones alongside per-km.
- **Customers** (`Customers` tab): list, quick-fill on the order form by name/phone, order history.
- **Payments**: advance / balance / paid in full, one-tap status chips (Orders → `Payment status`, `Balance (₹)`).

New tabs are created on first use with a bold frozen header; new Orders columns are appended to the right of the header row. Formula cells are never overwritten and existing cells never move (except via the opt-in "Make room"). New tabs need "Sign in with Google" mode (Apps Script mode reports them as sync notes).

## v4: product categories + manual cake sizes
- **Categories tab** (created on the first sync if missing, bold frozen header): `Category | Pricing | Unit price (₹) | Unit | Cost per unit (₹) | Active | Notes`, seeded with Cake (Calculator), Granola bar ₹90/bar, Cookies ₹90/cookie. Read on every load / ↻; edit in the sheet or in **Menu & prices → Categories** (add, rename, price, cost, show/hide, remove). Local mode / offline falls back to those three defaults.
- **New order** (`#/order/new`) starts with a category picker; each category has its own screen at `#/order/new/<Category>`. "Calculator" categories show the full cake form; "Per unit" categories show customer + quantity (price = quantity × unit price; optional home-delivery charge).
- **Cake size**: one field that is both a dropdown of preset sizes and directly editable, plus a g / kg toggle (remembered on the phone). A typed weight that matches a preset (e.g. "500", "500 g", "0.5 kg") is saved as that preset and uses its servings and selling price; any other weight like "750 g" is priced by the cost calculator, with servings estimated from the nearest preset ("~6-9"). Saved in Orders → Size as the preset name or the typed weight. The sheet's own Calculator tab is untouched.
- **Orders tab**: `Category`, `Quantity`, `Unit price (₹)` are appended to the right. Per-unit orders leave the cake-only columns blank; old rows with a blank Category read as Cake. Each order keeps the unit price it was sold at.
- Dashboard (this month by category), Profit (By category), order lists and order screens show the category. Cost for per-unit items = quantity × Cost per unit (₹), or 0 when that's blank.
