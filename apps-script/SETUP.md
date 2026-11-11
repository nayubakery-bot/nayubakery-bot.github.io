# Turn on live Google Sheet sync (about 5 minutes)

You need the Google account that owns the **Bakery Workbook**. You do this once.

## Option A (recommended): Apps Script web app. No Google Cloud needed.

1. Open the Bakery Workbook:
   https://docs.google.com/spreadsheets/d/1LgfA7HGWy-KA5PcA6qQ7d1omzuJwiWKXIgpjEIcvyQI/edit
2. Go to **Extensions → Apps Script**. A new tab opens with a file called `Code.gs`.
3. Delete everything in `Code.gs`. Paste in the full script. To get it, open the app and go to **More → Settings & export → Google Sheet sync → "Script to paste (Code.gs)"**, or add `apps-script/Code.gs.txt` to the end of the app's address. Select all and copy it. Then click **Save** (the disk icon).
4. Set your own sync token:
   **Project Settings** (gear icon, left) → **Script properties** → **Add script property**
   - Property: `APP_TOKEN`
   - Value: a code only you know, for example `rw-7391` (if you skip this, the token is `2468`)
   → **Save script properties**.
5. Deploy the script:
   **Deploy → New deployment →** gear icon next to "Select type" → **Web app**
   - Description: `Rise & Whisk sync`
   - Execute as: **Me**
   - Who has access: **Anyone**
   → **Deploy** → **Authorize access**, pick your Google account. If you see "Google hasn't verified this app", click **Advanced → Go to (project name) (unsafe)** (it's your own script), then **Allow**.
6. Copy the **Web app URL** (it ends in `/exec`).
7. In the app go to **More → Settings & export → Google Sheet sync**:
   - "How should the app talk to the sheet?" → **Apps Script web app**
   - Paste the URL
   - Sync token: the `APP_TOKEN` from step 4
   → **Save & connect**. The badge at the top should change to **Synced**.

Do steps 6 and 7 on each phone that uses the app.

**If you change Code.gs later**, go to **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. That keeps the same URL.
**To change the token**, edit `APP_TOKEN` in Script properties. Then type the new token into the app on each phone.

## Option B: Sign in with Google (direct Sheets API)

This option needs a Google Cloud OAuth client ID:
1. Go to https://console.cloud.google.com. Create a project, then **APIs & Services → Library → Google Sheets API → Enable**.
2. **OAuth consent screen**: choose External, fill in the app name and your email, and add your Gmail under **Test users**.
3. **Credentials → Create credentials → OAuth client ID → Web application**.
   Under **Authorized JavaScript origins**, add the app's address exactly as you open it, for example `https://<your-app>.ship.place`, with no trailing slash. If the app moves to a new address, add that one too.
4. Copy the client ID (`…apps.googleusercontent.com`).
5. In the app's Settings choose **Sign in with Google**, paste the client ID (the spreadsheet ID is filled in already), and tap **Sign in with Google**.

Google sign-in lasts about 1 hour. After that the badge shows **Sign in needed**; tap **Sign in with Google** again. Option A doesn't expire, which is why it is recommended.

## How it behaves

- **Sheet → app:** every time the app is opened or reloaded, it reads all six tabs fresh. The ↻ button does the same. There is no background polling. If you edit the sheet while the app is open, tap ↻ to see the change.
- **App → sheet:** each change is written within about a second, and only to the cells that changed (rows are found by ID / item name / setting label). New orders, raw materials and recipe lines are added as new rows at the bottom.
- **Formula cells are never overwritten.** Cost per unit, inventory current stock and the totals stay as formulas, and the app calculates the same numbers itself.
- **Offline:** changes wait on the phone (the badge shows *Offline · N waiting*) and are sent once the phone is back online.
- **Conflicts:** for orders, the newer edit wins (using the **Updated at** column; the script stamps it when you edit an order by hand). For prices, stock and settings, if the sheet cell changed after the app loaded it, the sheet's value is kept and the app lists the skipped change under Settings.
- **Delete** in the app sets the order's Status to `Deleted`. The row stays in the sheet.
