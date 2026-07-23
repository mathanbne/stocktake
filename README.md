# Asset Stocktake PWA

Offline-first stocktake app. Scans 1D barcodes, classifies each scan instantly against a locally cached register, persists every scan to IndexedDB at capture, and writes results to a workbook in Excel Online. "Missing" is derived only at session close, behind a preview and explicit confirmation.

## How it works

1. Open the app, **sign in with your work Microsoft account**
2. **Pick the stocktake workbook** — searches your OneDrive and any SharePoint site you can open
3. **Download the register** once, then scan — online or off

There is no server and no shared secret. The app is a static site that talks to Microsoft Graph directly as the signed-in user, so each scanner reaches exactly the files they already have access to, and every write is attributed to them in the workbook's version history.

```
phone  ──sign in──▶  Microsoft identity
   │
   └──Graph──▶  your workbook in Excel Online
```

Scanning never waits on the network. The register is cached in IndexedDB, every scan is written to disk before the verdict appears on screen, and the upload queue drains whenever a connection is available.

### Entra app registration

The one piece of setup. A **public client** — a client id, no secret, no admin consent.

1. [entra.microsoft.com](https://entra.microsoft.com) → App registrations → New registration
   - Supported account types: **accounts in this organizational directory only**
   - Redirect URI: **Single-page application (SPA)** → `https://<your-app>.vercel.app`
     Add `http://localhost:5173` too if you want to run it locally.
2. Copy the **Application (client) ID** → that's `VITE_MS_CLIENT_ID`
3. API permissions → Microsoft Graph → **Delegated** → add `Files.ReadWrite.All` and `User.Read`

The client id ships in public JavaScript, and that is fine — it identifies the app rather than authorising it. A public client holds no secret and proves itself with PKCE; the user's own sign-in is what grants access. Delegated `Files.ReadWrite.All` reaches only what the signed-in user can already open, and unlike the *application* permission of the same name it needs **no admin consent**.

## Excel workbook structure

The app reads the real column set in `stock.xlsx` and adds only what it must in order to record a result. Nothing existing is renamed or overwritten — in particular your own `Status` column is read-only to this app.

**Table `AssetRegister`** (sheet "Register")

| Column header | Source | Notes |
|---|---|---|
| Asset Id | existing | The key for every write. Must be unique and non-empty. |
| Old Asset Id | existing | Matched as a scan alias, for assets still wearing a pre-relabel sticker. |
| Asset Name | existing | Shown on the scan verdict. |
| Department | existing | |
| Barcode | existing | Matched first where populated. |
| Facility Group | existing | e.g. `A BLOCK` — the coarse scope. |
| Facility | existing | e.g. `01G04` — the fine scope. |
| Make / Model | existing | Shown alongside Asset Name. |
| Serial Number | existing | Matched as a scan alias. |
| Condition | existing | |
| Status | existing | **Never written by the app.** |
| Stocktake Status | **add** | The app writes `Cited` or `Missing`. |
| Last Verified Date | **add** | ISO 8601 — format the column as **Text**. |
| Last Verified By | **add** | |

The three added columns must be **adjacent**; they are written as a single range.

There is no "Expected Location" column. The app derives one as `Facility Group / Facility` (→ `A BLOCK / 01G04`) and scope-matches it as a case-insensitive substring, so a session scoped to `A BLOCK` covers the whole block and one scoped to `01G04` covers the single room — with no extra column to maintain.

**Table `ScanLog`** (sheet "ScanLog") — append-only.

| Column header | Type |
|---|---|
| Session ID | Text (UUID) |
| Scan ID | Text (UUID, unique — the idempotency key) |
| Asset Id | Text (empty for an unknown tag) |
| Scanned Value | Text (exactly what came off the scanner) |
| Result | Text (`sighted`, `sighted_wrong_location`, `unknown_tag`, `duplicate`, `manual_entry`, `missing`) |
| Scanned By | Text |
| Timestamp | Text (ISO 8601) |
| Actual Location | Text |
| Notes | Text |

### Setting the workbook up

1. Add `Stocktake Status`, `Last Verified Date`, `Last Verified By` to the right of `Status`.
2. Format `Asset Id`, `Barcode`, `Serial Number` and `Last Verified Date` as **Text** — otherwise Excel strips leading zeros from numeric barcodes and coerces dates to serial numbers.
3. Select the whole range → Insert → Table → "My table has headers" → Table Design → rename to **`AssetRegister`**. A plain range will not work: Graph's table API only sees formatted tables.
4. Rename sheet `Sheet1` → `Register`.
5. Add sheet `ScanLog` with the headers above, also formatted as a table named **`ScanLog`**.

The app checks for both tables when you pick a file, so a wrongly-shaped workbook is rejected at setup rather than mid-stocktake.

### Tag matching

A scan is matched against `Barcode` → `Asset Id` → `Serial Number` → `Old Asset Id`, in that order, so it resolves whichever value the physical label actually encodes and tolerates a partly-populated `Barcode` column. Aliases resolve to a single `Asset Id`, so the same asset read twice by different columns is still one sighting.

Both sides are normalised the same way — whitespace removed, leading `#` and `'` stripped, uppercased — so the register's `#9339100142` matches a scan of `9339100142` and vice versa.

## Setup

```bash
npm install
cp .env.example .env    # then fill in VITE_MS_CLIENT_ID
npm run dev             # camera and sign-in both work on localhost
npm run build           # production build in dist/
```

Every variable is a `VITE_*` client value. There are no secrets to manage.

## Deploying

1. Push to GitHub.
2. Import the repo in Vercel — the Vite preset is detected automatically. It's a purely static build; there are no serverless functions.
3. Set `VITE_MS_CLIENT_ID` (and optionally `VITE_MS_TENANT_ID`) for **Production and Preview**. Vite inlines these at build time, so a change needs a redeploy, not just a restart.
4. Add the deployed origin as a **SPA redirect URI** on the app registration.
5. Vercel serves HTTPS by default — mandatory, since the camera, service worker and MSAL all require it.

`vercel.json` keeps `sw.js` and the manifest revalidating so `registerType: 'autoUpdate'` picks up new deploys; hashed assets keep Vercel's immutable caching.

First run on a phone: open the URL, sign in, pick the workbook, tap "Download register", then Add to Home Screen. After that the app launches and scans with zero connectivity.

## Sign-in and offline

Tokens are cached in `localStorage`, so a scanner stays signed in between shifts. Refreshing a token needs a network connection — but so does syncing, so this never blocks scanning. Offline, the app classifies against the cached register and queues scans; when the connection returns, the queue drains.

If sign-in has genuinely lapsed, MSAL redirects to Microsoft to re-authenticate. Nothing captured is lost: scans live in IndexedDB until the workbook confirms them. A session that is already open stays reachable even when signed out, so captured work is never stranded behind a login.

## Idempotency

Every scan carries a client-generated `Scan ID`. Uploads read the existing IDs before appending, so a lost response costs nothing — the identical batch is re-sent and every already-logged row is skipped.

Closing a session writes register cells by range, keyed on `Asset Id`, which makes a replay a no-op rewrite of the same values. If a close fails partway, the session stays in `closing`, the exact payload stays on disk, and the sync loop replays it — automatically, including after a reboot.

## Test checklist

**Sign-in and workbook**
- [ ] First launch shows "Sign in with Microsoft"; after signing in, your name appears.
- [ ] Search finds `stock.xlsx`; picking it stores the choice and survives an app restart.
- [ ] Picking a workbook *without* the two tables is rejected with a clear message, not accepted.
- [ ] Sign out, then relaunch → back to the sign-in screen, with any open session still resumable.

**Instant feedback**
- [ ] With airplane mode ON, scan a known tag: verdict flash renders immediately (no spinner, no delay).
- [ ] Scan a tag whose Facility differs from the Current location field → WRONG LOCATION, amber.
- [ ] Scan a value not in the register → UNKNOWN TAG, red.
- [ ] Scan the same tag twice (waiting out the cooldown) → ALREADY SCANNED.
- [ ] Type a tag manually → MANUAL ENTRY, and the ScanLog row carries the manual note.
- [ ] Scan five tags in a row → the camera never stalls or re-initialises between scans.

**Tag matching**
- [ ] A register value of `#9339100142` matches a scan of `9339100142`, and the reverse.
- [ ] An asset with an empty `Barcode` still resolves via `Asset Id`.
- [ ] Reading one asset by barcode and then by serial counts as one sighting; the second is ALREADY SCANNED.

**Durable capture**
- [ ] Scan 5 tags in airplane mode, force-quit the browser from the app switcher, relaunch → resume prompt shows all 5.
- [ ] Kill the tab within ~1s of a scan flash → at most that one scan is absent after relaunch.
- [ ] Let the phone die mid-session, recharge, relaunch → session resumes with counters intact.

**Resumable sessions**
- [ ] Close the tab, reopen hours later → "Open session found" with correct scope, scanner name, and scan count.
- [ ] Verify no timeout ever closes a session (leave one open overnight).

**Batched, idempotent sync**
- [ ] Scan 10 tags offline; go online → header counts down to "Synced"; ScanLog has exactly 10 new rows, in capture order.
- [ ] In DevTools, block `graph.microsoft.com` after a request is sent → client retries; ScanLog has no duplicate Scan IDs.
- [ ] Change `VITE_SYNC_INTERVAL_MS`, rebuild, and confirm the cadence changes.

**Safe reconciliation**
- [ ] Close a session with unscanned in-scope assets → preview lists each with Asset Id, name, expected location; nothing is written before Confirm.
- [ ] Confirm while offline → session enters "closing"; force-quit, relaunch, go online → the close completes **without** manually tapping Resume.
- [ ] Re-open the register: every scanned asset has `Stocktake Status` = `Cited` with Last Verified fields set, every previewed asset is `Missing`, and the original `Status` column is untouched.

**Scoping**
- [ ] A session scoped to `A BLOCK` includes assets whose Facility is `01G04`.
- [ ] A session scoped to `01G04` includes only that room.

**Storage limits**
- [ ] In Chrome DevTools → Application → Storage, simulate a filled quota → the app shows the blocking "Storage full" screen; no scan silently vanishes.

**Scanning fallback**
- [ ] Android Chrome uses BarcodeDetector (ZXing chunk never executes).
- [ ] iOS Safari falls back to ZXing and still decodes Code 128/39 tags.
- [ ] Deny camera permission → manual entry still works end to end.

**Installability**
- [ ] Install to home screen, enable airplane mode, launch from the icon → app opens and scans.
- [ ] Sign-in redirect returns correctly into the installed app rather than a browser tab.
