# Asset Stocktake PWA

Offline-first stocktake app. Scans 1D barcodes, classifies each scan instantly against a locally cached register, persists every scan to IndexedDB at capture, and syncs batches to a Vercel function that owns all Excel Online writes. "Missing" is derived only at session close, behind a preview and explicit confirmation.

The phone never authenticates to Microsoft and never holds a credential.

## How access to Excel Online works

```
mobile browser  ──▶  /api/excel  (Vercel Function)  ──▶  Microsoft Graph  ──▶  stock.xlsx
   no secrets            credentials live here                                Excel Online
```

This is the only shape that works for a browser-launched scanner. Anything a Vite client can read (`VITE_*`) is inlined into public JavaScript, so a flow URL or API key placed there is readable by anyone who opens the site. And a sign-in on the phone would break the core promise — a token that needs refreshing over the network can't gate a scanner in airplane mode.

So the app scans entirely against its local IndexedDB copy and only contacts `/api/excel` when there is connectivity. Devices are gated by a short access code, typed once on the Start screen and stored locally; it never affects offline scanning.

Auth is **delegated** — the app acts as you. You sign in once from a laptop, and the server keeps the resulting token and reuses it. Nothing is granted that you couldn't already open yourself, which is why **no admin consent is required**. (App-only auth would give the app its own tenant-wide identity, and that does need an administrator.)

Two consequences worth knowing up front: every write appears in the workbook's version history under **your** name, and the connection stops working if your account is disabled. Re-connecting is one visit to `/api/auth/start`.

### Entra app registration

1. Entra portal → App registrations → New registration.
   - Supported account types: **accounts in this organizational directory only**.
   - Redirect URI: **Web** → `https://<your-app>.vercel.app/api/auth/callback`
2. Certificates & secrets → New client secret → copy the **Value** (shown once). Note its expiry — the app reports a distinct "sign in again" error when it lapses, but a calendar reminder is cheaper.
3. API permissions → Microsoft Graph → **Delegated** permissions → add **`Files.ReadWrite.All`** and **`offline_access`**.

Delegated `Files.ReadWrite.All` reads and writes only what your own account can already reach, across both OneDrive for Business and any SharePoint library you have access to. Despite the name it does **not** require admin consent — unlike the *application* permission of the same name. `offline_access` is what makes the refresh token possible; without it the connection would die within the hour.

If your tenant blocks self-service app registration, this is the one thing you'll need IT for — and it's a far smaller ask than tenant-wide application permissions.

### Finding the ids

Run these in [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) signed in as yourself:

```http
# Workbook in your own OneDrive for Business — leave EXCEL_DRIVE_ID blank
GET /me/drive/root/search(q='stock.xlsx')          → EXCEL_ITEM_ID

# Workbook in a SharePoint / Teams library
GET /sites/{hostname}:/sites/{site-name}           → site id
GET /sites/{site-id}/drives                        → EXCEL_DRIVE_ID
GET /drives/{drive-id}/root/search(q='stock.xlsx') → EXCEL_ITEM_ID
```

### Connecting

After the env vars are set and deployed, visit once from a laptop:

```
https://<your-app>.vercel.app/api/auth/start?code=<STOCKTAKE_ACCESS_CODE>
```

Sign in, accept the consent prompt, and you should see "Excel connected". The refresh token is stored in KV and rotated automatically on every use — which is why a KV store is needed rather than an env var. Vercel env vars aren't writable at runtime.

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

**Table `ScanLog`** (sheet "ScanLog") — append-only; the server only ever adds rows.

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

### Tag matching

A scan is matched against `Barcode` → `Asset Id` → `Serial Number` → `Old Asset Id`, in that order, so it resolves whichever value the physical label actually encodes and tolerates a partly-populated `Barcode` column. Aliases are resolved to a single `Asset Id`, so the same asset read twice by different columns is still one sighting.

Both sides are normalised the same way — whitespace removed, leading `#` and `'` stripped, uppercased — so the register's `#9339100142` matches a scan of `9339100142` and vice versa.

## Setup

```bash
npm install
cp .env.example .env
npm run dev             # local dev (camera works on localhost)
npm run build           # production build in dist/
```

`.env` holds client tuning knobs only (`VITE_SYNC_INTERVAL_MS`, `VITE_SYNC_BATCH_SIZE`, `VITE_SCAN_COOLDOWN_MS`). Every credential is a server-side variable set in the Vercel dashboard — see `.env.example` for the full list.

To exercise `/api/excel` locally you need `vercel dev` (plain `vite` does not run the functions).

## Deploying

1. Push to a **private** GitHub repo. `.gitignore` already excludes `node_modules/`, `dist/`, and every `.env` except the example.
2. Import the repo in Vercel — the Vite preset is detected automatically, and `api/*.ts` becomes Node functions with no extra config.
3. Add **Upstash Redis** from the Vercel Marketplace (Storage tab). It injects `KV_REST_API_URL` and `KV_REST_API_TOKEN` automatically — that's where the refresh token lives.
4. Set the remaining server env vars from `.env.example` for **Production and Preview** both: `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `EXCEL_ITEM_ID`, `STOCKTAKE_ACCESS_CODE`, plus `EXCEL_DRIVE_ID` / `MS_TENANT_ID` if you need them.
5. Register `https://<your-app>.vercel.app/api/auth/callback` as a redirect URI on the Entra app, then visit `/api/auth/start?code=…` once from a laptop to connect.
6. Vercel serves HTTPS by default — mandatory, since both the camera and the service worker require it.

`vercel.json` keeps `sw.js` and the manifest revalidating so `registerType: 'autoUpdate'` actually picks up new deploys; hashed assets keep Vercel's immutable caching.

First run on a phone: open the URL, enter the access code, tap "Download register" once while connected, then Add to Home Screen. After that the app launches and scans with zero connectivity.

## Idempotency

Every scan carries a client-generated `Scan ID`. `submitScans` reads the existing IDs before appending, so a lost response costs nothing — the identical batch is re-sent and every already-logged row is skipped. The client marks a batch synced only on a 200.

`closeSession` writes register cells by range, keyed on `Asset Id`, which makes a replay a no-op rewrite of the same values. If the close fails partway, the session stays in `closing`, the exact payload stays on disk, and the sync loop replays it — automatically, including after a reboot.

## Test checklist

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
- [ ] Reading one asset by barcode and then by serial counts as one sighting, second read ALREADY SCANNED.

**Durable capture**
- [ ] Scan 5 tags in airplane mode, force-quit the browser from the app switcher, relaunch → resume prompt shows all 5.
- [ ] Kill the tab within ~1s of a scan flash → at most that one scan is absent after relaunch.
- [ ] Let the phone die mid-session, recharge, relaunch → session resumes with counters intact.

**Resumable sessions**
- [ ] Close the tab, reopen hours later → "Open session found" with correct scope, scanner name, and scan count.
- [ ] Verify no timeout ever closes a session (leave one open overnight).

**Batched, idempotent sync**
- [ ] Scan 10 tags offline; go online → header counts down to "Synced"; ScanLog has exactly 10 new rows, in capture order.
- [ ] In DevTools, block `/api/excel` after the request is sent (simulate a lost response) → client retries; ScanLog has no duplicate Scan IDs.
- [ ] Change `VITE_SYNC_INTERVAL_MS` and confirm the cadence changes without code edits.

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

**Security**
- [ ] `grep -rniE "client_secret|MS_CLIENT|MS_TENANT|EXCEL_DRIVE" dist/` returns nothing.
- [ ] A request to `/api/excel` without a valid `x-stocktake-code` header returns 401.

**Installability**
- [ ] Install to home screen, enable airplane mode, launch from the icon → app opens and scans.
