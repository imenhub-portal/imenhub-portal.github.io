# i-Nventori Ofis — read this first

Everything a fresh session needs is in this file. It exists because the useful context —
why this is not a Laravel app, what deploys how, where the secret lives, and which
"obvious improvements" are actually decisions already taken — is otherwise scattered across
a long chat history on one machine, or sitting in files that **do not sync** to a fresh
`git pull` (see [Sync model](#sync-model)).

Keep it updated. If you change behaviour worth a future session knowing, edit this file in
the same commit — and when you retire a feature, **delete the paragraph describing it**
rather than adding a newer one beside it. This document was once patched incrementally
until it claimed both that `is_portable` had been added and that it had been removed.

---

## 1. What this is

A Google Apps Script backend + GitHub Pages frontend + Google Sheets database for tracking
IMEN's office supplies and portable equipment. **Two categories, and only two:**

| Category | Enum key | Nature |
|---|---|---|
| **Alat Tulis** | `consumable` | Kertas, pen, dakwat. A quantity of interchangeable units. Runs out, gets topped up. **Issued, never returned.** No asset tag. |
| **Aset Alih** | `fixed_asset` | Laptop, projektor, mikrofon, penunjuk laser. One physical thing each. **Borrowed and returned.** Carries an asset tag. |

It answers two questions and only two: **where is each asset**, and **how much stock is
left**. It is deliberately **not** an accounting system — there is no cost, no
depreciation, no book value, no warranty tracking anywhere in it.

Two audiences, one HTML file:

- **Pemohon (public, no password)** — browses availability and submits requests.
- **Admin (one shared password)** — approves requests and manages everything.

### Pieces

- **Backend**: `Code.gs` — Apps Script bound to a Google Sheet. Deployed as a web app.
- **Frontend**: `index.html` — one self-contained file: hand-written CSS, vanilla JS, no
  framework, no build step. Served two ways at once, same as i-nstrumen: GitHub Pages, and
  Apps Script's own `doGet` which fetches this same file live from
  `raw.githubusercontent.com` on every request (single source of truth), falling back to a
  stub if GitHub is unreachable.
- **Database**: the Google Sheet, read and written only through `Code.gs`.

Part of the `imenhub` monorepo (remote: `imenhub-portal/imenhub-portal.github.io`, a
**public** GitHub Pages repo). **Because the repo is public, no real secret value may ever
appear in a committed file** — see [Secrets](#secrets).

---

## 2. Why this is not Laravel

The original specification asked for Laravel 11 + MySQL. That cannot run here: GitHub Pages
serves static files only and has no PHP runtime, and the whole monorepo is built on the
Apps Script + Sheets pattern. **This was raised with the owner, who chose to build on the
imenhub stack rather than change hosting.** The Laravel names were kept where they map, so
the code stays reviewable against the original spec:

| Spec (Laravel) | Here (`Code.gs`) |
|---|---|
| Migrations | `ensureSheets_()` — idempotent, with header auto-heal |
| Eloquent model + `$casts` | `_readTable_()` |
| Query scopes | the `SCOPES` object |
| `DB::transaction()` | `_txn_()` (LockService) |
| FormRequest validation | `_validate_(schema, payload)` |
| Routes + `auth` middleware | `doPost` switch + `ACTIONS_ADMIN` |
| `InventoryService` | the `svc*` functions (16 of them) |
| Scheduler | `checkDates()` + `installTriggers()` |
| Blade components | `itemsTable()` / `pill()` / `openDrawer()` |
| SoftDeletes | the `deleted_at` column, filtered on every read |

---

## 3. Live deployment

Already wired in — do not re-enter these:

| Constant | Value | Where |
|---|---|---|
| `SPREADSHEET_ID` | `1Xk1aKMmWR3AFTvWlMJDKUmv-5Ik_GSvyZeqXX30j_6E` | `Code.gs` |
| `FOLDER_ID` | `1NQUV2EHXpQhlWzcCm2aaIEOUvZ4CGOE7` | `Code.gs` |
| `ADMIN_EMAIL` | `imenmakmal@gmail.com` | `Code.gs` — **fallback only**, see §6 |
| `API_URL` (`/exec`) | deployment `AKfycbyY2fSJbt6…` | `index.html`, in the shim |

**Two different URLs, endlessly confused:**

- `https://script.google.com/macros/s/AKfycbyY2fSJbt6…/exec` — the **API**, from the Apps
  Script deployment. The page calls this for data.
- `https://imenhub-portal.github.io/i-Nventoriofis/` — the **link people open**. This is
  what gets shared.

### Naming and the URL rename

The system is **i-Nventori Ofis**. It lived at `/i-Nventori/` until the owner renamed it;
the folder is now `i-Nventoriofis/`. `i-Nventori/index.html` is kept as a **redirect stub**
so links shared before the rename still land somewhere sensible instead of a bare 404. It
can be deleted once nobody uses the old address.

### <a name="secrets"></a>Secrets

| Property | Used by | Effect if missing |
|---|---|---|
| `ADMIN_PASSWORD` | `_adminPassword_()` | **Fails closed** — nobody can log in and every gated action is denied. Deliberate: a missing property must never mean "no password required". The login names this specific cause rather than saying "wrong password". |

Current value is `hazde`, set by the owner in **Project Settings → Script Properties**.

---

## 4. <a name="sync-model"></a>Sync model — read before assuming a file is "in the repo"

Run `git ls-files i-Nventoriofis/` to see exactly what syncs. As of this writing:

```
i-Nventoriofis/CLAUDE.md          (this file)
i-Nventoriofis/Code.gs
i-Nventoriofis/appsscript.json
i-Nventoriofis/index.html
i-Nventoriofis/tests/test_backend.js
i-Nventoriofis/tests/check_html.js
```

`.claude/` is **excluded from git** via `.git/info/exclude` at the repo root, along with
`_local/` and `_backend/`. Note that last rule: **do not put anything you want to keep in a
folder named `_backend/`** — that is why the tests live in `tests/`.

To recreate the local dev-server config on a fresh machine (`.claude/launch.json`):

```json
{
  "version": "0.0.1",
  "configurations": [
    {
      "name": "i-nventori-static",
      "runtimeExecutable": "npx",
      "runtimeArgs": ["--yes", "http-server", "-p", "8125", "-c-1", "."],
      "port": 8125
    }
  ]
}
```

**Why `Code.gs` is tracked**: committing it is how work moves between machines — `git pull`
on another PC fetches the latest, you edit it locally, then **paste it into the Apps Script
editor** to deploy. `git push` does *not* deploy the backend. This is only safe because the
file contains no literal secrets.

---

## 5. Deploying a change

Frontend and backend deploy completely differently. Mixing them up wastes an hour, and has.

1. **`index.html`** — commit + push. GitHub Pages publishes it automatically, and Apps
   Script's `doGet` picks it up on the next request. **No redeploy needed.**
2. **`Code.gs`** — push does **nothing** to the live backend. Paste the file into the Apps
   Script editor, then **Deploy → Manage deployments → ✏️ Edit existing deployment →
   Version: New version → Deploy**. ⚠️ **Never "New deployment"** — that mints a *new*
   `/exec` URL and `API_URL` in `index.html` must keep matching it.

New sheet tabs and columns appear by themselves on the next request; `ensureSheets_()`
creates and header-heals but **never deletes**.

**The most common support question** is "the page says it cannot load the list". That is
almost always `index.html` (published instantly) being newer than the deployed `Code.gs`.
The page detects `Unknown API function` specifically and says so, naming the redeploy
steps — do not replace that with a generic "please reload".

---

## 6. Architecture

### HTTP entry points

**`doGet(e)`** — `?format=json` returns the raw payload; otherwise fetches and serves
`index.html` from GitHub.

**`doPost(e)`** — the JSON API. Dispatch is an explicit `switch`; **eight** entry points are
reachable and nothing else:

| Function | Auth |
|---|---|
| `getInitialData(pass)` | admin payload; anonymous callers get a stripped version |
| `getPublicCatalog()` | **none** — see the privacy note below |
| `getReturnContext(token)` | **token** |
| `startReturnUpload(token, …)` | **token** |
| `handleAction(action, payload)` | per-action, see below |
| `adminLogin(pass)` | none (it *is* the login) |
| `getItemHistory(id, pass)` | admin |
| `startResumableUpload(…, pass)` | admin |

**`handleAction(actionType, payload)`** is the single dispatch point for writes.
Authorization is enforced *before* the switch, so no individual handler can forget it.

- `ACTIONS_PUBLIC` — `Ping`, `SubmitRequest`, `SubmitReturnProof`. These create pending
  rows only; nothing moves without an admin.
- `ACTIONS_ADMIN` — everything else, checked against `_adminPassword_()` server-side on
  every call. A client-side boolean is never trusted, because `doPost` is callable
  anonymously from anywhere.

Every write is wrapped in `_txn_` (LockService) and appends exactly one `Transactions` row.

### The sheet

| Tab | Purpose |
|---|---|
| `Items` | Both categories, distinguished by `item_type` |
| `Transactions` | **Immutable ledger** |
| `Requests` | One row per requested **line**, grouped by `group_id` |
| `Custodians` | Auto-built directory of anyone who borrowed or received something |
| `Locations` | Storage places |
| `Config` | Key/value settings (the Pentadbir Inventori) |
| `Categories` | **Legacy, unused.** Kept because `ensureSheets_()` never deletes |

Retired columns that remain in the sheet, harmless and no longer read: `unit_cost`,
`salvage_value`, `useful_life_years`, `date_warranty_expiry`, `is_portable`, `category_id`,
`requester_department`.

### The Pentadbir Inventori

The officer who approves requests and receives every notification. Held in `Config`
(`admin_name`, `admin_email`, `admin_cc`) so changing the officer needs **no redeploy** —
it is edited in Tetapan. `_adminEmail_()` / `_adminName_()` / `_adminCc_()` resolve from
there, falling back to the `ADMIN_EMAIL` constant so mail never silently goes nowhere on a
sheet that predates this.

Note this is **unrelated** to `Custodians`. Custodians are the people who *borrow and
receive*; they approve nothing.

### Admin UI

Nav: **Permohonan · Alat Tulis · Aset Alih · Ringkasan · Laporan · Tetapan**, mirrored in
the icon rail. Login lands on **Permohonan** — pending requests are what need acting on.

Alat Tulis and Aset Alih are the two top-level sections, and that is the organising idea.
Everything applying to only one lives inside it as a sub-tab:

- **Alat Tulis** → Senarai | Stok | Log
- **Aset Alih** → Senarai | Pinjaman | Log

`viewSection()` owns the shell; `viewItems` / `viewStock` / `viewLoans` / `viewLedger` each
take an `embed` flag so they render without their own page header when a section supplies
one. Tetapan holds **Lokasi · Pentadbir Inventori**.

---

## 7. Decisions already taken — do not "fix" these without asking

Each of these looks like an oversight and is not.

### Model and terminology

- **There is no money in this app.** No unit cost, depreciation, book value or warranty.
  Removed at the owner's request: it is a location/stock tracker, not an asset register for
  accounting. Do not reintroduce a "total value" figure — for stationery it was meaningless
  anyway.
- **There are exactly two categories, and they *are* the item type.** The user-managed
  `Categories` table and the `is_portable` flag were both removed: a second classification
  had nothing left to hold, and every asset is movable by definition now. `catName()`
  survives only so the CSV keeps exporting whatever legacy rows already carry.
- **Only Aset Alih carries an `asset_tag`.** A tag names one physical thing you can stick a
  label on. Alat Tulis is a quantity of interchangeable units, so `svcAddItem` leaves its
  tag blank and the UI shows an "Alat Tulis" pill in that column instead. Alat Tulis
  therefore does not consume numbers in the `AST-YYYY-NNNN` sequence.
- **Asset tags widen past 9999** (`AST-2026-10000`) rather than wrapping, which would
  collide with `AST-2026-0001`. Soft-deleted items keep their tag reserved.

- **Padam and Lupuskan are two different things, worded apart on purpose.**
  *Lupuskan* (`svcDecommission`) records the disposal of something that genuinely
  existed — it stays in the register with a ledger row. *Padam* (`svcDeleteItem`) removes
  an item that should never have been registered: a duplicate or a typo.
  Padam is a **hard** delete: the item row and all its ledger rows go. That is deliberate
  — see the ledger note above. Guards, in order: an item on **open loan** is refused
  outright (somebody is holding it); and if the item has any movement beyond its
  registration `stock_add`, the first attempt is **refused** with the count, and only a
  second call carrying `confirm_history: true` proceeds. The UI backs that with a
  checkbox the admin must tick, and offers Lupuskan as the alternative. Because nothing
  remains, a purged asset tag **is** reissued — unlike the soft delete this replaced.
- **`deleted_at` survives with nothing writing it.** `_readTable_` still honours it, so an
  admin can hide a row by typing a date into the Sheet by hand. Keep the filter.

### The ledger

- **A ledger row is never EDITED. It is appended, or — in exactly one case — purged.**
  `_update_()` **throws** if handed the `Transactions` tab, so nothing can rewrite history
  in place. A check-in appends a *new* row; it never touches the original `check_out`.
  Open loans are **derived** by replaying the ledger (`_openLoans_`) rather than stored, so
  the two cannot drift.
  The single exception is `_purgeLedgerRows_`, called only by `svcDeleteItem`: a
  mis-entered item's rows describe stock that never arrived, and keeping them does not
  preserve history, it fabricates it. Rows are deleted bottom-up, because deleting a row
  shifts everything below it up by one. A test asserts that every recorded ledger mutation
  is a `deleteRow` and that no `setValue` ever happens.
- **`photo_borrower` and `photo_admin` are separate columns**, not one shared photo field,
  so whose evidence a `check_in` row holds is never ambiguous.

- **Restocking has a button on the Alat Tulis row itself, not only in the ⋮ menu.**
  Topping up is the single most common thing done to a consumable, and burying the one
  action the list exists for behind a three-dot menu made it read as missing. Aset Alih
  gets no such button — it is borrowed, not topped up.

- **An Alat Tulis cannot be borrowed.** `svcCheckOut` refuses `item_type === 'consumable'`
  outright and the row menu no longer offers Daftar Keluar for one. A consumable goes out
  and does not come back, so a loan against it could never be closed by a return and would
  sit open for ever; Keluar Stok is the correct route. The server enforces this rather than
  trusting the hidden button — the UI is not a guard.

- **The `maintenance` status renders as "Perhatian" for Alat Tulis and
  "Penyelenggaraan" for Aset Alih.** The status itself is set purely by quantity reaching
  zero, and that reads completely differently per category: a projector really can be away
  for servicing, but a pen at zero has simply run out. `pill(status, type)` takes an
  optional type, and `statusLabel()` gives the CSV the same text so the export never
  disagrees with the screen. The Laporan status breakdown counts both categories under one
  heading, so it passes no type and keeps the neutral label.

- **The duplicate finder matches on a normalised name** — trimmed, lowercased, inner
  whitespace collapsed — so "Pen  Biru" and "pen biru" group together. Deliberately **not**
  fuzzy: a near-match that silently grouped two genuinely different items would be worse
  than missing one. It labels each row Selamat / Ada rekod / Dipinjam so the admin can see
  at a glance which copy is safe to remove, and offers no delete button at all on one that
  is out on loan.

### Stock movements

- **A stock issue records `custodian_id`, and that is the whole point of it.** Before this,
  "who took the pens" could only be free text — unsearchable and impossible to total. The
  field is *optional* because stock also leaves for damage or loss; those units count
  toward total issuance but are excluded from per-staff figures.
- **A topup records when it arrived and where it came from; a withdrawal records neither.**
  `svcStockChange` takes `received_date` and `source` on `stock_add` only and **throws** if
  either appears on a `stock_remove` — an issue happens at the counter in the moment, so
  allowing them would let a withdrawal be quietly backdated or attributed to a supplier.
  Backdating a *delivery* is normal; a **future** date is rejected because the stock is not
  physically there and counting it would make the balance a lie. `source` is free text with
  a `<datalist>` of values already used, built from `D.transactions` — no Suppliers table
  to maintain.
- **Quantity is not editable on the item edit form.** Stock only moves through ledgered
  actions, so the item row can never disagree with the ledger.
- **A fixed asset's open-loan check runs *before* its stock check** in `svcCheckOut`. An
  asset on loan also has `quantity_available: 0`, and "stok tidak mencukupi" is a true but
  useless answer to "why can't I borrow this?".

### Requests

- **A request carries many items.** `Requests` holds one row per line with a shared
  `group_id`, rather than a header/detail pair of tabs — that keeps `svcDecideRequest`
  working per row, makes partial approval natural, and adds no join.
- **`svcSubmitRequest` validates every line before writing anything**, so "2 rim kertas and
  4 pen" where the pens are short is rejected **whole** rather than half-recorded. It sends
  **one** acknowledgement listing all lines: four pens must not mean four emails.
- **`svcDecideGroup` delegates to `svcDecideRequest`**, which calls the real `svcCheckOut` /
  `svcStockChange`. There is deliberately no second handover implementation to drift.
- **Availability is re-checked server-side at submit *and* at approve**, because the
  browser's catalog can be minutes stale.
- **The public page is two big buttons, then a form.** Choosing Alat Tulis or Aset Alih
  first lets each form be shaped for its kind — quantities and stock for stationery,
  availability and return dates for equipment. Availability is baked into each dropdown
  option so a requester never looks it up elsewhere. **There is no Jabatan field.**

### Privacy

- **`getPublicCatalog` is a deliberately separate, narrower payload — do not "simplify" it
  by reusing `getInitialData`.** It is reachable without any password, so it carries
  availability and location but **never** the name or email of whoever holds an asset,
  never the ledger, never the custodian directory. An unavailable asset reports its
  expected return date and nothing more. Tests assert the serialised payload contains no
  `@` and no custodian reference. The same assertion covers `getReturnContext`.
- **`config` is blanked for anonymous callers** — the officer's address is contact detail,
  not public data.

### Return proof

- **The borrower's link carries a capability token, not a password.**
  `startResumableUpload` requires admin; for a borrower to upload, that gate must open, and
  an unauthenticated upload endpoint lets anyone with the URL fill the Drive folder. So the
  check-out email carries `?pulang=<token>` granting exactly three things scoped to one
  loan: read that item's public details, one Drive upload, attach a photo URL.
  `_itemByReturnToken_` is the single place that rule lives. The token is minted at
  check-out, stored on the `Items` row, and **cleared at check-in** — closing the loan is
  what expires the link.
- **Submitting proof deliberately does not close the loan.** The admin still confirms
  physical receipt, and the officer's email says so explicitly.
- **`capture="environment"` on the file input *is* the camera feature.** One attribute makes
  a phone open the camera instead of a file picker — no getUserMedia, no extra library.
  `dropZone(key, camera, label)` takes it as a flag; `doUpload()` routes through
  `startReturnUpload` when `RETURN_TOKEN` is set and `startResumableUpload` otherwise.

### Frontend mechanics

- **Instant load is a contract, not an accident.** On a repeat visit `boot()` paints
  straight from the `localStorage` cache — measured at ~5 ms, no spinner, no "stale data"
  banner — then refreshes in the background and **skips the repaint entirely if nothing
  changed**, so a background fetch can never flicker the view someone is reading. A failed
  background refresh is silent by design. Do not add a loading state to the cached path.
- **The `<head>` prefetch is tagged with the password it used (`__earlyPass`), and `boot()`
  only consumes it when that matches the current session.** This is a shipped bug that got
  fixed: on a first visit the prefetch runs anonymously, so the server correctly answers
  `is_admin:false`; `boot()` then read that stale payload *after* a successful login,
  concluded the session had expired, and bounced the user back to the login screen. On a
  public visit the prefetch is skipped entirely. `__earlyPending` distinguishes "in flight"
  from "never issued", so a prefetch that failed to start cannot leave `boot()` waiting
  forever.
- **`_readTable_` keys blank-row detection off `id`, falling back to the first column.**
  `Config` is key/value with no `id`; without that fallback every row in such a tab is
  silently discarded — which is exactly what happened when `Config` was added, making the
  officer setting appear to save and then vanish.
- **Every colour is a CSS custom property.** That is what makes the dark theme a pure token
  swap under `[data-theme="dark"]`. A hard-coded hex silently breaks dark mode for that
  element. The theme is applied by a tiny inline script in `<head>` **before first paint**,
  so dark-mode users never see a white flash.
- **Icons need `viewBox="0 0 24 24"`.** The sprite is drawn on a 24×24 grid; without a
  viewBox the browser maps one unit to one pixel and clips every icon below 24px to its
  top-left corner. This shipped once. `tests/check_html.js` now fails the build if any
  sprite `<svg>` lacks it.
- **No Tailwind.** i-nstrumen flags its CDN JIT as a known anti-pattern, and it conflicts
  with the single-self-contained-file, paste-to-deploy workflow.
- **CDN libraries are lazy-loaded on first use, never at boot** (`qrcodejs`,
  `html5-qrcode`). `tests/check_html.js` fails the build if a `<script src>` appears in the
  document.

### Repo hygiene

- This is a shared monorepo with independent sessions working on other apps. Always
  `git pull --rebase` before pushing, and keep commits scoped to `i-Nventoriofis/`.

---

## 8. How to verify a change before considering it done

There is no local Apps Script emulator, so verification uses three techniques.

### 1. Backend unit tests — run before any deploy

```bash
node tests/test_backend.js
```

**277 assertions.** Loads `Code.gs` into a `vm` context with hand-written mocks for
`SpreadsheetApp`, `LockService`, `PropertiesService`, `CacheService`, `MailApp`,
`UrlFetchApp` and `ScriptApp`, backed by an in-memory spreadsheet, then drives the **real**
service functions. It covers asset-tag generation, stock round-trips, `SCOPES` boundaries
(29/30/31 days, 364/365/366 days), authorization including fail-closed with no password,
header auto-heal, ledger immutability, the grouped request lifecycle, the Config fallback,
and the return-token lifecycle.

Two harness quirks worth knowing:

- Values created inside the vm come from **another realm**, so `x instanceof Date` is false
  in the test file. Use the `isDate()` helper.
- Dates the server stores are **local-day**; `iso()` renders UTC. Use `fmtISO()` when
  comparing a stored date against an expected calendar day, or the timezone offset shifts
  the expectation across midnight.

### 2. Frontend static checks

```bash
node tests/check_html.js
```

**21 checks.** Parses every inline `<script>`, verifies tag balance (the "page won't load"
bug class that has bitten i-print), that every `on*=` handler resolves to a defined
function, that every icon reference exists in the sprite **and declares a viewBox**, and
that no CDN script loads at boot.

Two things to watch:

- The tag-balance check counts literal `<div` text **in comments too** — never write
  example markup in a comment.
- It counts both branches of a conditional return, so a function whose early-return path
  closes a card the main path also closes will read as unbalanced. Build the shared markup
  into a variable and have **each** return path emit its own complete element. `viewLedger`,
  `viewRequests` and `renderRequestForm` all had to be restructured this way.

### 3. Live browser testing

Start the static server (`.claude/launch.json` above) and drive the UI with the Browser
preview tools. There is no local backend, so the productive technique is to **replace
`window.google` with a mock** exposing `getInitialData` / `getPublicCatalog` /
`getReturnContext` / `handleAction` / `adminLogin`, feed it a synthetic payload, then call
`boot()` or `showPublic()`. That exercises every render path for real.

The shim defines `google.script.run` as a **non-configurable getter**, so replace the whole
`window.google` object rather than trying to redefine the property.

⚠️ **When the Browser pane is not displayed the page does not composite frames, so CSS
transitions never advance** and `getComputedStyle` returns the transition's *first* frame.
Anything with a `transition` — the active nav pill's background, the drawer slide, the
mobile rail — will measure as its starting value and look broken when it is not. Inject
`*{transition:none !important}` before measuring animated properties.

### 4. Live probe after deploying the backend

```bash
curl -s -L -H 'Content-Type: text/plain;charset=utf-8' \
  -d '{"fn":"getPublicCatalog","args":[]}' \
  'https://script.google.com/macros/s/AKfycbyY2fSJbt6FMYtH8fIun8D_O4JWbhBZlN9hfPG-QVLs0T6m0OePE5_WxVr7XSRawaGE/exec'
```

Re-assert the payload contains no `@`. This is the cheapest way to confirm a deploy landed
and that the privacy boundary survived it. `{"fn":"getReturnContext","args":["bad-token"]}`
should error rather than leak anything.

Always syntax-check both files before considering an edit done — `Code.gs` with plain
`vm.Script`, and each inline `<script>` block of `index.html` (which is what
`tests/check_html.js` does).

---

## 9. What has not been proven

Honest gaps, so nobody assumes otherwise:

- **Email delivery** is tested against a mock, never against real Gmail. Structure,
  recipients, escaping and failure handling are all verified; actual delivery is not. Mail
  sending is best-effort and wrapped in `try/catch` everywhere — a quota failure must never
  roll back a handover that physically happened, and every action returns
  `mailed: true|false` so the UI reports honestly instead of claiming an email that never
  left.
- **The phone camera** is verified only as far as the `capture="environment"` attribute
  being present. Whether a given handset opens the camera is untested.
- **Drive uploads** run against a mocked `UrlFetchApp`. The resumable-session logic is
  exercised; a real file has never been PUT through it from this test suite.
