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

> **Start with [§4](#4-known-faults-on-the-live-deployment--read-this-before-debugging).**
> One deployment-side fault is live, and two things in the code look like mistakes but are
> workarounds for it. A full session was lost re-deriving that from scratch.

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
| `ADMIN_EMAIL` | `imenmakmal@gmail.com` | `Code.gs` — **fallback only**, see §7 |
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

## 4. Known faults on the live deployment — read this before debugging

**One deployment-side fault is live right now, and two workarounds in the code exist only
because of it.** A whole session was lost re-deriving this. Do not repeat that: run the two
commands below first, and believe the result over any theory.

### The fault: this deployment answers no POST, and no ContentService response

```bash
EXEC='https://script.google.com/macros/s/AKfycbyY2fSJbt6FMYtH8fIun8D_O4JWbhBZlN9hfPG-QVLs0T6m0OePE5_WxVr7XSRawaGE/exec'

# Does POST work? Prints OK only when the reply is JSON, which is the
# only thing that matters. A broken deployment returns Google's HTML shell.
R=$(curl -s -L -m 60 -H 'Content-Type: text/plain;charset=utf-8'  -d '{"fn":"handleAction","args":["Ping",{}]}' "$EXEC")
case "$R" in '{'*) echo "POST OK: $R";; *) echo 'POST BROKEN (HTML, not JSON)';; esac

# Does the page itself still serve? Expect 200 and ~260KB.
curl -s -o /dev/null -w 'GET %{http_code} %{size_download}\n' -L "$EXEC"
```

Today those print `POST BROKEN (HTML, not JSON)` and `GET 200 267006`. That is the whole
shape of it:

| Path | Response type | Result |
|---|---|---|
| `GET /exec` | HtmlService | ✅ serves the full app |
| `GET /exec?format=json` | ContentService | ❌ Page not found |
| `POST /exec` (any content type) | ContentService | ❌ Page not found |

Both broken paths are ContentService; the working one is HtmlService, and all three live in
the same deployed file. **The code is not at fault** — `doPost` is present, parses, and is
covered by tests (§9), and the same `doPost` served live traffic earlier the same day. Do
not go looking for a truncated paste or a missing function; that theory was chased and
disproved.

**This does not need a new deployment, and the `/exec` URL does not change.** POST worked
on this exact URL earlier the same day — a full round trip of `AddItem`, `StockAdd`,
`StockRemove` and `DeleteItem` went through it — so the deployment and the URL are sound.
What is wrong is the *published version*. Republishing the existing deployment
(**Manage deployments → ✏️ Edit → Version: New version**) is the fix, and it keeps the URL
that `API_URL` already points at.

An earlier version of this file recommended creating a **new** deployment. That was wrong
and cost the owner real time: it was a guess made after a few failed attempts, not a
conclusion from evidence. Do not repeat it. If a session ever does deliberately replace the
deployment, `API_URL` in `index.html` must be updated to match the new URL in the same
commit.

Once POST answers again: delete the embed block described below and re-check with the two
commands above.

### Workaround 1 — the Pages build embeds `/exec` in a full-page iframe

The Pages build reaches the backend only over `doPost`, so with POST dead it cannot load at
all. Served from `/exec` the same file works, because Apps Script provides
`google.script.run` natively and that RPC channel never touches `doPost`.

So `index.html` sets `window.__EMBED` when `location.hostname` contains `github.io`, and at
`DOMContentLoaded` swaps the body for a full-page iframe pointing at `/exec`. The head
prefetch and `boot()` both check that flag and stand down, so nothing runs against an
endpoint that cannot answer.

- It embeds rather than **redirects** on purpose: a redirect worked, but moved the address
  bar to `script.google.com`, so the Pages URL stopped being the app's address — which is
  the point of hosting it there. `doGet` already sets `XFrameOptionsMode.ALLOWALL`, which
  is what permits framing. `allow="camera"` is forwarded so return-proof capture still
  works inside the frame.
- Guarded on hostname, **not** feature detection: inside Apps Script the file is served
  from `googleusercontent.com`, so the flag is false there and it cannot nest.
- An earlier version used `document.open()`/`write()`/`close()`; mid-parse document
  replacement left the original head in place. A flag is deterministic, that was not.
- **Delete this block once POST works.** It costs the instant-load behaviour that Pages
  hosting exists for — every open now goes through Apps Script.

### Workaround 2 — none. The empty admin screen was a real bug, and is fixed

Do not confuse the two. With the embed in place the app loaded but the admin view stayed
empty, which looked like the same fault and was not. `getInitialData` was returning `null`.

It was diagnosed with the **in-app diagnostic** — the "Jalankan diagnostik" button on the
no-data screen, which calls three server functions in increasing payload size and prints
what actually came back. It uses only functions the deployed script already has, so it runs
without a redeploy. That is the tool to reach for when the backend cannot be called from a
terminal:

```
1. adminLogin       (kecil)       -> object | 11 aksara    | 1006ms
2. getPublicCatalog (sederhana)   -> object | 24587 aksara | 6892ms | items=121
3. getInitialData   (besar)       -> NULL   | 4 aksara     | 8967ms
```

Same session, same channel, seconds apart: the small responses arrived, the large one did
not. The call *succeeded* and took nine seconds — it neither failed nor timed out — but the
value never came back. The payload was ~141KB for 121 items and 125 ledger rows.

Splitting the ledger out (`getLedger()`) was the first cut — 141KB → ~90KB — but a later
diagnostic showed `getInitialData` **still** returned `null` at ~90KB, so that alone was not
enough. The full fix (see §8, *"The initial payload is lean"*) does three things and gets
the critical path to **~30KB, ~1.1× the proven-good `getPublicCatalog`**:

- **Lean items** — `getInitialData` sends only the ~11 fields the list renders
  (`_itemForList_`); full detail (notes, serial, photos, the other dates) is fetched on
  demand by `getItem()` when a drawer opens.
- **No `Date` objects** — the whole payload is `_jsonSafe_`'d (JSON round-trip) so every date
  is an ISO string. `getPublicCatalog` delivered partly *because* it carried no Dates; the
  admin payload carried one per item. This is likely as much the cause as raw size.
- **Alerts as counts, not lists** — `inspection`/`audit` ship as `inspection_count`/
  `audit_count`; `low_stock`/`overdue` are derived on the client (`deriveAlerts`) from the
  lean items. The per-item `audit` array — one entry per never-audited item — was ~22KB, the
  single largest chunk, for data nothing lists.

**If a payload ever silently returns `null` again, suspect its size AND its Date objects.**
Verify with the in-app diagnostic after deploying: `getInitialData` should read ~30KB, not
`NULL`.

---

## 5. <a name="sync-model"></a>Sync model — read before assuming a file is "in the repo"

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

## 6. Deploying a change

Frontend and backend deploy completely differently. Mixing them up wastes an hour, and has.

1. **`index.html`** — commit + push. GitHub Pages publishes it automatically, and Apps
   Script's `doGet` picks it up on the next request. **No redeploy needed.**
2. **`Code.gs`** — push does **nothing** to the live backend. Apps Script has no
   connection to the repo at all: it never clones, pulls or watches, and runs only the copy
   in its own editor, frozen at the last published version.

   Preferred: double-click **`tools/2-DEPLOY.cmd`**. It pushes `Code.gs` via clasp,
   republishes **the existing deployment by ID**, then POSTs a `Ping` and confirms the
   server answers `pong` before reporting success — because deploying and working are not
   the same thing, and the failure this stack actually hits looks perfectly fine at the
   deploy step. Run `tools/1-SETUP-clasp.cmd` once first (enables the Apps Script API,
   `clasp login`, and asks for the Script ID → writes `.clasp.json`).

   By hand, if clasp is not set up: paste into the editor, then **Deploy → Manage
   deployments → ✏️ Edit → Version: New version → Deploy**. ⚠️ **"New deployment" mints a
   different `/exec` URL**, and `API_URL` in `index.html` must be updated to match — so use
   it only when deliberately replacing the deployment (which, per §4, is currently the
   thing that needs doing).

New sheet tabs and columns appear by themselves on the next request; `ensureSheets_()`
creates and header-heals but **never deletes**.

**"The page cannot load the list"** is usually `index.html` (published instantly) being
newer than the deployed `Code.gs`. The page detects `Unknown API function` specifically and
says so, naming the redeploy steps — do not replace that with a generic "please reload".
If the message is a **404** or the data is simply empty, that is a different problem: see
§4 before assuming a stale backend.

### Division of labour — agreed with the owner, do not change it

**The owner deploys `Code.gs` and manages deployments. Claude does not.** Pasting into the
Apps Script editor and publishing versions is deliberately a human step, and that stays.

So do **not** propose or build a GitHub Action that runs clasp on push. It was offered and
declined. It would also mean storing Google OAuth tokens as a secret in a public repo,
which is the owner's call to make and they have made it.

`tools/2-DEPLOY.cmd` exists as a convenience for the owner to run, not as automation for
Claude to trigger. What Claude does is: change `Code.gs`, run the tests, commit, push, and
then **say plainly that a deploy is needed** — never imply the backend is live when only
the repo has moved.

### The tooling, and why it is `.cmd` and not `.ps1`

`CurrentUser` execution policy here is `RemoteSigned`, which blocks right-click-run on an
unsigned `.ps1`. The `.cmd` wrappers pass `-ExecutionPolicy Bypass` and pause so the output
stays readable. The `.ps1` files are CRLF with a BOM and avoid backtick line-continuation
and `\"` sequences — both parsed wrong when first generated, and a broken script fails at
the user's double-click, where there is nobody to debug it. Parse them with
`[System.Management.Automation.Language.Parser]::ParseFile` after editing.

`.claspignore` keeps everything except `Code.gs` and `appsscript.json` out of the script
project, so the only copy of the frontend stays the one Pages serves. clasp's OAuth tokens
land in `.clasprc.json`, which is gitignored — **this repo is public**.

---

## 7. Architecture

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

## 8. Decisions already taken — do not "fix" these without asking

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

- **The log is split by inventory and never mixes.** Alat Tulis (consumable) and Aset Alih
  (fixed_asset) are two separate inventories sharing one webapp, and their audit logs must
  stay apart. `viewLedger(embed, type)` **always** narrows to the section's `item_type` (via
  the `t.item_id → itemById().item_type` join) — there is deliberately no "show everything"
  mode, and `exportLedgerCSV(type)` filters the same way and names the file per type
  (`…_alat-tulis_…` / `…_aset-alih_…`). Do not reintroduce a combined log or a sticky
  cross-section scope: an earlier `LEDGER_SCOPE='all'` global, set by "Log penuh" buttons,
  leaked both types into one list and persisted across sections — that was the bug. The
  combined overview is the **Ringkasan** dashboard, not the log. (A transaction whose item is
  not in `D.items` — only possible via a manual soft-delete, since the app hard-purges —
  falls out of both logs; denormalising `item_type` onto the row would close that, but needs
  a deploy and has not been done.)

- **Clicking an item row opens a non-modal quick-peek** (`openQuickPeek` / `rowPeek`): a
  compact movement history (quantity, who, date) beside the list, rendered **instantly from
  `D.transactions`** already in memory — no server round-trip — with a fallback to
  `getItemHistory` if the ledger has not merged yet. It reuses the `#drawer` element but
  leaves the scrim hidden, so you can click straight to the next row; `Sejarah Penuh` opens
  the full modal history. `rowPeek` ignores clicks inside `.rowmenu` so the ⋮ menu still
  works. Don't make it fetch per-click — the whole point is that the ledger is already local.

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

- **`doPost` is covered by tests, and the mock `TextOutput` is faithful.** It was not,
  for a long time: every test called `handleAction` directly, so the one function the
  entire frontend goes through was never exercised. That is the worst possible gap,
  because a missing `doPost` leaves `doGet` working — the app loads normally and only its
  API calls fail, which reads as a network fault rather than a missing function. Section 23
  asserts the round-trip, that every API name the page can call is routed, that thrown
  errors come back as a normal `{ok:false}` envelope rather than a transport failure, that
  malformed input is answered rather than crashing, and that the response declares JSON
  (the shim calls `r.json()` on it). It bails immediately with a named error if `doPost` is
  absent, rather than dying with a stack trace.

- **The initial payload is lean; item detail and the ledger are fetched on demand.** This
  is what makes `getInitialData` deliver over `google.script.run` at all (see §4). Three
  rules, none of which should be "optimised" back:
  - `getInitialData` sends items via `_itemForList_` — ~11 grid fields only. The full row
    (notes, serial, photos, lifecycle dates) comes from `getItem(id)` when the edit/history
    drawer opens. **The edit form pre-fills from that fetch, not the list** — pre-filling
    from the lean row and saving would blank the missing fields. `openItemDrawer`/
    `openHistory` therefore fetch first; `getItemsForExport` does the same for the CSV.
  - The whole payload is `_jsonSafe_`'d — **no `Date` objects cross the bridge**, only ISO
    strings. `toDate()`/`inputDate()` already parse strings (the cache path always did), so
    this changed nothing on the client.
  - Alerts travel as **counts** (`inspection_count`, `audit_count`); `low_stock` and
    `overdue` are rebuilt client-side by `deriveAlerts()` from the lean items. Keep
    `deriveAlerts` and the server counts in agreement with `SCOPES` — a test asserts the
    derived counts match the `checkDates` digest.
  - Target size: ~30KB, ~1.1× `getPublicCatalog`. If you add a field to `_itemForList_`,
    watch that ratio.

- **Reads do not re-heal the schema or re-open the spreadsheet per tab.** `getInitialData`/
  `getPublicCatalog` call `ensureSheetsOnce_()` (a 5-min `CacheService` gate) instead of
  `ensureSheets_()` on every request; writes still heal fully. `_ss_()` memoises `openById`
  for the execution. With the payload now under 100KB, the `CacheService` payload cache
  finally succeeds too, so a repeat load inside 45s skips the sheet reads entirely. Together
  these cut a ~5s load (the diagnostic's figure) down sharply.

- **Mobile: sized in `dvh`, touch targets ≥44px, safe-area insets.** The mobile touch
  offset was dominated by the iframe workaround (§4) plus `100vh` overrunning the visible
  viewport; the embed and the ≤820px rules now use `100dvh`. A `@media (pointer:coarse)`
  block gives the icon rail (the only nav on a phone) and the icon buttons 44px targets, and
  fixed elements pad for `env(safe-area-inset-*)`. The **real** fix for the offset is the
  direct-serve deployment (§4) — these make the framed interim usable.

- **The ledger is fetched separately from the initial payload.** As one response the admin
  payload reached 141KB against a real inventory, took nine seconds, and came back `null` —
  while `getPublicCatalog` (24KB) succeeded over the same channel seconds earlier and
  `adminLogin` (11 bytes) was instant. Small responses arrived, the large one did not.
  `getInitialData` now omits `transactions` and reports `tx_total` instead; `getLedger()`
  serves the ledger on its own, admin-gated, and the client merges it in and re-renders a
  moment later. That puts ~90KB on the critical path instead of 141KB, and the app is
  usable as soon as the items land — the ledger only feeds Laporan, the Log tab and a few
  counts, none of which are on screen at boot.

  No consumer needed changing: they all already read `(D.transactions || [])`, so before
  the merge they see an empty ledger, which is what they saw during loading anyway. The one
  exception is the Tetapan transaction count, which reads `tx_total` so it is correct in
  the gap. `_row` — a sheet coordinate the client never used — is now stripped from every
  table on the way out.

- **`return_token` never leaves the server, and dead columns never reach the wire.**
  `_itemForClient_()` strips seven fields from every item in the payload. Six are dead
  weight — retired financial columns (`unit_cost`, `salvage_value`, `useful_life_years`,
  `date_warranty_expiry`) plus `updated_at`/`deleted_at` — about a fifth of the payload on
  a real inventory, sent on every load. The seventh matters more: `return_token` is the
  capability that lets a borrower upload return proof **without logging in**, and it was
  being handed to every client that requested the payload despite nothing on the page
  reading it. It stays on the sheet, so borrower links keep working; it just never goes
  out. Deliberately a denylist — a new sheet column should reach the client by default,
  because the alternative is a field silently missing from the UI with nothing to explain
  why.

- **An empty response is retried once.** `getInitialData` returning nothing is not the same
  as failing, and the first request can land while the transport is still settling —
  especially when the page is framed. `refresh()` retries once, then reports. `render()`
  also no longer returns silently when `D` is null: it used to leave the boot placeholder
  ("Memuatkan data inventori…") on screen for ever, so a payload that never arrived looked
  identical to one still in flight.

- **A failed `<head>` prefetch retries; it is not a dead end.** `boot()` used to call
  `accept(null)` when the prefetch came back empty and stop there — no retry, no real error,
  just "Tidak dapat memuatkan data". But that prefetch fires from `<head>`, before the body
  parses and before the transport is necessarily ready, so its failure says nothing about
  whether a normal fetch would work. That is precisely why the public catalog (fetched
  after `DOMContentLoaded`) loaded fine while the admin view sat empty on the same
  deployment. It now retries once via `refresh(true)` before giving up, and the banner
  reports `window.__earlyErr` — the actual server message — rather than a generic
  "no answer". Verified both ways: prefetch-fails-then-succeeds recovers silently, and
  both-fail stops after exactly one retry with the real error named.

- **A failed refresh is never silent.** `refresh(true)` used to swallow its failure
  whenever the localStorage cache had already painted something, on the reasoning that the
  user still had usable data. In practice that turned every backend outage into "the
  database did not load": stale or empty data on screen, no error, nothing to report. The
  `silent` flag now suppresses only the transient toast — a banner above the content names
  the failure, says how many items are actually in the current view, and offers a retry. It
  lives **outside** `#content` because `render()` rewrites that element wholesale, and it
  clears on the next successful refresh.

- **Every notification is sent under `MAIL_SENDER` = "i-Nventori Pejabat IMEN", and
  `_sendMail_()` is the only place that touches `MailApp`.** MailApp defaults the sender's
  display name to the owning account's own name, so recipients were seeing the raw Gmail
  address. The helper sets `name` on every send; the subject prefix and the email header
  and footer use the same constant, so the From line and the body cannot disagree.
  **The address cannot be changed from code** — Apps Script always sends as the account
  that owns the script — so `imenmakmal@gmail.com` is still there for anyone who expands
  the header. Hiding it needs a Gmail "Send mail as" alias on that account, which is
  configured in Gmail, not here. Tests assert every sent mail carries the name and that
  exactly one call site references `MailApp` directly; the check is proven to fail when a
  send bypasses the helper.

- **Alat Tulis shows one number: the current balance.** The list used to read `48/50`,
  but `quantity_total` and `quantity_available` move together in every path that touches
  them — registration sets both, stock add/remove applies the same delta to both, and the
  audit adjust moves both by the same variance — so for a consumable the two are always
  equal and the pair was printing the same number twice. The balance is what the system is
  for: 50 in, 2 used, 48 left; add 5, it is 53. `quantity_total` still exists in the sheet
  and is still maintained — nothing in the ledger changed — it is simply no longer shown.
  Removed from the items table, the item history, the CSV export, and the registration
  form (which now asks for **Stok Semasa**, the opening balance).

- **Handler strings passed to `mi()` must not contain a double quote.** `mi()` drops the
  handler into an onclick attribute delimited by double quotes, so a quote inside closes
  the attribute early and the browser keeps a truncated fragment — `openStock(1,` — giving
  a menu entry that silently does nothing, with no console error. Tambah Stok and Keluar
  Stok shipped that way for a while, because they are the only entries that pass a string
  argument. `check_html.js` now evaluates every `mi()` call site and reads the result the
  way a browser would; the check is proven to fail when the quoting is reverted. A static
  scan cannot catch this — the attribute template is in `mi()` and the argument is at the
  call site — which is why the guard evaluates rather than greps.

- **Stock actions live only in the row's ⋮ menu.** An inline row button was tried and
  removed at the owner's request: with the menu working there were two Tambah Stok
  controls on the same row.

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
  a `<datalist>` of values already used, built from `D.transactions` once the ledger
  has merged in (see §4) — no Suppliers table
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

## 9. How to verify a change before considering it done

There is no local Apps Script emulator, so verification uses three techniques.

### 1. Backend unit tests — run before any deploy

```bash
node tests/test_backend.js
```

**376 assertions.** Loads `Code.gs` into a `vm` context with hand-written mocks for
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

**23 checks.** Parses every inline `<script>`, verifies tag balance (the "page won't load"
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

## 10. What has not been proven

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

---

## 11. Session log — what changed and why, most recent first

Only entries a future session would be misled without. Not a changelog; the git history is
the changelog.

### 2026-08-21 — the load actually fixed, and a mobile pass

A diagnostic run showed `getInitialData` **still** returned `null` even with the ledger
split deployed (the payload was ~90KB). So the split was necessary but not sufficient. The
real fix (§4, §8): the initial payload is now lean (`_itemForList_`, ~11 fields), Date-free
(`_jsonSafe_`), and its alerts are counts — ~30KB total, ~1.1× the `getPublicCatalog` that
always delivered. Full item detail moved to `getItem(id)` fetched when a drawer opens, so
the edit form had to fetch-before-prefill or it would have blanked serial/notes on save.
Slow loads were separately addressed: `ensureSheetsOnce_`, a memoised spreadsheet handle,
and — now that the payload fits under 100KB — a `CacheService` cache that finally works.

Same day, a mobile pass: `dvh` instead of `vh` (the address bar made `100vh` overrun and
dragged taps off), 44px touch targets for the icon rail (the only nav on a phone), and
safe-area insets. The deeper touch offset is the nested-iframe workaround; its real fix is
the direct-serve deployment (§4), which is the owner's task.

Awaiting the owner's `Code.gs` deploy to take effect — verify with the in-app diagnostic.

### 2026-08-19 — the day the admin screen would not load

Four separate faults were tangled together and looked like one. Untangling them took a
whole session, so the order matters:

1. **Tambah Stok and Keluar Stok were dead buttons.** `mi()` drops its handler into an
   `onclick="…"` attribute, and these two were the only menu entries passing a string
   argument — written with double quotes, which closed the attribute early. The browser
   kept `openStock(1,` and clicking did nothing, silently. Fixed with escaped single
   quotes; `check_html.js` now evaluates every `mi()` call site and reads the result the
   way a browser would.
2. **`svcCheckOut` had no type guard**, so an Alat Tulis could be booked out as a loan that
   nothing could ever close. Now refused server-side, and the menu no longer offers it.
3. **The Pages build could not reach the backend at all** — POST and every ContentService
   response return "Page not found" on this deployment. Still true. See §4.
4. **`getInitialData` returned `null`** because the payload had grown to ~141KB. Split:
   the ledger now travels separately. See §4 and §8.

Also that day, in passing: `return_token` was being sent to every client that asked for the
payload despite nothing reading it; the admin's `<head>` prefetch was a dead end when it
failed; a failed refresh was silent whenever the cache had painted; and `render()` left the
boot placeholder up for ever when data never arrived — so "never arrived" and "still
loading" looked identical. All fixed, all in §8.

The lesson worth keeping: **the app could not say what was wrong.** Every failure mode
above presented as a blank screen. The banner, the named errors, and the in-app diagnostic
exist because of that, and are worth more than any single fix on this list.
