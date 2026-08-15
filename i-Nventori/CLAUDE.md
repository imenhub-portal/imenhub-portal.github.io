# i-Nventori — project context for a fresh session

Read this before doing anything else in this folder. It exists because the useful
context — why this is not a Laravel app, what deploys how, where the secret lives — is
otherwise scattered across a chat history on one machine, or sitting in files that
**do not sync** to a fresh `git pull` (see "Sync model"). This file does sync — keep it
updated whenever you make a change worth a future session knowing about.

## What this is

A Google Apps Script backend + GitHub Pages frontend + Google Sheets database, for
tracking IMEN's office inventory: fixed assets (laptops, monitors, furniture) and
consumables (stationery, cartridges), with check-out/check-in lifecycle, full date
tracking, an immutable transaction ledger, and a banking-style dashboard.

- **Backend**: `Code.gs` — Apps Script bound to a Google Sheet (tabs: `Items`,
  `Categories`, `Locations`, `Custodians`, `Transactions`). Deployed as a web app.
- **Frontend**: `index.html` — one self-contained file (hand-written CSS, vanilla JS,
  no framework, no build step). Served two ways simultaneously, same as i-nstrumen:
  GitHub Pages at `https://imenhub-portal.github.io/i-Nventori/`, and Apps Script's own
  `doGet`, which fetches this same file live from `raw.githubusercontent.com` on every
  request (single source of truth), falling back to a stub if GitHub is unreachable.
- **Database**: the Google Sheet, read/written only through `Code.gs`.

Part of the `imenhub` monorepo (remote: `imenhub-portal/imenhub-portal.github.io`, a
**public** GitHub Pages repo). **Because the repo is public, no real secret value may
ever appear in a committed file** — see Secrets below.

## Why this is not Laravel

The original specification asked for Laravel 11 + MySQL. That cannot run here: GitHub
Pages serves static files only and has no PHP runtime, and the whole monorepo (10 sibling
apps) is built on the Apps Script + Sheets pattern. **This was raised with the project
owner, who chose to build the full specification on the imenhub stack rather than change
hosting.** Nothing in the spec was dropped — every Laravel construct has a direct
equivalent, and the names were kept so the code stays reviewable against the spec:

| Spec (Laravel) | Here (`Code.gs`) |
|---|---|
| Migrations | `ensureSheets_()` — idempotent, with header auto-heal |
| Eloquent model + `$casts` | `_readTable_()` |
| `scopeExpiringWarranty()` / `scopeOverdue()` | the `SCOPES` object |
| `DB::transaction()` | `_txn_()` (LockService) |
| `StoreItemRequest` etc. | `_validate_(schema, payload)` |
| Routes + `auth` middleware | `doPost` switch + `ACTIONS_ADMIN` |
| `InventoryService` | the `svc*` functions |
| Scheduler `inventory:check-dates` | `checkDates()` + `installTriggers()` |
| `<x-table>` / `<x-badge>` / `<x-slide-over>` | `itemsTable()` / `pill()` / `openDrawer()` |
| SoftDeletes | the `deleted_at` column, filtered on every read |

## Sync model — read this before assuming a file is "in the repo"

Run `git ls-files i-Nventori/` to see exactly what syncs. As of this writing:

```
i-Nventori/CLAUDE.md          (this file)
i-Nventori/Code.gs
i-Nventori/appsscript.json
i-Nventori/index.html
i-Nventori/tests/test_backend.js
i-Nventori/tests/check_html.js
```

`.claude/` is **excluded from git** via `.git/info/exclude` at the repo root (along with
`_local/` and `_backend/`). Note that rule: **do not put anything you want to keep in a
folder named `_backend/`** — that is why the tests live in `tests/` and not `_backend/`
like i-print's do.

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

**Why `Code.gs` is tracked**: committing it is how work moves between machines — `git
pull` on another PC fetches the latest, you edit it locally, then **manually paste it
into the Apps Script editor** to deploy. `git push` does *not* deploy the backend. This
is only safe because the file contains no literal secrets.

## Live deployment

Already wired in — do not re-enter these:

| Constant | Value | Where |
|---|---|---|
| `SPREADSHEET_ID` | `1Xk1aKMmWR3AFTvWlMJDKUmv-5Ik_GSvyZeqXX30j_6E` | `Code.gs:25` |
| `FOLDER_ID` | `1NQUV2EHXpQhlWzcCm2aaIEOUvZ4CGOE7` | `Code.gs:26` |
| `ADMIN_EMAIL` | `imenmakmal@gmail.com` | `Code.gs:27` |
| `API_URL` (`/exec`) | deployment `AKfycbyY2fSJbt6…` | `index.html`, in the shim |

**Two different URLs, easy to confuse:**

- `https://script.google.com/macros/s/AKfycbyY2fSJbt6…/exec` — the **API**, from the
  Apps Script deployment. The page calls this for data.
- `https://imenhub-portal.github.io/i-Nventori/` — the **link people open**, from GitHub
  Pages. This is what you share.

Remaining one-time steps in the Apps Script editor:

1. Paste `Code.gs` into the bound project.
2. **Project Settings → Script Properties → add `ADMIN_PASSWORD`.** Never put this in a
   file — the repo is public.
3. Run `setup()` once. It creates the sheet tabs, seeds starter categories/locations, and
   installs the daily 07:00 trigger.
4. Deploy → Manage deployments → Edit → New version → Deploy.

## Secrets

| Property | Used by | Effect if missing |
|---|---|---|
| `ADMIN_PASSWORD` | `_adminPassword_()` | **Fails closed** — nobody can log in and every gated action is denied. This is deliberate: a missing property must never mean "no password required". |

## Deploying a change

Frontend and backend deploy completely differently — mixing them up wastes an hour:

1. **Frontend (`index.html`)**: commit + push. GitHub Pages publishes it automatically,
   and Apps Script's `doGet` picks it up on the next request (no redeploy needed).
2. **Backend (`Code.gs`)**: push does **nothing** to the live backend. Paste the file
   into the Apps Script editor, then **Deploy → Manage deployments → ✏️ Edit existing
   deployment → Version: New version → Deploy**. ⚠️ **Never click "New deployment"** —
   that mints a *new* `/exec` URL, and `API_URL` in `index.html` must keep matching it.

Either order is safe mid-rollout as long as the JSON contract does not change
(`doPost`'s `{fn, args}` request shape, `{ok, result}` / `{ok:false, error}` response).

## Architecture summary

**`doGet(e)`** — `?format=json` returns the raw payload; otherwise fetches and serves
`index.html` from GitHub.

**`doPost(e)`** — the JSON API. Five entry points, explicitly allowlisted; anything else
is rejected: `getInitialData`, `handleAction`, `adminLogin`, `getItemHistory`,
`startResumableUpload`.

**`handleAction(actionType, payload)`** — the single dispatch point. Authorization is
enforced *before* the switch, so no individual handler can forget it. Every write is
wrapped in `_txn_` and appends exactly one `Transactions` row.

**Authorization** — one shared admin password, checked **server-side on every gated
action**. A client-side boolean is never trusted, because `doPost` is callable
anonymously from anywhere. Anonymous callers can read inventory but receive **no
custodian records and no ledger** — custodian names/emails are personal data (this is
the `GetAllIMenianUsers` lesson from i-nstrumen, applied up front).

**Data model** — see `SCHEMA` at the top of `Code.gs`; column names are the spec's,
verbatim. Two columns were *added* beyond the spec because its own depreciation formula
needs them: `salvage_value` and `useful_life_years`.

**The ledger is genuinely immutable, not just by convention.** `_update_()` throws if
handed the `Transactions` tab, so there is no code path that can edit or delete a ledger
row. A check-in appends a **new** `check_in` row; it does not touch the original
`check_out` row. Open loans are *derived* by replaying the ledger (`_openLoans_`) rather
than stored, so the two can never drift apart. A test asserts this mechanically.

**Date engine** — `checkDates()` runs daily at 07:00 via a time-driven trigger and emails
one digest. It uses the same `SCOPES` predicates the dashboard badges use, so the email
and the UI can never disagree about what is overdue. A test asserts they match.

## Gotchas / things not to "fix" without asking first

- **`bookValue()` in `index.html` deliberately mirrors `_bookValue_()` in `Code.gs`.**
  The duplication is intentional — it powers the live "nilai buku" preview in the item
  drawer before anything is saved. If you change one, change both.
- **Quantity is not editable on the item edit form**, by design. Stock only moves through
  the ledgered actions (Tambah/Kurang Stok, check-out/in, audit), so the item row can
  never disagree with the ledger. Making the field editable would reintroduce that drift.
- **A fixed asset's open-loan check runs before its stock check** in `svcCheckOut`. An
  asset out on loan also has `quantity_available: 0`, and "stok tidak mencukupi" is a
  true but useless answer to "why can't I borrow this?". Order matters; a test covers it.
- **Asset tags widen past 9999** (`AST-2026-10000`) rather than wrapping, which would
  collide with `AST-2026-0001`. Soft-deleted items keep their tag reserved.
- **CDN libraries are lazy-loaded on first use, never at boot** (`qrcodejs`,
  `html5-qrcode`). `tests/check_html.js` fails the build if a `<script src>` appears in
  the document. Keeping boot dependency-free is why the page loads instantly.
- **No Tailwind.** i-nstrumen flags its Tailwind CDN as a known anti-pattern; this app
  uses hand-written CSS custom properties instead, so the "single self-contained file,
  paste to deploy" workflow keeps working.
- **The instant-load behaviour is a deliberate contract, not an accident.** On a repeat
  visit `boot()` paints straight from the `localStorage` cache — measured at ~5 ms, with
  no spinner and no "stale data" banner — then refreshes in the background. The refresh
  **compares the new payload against the cached one and skips the repaint entirely if
  nothing changed**, so a background fetch can never flicker the view someone is reading.
  A failed background refresh is silent by design (`refresh(true)`), because the user
  still has usable data and any *write* would surface its own error. Do not "fix" this by
  adding a loading state to the cached path.
- **The `<head>` prefetch is tagged with the password it used (`__earlyPass`),
  and `boot()` must only consume it when that matches the current session.**
  This is not defensive coding — it is a shipped bug that got fixed. On a first
  visit the prefetch runs anonymously, so the server correctly answers
  `is_admin:false`; `boot()` then read that stale anonymous payload *after* a
  successful login, concluded the session was expired, and bounced the user
  straight back to the login screen ("blinks in, then returns to login").
  `__earlyPending` covers the case where the prefetch was never issued at all,
  which would otherwise leave `boot()` waiting on a result that never arrives.
  The logout branch is also guarded on `PASS` being non-empty, so an anonymous
  payload can never trigger a logout. Five scenarios are covered by the browser
  walk-through in "How to verify" — re-run them if you touch boot/login.
- **Every colour is a CSS custom property**, which is what makes the dark theme a pure
  token swap under `[data-theme="dark"]` with no duplicated layout rules. If you add a
  hard-coded hex anywhere, dark mode silently breaks for that element. The theme is
  applied by a tiny inline script in `<head>` **before first paint**, so dark-mode users
  never see a white flash.
- **The dashboard is split into Aset Tetap and Inventori Guna Habis on
  purpose — do not merge them back into one summary.** The two are measured
  in different units: an asset's worth is its depreciated *book value*, while
  a box of pens is simply *unit cost x quantity on hand* (consumables have
  `useful_life_years: 0`, so `_bookValue_` correctly returns cost unchanged).
  Averaging them produced a headline number that meant nothing for either,
  which is what prompted the split.
- **Terminology is the owner's, and the two words are not interchangeable.**
  *Aset* = kerusi, meja, komputer, alat besar — tracked one by one, depreciated,
  and loanable if movable. *Inventori* = alat tulis, kertas, kartrij — tracked
  by quantity, issued not borrowed, never depreciated. The DB enum stays
  `fixed_asset`/`consumable`; only the labels changed. Do not reintroduce
  "Aset Tetap" or "Bahan Guna Habis".
- **`is_portable` gates who can be handed an asset.** Blank counts as movable
  so rows predating the column keep working; only an explicit `FALSE` (meja,
  kerusi) blocks check-out. Enforced server-side in `svcCheckOut`, and the
  check-out list filters to movable assets only so the admin is never offered
  a choice the server will reject. Inventory is excluded from that list
  entirely — pens are issued through Keluar Stok and never come back.
- **A handover can name a recipient who does not exist yet.** Typing a name +
  email in the check-out form creates (or matches by lowercased email) a real
  Custodian row, so the ledger keeps a genuine foreign key and the person is
  reusable — rather than the name living as loose text. `_resolveRecipient_`
  owns this.
- **Both ends of a handover send email**, each carrying a TRX reference, full
  timestamps, and — on return — the duration held and whether it was late.
  Sending is best-effort and wrapped in try/catch: the ledger row is already
  written, and a mail-quota failure must never roll back a handover that
  physically happened. Both actions return `mailed:true|false` so the toast
  reports honestly instead of claiming an email that never left.
- **A stock issue records `custodian_id`, and that is the whole point of it.**
  Before this, "who took the pens" could only be free text in `reason_notes`
  — unsearchable and impossible to total per person. The field is *optional*
  because stock also leaves for damage or loss with no recipient; those units
  count toward total issuance but are excluded from per-staff figures.
  `issueStats()` in `index.html` derives all per-staff and per-item
  consumption from the ledger alone, so it can never drift from it. A
  recipient on a stock *addition* is rejected — nobody receives a restock.
- **The dashboard chart is real data, not decoration.** `monthlyFlow()` derives value in
  vs. value out per month by joining the ledger against each item's `unit_cost` — the
  asset-register analogue of a profit/loss chart.
- This is a shared monorepo with independent sessions working on other apps. Always
  `git pull --rebase` before pushing, and keep commits scoped to `i-Nventori/`.

## How to verify a change before considering it done

There is no local Apps Script emulator, so verification uses three techniques:

1. **Node unit tests — run these before any deploy.**
   ```bash
   node tests/test_backend.js
   ```
   Loads `Code.gs` into a `vm` context with hand-written mocks for `SpreadsheetApp`,
   `LockService`, `PropertiesService`, `CacheService` and `MailApp`, backed by an
   in-memory spreadsheet, then drives the **real** service functions. 125 assertions
   covering depreciation edge cases, asset-tag generation, stock round-trips, all
   `SCOPES` boundaries (29/30/31 days, 364/365/366 days), authorization (including
   fail-closed with no password set), header auto-heal, and ledger immutability.
   *Note: values created inside the vm are from another realm, so `x instanceof Date` is
   false in the test file — use the `isDate()` helper.*

2. **Frontend static checks.**
   ```bash
   node tests/check_html.js
   ```
   Parses every inline `<script>`, checks tag balance (the "page won't load" bug class
   that has bitten i-print before), verifies every `on*=` handler resolves to a defined
   function, every icon reference exists in the sprite, and no CDN script loads at boot.
   *Watch out: the tag-balance check counts literal `<div` text in comments too — do not
   write example markup in a comment.*

3. **Live browser testing.** Start the static server (`.claude/launch.json` above) and
   drive the UI with the Browser preview tools. Because there is no local backend, the
   productive technique is to **replace `window.google` with a mock** exposing
   `getInitialData` / `handleAction` / `adminLogin` / `getItemHistory`, feed it a
   synthetic payload, then call `boot()`. That exercises every render path for real. Note
   the shim defines `google.script.run` as a *non-configurable* getter, so you must
   replace the whole `window.google` object rather than redefine the property.
   ⚠️ **When the Browser pane is not displayed the page does not composite frames, so CSS
   transitions never advance** and `getComputedStyle` returns the transition's first
   frame. Anything with a `transition` — the active nav pill's background, the drawer's
   slide, the mobile rail — will measure as its *starting* value and look broken when it
   is not. Inject `*{transition:none !important}` before measuring animated properties.

Always syntax-check both files before considering an edit done — `Code.gs` with plain
`vm.Script`, and each inline `<script>` block of `index.html` (that is what
`tests/check_html.js` does).
