# i-Nstrumen — project context for a fresh session

Read this before doing anything else in this folder. It exists because most of what
you'd otherwise discover by exploring is either scattered across a long chat history
on one machine, or sitting in `_backend/`/`.claude/` files that **do not sync** to a
fresh `git pull` (see "Sync model" below). This file does sync — keep it updated
whenever you make a change worth a future session knowing about.

## What this is

A Google Apps Script backend + GitHub Pages frontend + Google Sheets database, for
booking and managing shared lab equipment (UKM's IMEN lab). Three pieces:

- **Backend**: `Code.gs` — a Google Apps Script project bound to a Google Sheet
  (tabs: `Equipment`, `Bookings`, `Logs`, `Config`). Deployed as a web app.
- **Frontend**: `index.html` — a single self-contained file (Tailwind CDN, Phosphor
  Icons, vanilla JS, no build step, no framework). Served two ways simultaneously:
  - **GitHub Pages** at `https://imenhub-portal.github.io/i-nstrumen/` — the primary
    entry point for students/PICs (no Google banner, no `authuser` issues).
  - **Apps Script's own `doGet`** — fetches the *same* `index.html` live from
    `raw.githubusercontent.com` on every request (single-source-of-truth UI; falls
    back to a locally-pasted emergency copy only if GitHub is unreachable).
- **Database**: the Google Sheet itself, read/written entirely through `Code.gs`.

Part of the `imenhub` monorepo (remote: `imenhub-portal/imenhub-portal.github.io`,
a **public** GitHub Pages repo). Sibling apps live alongside this one: `i-print/`,
`i-menian/`, `i-office/`, `i-survey/`, `i-KKPelanggan/`, `i-elemen/`,
`i-surveycafe/`, `i-staff/`. **Because the repo is public, no real secret value
(API key, admin PIN, password) may ever appear in a committed file — see Secrets
below for where they actually live.**

## Sync model — read this before assuming a file is "in the repo"

Run `git ls-files i-nstrumen/` to see exactly what syncs. As of this writing, only:

```
i-nstrumen/Code.gs
i-nstrumen/index.html
i-nstrumen/guideline1.html
i-nstrumen/guideline2.html
i-nstrumen/logoimen.png
i-nstrumen/CLAUDE.md          (this file)
```

Everything else physically present in this folder — `.claude/` (dev-server config,
local settings) and `_backend/` (a stale duplicate `Code.gs`, an old technical
reference `full.md`, generic AI-conduct notes in `4CLAUDE.md`, an Android WebView
wrapper project, HTML glossary/workflow pages, a CSV seed file) — is **excluded from
git** via `.git/info/exclude` at the repo root (`_local/`, `.claude/`, `_backend/`).
None of it will be present after a fresh `git clone`/`git pull` on another machine.
**Treat `_backend/Code.gs` as stale junk, not a reference** — the tracked
`i-nstrumen/Code.gs` is the one that matches (or should match, once redeployed) the
live Apps Script project.

**Why `Code.gs` is tracked at all** (this is intentional, confirmed with the user,
not an oversight): committing it is how the user moves work between machines —
`git pull` on a different PC/laptop fetches the latest `Code.gs`, they edit it
locally with an AI coding tool, then **manually paste it into the Apps Script
editor** to actually deploy (`git push` does *not* auto-deploy the backend). This
only works safely because the file contains **no literal secrets** — see below.

To recreate the local dev-server config on a fresh machine (`.claude/launch.json`,
used to browser-preview this app locally via the Browser preview tools):

```json
{
  "version": "0.0.1",
  "configurations": [
    {
      "name": "i-nstrumen-static",
      "runtimeExecutable": "npx",
      "runtimeArgs": ["--yes", "http-server", "-p", "8123", "-c-1", "."],
      "port": 8123
    }
  ]
}
```

This serves the current directory as static files on `localhost:8123` — good enough
to exercise `index.html`'s UI logic, but note it still calls the **live** production
`/exec` backend (there's no local Apps Script emulator), so treat any write actions
during testing as hitting real data.

## Deploying a change

Frontend and backend deploy completely differently — mixing them up is the single
easiest way to waste an hour:

1. **Frontend (`index.html`)**: commit + push to this repo. GitHub Pages publishes
   it automatically. Apps Script's `doGet` also picks it up automatically (it fetches
   the file live from GitHub on every request — no redeploy needed on that side).
2. **Backend (`Code.gs`)**: git push does **nothing** to the live backend. You must
   open the Apps Script editor, paste in the current file contents, then
   **Deploy → Manage deployments → ✏️ Edit existing deployment → Version: New
   version → Deploy**. ⚠️ **Never click "New deployment"** — that mints a *new*
   `/exec` URL, and `WEB_APP_URL` (`Code.gs:4`) / `API_URL` (`index.html:16`, inside
   the `google.script.run` shim) must stay identical between both files or the whole
   app breaks.

Either order is safe mid-rollout as long as you don't change the JSON contract
(`doPost`'s `{fn, args}` request shape / `{ok, result}` or `{ok:false, error}`
response shape) — old frontend can talk to new backend and vice versa.

## One-time Apps Script setup (Script Properties)

Some things only live in the Apps Script project's own settings (**Project
Settings → Script Properties** in the editor) — never in any file, so they never
touch git. If you're setting up a fresh deployment or these are missing, things
fail *closed* (safe) rather than open, but silently:

| Property | Used by | Effect if missing |
|---|---|---|
| `MASTER_CREDS_JSON` | `adminLogin()` (`Code.gs`) | Master/admin login always returns "Access denied" (regular PIC email-based login still works) |
| `GEMINI_KEY` | `_callGemini()` | SmartMatch AI chat throws, falls through to next provider |
| `MISTRAL_KEY` | `_callMistral()` | Same, falls through to Groq |
| `GROQ_KEY` | `_callGroq()` | If all three are missing, SmartMatch fails entirely |

`MASTER_CREDS_JSON`'s shape: a JSON object mapping a PIN or email (lowercase) to a
display name, e.g. `{"<pin>":"<name>", "<email>":"<name>"}` — ask the project owner
for the actual value; it is deliberately not written here.

## Architecture summary

**`doGet(e)`** (`Code.gs`) — serves the app. Modes, checked via query params:
`?format=json` → raw equipment/labs JSON for external tools; `?page=glossary` →
serves the `ijana` HTML file; default → fetches and serves `index.html` from
GitHub (see Sync model), injecting `?lab=`/`?view=` deep-link params, with a
locally-pasted fallback if GitHub is unreachable.

**`doPost(e)`** — the JSON API the frontend actually talks to. Three entry points,
explicitly allowlisted (nothing else is callable):
- `getInitialData` — full app payload (equipment, bookings, logs, config). Cached
  server-side via `CacheService` for 45s (gzipped+chunked, see `SYSDATA_CACHE_*`
  constants), invalidated immediately on any write (`SYSDATA_MUTATING_ACTIONS`).
- `fetchUpdates(lastTimestamp)` — incremental poll (last 50 logs + all bookings,
  filtered by timestamp), always live/uncached. Driven by the frontend's 60-second
  `setInterval` (`startPolling()` in `index.html`) — **this is intentional and still
  needed**; it's what lets a booking approval/rejection or a new log entry show up
  for other open tabs without a manual refresh, and it self-pauses whenever a modal
  or the registration form is open so it can never disrupt someone mid-edit.
- `handleFrontendAction(actionType, payload)` — everything else, dispatched via a
  `switch`. Current action list (`Code.gs`, search `case '...'` inside this
  function): `AdminLogin`, `Usage`, `Book`, `UpdateBooking`, `MarkNoShow`, `Report`,
  `EditEquipment`, `DeleteEquipment`, `EditLab`, `AddLab`, `DeleteLab`,
  `EditCoordinator`, `EditTechStaff`, `Archive`, `CancelBooking`, `FindMyBookings`,
  `GetSmartMatch`, plus the i-Menian Crossbridge actions `SearchIMenian`,
  `GetIMenianPhoto`, `GetIMenianPhotoBatch`, `SearchAllLabsIMenian`, `GetLabData`.

**Authorization** — two tiers, enforced server-side (not just client-side):
- `ADMIN_ONLY_ACTIONS` (edit/delete equipment, labs, coordinators, tech staff,
  archive) — any valid logged-in session (PIC or Master), no per-lab restriction.
- `LAB_SCOPED_ACTIONS` (`UpdateBooking` — approve/reject) — valid session **and**
  (Master, or the booking's lab is in that PIC's assigned labs).
- A session comes from `adminLogin(credential)`: matches `MASTER_CREDS_JSON` first
  (role `master`, all labs), else matches a `Config` sheet coordinator by email
  (role `pic`, that person's lab(s)). Issues a random token stored in
  `PropertiesService` (not `CacheService` — its 6h max TTL is too short) with an
  8-hour expiry (`ADMIN_TOKEN_TTL_MS`). The frontend attaches this token as
  `payload.__adminToken` automatically for gated actions (see
  `ADMIN_GATED_ACTIONS` in `index.html`'s `backgroundSync()`) and does a soft
  logout if the server ever reports the session expired/invalid.
- **Known limitation, not a bug**: PIC login has no real password — just a
  coordinator-email match. This was a deliberate scope decision (adding real
  per-PIC passwords is future work, not assumed broken).

**Data model** (Sheets, one tab each): `Equipment` (name, lab, accessMode
`direct`/`booking`/`both`, status, imageUrl, processCapabilities,
materialsOptions, maintenance info), `Bookings` (id, equipmentName, lab, userName,
userEmail, userPhone, userId, affiliation, supervisor, date, duration, samples,
materials, process, variant, status, remarks, timestampCreated, timestampActioned,
paymentType, paymentRef), `Logs` (usage/report/rejection history, up to 5000 rows
served on initial load), `Config` (A2:A = lab list, C2 = coordinators JSON, D2 =
tech staff JSON, E2:F = key/value settings incl. `officialEmail`/`lastArchive`).

**i-Menian Crossbridge**: this app reads (read-only) a *separate organization's*
Google Sheet (member directory: name/email/phone/matric/photo) by hardcoded ID, to
autocomplete supervisor/PIC name lookups. Deliberately cross-lab (a supervisor can
be registered under a different lab than the one being booked). As of the PDPA
pass below, this is a server-side *search* (`SearchAllLabsIMenian`, query-filtered,
top 12 results) — never a bulk directory dump to the client.

## Recent major changes (most recent session first)

- **Added "Volume (ml)" tracking unit + fixed historical unit-mixing bug** —
  `trackingUnit` (per-equipment dropdown: Hour / Quantity (pcs) / Weight (mg) /
  Weight (g), now also Volume (ml)) used to be a purely *live* display label —
  a historical Usage log had zero memory of what unit it was recorded under, so
  the Equipment Utilization Summary just summed ALL of an equipment's historical
  `samples` and relabeled the whole total under whatever unit is currently
  configured (e.g. switching Weight (g) → Volume (ml) silently merged old grams
  into a number now shown as "ml", with no indication a change ever happened).
  Fixed: `saveLog()` now writes a `unit` column (col 17/Q, auto-heals its header
  the same way `userEmail`/col 16 already does) captured from `eq.trackingUnit`
  at the moment each Usage log is created; `_buildUtilizationData()` now groups
  samples by (equipment, unit) instead of equipment alone, so an equipment with
  logs spanning two different units gets two separate rows instead of one merged
  one. Logs written before this shipped have no `unit` and are bucketed as
  "Unspecified (legacy)" rather than guessed at — old data isn't and can't be
  retroactively relabeled, but nothing new will ever be silently mixed again.
- **PIC name shown in the lab page header** — `getCoordinatorForLab()` (was already
  used internally for booking-email routing) is now also rendered for users to see,
  with a "No PIC assigned" fallback.
- **Booking Requests admin panel redesign** (`renderAdmin()` in `index.html`):
  summary stats bar (Pending/Approved/Rejected/Cancelled counts), History
  sub-filters now show a count per status, a consistent completeness dot replaces
  an ad-hoc warning string, Supervisor/Qty/Material collapse behind a per-row
  "Show more" toggle, equipment avatars show the real `imageUrl` photo with an
  icon fallback. Purely visual — no backend change, no new data tracked.
- **Cancel Booking button hidden for walk-in-only equipment** — walk-in
  (`accessMode: 'direct'`) never creates a booking record, so there was nothing to
  cancel; now gated behind the same `mode === 'booking' || mode === 'both'` check
  the Book button already uses.
- **PDPA compliance security pass** (the biggest change): admin/PIC write actions
  previously had *zero* server-side check — `state.adminAuthenticated` was just a
  client-side JS boolean, trivially bypassable via devtools regardless of login.
  Added: real server-issued session tokens (see Authorization above); split the
  public "missed check-in" self-service path out of `UpdateBooking` into a new
  `MarkNoShow` action (so gating `UpdateBooking` to PICs doesn't break that public
  flow — the server independently re-validates the booking window actually
  passed); replaced `GetAllIMenianUsers` (shipped another org's entire member
  directory to any anonymous visitor) with `SearchAllLabsIMenian` (server-side
  filtered search); moved `MASTER_CREDS` out of `index.html` (readable via View
  Source) **and** out of `Code.gs` itself (also public, since it's git-tracked —
  see Sync model) into Script Properties.
- **Page-load speed** (from ~3-4s to near-instant on repeat visits): frontend
  stale-while-revalidate cache in `localStorage` (`instrumen_data_v1`, no
  expiry check currently — acceptable since it's display-only, all writes still
  hit the live backend), an early `fetch()` fired from `<head>` before the rest of
  the page parses, `pdfmake`/`vfs_fonts` (~1MB) now lazy-loaded on first PDF export
  instead of blocking boot, a `CacheService`-backed 45s cache on `getInitialData`
  wrapping the previously-uncached `getSystemData()`, batched `Config` sheet reads
  (was 4 separate round-trips), and a fast-path date formatter avoiding
  `Utilities.formatDate`'s per-cell JS↔Java bridge cost (falls back to the slow
  path if the sheet's timezone ever differs from the script's).

## Gotchas / things not to "fix" without asking first

- `DEMO_PASSWORD` in `index.html` is unrelated to the admin auth system above —
  it's a separate, pre-existing "demo mode" toggle that hard-blocks all writes
  regardless (`backgroundSync()`'s first check), so its own weak client-side gate
  is low-stakes by design.
- Tailwind is still loaded via CDN JIT (`cdn.tailwindcss.com`) — a known
  production anti-pattern, intentionally not yet replaced (would need a build
  step, which conflicts with the "single self-contained HTML file, paste to
  deploy" workflow this project deliberately uses).
- Any authenticated PIC can currently edit *any* lab's equipment/coordinators, not
  just their own — the UI already scopes the booking-approval list to a PIC's own
  lab(s), but equipment/coordinator edits don't have that restriction. This was a
  conscious choice (kept as-is rather than tightened) during the PDPA pass, not an
  oversight — ask before changing.
- This is a shared monorepo with independent sessions working on other apps.
  Always `git pull --rebase` before pushing, and keep commits scoped to
  `i-nstrumen/` files unless a cross-app change is genuinely intended.

## How to verify a change before considering it done

No local Apps Script emulator exists, so verification leans on three techniques
used throughout this project's history:

1. **Live browser testing**: start the `i-nstrumen-static` dev server (see Sync
   model above) and drive the UI through the Browser preview tools — it talks to
   the real production backend, so this exercises real data end-to-end.
2. **Direct backend probing**: from the browser console (or any JS environment),
   `fetch()` the `/exec` URL directly with `{fn:'handleFrontendAction', args:[...]}
   ` to test one backend action in isolation without going through the UI —
   useful for confirming exactly what's live vs. what's only in the local file.
3. **Node.js unit tests for backend logic before it ever touches Apps Script**:
   extract a function's source text out of `Code.gs` by regex and `vm.Script`/
   `vm.createContext` it with hand-written mocks for `SpreadsheetApp`,
   `PropertiesService`, `CacheService`, `Utilities` — cheap, fast, and catches
   logic errors (e.g. auth bypass, off-by-one date math) before a real deploy.

Always syntax-check both files before considering an edit done — `Code.gs` with
plain `vm.Script`, and each inline `<script>` block of `index.html` the same way
(skip the one containing the `<?!= ... ?>` GAS templating scriptlet, which isn't
valid standalone JS by design).
