# i-Print

Online poster printing request form for UKM IMEN.

## Repository layout

| Path | Purpose |
|------|---------|
| `index.html` | Public form UI. Static site served by GitHub Pages. Edit → commit → push (auto-deploys). |
| `_backend/Code.gs.txt` | **Source copy** of the Google Apps Script backend. NOT auto-deployed. Paste its contents into the Apps Script editor that owns the spreadsheet + Drive. |
| `_backend/appscript.json` | Apps Script manifest (OAuth scopes). |
| `printing_allocation_logic.md` | **Reference spec** for the 5-channel allocation logic. The dashboard implements its math; admin settings are the live source of truth. |

> `_backend/` is the canonical backend source. Any machine can `git pull` and re-paste
> into Apps Script, so the latest logic always travels with the repo.

## Backend deploy / sync workflow

1. **UI changes** → edit `index.html`, commit & push. GitHub Pages serves it automatically.
2. **Backend changes** → edit `_backend/Code.gs.txt`, then manually paste the updated
   functions into the Google Apps Script editor and re-deploy (or save + redeploy web app).
3. Keep `Code.gs.txt` in sync with what is live in Apps Script.

## Script Properties (Apps Script → Project Settings → Script Properties)

- `ADMIN_PASSWORD` — admin dashboard password. Falls back to a hardcoded default if unset;
  set this to keep the admin panel locked.
- `SPREADSHEET_ID` and `ADMIN_EMAIL` are currently hardcoded in `Code.gs.txt`. Move them to
  Script Properties if you want them out of the public source.

## Supervisor (Penyelia) field

- Section 1 **"Maklumat Pemohon"** includes an optional **Nama Penyelia** input
  (`name="penyelia_nama"`).
- `processForm` auto-creates the `Penyelia` header column in the sheet on the first
  submission — no manual column setup needed.
- The value is also included in the admin notification email.

## Admin dashboard (banking-style sidebar)

After login, the admin sees a sidebar with 3 sections:

| Pane | Purpose |
|------|---------|
| 📋 **Tempahan** | Existing orders table (Active / History) + stat cards. |
| 📊 **Peruntukan** | Financial allocation dashboard — revenue split into 5 funds. |
| ⚙️ **Harga & Kos** | Price & cost constant editor (saved to Config sheet). |
| 🧾 **Pembelian** | Restock / spend records. Logs a purchase against a fund; the matching channel balance is deducted. |

Customers only see the **Borang** (form) and **Semak Status** tabs. Admin/Allocation/Settings
are hidden behind the admin password.

## Printing cost & allocation logic

Implements `printing_allocation_logic.md` — every order's revenue is split into 5 channels:

1. **Paper Fund** — RM 6.67/A1 (×2 A0), target RM 300
2. **Ink Fund** — RM 20.00/A1 (×2 A0), target RM 3,000
3. **Maintenance** — RM 2.83/A1 (×2 A0), ongoing
4. **Printhead Fund** — RM 0.50/A1 (×2 A0), target RM 2,500
5. **Net Profit** — fixed RM 5.00 (remainder handling below)

### Allocation math
- **A0 = ×2 all four cost constants** (paper, ink, printhead, maintenance).
- **Profit is fixed** at RM 5.00 regardless of selling price.
- **Surplus** = selling price − costs − RM 5. If surplus > 0, it splits **50/50**
  into the Paper Fund and Ink Fund.
- **Breakeven** = costs + fixed profit. The admin settings **refuse to save** any
  price below breakeven.

### Storage — Config sheet (sheet #2 of the spreadsheet)
A key/value sheet auto-created on first access (`_configSheet_` / `_cfgAll_` / `_cfgSet_`).
Seeded with defaults from the MD. Holds three blocks:
- **Prices:** `price_A1_glossy`, `price_A0_glossy`, `price_A1_plain`, `enabled_plain`
- **Costs:** `cost_paper`, `cost_ink`, `cost_printhead`, `cost_maintenance`, `fixed_profit`
  + target cycles (`target_paper`, `target_ink`, `target_printhead`)
- **Ledger:** `total_prints` + `ledger_paper`/`ledger_ink`/`ledger_maintenance`/
  `ledger_printhead`/`ledger_profit` (net balances = allocated − spent)

### Storage — Purchases sheet (restock spends)
A separate sheet auto-created on first purchase. Schema: `[date, channel, amount, note]`.
- `channel` is one of the 5 fund keys (`paper`/`ink`/`maintenance`/`printhead`/`profit`).
- `amount` is the RM spent (positive); the matching fund balance is **deducted**.
- **Recompute-safe:** `recomputeLedger` replays order revenue AND re-sums purchases, so
  spends are never overwritten away by a recompute.
- Deleting a purchase returns its amount to the fund.
- A channel can go **negative** (deficit, shown in red) — e.g. restocking before enough
  revenue has accumulated.
- A channel can **exceed its target** (e.g. 400/300); the bar overflows past 100% in green
  and a `+RM X surplus` label appears. It never stops accumulating.

### Dynamic pricing
The form no longer hardcodes prices. On load it calls `getPrices()` and stores the result
in `window.PRICES`; `calcPrice()` reads from it. If the fetch fails, it falls back to the
legacy defaults (50/80/30) so the form never breaks.

## Backend API endpoints (`doPost`)

| `fn` | Function | Gated |
|------|----------|-------|
| `getPrices` | `getPrices()` | no |
| `getAllocationConfig` | `getAllocationConfig(pass)` | yes |
| `saveAllocationConfig` | `saveAllocationConfig(pass, cfg)` | yes |
| `getAllocationLedger` | `getAllocationLedger(pass)` | yes |
| `recomputeLedger` | `recomputeLedger(pass)` | yes |
| `addPurchase` | `addPurchase(pass, channel, amount, note)` | yes |
| `getPurchases` | `getPurchases(pass)` | yes |
| `deletePurchase` | `deletePurchase(pass, rowIndex)` | yes |

(plus the existing `startResumableUpload`, `processForm`, `updatePosterStatus`,
`deleteRow`, `getAdminData`, `checkUserStatus`).

## Deploy steps (after backend changes)

1. Paste updated `Code.gs.txt` into the Apps Script editor.
2. **Run `recomputeLedger` once** (or click ↻ in the Allocation pane) to seed the ledger
   from existing order history.
3. Redeploy the web app (new version) so `doPost` learns the new routes.
4. The form auto-fetches prices on next load.
