# i-Print

Online poster printing request form for UKM IMEN.

## Repository layout

| Path | Purpose |
|------|---------|
| `index.html` | Public form UI. Static site served by GitHub Pages. Edit → commit → push (auto-deploys). |
| `_backend/Code.gs.txt` | **Source copy** of the Google Apps Script backend. NOT auto-deployed. Paste its contents into the Apps Script editor that owns the spreadsheet + Drive. |
| `_backend/appscript.json` | Apps Script manifest (OAuth scopes). |

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
