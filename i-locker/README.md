# i-Locker 3D

Interactive 3D locker-booking interface for students — **4 banks (A/B/C/D) × 3×3 = 36 lockers**, built with Three.js.

> **Stage: live demo (UI only).** All data is in-memory and resets on refresh.
> There is **no backend wired up yet** — booking/release actions are dummy
> operations so the interface can be shown and tested.

## Live

https://imenhub-portal.github.io/i-locker/

## Status colours

| Colour | Meaning        |
|--------|----------------|
| Green  | Available      |
| Orange | Booked         |
| Red    | Overdue        |
| Grey   | Unavailable    |

## Features

- Full 3D locker bank with recessed panels, handles and slot-number decals
- Orbit / zoom / pan (drag, scroll, right-drag); az/ polar angles clamped so the bank never leaves view
- Hover nudge + tooltip (desktop); tap-to-open (mobile)
- Click a locker → door swings open, camera flies in, details drawer slides in
- Book flow: student name + ID + duration → door closes, locker turns orange
- Legend with quick filters (All / Available / Booked / Overdue)
- Auto-framing for phone, tablet and desktop; bottom-sheet drawer on phones
- Runs offline-friendly as a single file (Three.js + Tailwind from CDN)

## Running locally

Open `index.html` in a browser, or serve the folder:

```sh
python3 -m http.server 8000
# then visit http://localhost:8000
```

## Hosting

This folder is published as part of the `imenhub-portal.github.io` GitHub Pages
site. Pushing to `main` publishes it automatically.

## Planned backend (not built yet)

The intended design follows the existing `i-nstrumen` / `i-Nventoriofis` pattern:

- **UI** stays on GitHub Pages (single source of truth).
- **Data** moves to a Google Apps Script project (`Code.gs`) + a Google Sheet.
- The frontend will get a small `google.script.run` shim so the same code works
  both on GitHub Pages (via `fetch` to the Apps Script JSON API) and inside
  Apps Script (native `google.script.run`).

When that work starts, `Code.gs` and `appsscript.json` will be added here and
pushed to Apps Script with `clasp` (see `.claspignore` — only the backend is
uploaded, never a second copy of `index.html`).
