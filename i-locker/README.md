# i-Locker 3D

Interactive 3D locker-booking interface for students — **4 locker banks (L1–L4) × 3×3 = 36 cabinets (K1–K9)**, built with Three.js.

## Naming

A cabinet is addressed as **`L{n}-K{n}`**:

- **`L{n}`** — the locker bank, `L1`–`L4` (shown on the label floating above each bank)
- **`K{n}`** — the cabinet slot, `K1`–`K9` (shown on the door decal)
- Ordered top-left → across → down: `K1..K3` top row, `K4..K6` middle, `K7..K9` bottom

So `L2-K5` = bank 2, middle-row centre cabinet.

Internally there are two separate identities, on purpose:

| Field  | Example | Used for |
|--------|---------|----------|
| `key`  | `25` (number) | every lookup — Maps, mesh userData, filters |
| `id`   | `L2-K5` | display only (drawer, tooltip, toasts) |
| `code` | `K5` | door decal |

Because the **lookup key is a number**, renaming the display text can never
silently break a lookup. A boot-time integrity check asserts all 36 cabinets
round-trip and fails loudly if they ever don't.


> **Stage: local demo (UI only).** Bookings reset on refresh; fob mappings and theme preferences persist in this browser.
> There is **no backend wired up yet** — booking/release actions are dummy
> operations so the interface can be shown and tested.

## Live

https://imenhub-portal.github.io/i-locker/

## Status colours

| Colour | Meaning        |
|--------|----------------|
| Green  | Available      |
| Orange | Booked         |
| Red    | Overdue (pulses / breathes) |
| Blue   | On Hold — key collected |

Doors are painted in bold status colours (not just tinted grey), so the state
reads from any angle. Overdue lockers breathe — the door glow and edge tag swell
and fade on a ~2.9s cycle, matching the pulsing dot in the header.

Booked means awaiting collection. The due clock starts at collection; an On Hold
cabinet automatically becomes Overdue when its due date passes. Collection and
return by physical scan are deferred; scans currently register fobs only.

## Controls

- **Reset View** (top bar) — smoothly returns the camera to the fitted home view
  and closes any open door.
- **Light / Dark** (top bar) — flips the whole theme: 3D scene, lighting and UI.
  The choice is saved in `localStorage` and honoured on the next visit; a first
  visit follows the OS `prefers-color-scheme`.

## Features

- Full 3D locker bank with recessed panels, handles and slot-number decals
- Fixed front view with a slight downward angle; left/right buttons slide the view
- Hover nudge + tooltip (desktop); tap-to-open (mobile)
- Click an available cabinet → door opens and camera zooms in. Held cabinets stay shut and shake with a red flash. Close returns to the overview.
- Book flow: student name + ID + duration → door closes, locker turns orange
- Legend with quick filters (All / Available / Booked / On Hold / Overdue)
- Auto-framing for phone, tablet and desktop; bottom-sheet drawer on phones
- Single HTML file; internet is required for CDN dependencies

## Privacy / PDPA

This is a public static demo, not a certification of PDPA compliance. Use fictitious
student details and test fobs only. Browser-side masking and a PIN are not access control.

- The demo uses non-identifying pseudonyms (`Student 01`…), not real people.
- Student references are synthetic and always displayed **masked**
  (`STU-••••81`). Full values are never written into the DOM.
- The locker detail panel shows occupancy (Booked / Overdue / dates) but marks
  the holder's identity **RESTRICTED** and hides it.
- Booking toasts never echo the holder's name.
- Form entries remain in this tab's memory until refresh. They are not sent to a backend.
- Fob mappings are stored in localStorage and accessible to anyone using that browser profile.

Any future backend must keep real records **server-side** (Apps Script / Sheet)
and reveal identity only to an authenticated owner or admin — never to this
public page.

## Touch / mobile

- Tap targets are ≥44px (Apple HIG / Material).
- Tap to inspect; use the left/right buttons to slide. Navigation is disabled while inspecting a cabinet.
- Bottom-sheet drawer, collapsible legend, safe-area insets and PWA meta for
  Add to Home Screen.

## PC admin and USB reader

Admin appears only for a fine-pointer, hover-capable viewport at least 1024px wide.
Unlock using demo PIN **1234**. This is a cosmetic SHA-256 PIN check, not real authentication.
Select a cabinet, choose **Register fob**, then tap the USB keyboard-wedge reader.
The focused browser receives digits followed by Enter; no USB bridge or port is used.
Input normalization matches the reference: digits only, last ten digits, minimum eight.
There is a 600ms inter-key buffer reset and a manual entry fallback.

Bindings are one-to-one. Registering replaces a cabinet's old binding and moves an
already-bound fob from its former cabinet. `ilocker-fobs` persists only in this
browser/origin; it does not synchronize between devices. The reference file was
removed and is not part of the app. Real-reader testing on the focused kiosk PC
is still required. GitHub Pages can receive these standard keyboard events.

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
