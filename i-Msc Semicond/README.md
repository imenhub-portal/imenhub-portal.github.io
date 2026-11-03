# i-Msc Semicond — Marketability Study

Interactive questionnaire and analysis dashboard for the proposed
**Master of Science in Semiconductor Engineering** at
**IMEN, Universiti Kebangsaan Malaysia (UKM)**.

Built with plain HTML + Tailwind (CDN) + Alpine.js + three.js (r128), and a
Google Apps Script backend that writes to Google Sheets.

---

## Folder contents

| File | Purpose |
|---|---|
| `index.html` | Respondent survey — hero, programme infographic, profile + 3-step survey |
| `dashboard.html` | Password-gated analysis dashboard (English only) |
| `assets/data.js` | All survey content, bilingual UI strings, config |
| `assets/scenes.js` | three.js hero + per-course illustrations — interactive (drag to orbit, pinch/wheel to zoom, inertia) |
| `assets/survey.js` | Alpine survey application (course deck, draft persistence) |
| `assets/dashboard.js` | Dashboard aggregation, charts, CSV export |
| `assets/imen.css` | Shared visual theme (Inter body, Plus Jakarta Sans display) |
| `assets/logo.png` | IMEN logo |
| `assets/img/programme-flow.svg` | Programme-flow artwork |
| `Code.gs` | Google Apps Script backend |

---

## Respondent flow

1. **Hero (a toy, not a gate)** — the hero shows a static silicon wafer with a
   draggable gold die. Dropping it on the socket fires an electrical-pulse
   wavefront across the wafer; nothing is locked either way.
2. **Profile** — email, name, category, selection criteria → **Proceed**
   (only the profile is validated).
3. **Survey — 3 steps**
   1. **Courses** — the 14 course cards, **one at a time** (summary, "What you
      will learn" bullets, collapsible overview, "Card X of 14"). Previous /
      "Got it — next course"; the last card shows "Continue to questions".
      Each card has a hyper-real 3D illustration of its own (14 scenes, one file each in `assets/courses/`, built on the shared `assets/course-kit.js`; drag to orbit) and a required 1–5 suitability rating (stored as `core_1`…`elective_5`). Cards are marked read (`msc_semicond_read_v1`). The progress bar spans all 16 units (14 cards + preferences + review) and the label reads "Card X of 14" during the deck.
   2. **Preferences** — study mode, fee range, interest, feedback.
   3. **Review** → **Submit**.
   In-progress answers are saved as a draft (`msc_semicond_draft_v1`) and
   restored on reload; cleared on submit/reset.

---

## Run locally

```bash
cd "i-Msc Semicond"
python3 -m http.server 8080
```

Open <http://localhost:8080>.

---

## Live backend (Google Sheets)

The backend is wired and live:

- **Sheet:** `1Utai_Q0oAeJWkJeL9ux8SXCovXTUYProwbCMup-gUN8`
- **Web app:** `https://script.google.com/macros/s/AKfycbyVrdk1_LxoJn4RfU7ImQZ_CMJxDdp_nYld-T7pxj7ZduaPrimt4yCSYVJhWE3H6LG6/exec`
- `APP_CONFIG.DEMO_MODE` is `false`; submissions go straight to the
  `Responses_MSc_Semicond` sheet (created automatically with headers on the
  first write, or run `testSetup()` in the Apps Script editor).

### One remaining manual step — dashboard password

In <https://script.google.com> open the project and run once:

```js
setAdminPassword('YourStrongPassword');
```

Only the SHA-256 hash is stored as a Script Property; the plaintext never
touches the static files. Share the chosen password only with dashboard users.

To run locally in demo mode instead, set `DEMO_MODE: true` in
`assets/data.js` — responses then stay in `localStorage`, the dashboard demo
password is `imen-demo`, and 32 sample responses are seeded for the charts.

---

## Data model

Each response stores:

`timestamp, email, name, language, category, category_other, criteria,
study_mode, fee_range, fee_other, interest, feedback, user_agent`

- `criteria` is pipe-separated (`A | B | C`).
- Each course is rated 1–5 on its card; the ratings are the `core_1`…`core_9` / `elective_1`…`elective_5` columns. `Code.gs` extends an older sheet's header row automatically.
- `language` records the interface language (`en` / `ms`) but **does not change
  the answer keys**, so the dashboard aggregates every response into one
  English-only analysis — no double counting and no English-only filtering.

---

## Dashboard access

- Separate page: `dashboard.html`, not linked from the survey.
- Password is verified **server-side** by Apps Script. The client sends a
  SHA-256 hash; the plaintext password never appears in the static files.
- `noindex,nofollow` is set on the page.

---

## Deploy to GitHub Pages

The site lives at:
`https://imenhub-portal.github.io/i-Msc%20Semicond/`

Commit and push the folder to `imenhub-portal/imenhub-portal.github.io` on the
`main` branch. GitHub Pages serves it automatically.

---

## Content policy

- Every question and course description is taken **verbatim** from the official
  questionnaire document.
- Only obvious typos/spacing were corrected:
  - `Private Company-Employer/Executive/ Staff` → `Private Company – Employer / Executive / Staff`
  - `Self-Employed/ Freelancer/ Business owner` → `Self-Employed / Freelancer / Business Owner`
  - `intend to offer` → `intends to offer`
- The interface chrome is bilingual (English default, Bahasa Melayu toggle).
  Questions stay in English.

---

## Course illustrations (three.js)

- `assets/course-kit.js` — studio environment map, PBR materials, procedural textures and helpers.
- `assets/courses/<key>.js` — one scene per course (`device`, `fab`, `materials`, `package`, `systems`, `research`, `management`, `project1`, `project2`, `vlsi`, `mems`, `photonics`, `failure`, `supply`). Each registers `IMEN.courseScenes.<key>`.
- `course-preview.html?scene=<key>` — developer preview page with draw-call / triangle counts. Not linked from the site.

---

## Analysis dashboard

- `dashboard.html?preview=1` — opens the full dashboard with **generated sample data** (no password, nothing real shown). Use it to see what the analysis looks like.
- `dashboard.html` — real dashboard, unlocked with the admin password (`setAdminPassword()` in Apps Script; `imen-demo` in demo mode).
- Analyses: key findings, 8 KPIs, per-course ratings (mean ± 95% CI, SD, favourable %, distribution, verdict), course × category heat-map, interest / study mode / fee, interest by category, mode × fee, market criteria (+ by category), respondent mix, language, timeline, feedback keywords and comments, searchable response table. Filters: category, mode, interest, language, period.
- Downloads (all respect the active filters): Excel workbook (18 sheets), PDF report (print view), raw responses CSV, analysis summary CSV, JSON. Raw exports leave out email and name.
