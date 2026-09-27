# i-siswazah — Sistem Pengurusan Siswazah (context untuk sesi baharu)

Baca fail ini dahulu sebelum membuat sebarang perubahan dalam folder ini. Ia ditulis
supaya sesi AI pada PC/laptop lain boleh menyambung kerja tanpa perlu meneliti
sejarah chat yang panjang.

## Apa ini

Webapp **frontend sahaja** (mod demo) untuk pengurusan operasi siswazah / pascasiswazah
IMEN, UKM. Ia menggantikan fail Excel yang mempunyai banyak tab dengan satu interface
sidebar moden gaya sistem perbankan, di mana **setiap CSV menjadi satu workspace
berbentuk jadual yang boleh diedit**.

- **Frontend**: `index.html` + `styles.css` + `app.js` — tiada build step, tiada
  framework, vanilla JS sahaja. Fail CSS dan JS diasingkan supaya mudah diselenggara.
- **Backend**: `Code.gs` (Google Apps Script) — **BELUM dibina** dalam folder ini.
  Pengguna akan memindahkan `Code.gs` secara manual ke editor Apps Script kemudian.
- **Data**: mod demo menggunakan data contoh kecil dalam `app.js`, disimpan pada
  `localStorage` browser. **Tiada data sebenar** dalam demo.

Sebahagian daripada monorepo `imenhub` (remote:
`imenhub-portal/imenhub-portal.github.io`, repo **awam**). Aplikasi bersaudara:
`i-nstrumen/`, `i-print/`, `i-menian/`, `i-office/`, `i-survey/`, `i-KKpelanggan/`,
`i-elemen/`, `i-surveycafe/`, `i-staff/`.

**Kerana repo ini awam, jangan sesekali commit sebarang rahsia** (kata laluan, API key,
token). Mod demo ini tiada rahsia — ia tidak pernah menyentuh data sebenar.

## Fail penting

| Fail | Peranan |
|---|---|
| `index.html` | Struktur: sidebar, topbar (carian global), kandungan, footer kredit |
| `styles.css` | Semua gaya. Token warna/radius/bayang bermula di `:root` |
| `app.js` | Logik: definisi workspace, state, router, jadual, admin, export |
| `*.csv` | Data mentah **sumber** daripada Excel. Belum diimport ke app.js |

## Struktur data

### `WORKSPACES` dalam `app.js`
Setiap workspace mewakili satu tab Excel. Setiap kolum:
- `header` — teks **ASAL** daripada CSV (dikekalkan supaya pengendali tidak keliru).
  JANGAN tukar teks header tanpa persetujuan pengguna.
- `key` — id dalaman kolum.
- `type` — `text` | `date` | `number` | `select` | `textarea`.
- `identity: true` — medan boleh dicadangkan bila "Salin ke Workspace".
- `required: true` — wajib diisi sebelum boleh simpan.
- `noEdit: true` — tidak boleh diedit (contoh: BIL).

**12 workspace aktif**, dikelompokkan 5 kategori (sidebar):
- `kpi`, `pendaftaran` → KEMASUKAN
- `tambahMasa`, `tangguh` → PENGURUSAN PENGAJIAN
- `notis`, `pdpl`, `serahTesis`, `viva` → TESIS & PEPERIKSAAN
- `senat`, `graduan` → PENGIJAZAHAN
- `honorarium`, `jps` → PENTADBIRAN

### State (`state` object)
```js
{
  records: { wsId: [ {id, bil, ...kolum} ] },  // data jadual
  reminders: [ {id, wsId, title, date, priority, note, done} ],
  history: [ {id, wsId, recordId, field, oldVal, newVal, at} ], // 300 entri terakhir
  trash: [ {id, wsId, record, at} ],           // Arkib Padaman (boleh pulih)
  lastVisit: 'wsId',
  seq: number,
  settings: { ...defaultSettings() }           // boleh edit via tab Admin
}
```
Disimpan ke `localStorage` key `sps_demo_state_v1` via `save()` / `load()`.

### Tetapan (`defaultSettings()`)
Boleh diedit pengguna melalui **tab Admin** (tiada perlu ubah kod):
- `orgName`, `orgSub`, `operatorName`, `operatorRole`, `operatorInitials`
- `pageSize` (baris setiap halaman)
- `trashRetentionDays`, `confirmDelete`
- `showDemoBadge`
- `attentionThresholds`: `{ vivaSoonDays, upcomingDays }`
- `attentionRules`: toggle 5 peraturan "Perlu Perhatian"

`ensureSettings()` akan mengisi tetapan yang hilang untuk state lama.

## Ciri utama

1. **Dashboard** — 3 kad gradien (Perlu Perhatian / Peringatan / Acara 60 Hari) +
   2 kad metrik mini. Kad "Perlu Perhatian" bertukar **merah lembut berpulsa**
   (`g-alert`) bila ada item, kekal biru bila tiada.
2. **Sidebar kategori** — setiap kategori adalah kad dengan badge gradien, status
   "X perlu perhatian", dan kiraan rekod. Item ada titik kuning amaran.
3. **Jadual boleh edit** — klik **Edit** → ubah → semak (medan penting) → Simpan/Batal.
   Kolum BIL dikecilkan (38px, kelas `col-bil`).
4. **Format tarikh penuh** — `fmtDate()` menghasilkan `11 Ogos 2026` (bukan "Ogo").
   `fmtDateLong()` menambah nama hari untuk tooltip.
5. **Carian global** — merentas semua workspace ikut nama / no. matrik.
6. **Salin ke Workspace** — cadangkan medan `identity` sahaja; tarikh/kelulusan tidak disalin.
7. **Padam (boleh pulih)** — menu tindakan "Padamkan" menyimpan rekod ke **Arkib
   Padaman** (`state.trash`). Pulihkan di **Admin → Arkib Padaman**. Boleh padam kekal.
8. **Reminder manual** — dengan countdown (hari lagi / hari ini / lewat).
9. **Sejarah perubahan** — nilai sebelum/selepas bagi setiap edit.
10. **Admin & Tetapan** — edit identiti sistem, pengendali, paparan, peraturan
    perhatian, arkib; eksport JSON penuh; reset demo.
11. **Eksport CSV** per workspace (header asal dikekalkan).
12. **Responsif** — PC/tablet/telefon. Sidebar jadi menu ☰ pada ≤900px.
    `@media (pointer: coarse)` membesarkan sasaran sentuhan 44px.

## Prinsip reka bentuk (JANGAN langgar)

- **Header jadual kekal seperti Excel asal** — pengendali mesti kenali istilah asal
  (PDPL, SMP, PTSL, JPS, dll).
- **Wording yang dibina bersama jangan diubah** tanpa kebenaran.
- Gaya **financial dashboard**: sidebar navy gelap, latar off-white `#f7f8fc`, kad
  gradien biru/ungu/teal, IBM Plex Sans, ikon SVG outline (bukan emoji).
- Token warna/radius di `:root` — guna token, bukan nilai hardcoded.
- Data demo, **bukan** data sebenar. Label MOD DEMO sentiasa dipaparkan.

## Ujian

Tiada test runner formal. Ujian manual/otomatik menggunakan **jsdom** (dipasang
`--no-save`, tidak disimpan ke repo):

```bash
npm install jsdom --no-save
$env:NODE_PATH = (Resolve-Path node_modules).Path   # PowerShell
node <test-script>.js
```

Corak ujian yang digunakan: muat `index.html` + `app.js` dalam jsdom, tambah
`window.__t = {init, go, getState, ...}`, kemudian sahkan render dan interaksi.
**Jalankan sekurang-kurangnya semakan ini sebelum push:**
- `node --check app.js` (syntax)
- Semak kurungan CSS seimbang
- Konfirmasi tarikh penuh: `fmtDate('2026-08-11') === '11 Ogos 2026'`
- Header asal masih wujud dalam DOM
- Tambah rekod → baris masuk ke **bawah** (bukan atas)
- Padam → masuk Arkib → boleh pulih

## Aliran kerja go-live (pindah ke backend sebenar)

1. **Bersihkan data CSV** — buang baris/kolum kosong, betulkan format tarikh bercampur,
   asingkan medan gabungan (nama + no. matrik), sahkan format no. matrik/IC.
2. **Sediakan `Code.gs`** — Apps Script membaca/menulis Google Sheets. Setiap rekod
   mesti ada **ID kekal**; kemas kini ikut ID, bukan nombor baris.
3. **Hoskan frontend** — melalui GitHub Pages dan/atau Apps Script `doGet`
   (rujuk corak `i-nstrumen/CLAUDE.md`).
4. **Ganti `seedData()`** dengan panggilan API sebenar.

**Penting**: `git push` TIDAK auto-deploy backend Apps Script. Pengguna memindahkan
`Code.gs` ke editor Apps Script secara manual.

## Deployment

GitHub Pages monorepo. Untuk push:
```bash
git add i-siswazah/
git commit -m "i-siswazah: <ringkasan>"
git push origin main
```
URL selepas deploy: `https://imenhub-portal.github.io/i-siswazah/`

## Kredit

Developed by: **Hab Digital IMEN**
