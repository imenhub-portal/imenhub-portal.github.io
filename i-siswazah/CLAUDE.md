# i-siswazah — Sistem Pengurusan Siswazah (context untuk sesi baharu)

Baca fail ini dahulu sebelum membuat sebarang perubahan dalam folder ini. Ia ditulis
supaya sesi AI pada PC/laptop lain boleh menyambung kerja tanpa perlu meneliti
sejarah chat yang panjang.

## Status semasa — kerja tempatan 28 September 2026

### Kemas kini import tiga fail (menggantikan butiran import satu fail di bawah)

Pengguna mengesahkan fail tanpa nombor ialah fail 1. Ketiga-tiga fail digunakan:

| Sumber | Penyelia | Entri pelajar | Penyelia kosong |
|---|---:|---:|---:|
| `senarai_pelajar_Penyelia.csv` | 33 | 95 | 4 |
| `senarai_pelajar_Penyelia2.csv` | 33 | 94 | 4 |
| `senarai_pelajar_Penyelia3.csv` (semakan 18 Jun 2026) | 33 | 86 | 5 |

- Union: **34 penyelia unik, 95 pelajar unik, 5 penyelia kosong**. 275 kemunculan pelajar → 180 kemunculan bertindih. 22 pelajar mempunyai 24 perubahan medan bukan kosong (termasuk satu pertukaran parent); ini disimpan sebagai `conflicts`, bukan dibuang. Tiada nama tepat sama dengan matrik berlainan dalam audit; nama mirip tidak digunakan untuk menyatukan identiti.
- **91 aktif, 4 Rekod Lepas: 2 graduasi, 2 diberhentikan, 0 menarik diri.** Graduasi: P86119 (`GRADUAN 2024`) dan P119396 (`KONVO 2024`), kedua-duanya `tahunGraduasi: 2024`, tarikh kosong, `needsDateReview: true`. Diberhentikan: P107661 dan P108824, berdasarkan nota jelas. Bukti lengkap di `statusEvidence`.
- `classify()` hanya menerima teks afirmatif keseluruhan `GRADUAN <tahun>`, `KONVO <tahun>`, `TELAH BERGRADUASI [tahun]`, awalan `DIBERHENTIKAN`, atau `MENARIK DIRI`/`TELAH MENARIK DIRI [DARIPADA PENGAJIAN]`. `AKTIF SEMULA` ialah reaktivasi eksplisit. BAKAL GRADUAN, pending Senat, pembetulan, tangguh, tidak mendaftar/tidak dapat dikesan tidak membuktikan graduasi/penamatan. Tiada tarikh direka. Bukti final lama dikekalkan jika sumber kemudian tiada reaktivasi/final baru.
- Fail 3 mempunyai 15 kolum termasuk ringkasan PhD/MSc; fail 1/2 mempunyai 12. Semua baris/header/sel asal termasuk jumlah 47/39/86 dan ringkasan disimpan di `sources`; semua versi pelajar di `sourceVersions`. Snapshot/medan semasa menggunakan nilai bukan kosong terakhir mengikut 1→2→3; kosong tidak memadam nilai lama. Fail 1/2 tiada tarikh semakan eksplisit; urutan nombor ialah aturan tie-break yang didokumenkan, bukan tarikh andaian.
- **9 pelajar tiada dalam fail 3 tetap dikekalkan**, ditanda `missingFromLatest`. P168677 berpindah ke Dr. Muhammad Asif Ahmad Khushaini (K025867); satu listing sahaja dalam seed. ID penyelia fail 1 dikekalkan; penyelia baharu menggunakan hash identiti, bukan nombor BIL yang berubah.
- Import browser lama kekal opt-in. Preview memaparkan laporan dan cadangan status/parent. Lalai mengekalkan semua nilai/user lifecycle; checkbox setiap padanan tunggal membolehkan pengendali memilih versi gabungan CSV dan parent. Versi sebelum adopt disimpan dalam `importBackups`, ID browser dikekalkan. Padanan berganda tidak diadopt secara automatik. Rekod yang tiada daripada sumber tidak dipadam.
- Graduan sejarah tanpa tarikh boleh diedit/disimpan tanpa mencipta tarikh. Peralihan manual baharu ke graduasi masih wajib tarikh. Sync undated ialah no-op; selepas tarikh disahkan, ownership v2 digunakan dengan `_syncSourceIds`. Import tidak memadam rekod graduan sedia ada.
- Ujian terkini **lulus**: `python verify_sources.py` (parser CSV bebas, semua provenance/precedence/ID/status), `node regression.cjs` (termasuk konflik browser, backup adopt/transfer tanpa duplikasi), `node browser-regression.cjs` (Edge file:// 1440/768/390). Dua graduan masih memerlukan pengesahan tarikh sebenar; konflik 22 pelajar boleh disemak dalam drawer. Tiada blocker fail sumber. Belum commit/push/deploy.

### Catatan pelaksanaan satu fail terdahulu (sejarah)

- **Belum commit/push/deploy sesi ini.** Status remote tidak disahkan. Backend masih belum dibina.
- Workspace penyelia menggunakan CSV sebenar `senarai_pelajar_Penyelia.csv`: **33 penyelia, 95 pelajar, 4 penyelia tanpa pelajar**. Workspace lain masih seed contoh. MOD DEMO merujuk frontend/localStorage, bukan jaminan bahawa semua data ialah rekaan.
- `import_penyelia.py` menggunakan Python `csv.reader` (quoted multiline), menjana `penyelia-seed.js`. Jalankan `python import_penyelia.py` selepas menukar CSV. Fail JS statik dimuat sebelum `app.js`, serasi `file://` dan hosting statik.
- `penyelia.js` dimuat selepas `app.js`: identiti nama/ID, nested workspace, snapshot semester, drawer, merge, typed trash, sync ownership. Ia menggantikan beberapa fungsi legacy dalam `app.js`; baca kedua-duanya sebelum mengubah logik.
- Browser baharu menerima dataset penuh. Browser lama menggunakan **Semak Import CSV → Gabung — kekalkan suntingan**. Padanan ID/matrik mengekalkan semua nilai lama; sumber CSV penuh disimpan dalam `csvSources[hash]`, boleh dibaca dalam drawer. Tiada reset automatik. Konflik tidak diselesaikan automatik: pengendali perlu menyemak dan menyunting nilai yang dipilih. Padanan berganda juga dikekalkan.
- `provenance.values`, `sourceRows`, `sourceValue` menyimpan nilai asal, termasuk baris ringkasan PhD/MSc dan nota semester. Ringkasan bukan penyelia/pelajar baharu. STATUS TERKINI berasingan daripada status aliran kerja. Hanya dua nota jelas “Diberhentikan…” dipetakan; tiada tarikh atau graduasi direka daripada nota. “GRADUAN 2024”/“KONVO 2024” sejarah kekal teks untuk semakan manusia, tanpa tarikh graduasi andaian.
- Paparan nama + ID di bawah, tiada kurungan/kolum ID visual. ID masih dalam carian, edit berlabel dan eksport. Hanya suffix terminal `(Pdigits)` diekstrak; IC tidak dinormalisasi menjadi matrik.
- Program: **tepat dua pilihan sahaja** — **Doktor Falsafah** / **Sarjana Sains Kejuruteraan Mikro dan Nanoelektronik**. Semua kolum program di semua workspace ialah `type: 'select'` dengan `PROGRAM_OPTIONS` (app.js) dan dikuatkuasakan penyelia.js. `programValue()` memetakan legacy (SARJANA/Sarjana Sains/MSc → pilihan 2; KEDOKTORAN/PhD → pilihan 1); nilai lama disimpan dalam `_legacyProgram` bila migrasi. Importer (`import_penyelia.py`) juga memetakan dan akan gagal jelas jika program tidak dikenali.
- Papar sesi memilih snapshot sebenar. Tiga sesi terakhir yang direkodkan/aktif sehingga sesi pilihan dipaparkan; sesi lama kekal dalam dropdown. Tiada pengiraan semester global atau extrapolasi apabila snapshot tiada. Semester lama tanpa sesi kekal dalam drawer sehingga pengendali menetapkan sesi.
- Sync graduan v2 mempunyai `_syncOrigin: penyeliaPelajar:v2`, `_syncPelajarId`, `_syncParentId` dan ID stabil `graduan-<studentId>`. Penyelia utama ialah parent. Padanan manual/legacy tidak diambil alih. Pembalikan hanya membuang rekod milik v2; `_syncPelajarId` lama sahaja tidak membuktikan ownership dan rekod itu dipelihara.
- Padaman nested mempunyai `kind: student|supervisor`; pelajar dipulihkan ke parent. Jika parent dipadam, pulihkan parent dahulu. Padaman penyelia membawa semua pelajar dalam satu arkib. Padaman nested tidak membalikkan status graduasi.
- Ujian sebenar: `regression.cjs` (jsdom), `browser-regression.cjs` (Playwright + Edge), syntax JS dan `git diff --check` lulus. Kiraan CSV bebas 33/95; migrasi subset lama 8→95 tanpa menimpa nilai; merge ulangan tidak menggandakan. Edge `file://` diuji 1440/768/390, tiada page overflow, drawer muat, global search membuka rekod lepas, graduasi wajib tarikh, persistence dan tiada page errors. GitHub Pages sebenar belum diuji/deploy.
- Dependencies ujian dipasang **hanya** di `C:\Users\ruxxz\AppData\Local\Temp\opencode`; jalankan `$env:NODE_PATH='C:\Users\ruxxz\AppData\Local\Temp\opencode\node_modules'; node regression.cjs` dan `node browser-regression.cjs`. Skrip browser memerlukan Edge; menggunakan context ujian terasing.
- Had: storan masih satu browser/localStorage; konflik import perlu semakan manusia, bukan wizard field-by-field. Semasa sesi, fail pengguna untracked berubah daripada `Q2 17062026_...csv` kepada `senarai_pelajar_Penyelia2.csv` dan `senarai_pelajar_Penyelia3.csv`; agent tidak menulis/memadamnya dan tidak menggunakannya sebagai pengganti sumber yang diminta.

## Status terdahulu (rujukan sejarah; digantikan ringkasan di atas)

- **Fasa**: Prototaip demo **sudah go-live**. Frontend siap dan diuji; pengesahan
  workflow dengan pengendali sedang berjalan.
- **Live URL**: `https://imenhub-portal.github.io/i-siswazah/`
- **Committed**: `index.html`, `styles.css`, `app.js`, `CLAUDE.md`, dan 12 fail CSV.
- **Git**: selaras dengan `origin/main` (`imenhub-portal/imenhub-portal.github.io`).
- **Belum dibuat**: `Code.gs` (backend Apps Script) — rancangan penuh di bawah.
- **Data**: masih **data contoh kecil** dalam `app.js` + `localStorage`. Belum import
  penuh CSV ke jadual, belum cleansing data sebenar.

### Langkah seterusnya yang dijangkakan
1. Tunjuk demo kepada pengendali, kumpul maklum balas workflow.
2. Selaraskan header/kolum dengan fail Excel sebenar (kekal teks asal).
3. Bila workflow dipersetujui → bersihkan CSV → bina `Code.gs` → sambung backend.

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
  records: { wsId: [ {id, bil, locked?, lockedAt?, ...kolum} ] },  // data jadual
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
13. **Penyelia berbilang** — 10 kolum penyelia ditanda `multi: true`. Mod Edit
    memaparkan senarai boleh ulang (Tambah/Buang baris). Nilai disimpan sebagai
    teks dipisah `\n` (kekal serasi CSV).
14. **Auto-label penyelia (PENGIJAZAHAN)** — `senat` dan `graduan` ada `autoLabel: true`.
    `extractPenyeliaNames()` buang label lama → `formatPenyelia()` jana semula:
    1 → Penyelia Utama; 2 → Utama + Bersama; 3+ → Pengerusi Jawatankuasa Penyeliaan
    + Ahli (i, ii, iii…). Dijana semula **setiap kali simpan**.
    **Pengecualian penting:** medan `PENYELIA BERSAMA` (workspace Senarai Pelajar
    Mengikut Penyelia) guna `formatPenyeliaBersama()` — nama **sentiasa** dilabel
    `Penyelia Bersama:` (1 orang pun), bukan `Penyelia Utama`, kerana jadual itu
    sendiri sudah mewakili penyelia utama. `prepareStudentData()` memigrasi label lama.
15. **Kunci rekod (`locked`)** — layer kedua atas Edit. Menu 3-titik: **Kunci Rekod**
    / **Buka Kunci** (minta pengesahan). Rekod dikunci **dipindah ke panel
    collapsible "Rekod Dikunci"** di bawah jadual utama (default tertutup,
    `ui.lockedOpen[wsId]`). Jadual utama hanya rekod aktif. Baris terkunci amber
    lembut (`is-locked`). Klik Edit pada rekod terkunci → pengesahan buka kunci
    kemudian terus masuk mod Edit. Medan: `locked`, `lockedAt`. Terpakai semua 12
    workspace. Rekod terkunci **tidak** dikira "Perlu Perhatian".
16. **Pemeriksa PL & PD berasingan** (`pdpl`) — kolum `pemeriksa` lama dipecahkan
    kepada **`pemeriksaLuar` (PL = Pemeriksa Luar)** dan **`pemeriksaDalam`
    (PD = Pemeriksa Dalam)**, kedua-duanya `multi: true`. `migrateRecords()` +
    `splitPemeriksa()` memisahkan data lama secara automatik semasa `load()`
    (nama berlabel `(PD)` → Dalam; label universiti UM/USM/UiTM/INOR dll → Luar).
17. **Senarai Pelajar Mengikut Penyelia** (`penyeliaPelajar`) — workspace **bersarang**
    (kategori baharu PENYELIA-PELAJAR). Setiap penyelia = panel collapsible berisi
    pelajar aktif (default terbuka) + sub-collapsible **"Rekod Lepas"**. Toggle
    **Aktif/Semua**. Setiap pelajar ada butang **Edit** (drawer) dan **Ubah Status**
    (dropdown: Aktif / Telah Bergraduasi / Menarik Diri / Diberhentikan). Graduasi
    wajib tarikh; lain optional. `syncGraduan()` **auto-sync** pelajar bergraduasi ke
    workspace `graduan` (padanan `noPelajar` + `_syncPelajarId`); bila status ditukar
    balik ke Aktif, rekod graduan yang dijana auto **dibuang**.
18. **Semester Master** — `settings.semesterAktif = { sesi, semesterPengajian }`.
    **Had terkini**: semua dropdown sesi dalam worksheet (Papar sesi, sejarah
    semester, pemilih sesi) **tidak boleh melebihi sesi aktif**. `generateSemesterOptions()`
    menjana dari sesi aktif ke belakang (6 semester) sahaja — tiada advance ke hadapan.
    `allSessions()` juga ditapis pada sesi aktif. Pemilih **sesi aktif di Admin** guna
    `generateAdminSemesterOptions()` yang boleh jangkau sampai tahun sistem semasa
    (supaya pengendali boleh memajukan sesi dari semasa ke semasa). Tukar sesi aktif
    akan menetapkan semula sesi paparan. `sesiOrdinal()` digunakan untuk perbandingan.
19. **Panel Analisis dashboard** (`renderAnalytics()` + `computeAnalytics()`) — nisbah
    pelajar:penyelia, bar kesihatan status (aktif/graduasi/menarik diri/diberhentikan),
    kemasukan (tawaran/terima/tolak & peratusan), dan graduasi tahun semasa + pecahan
    mengikut tahun. Dikira daripada data Semua Pelajar (termasuk rekod lepas).
20. **Semester-Sesi (2 kotak)** — 10 kolum `semester` ditanda `semesterSesi: true`.
    Semasa edit, papar **2 kotak**: nombor semester + dropdown sesi (dihadkan sesi
    aktif). Format simpanan kanonikal `"Sem N · S/YYYY-YYYY"` (`formatSemesterSesi()`),
    dihurai oleh `parseSemesterSesi()`. Nilai lama tak dikenali dikekalkan sebagai
    `legacy` dan dipaparkan sebagai "Nilai lama:" tanpa hilang. `collectDraftEdits()`
    menggabungkan kedua-dua kotak.

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

## Rancangan `Code.gs` (backend Google Apps Script — AKAN DIBINA)

`Code.gs` **belum wujud** dalam folder ini. Bila sedia, ia akan dibina di sini dan
kemudian **dipindah secara manual** oleh pengguna ke editor Apps Script. `git push`
TIDAK auto-deploy backend. Rancangan penuh:

### Peranan `Code.gs`
Backend yang membaca/menulis **Google Sheets** (satu sheet = satu workspace, nama tab
ringkas seperti `pendaftaran`, `viva`, `pdpl`). Berfungsi sebagai API untuk `app.js`.

### Corak yang akan digunakan (selaras `i-nstrumen/`)
- `doGet(e)` — hidangkan `index.html` (fetch dari `raw.githubusercontent.com` supaya
  satu sumber; fallback salinan tempatan).
- `doPost(e)` — kemas kini data. Body JSON `{ action, wsId, payload }`.
- Sumber data **satu**: Google Sheet yang diikat pada projek Apps Script.

### API yang diperlukan oleh frontend
| Fungsi frontend sekarang | API `Code.gs` nanti |
|---|---|
| `seedData()` | `getAllData()` → pulangkan semua rekod + tetapan |
| `commitEdits()` | `updateRecord(wsId, id, changes)` |
| `addRow()` | `createRecord(wsId, record)` |
| `doDelete()` | `deleteRecord(wsId, id)` (soft delete ke tab `Trash`) |
| `restoreTrash()` | `restoreRecord(wsId, id)` |
| `save()` (tetapan) | `saveSettings(settings)` |

### Peraturan data WAJIB
- Setiap rekod mesti ada **ID kekal** (`id`, jana di backend). Kemas kini ikut **ID**,
  **bukan nombor baris** — baris boleh berubah bila disusun.
- `bil` ialah nombor paparan sahaja, bukan kunci.
- Simpan no. matrik / IC / telefon sebagai **teks** (elak sifar di hadapan hilang dan
  notasi saintifik seperti `8.80927E+11`).
- Tabar `history` dan `trash` sebagai tab berasingan untuk jejak audit.
- **Tiada rahsia dalam kod** — repo awam. Token/kata laluan disimpan dalam
  Script Properties, bukan dalam fail.

### Aliran kerja go-live
1. **Bersihkan data CSV** — buang baris/kolum kosong, betulkan format tarikh bercampur
   (contoh `15/092026`, tahun 2006/2002 di tengah rekod 2026), asingkan medan gabungan
   (nama + no. matrik), sahkan format no. matrik/IC.
2. **Sediakan `Code.gs`** mengikut spesifikasi di atas.
3. **Pindah Google Sheets + `Code.gs`** ke editor Apps Script, deploy sebagai Web App.
4. **Tukar sumber frontend** — ganti `seedData()`/`save()` dengan panggilan API sebenar
   (Frontend kekal GitHub Pages; backend panggil guna `fetch`).
5. **Uji** — tambah/edit/padam rekod melalui webapp, pastikan Google Sheets terkemas kini.

### Titik sambungan dalam kod sekarang
- `seedData()` — ganti dengan muat dari API.
- `save()` / `load()` — ganti `localStorage` dengan panggilan API.
- Semua mutasi (`commitEdits`, `addRow`, `doDelete`, `restoreTrash`) — tambah panggilan
  API di samping kemas kini state tempatan.
- `API_URL` belum wujud — akan ditambah sebagai pemalar di atas `app.js`.

## Deployment

GitHub Pages monorepo. Untuk push:
```bash
git add i-siswazah/
git commit -m "i-siswazah: <ringkasan>"
git push origin main
```
URL selepas deploy: `https://imenhub-portal.github.io/i-siswazah/`

Sama seperti `i-nstrumen/`, cara pindah kerja antara mesin ialah `git push` di sini,
`sit pull`/`git pull` di mesin lain. **Kod backend berasingan**: `Code.gs` mesti
dipindah ke editor Apps Script secara manual — `git push` tidak mendeploy-nya.

## Cara menyambung sesi di PC/laptop lain

1. `git pull` — fail frontend + `CLAUDE.md` + CSV akan muncul.
2. Minta AI **baca `CLAUDE.md` ini dahulu** sebelum mengubah apa-apa.
3. Baca "Status semasa" di atas untuk tahu fasa kerja terkini.
4. Ikut "Prinsip reka bentuk" — jangan ubah header/wording tanpa kebenaran pengguna.
5. Jalankan semakan dalam "Ujian" sebelum push semula.

## Log sesi (ringkas — tambah semasa sesi bermakna)

- **Sesi 1** — Bina frontend demo penuh: `index.html`, `styles.css`, `app.js`.
  12 workspace jadual boleh edit, dashboard, carian, salin workspace, reminder,
  sejarah, Arkib Padaman, tab Admin & Tetapan, responsif PC/tablet/telefon,
  format tarikh penuh Melayu. Go-live ke GitHub Pages. `Code.gs` sengaja belum dibina.
- **Sesi 2** — (a) Penyelia berbilang baris (10 kolum `multi: true`) dengan auto-label
  penyelia untuk pengijazahan (`senat`, `graduan`). (b) Ciri **kunci rekod**:
  rekod dikunci dipindah ke panel collapsible "Rekod Dikunci" di bawah jadual utama;
  jadual utama hanya rekod aktif; buka kunci perlu pengesahan; Edit pada rekod
  terkunci minta buka kunci dahulu.
- **Sesi 3** — (a) Pisah pemeriksa `pdpl` kepada **PL (Pemeriksa Luar)** dan
  **PD (Pemeriksa Dalam)**, kedua-duanya multi. Migrasi automatik data lama.
- **Sesi 4** — Workspace bersarang **"Senarai Pelajar Mengikut Penyelia"** (kategori
  PENYELIA-PELAJAR): panel penyelia collapsible, toggle Aktif/Semua, drawer edit
  pelajar, status pelajar (graduasi/menarik diri/diberhentikan), auto-sync graduan,
  semester master di Admin.
- **Sesi 5** — (a) Panel **Analisis Keseluruhan Siswazah** pada dashboard: nisbah
  pelajar:penyelia, bar kesihatan status, kadar kemasukan (tawaran/terima/tolak),
  graduasi tahun semasa. (b) Kolum **Semester-Sesi 2 kotak** pada 10 kolum semester.
  (c) Had sesi: dropdown tidak melebihi sesi aktif Admin.

### Baki dropdown & format
Permintaan yang belum dilaksanakan (lihat perbincangan terakhir):
- `BENTUK PENDAFTARAN` → dropdown (Sepenuh masa / Separuh masa)
- `PROGRAM PENGAJIAN` → **dua pilihan sahaja dilaksanakan & dikoherensikan untuk semua workspace** (28 September 2026). Hanya **Doktor Falsafah** dan **Sarjana Sains Kejuruteraan Mikro dan Nanoelektronik**.
- `SEMESTER PENGAJIAN` → dropdown format `1/2025-2026` (pilihan C), semua 9 kolum
- Enforce nilai dropdown sahaja; data lama akan dikemas kini manual kemudian

## Kredit

Developed by: **Hab Digital IMEN**
