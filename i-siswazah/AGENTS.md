# AGENTS.md — i-siswazah

Dokumen konteks penuh projek ini ada dalam **`CLAUDE.md`** dalam folder yang sama.

**Sila baca `CLAUDE.md` dahulu** sebelum mengubah apa-apa dalam folder ini. Ia
mengandungi:

- Status semasa & fasa kerja
- Struktur data (12 workspace, state, tetapan)
- Prinsip reka bentuk yang **tidak boleh dilanggar** (header asal, wording dibina bersama)
- Cara ujian
- Rancangan `Code.gs` (backend Apps Script yang akan dibina kemudian)
- Cara menyambung sesi antara PC/laptop

Ringkas: frontend/localStorage sahaja. `penyelia.js` melanjutkan `app.js` dan
`penyelia-seed.js` memuat union 34 penyelia/95 pelajar daripada tiga CSV
(tanpa nombor = fail 1, kemudian 2 dan 3): 91 aktif, 2 graduasi sejarah tanpa
tarikh tepat, 2 diberhentikan. Semua provenance/konflik dikekalkan; workspace
lain masih contoh. Baca status terkini CLAUDE.md (bahagian lama ialah sejarah).
Jana seed: `python import_penyelia.py`. Audit bebas: `python verify_sources.py`. Ujian: `regression.cjs` dan
`browser-regression.cjs` dengan dependency temp (lihat CLAUDE.md).
Browser lama mesti semak/gabung import; jangan reset atau menimpa konflik.
Sync v2 hanya boleh memadam rekod dengan ownership eksplisit; arkib nested
mesti mengekalkan jenis dan parent. Kerja sesi ini belum commit/push/deploy.
URL hosting sedia ada: `https://imenhub-portal.github.io/i-siswazah/`.
Backend `Code.gs` **belum dibina** dan akan dipindah ke Apps Script secara manual.

Developed by: **Hab Digital IMEN**
