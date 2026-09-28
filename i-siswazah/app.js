/* ============================================================
   SISTEM PENGURUSAN SISWAZAH - MOD DEMO (frontend sahaja)
   Tiada backend. Data disimpan dalam localStorage.
   ============================================================ */

const PROGRAM_OPTIONS = [
  'Doktor Falsafah',
  'Sarjana Sains',
  'Sarjana Sains Kejuruteraan Mikro dan Nanoelektronik'
];

/* ---------- STATUS PELAJAR (workspace penyeliaPelajar) ---------- */
const STATUS_PELAJAR_OPTIONS = [
  { value: '', label: 'Aktif' },
  { value: 'graduasi', label: 'Telah Bergraduasi' },
  { value: 'menarik_diri', label: 'Menarik Diri' },
  { value: 'diberhentikan', label: 'Diberhentikan' }
];

function statusPelajarLabel(v) {
  const o = STATUS_PELAJAR_OPTIONS.find(function (x) { return x.value === v; });
  return o ? o.label : 'Aktif';
}

/* Kolum: header = teks ASAL (dikekalkan), key = id dalaman, type, identity = boleh disalin */
const WORKSPACES = [
  { id: 'kpi', name: 'KPI Hal Ehwal Siswazah', group: 'KEMASUKAN', icon: 'target',
    desc: 'Pemantauan permohonan kemasukan dan kemas kini data',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'tarikh', header: 'TARIKH', type: 'date' },
      { key: 'semester', header: 'SEMESTER', type: 'text', identity: true },
      { key: 'nama', header: 'NAMA PELAJAR', type: 'text', identity: true, required: true },
      { key: 'cadanganPenyelia', header: 'CADANGAN PENYELIA', type: 'text', identity: true, multi: true },
      { key: 'noPelajar', header: 'NO. PELAJAR/N IC/Passport', type: 'text', identity: true },
      { key: 'program', header: 'PROGRAM PENGAJIAN', type: 'select', identity: true, options: PROGRAM_OPTIONS },
      { key: 'tajuk', header: 'TAJUK TESIS', type: 'textarea' },
      { key: 'bentuk', header: 'BENTUK PENDAFTARAN', type: 'text' },
      { key: 'tarikhTerima', header: 'TARIKH TERIMA PERMOHONAN', type: 'date' },
      { key: 'tarikhLulus', header: 'TARIKH LULUS PERMOHONAN', type: 'date' },
      { key: 'tarikhKemaskini', header: 'TARIKH KEMASKINI DATA DALAM join.UKM', type: 'date' }
    ] },
  { id: 'pendaftaran', name: 'Pendaftaran', group: 'KEMASUKAN', icon: 'userPlus',
    desc: 'Rekod pendaftaran pelajar baharu dan status tawaran',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'tarikh', header: 'TARIKH', type: 'date' },
      { key: 'nama', header: 'NAMA PELAJAR', type: 'text', identity: true, required: true },
      { key: 'noPelajar', header: 'NO. PELAJAR/N IC/Passport', type: 'text', identity: true },
      { key: 'program', header: 'PROGRAM PENGAJIAN', type: 'select', identity: true, options: PROGRAM_OPTIONS },
      { key: 'tajuk', header: 'TAJUK TESIS', type: 'textarea' },
      { key: 'penyelia', header: 'PENYELIA UTAMA', type: 'text', identity: true, multi: true },
      { key: 'bentuk', header: 'BENTUK PENDAFTARAN', type: 'text' },
      { key: 'tarikhMendaftar', header: 'TARIKH MENDAFTAR', type: 'date' },
      { key: 'semester', header: 'SEMESTER', type: 'text', identity: true },
      { key: 'statusTawaran', header: 'STATUS TAWARAN', type: 'select', options: ['Tawar', 'Terima Tawaran', 'Tolak Tawaran'] }
    ] },
  { id: 'tambahMasa', name: 'Permohonan Tambah Masa', group: 'PENGURUSAN PENGAJIAN', icon: 'clock',
    desc: 'Permohonan lanjutan tempoh pengajian',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'tarikhPermohonan', header: 'TARIKH PERMOHONAN', type: 'date' },
      { key: 'nama', header: 'NAMA PELAJAR', type: 'text', identity: true, required: true },
      { key: 'noMatrik', header: 'NO. MATRIK', type: 'text', identity: true },
      { key: 'penyelia', header: 'PENYELIA', type: 'text', identity: true, multi: true },
      { key: 'alasan', header: 'ALASAN PERMOHONAN', type: 'textarea' },
      { key: 'semester', header: 'SEMESTER PENGAJIAN', type: 'text', identity: true },
      { key: 'tarikhTindakan', header: 'TARIKH TINDAKAN DIAMBIL', type: 'date' },
      { key: 'tarikhLulus', header: 'TARIKH DILULUSKAN', type: 'date' }
    ] },
  { id: 'tangguh', name: 'Permohonan Tangguh Pengajian', group: 'PENGURUSAN PENGAJIAN', icon: 'pause',
    desc: 'Permohonan penangguhan pengajian',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'tarikhPermohonan', header: 'TARIKH PERMOHONAN', type: 'date' },
      { key: 'nama', header: 'NAMA PELAJAR', type: 'text', identity: true, required: true },
      { key: 'noMatrik', header: 'NO. MATRIK', type: 'text', identity: true },
      { key: 'penyelia', header: 'PENYELIA', type: 'text', identity: true, multi: true },
      { key: 'alasan', header: 'ALASAN PERMOHONAN', type: 'textarea' },
      { key: 'semester', header: 'SEMESTER PENGAJIAN', type: 'text', identity: true },
      { key: 'tarikhLulus', header: 'TARIKH DILULUSKAN', type: 'date' }
    ] },
  { id: 'notis', name: 'Notis Serah Tesis', group: 'TESIS & PEPERIKSAAN', icon: 'fileText',
    desc: 'Penerimaan notis serah tesis dan pencalonan',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'tarikhPermohonan', header: 'TARIKH PERMOHONAN DITERIMA', type: 'date' },
      { key: 'nama', header: 'PELAJAR', type: 'text', identity: true, required: true },
      { key: 'noMatrik', header: 'NO. MATRIK', type: 'text', identity: true },
      { key: 'penyelia', header: 'PENYELIA', type: 'text', identity: true, multi: true },
      { key: 'program', header: 'PROGRAM', type: 'text', identity: true },
      { key: 'semester', header: 'SEMESTER PENGAJIAN', type: 'text', identity: true },
      { key: 'tarikhTerima', header: 'TARIKH TERIMA', type: 'date' },
      { key: 'tarikhTerimaPencalonan', header: 'TARIKH TERIMA PENCALONAN', type: 'date' },
      { key: 'tarikhLulus', header: 'TARIKH LULUS', type: 'date' }
    ] },
  { id: 'pdpl', name: 'Pencalonan PDPL', group: 'TESIS & PEPERIKSAAN', icon: 'users',
    desc: 'Pencalonan pemeriksa dan penghantaran tesis kepada PDPL',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'tarikhPermohonan', header: 'TARIKH PERMOHONAN DITERIMA', type: 'date' },
      { key: 'nama', header: 'PELAJAR', type: 'text', identity: true, required: true },
      { key: 'noMatrik', header: 'NO. MATRIK', type: 'text', identity: true },
      { key: 'penyelia', header: 'PENYELIA', type: 'text', identity: true, multi: true },
      { key: 'semester', header: 'SEMESTER PENGAJIAN', type: 'text', identity: true },
      { key: 'tarikhTerima', header: 'TARIKH TERIMA PENCALONAN', type: 'date' },
      { key: 'tarikhLulus', header: 'TARIKH PENCALONAN DILULUSKAN', type: 'date' },
      { key: 'pemeriksaLuar', header: 'PEMERIKSA LUAR (PL)', type: 'text', multi: true },
      { key: 'pemeriksaDalam', header: 'PEMERIKSA DALAM (PD)', type: 'text', multi: true },
      { key: 'tarikhHantar', header: 'TARIKH HANTAR TESIS KEPADA PDPL', type: 'date' },
      { key: 'tarikhUpdate', header: 'TARIKH UPDATE DALAM SMP', type: 'date' }
    ] },
  { id: 'serahTesis', name: 'Serah Tesis untuk Pemeriksaan', group: 'TESIS & PEPERIKSAAN', icon: 'send',
    desc: 'Penghantaran tesis dan penerimaan laporan pemeriksa',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'tarikhPermohonan', header: 'TARIKH PERMOHONAN DITERIMA', type: 'date' },
      { key: 'nama', header: 'PELAJAR', type: 'text', identity: true, required: true },
      { key: 'noMatrik', header: 'NO. MATRIK', type: 'text', identity: true },
      { key: 'penyelia', header: 'PENYELIA', type: 'text', identity: true, multi: true },
      { key: 'semester', header: 'SEMESTER PENGAJIAN', type: 'text', identity: true },
      { key: 'tarikhHantar', header: 'TARIKH HANTAR TESIS KEPADA PDPL', type: 'date' },
      { key: 'laporanPL', header: 'TARIKH TERIMA LAPORAN DARI PL', type: 'date' },
      { key: 'laporanPD', header: 'TARIKH TERIMA LAPORAN DARI PD', type: 'date' },
      { key: 'tarikhUpdate', header: 'TARIKH UPDATE DALAM SMP', type: 'date' }
    ] },
  { id: 'viva', name: 'Peperiksaan Lisan', group: 'TESIS & PEPERIKSAAN', icon: 'mic',
    desc: 'Peperiksaan lisan (viva) dan pemantauan tempoh',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'nama', header: 'PELAJAR', type: 'text', identity: true, required: true },
      { key: 'noMatrik', header: 'NO. MATRIK', type: 'text', identity: true },
      { key: 'program', header: 'PROGRAM', type: 'text', identity: true },
      { key: 'penyelia', header: 'PENYELIA', type: 'text', identity: true, multi: true },
      { key: 'semester', header: 'SEMESTER PENGAJIAN', type: 'text', identity: true },
      { key: 'tarikhHantar', header: 'TARIKH HANTAR TESIS KEPADA PDPL', type: 'date' },
      { key: 'tarikhTerimaLaporan', header: 'TARIKH TERIMA LAPORAN DARIPADA PDPL', type: 'date' },
      { key: 'tarikhViva', header: 'TARIKH VIVA', type: 'date' },
      { key: 'tarikhUpdate', header: 'TARIKH UPDATE DALAM SMP', type: 'date' },
      { key: 'catatan', header: 'Catatan', type: 'textarea' }
    ] },
  { id: 'senat', name: 'Pengesahan Senat', group: 'PENGIJAZAHAN', icon: 'checkCircle',
    desc: 'Proses pengesahan senat selepas pembetulan tesis',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'nama', header: 'PELAJAR', type: 'text', identity: true, required: true },
      { key: 'program', header: 'PROGRAM', type: 'text', identity: true },
      { key: 'penyelia', header: 'PENYELIA', type: 'textarea', identity: true, multi: true, autoLabel: true },
      { key: 'semester', header: 'SEMESTER PENGAJIAN', type: 'text', identity: true },
      { key: 'terimaTesis', header: 'TARIKH TERIMA TESIS SELEPAS PEMBETULAN', type: 'date' },
      { key: 'permohonanPTSL', header: 'TARIKH PERMOHONAN PENGESAHAN PTSL', type: 'date' },
      { key: 'pengesahanPTSL', header: 'TARIKH PENGESAHAN PTSL', type: 'date' },
      { key: 'permohonanJPS', header: 'TARIKH PERMOHONAN PENGESAHAN AHLI JPS', type: 'date' },
      { key: 'kelulusanJPS', header: 'TARIKH KELULUSAN AHLI JPS', type: 'date' },
      { key: 'hantarSenat', header: 'TARIKH HANTAR KE SENAT', type: 'date' },
      { key: 'tarikhUpdate', header: 'TARIKH UPDATE DALAM SMP', type: 'date' }
    ] },
  { id: 'graduan', name: 'Bakal Graduan', group: 'PENGIJAZAHAN', icon: 'award',
    desc: 'Senarai bakal graduan mengikut rujukan senat',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'nama', header: 'PELAJAR', type: 'text', identity: true, required: true },
      { key: 'program', header: 'PROGRAM', type: 'text', identity: true },
      { key: 'penyelia', header: 'PENYELIA', type: 'textarea', identity: true, multi: true, autoLabel: true },
      { key: 'semester', header: 'SEMESTER PENGAJIAN', type: 'text', identity: true },
      { key: 'catatan', header: 'Catatan', type: 'text' }
    ] },
  { id: 'honorarium', name: 'Honorarium', group: 'PENTADBIRAN', icon: 'wallet',
    desc: 'Pemantauan honorarium pemeriksa',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'nama', header: 'PELAJAR', type: 'text', identity: true, required: true },
      { key: 'noMatrik', header: 'NO. MATRIK', type: 'text', identity: true },
      { key: 'tarikhViva', header: 'TARIKH VIVA', type: 'date' },
      { key: 'pdpl', header: 'PDPL', type: 'text' },
      { key: 'jumlah', header: 'JUMLAH HONORARIUM', type: 'text' },
      { key: 'tarikhSedia', header: 'TARIKH SEDIA HONORARIUM', type: 'date' },
      { key: 'tarikhPos', header: 'TARIKH POS HONORARIUM', type: 'date' }
    ] },
  { id: 'jps', name: 'Mesyuarat JPS', group: 'PENTADBIRAN', icon: 'calendar',
    desc: 'Pemantauan mesyuarat JPS dan penyediaan minit',
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'perkara', header: 'PERKARA', type: 'text', required: true },
      { key: 'tarikhMsyt', header: 'TARIKH MSYT', type: 'date' },
      { key: 'tarikhMinit', header: 'TARIKH SIAP MINIT', type: 'date' },
      { key: 'serahKP', header: 'TARIKH SERAH KPD KP UTK SEMAKAN DAN PENGESAHAN', type: 'date' },
      { key: 'serahTP', header: 'TARIKH SERAH KPD TP UTK PENGESAHAN DAN KELULUSAN', type: 'date' },
      { key: 'emelJkuasa', header: 'TARIKH EMEL KE JKUASA', type: 'date' }
    ] },
  { id: 'penyeliaPelajar', name: 'Senarai Pelajar Mengikut Penyelia', group: 'PENYELIA-PELAJAR', icon: 'userCheck',
    desc: 'Paparan pelajar bagi setiap penyelia — status aktif, graduasi dan penarikan diri',
    nested: true,
    columns: [
      { key: 'bil', header: 'BIL', type: 'number', noEdit: true },
      { key: 'nama', header: 'PELAJAR', type: 'text', required: true },
      { key: 'program', header: 'PROGRAM', type: 'select', options: PROGRAM_OPTIONS },
      { key: 'noPelajar', header: 'NO. PELAJAR', type: 'text' },
      { key: 'penyeliaBersama', header: 'PENYELIA BERSAMA', type: 'text', multi: true },
      { key: 'semesterPengajian', header: 'SEM. PENGAJIAN', type: 'text' },
      { key: 'status', header: 'STATUS', type: 'select', options: STATUS_PELAJAR_OPTIONS }
    ] }
];

const GROUP_ORDER = ['KEMASUKAN', 'PENGURUSAN PENGAJIAN', 'TESIS & PEPERIKSAAN', 'PENGIJAZAHAN', 'PENTADBIRAN', 'PENYELIA-PELAJAR'];

/* ---------- STATE ---------- */
const STORAGE_KEY = 'sps_demo_state_v1';

let state = {
  records: {},
  reminders: [],
  history: [],
  trash: [],
  lastVisit: null,
  seq: 1,
  settings: null
};

/* Tetapan sistem lalai — boleh diedit melalui tab Admin */
function defaultSettings() {
  return {
    orgName: 'Sistem Pengurusan Siswazah',
    orgSub: 'IMEN · UKM',
    operatorName: 'Pentadbir Sistem',
    operatorRole: 'Pengendali',
    operatorInitials: 'PN',
    pageSize: 25,
    trashRetentionDays: 30,
    confirmDelete: true,
    showDemoBadge: true,
    attentionThresholds: {
      vivaSoonDays: 14,
      upcomingDays: 60
    },
    attentionRules: {
      serahTesis: true,
      viva: true,
      notis: true,
      jps: true,
      senat: true
    },
    semesterAktif: {
      sesi: '2/2025-2026',
      semesterPengajian: 1
    }
  };
}

/* Pastikan settings wujud & lengkap (untuk state lama) */
function ensureSettings() {
  if (!state.settings) state.settings = defaultSettings();
  const d = defaultSettings();
  Object.keys(d).forEach(function (k) {
    if (state.settings[k] === undefined) state.settings[k] = d[k];
  });
  if (!state.settings.attentionThresholds) state.settings.attentionThresholds = d.attentionThresholds;
  if (!state.settings.attentionRules) state.settings.attentionRules = d.attentionRules;
  if (!state.settings.semesterAktif) state.settings.semesterAktif = d.semesterAktif;
  return state.settings;
}

let ui = {
  view: 'dashboard',
  editingRow: null,
  draft: null,
  sort: {},
  filter: {},
  query: {},
  page: {},
  lockedOpen: {},
  penyeliaOpen: {},
  lepasOpen: {},
  pageSize: 25
};

/* ---------- ICONS ---------- */
const ICONS = {
  grid: '<rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/>',
  target: '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1.5"/>',
  userPlus: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M19 8v6M22 11h-6"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  pause: '<circle cx="12" cy="12" r="9"/><path d="M10 9v6M14 9v6"/>',
  fileText: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M9 13h6M9 17h6"/>',
  users: '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>',
  send: '<path d="M22 2 11 13M22 2l-7 20-4-9-9-4z"/>',
  mic: '<path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2M12 19v4M8 23h8"/>',
  checkCircle: '<circle cx="12" cy="12" r="9"/><path d="M9 12l2 2 4-4"/>',
  award: '<circle cx="12" cy="8" r="6"/><path d="M15.5 13.5 17 22l-5-3-5 3 1.5-8.5"/>',
  wallet: '<path d="M20 12V8H6a2 2 0 0 1 0-4h12v4"/><path d="M4 6v12a2 2 0 0 0 2 2h14v-4"/><path d="M18 12a2 2 0 0 0 0 4h4v-4z"/>',
  calendar: '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
  edit: '<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.12 2.12 0 0 1 3 3L12 15l-4 1 1-4z"/>',
  save: '<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><path d="M17 21v-8H7v8M7 3v5h8"/>',
  x: '<path d="M18 6 6 18M6 6l12 12"/>',
  copy: '<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
  bell: '<path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.7 21a2 2 0 0 1-3.4 0"/>',
  trash: '<path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>',
  history: '<path d="M3 3v5h5"/><path d="M3.05 13A9 9 0 1 0 6 5.3L3 8"/><path d="M12 7v5l4 2"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  chevronLeft: '<path d="M15 18l-6-6 6-6"/>',
  chevronRight: '<path d="M9 18l6-6-6-6"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 16v-4M12 8h.01"/>',
  alert: '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
  external: '<path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><path d="M15 3h6v6M10 14 21 3"/>',
  download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5M12 15V3"/>',
  more: '<circle cx="12" cy="5" r="1.6"/><circle cx="12" cy="12" r="1.6"/><circle cx="12" cy="19" r="1.6"/>',
  settings: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>',
  restore: '<path d="M3 3v5h5"/><path d="M3.05 13A9 9 0 1 0 6 5.3L3 8"/><path d="M12 8v4l3 2"/>',
  sliders: '<path d="M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3M1 14h6M9 8h6M17 16h6"/>',
  shield: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>',
  database: '<ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/>',
  lock: '<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
  lockOpen: '<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 9.9-1"/>',
  chevronDown: '<path d="M6 9l6 6 6-6"/>',
  userCheck: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M16 11l2 2 4-4"/>'
};

function icon(name, size) {
  const s = size || 18;
  return '<svg width="' + s + '" height="' + s + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' + (ICONS[name] || '') + '</svg>';
}

/* ---------- HELPERS ---------- */
function ws(id) { return WORKSPACES.find(function (w) { return w.id === id; }); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]; }); }
function uid(p) { return (p || '') + Math.random().toString(36).slice(2, 9) + (state.seq++); }
function todayISO() { const d = new Date(); return d.toISOString().slice(0, 10); }

function parseDate(v) {
  if (!v) return null;
  const s = String(v).trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}

function daysFromNow(v) {
  const d = parseDate(v);
  if (!d) return null;
  const t = new Date(); t.setHours(0, 0, 0, 0);
  d.setHours(0, 0, 0, 0);
  return Math.round((d - t) / 86400000);
}

function fmtDate(v) {
  const d = parseDate(v);
  if (!d) return v || '';
  const bulan = ['Januari', 'Februari', 'Mac', 'April', 'Mei', 'Jun', 'Julai', 'Ogos', 'September', 'Oktober', 'November', 'Disember'];
  const hh = String(d.getDate()).padStart(2, '0');
  const bl = bulan[d.getMonth()];
  const th = d.getFullYear();
  return hh + ' ' + bl + ' ' + th;
}

/* Format penuh dengan hari (untuk tooltip/paparan panjang) */
function fmtDateLong(v) {
  const d = parseDate(v);
  if (!d) return v || '';
  const hari = ['Ahad', 'Isnin', 'Selasa', 'Rabu', 'Khamis', 'Jumaat', 'Sabtu'];
  const bulan = ['Januari', 'Februari', 'Mac', 'April', 'Mei', 'Jun', 'Julai', 'Ogos', 'September', 'Oktober', 'November', 'Disember'];
  return hari[d.getDay()] + ', ' + String(d.getDate()).padStart(2, '0') + ' ' + bulan[d.getMonth()] + ' ' + d.getFullYear();
}

function save() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }
  catch (e) { console.warn('Gagal simpan', e); }
}

function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    if (parsed && parsed.records) { state = parsed; migrateRecords(); return true; }
  } catch (e) { console.warn('Gagal muat', e); }
  return false;
}

/* Migrasi data: pecahkan kolum lama 'pemeriksa' (pencalonan pdpl) kepada
   pemeriksaLuar (PL) dan pemeriksaDalam (PD). Idempotent. */
function migrateRecords() {
  const recs = (state.records && state.records.pdpl) || [];
  let changed = false;
  recs.forEach(function (r) {
    if (r.pemeriksa !== undefined) {
      const split = splitPemeriksa(r.pemeriksa);
      if (r.pemeriksaLuar === undefined) r.pemeriksaLuar = split.luar;
      if (r.pemeriksaDalam === undefined) r.pemeriksaDalam = split.dalam;
      delete r.pemeriksa;
      changed = true;
    }
  });
  if (changed) save();
}

/* Pisahkan teks pemeriksa lama: nama berlabel (PD) atau tanpa label universiti = Pemeriksa Dalam;
   nama berlabel universiti (UM/USM/UiTM/INOR dll) = Pemeriksa Luar. */
function splitPemeriksa(val) {
  const luar = [], dalam = [];
  String(val || '').split('\n').map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (line) {
    if (/\(\s*PD\s*\)/i.test(line)) {
      dalam.push(line.replace(/\s*\(\s*PD\s*\)\s*$/i, '').trim());
    } else if (/\((UM|USM|UiTM|INOR|UPM|UTM|UNITEN|UIA|IIUM|UNIMAS|UMS|USM|UTHM|UMPSA|Uni[^)]*)\)/i.test(line)) {
      luar.push(line.trim());
    } else {
      dalam.push(line.trim());
    }
  });
  return { luar: luar.join('\n'), dalam: dalam.join('\n') };
}

function toast(msg, type) {
  const el = document.createElement('div');
  el.className = 'toast' + (type ? ' is-' + type : '');
  const ic = type === 'success' ? 'checkCircle' : type === 'error' ? 'alert' : type === 'warn' ? 'alert' : 'info';
  el.innerHTML = icon(ic, 16) + '<span>' + esc(msg) + '</span>';
  document.getElementById('toastWrap').appendChild(el);
  setTimeout(function () { el.style.opacity = '0'; el.style.transition = 'opacity .3s'; setTimeout(function () { el.remove(); }, 320); }, 2600);
}

function nextId(wsId) {
  const recs = state.records[wsId] || [];
  let max = 0;
  recs.forEach(function (r) { const n = parseInt(r.bil, 10); if (!isNaN(n) && n > max) max = n; });
  return max + 1;
}

/* ---------- PENYELIA (senarai & auto-label) ---------- */
/* Ekstrak nama penyelia sahaja daripada teks tersimpan.
   Membuang label seperti "Penyelia Utama", "Pengerusi JK Penyeliaan", "i)" dsb. */
function extractPenyeliaNames(val) {
  if (!val) return [];
  const labels = /^(penyelia utama|penyelia bersama|pengerusi (jk|jawatankuasa) penyeliaan|ahli (jk|jawatankuasa) penyeliaan|penyelia)\s*:?\s*$/i;
  let s = String(val).replace(/\\n/g, '\n');
  let out = [];
  s.split('\n').forEach(function (line) {
    let t = line.trim();
    if (!t) return;
    if (labels.test(t)) return;
    t = t.replace(/^(i{1,3}|iv|v|vi{0,3}|ix|x)[.)]\s*/i, '');
    t = t.replace(/^[-•*]\s*/, '');
    t = t.replace(/^(dr|prof|ts|pm|ym|ir)\.?\s*$/i, '');
    t = t.trim();
    if (t) out.push(t);
  });
  return out;
}

/* Jana semula label penyelia mengikut bilangan:
   1 → Penyelia Utama
   2 → Penyelia Utama + Penyelia Bersama
   3+ → Pengerusi Jawatankuasa Penyeliaan + Ahli (i, ii, iii…) */
function formatPenyelia(names) {
  const list = (names || []).map(function (n) { return String(n).trim(); }).filter(Boolean);
  if (!list.length) return '';
  if (list.length === 1) {
    return 'Penyelia Utama\n' + list[0];
  }
  if (list.length === 2) {
    return 'Penyelia Utama\n' + list[0] + '\n\nPenyelia Bersama\n' + list[1];
  }
  const roman = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'];
  let s = 'Pengerusi Jawatankuasa Penyeliaan\n' + list[0] + '\n\nAhli Jawatankuasa Penyeliaan';
  for (let i = 1; i < list.length; i++) {
    s += '\n' + (roman[i - 1] || (i + 1)) + ') ' + list[i];
  }
  return s;
}

/* Baca nilai senarai penyelia daripada DOM untuk kolum multi */
function readMultiValue(key) {
  const wrap = document.querySelector('.penyelia-list[data-multi-key="' + key + '"]');
  if (!wrap) return null;
  const inputs = wrap.querySelectorAll('.penyelia-input');
  const names = [];
  inputs.forEach(function (inp) { const v = inp.value.trim(); if (v) names.push(v); });
  return names;
}

/* ---------- SEED (DATA CONTOH KECIL) ---------- */
function seedData() {
  state.records = {
    pendaftaran: [
      { id: uid('r'), bil: 1, tarikh: '', nama: 'SITI NASUHA BINTI MUSTAFFA', noPelajar: 'P170334', program: 'Doktor Falsafah', tajuk: '', penyelia: '', bentuk: '', tarikhMendaftar: '', semester: '2/2025-2026', statusTawaran: 'Tawar' },
      { id: uid('r'), bil: 2, tarikh: '', nama: 'NURUL LIYANA BINTI ABDULLAH', noPelajar: 'P170858', program: 'Doktor Falsafah', tajuk: '', penyelia: '', bentuk: '', tarikhMendaftar: '', semester: '', statusTawaran: 'Tawar' },
      { id: uid('r'), bil: 3, tarikh: '', nama: 'VASANTHAN A/L SUBRAMANIAM', noPelajar: 'P172682', program: 'Doktor Falsafah', tajuk: '', penyelia: '', bentuk: '', tarikhMendaftar: '', semester: '', statusTawaran: 'Terima Tawaran' }
    ],
    viva: [
      { id: uid('r'), bil: 1, nama: 'Syazwani Izrah binti Badrudin', noMatrik: 'P130001', program: 'PhD', penyelia: 'Dr. Rhonira Latif', semester: '', tarikhHantar: '', tarikhTerimaLaporan: '', tarikhViva: '2026-10-30', tarikhUpdate: '', catatan: 'pembetulan' },
      { id: uid('r'), bil: 2, nama: 'Mohd Faris Musawwi bin Ruslan', noMatrik: 'P125395', program: 'MSc', penyelia: 'Dr. Abdul Rahman Mohmad', semester: '', tarikhHantar: '', tarikhTerimaLaporan: '', tarikhViva: '2026-10-27', tarikhUpdate: '', catatan: 'akan masuk senat sept' },
      { id: uid('r'), bil: 3, nama: 'Aimi Nabila Efandi', noMatrik: 'P138067', program: 'MSc', penyelia: 'Prof. Dr. Mohd Yusri Abd. Rahman', semester: '', tarikhHantar: '2026-04-01', tarikhTerimaLaporan: '2026-05-06', tarikhViva: '2026-12-12', tarikhUpdate: '', catatan: 'akan masuk senat sept' }
    ],
    notis: [
      { id: uid('r'), bil: 1, tarikhPermohonan: '2026-08-01', nama: 'Syazwani Izrah bin Badrudin', noMatrik: '130001', penyelia: 'Dr. Rhonira Latif', program: 'PhD', semester: '', tarikhTerima: '2026-08-01', tarikhTerimaPencalonan: '2026-08-10', tarikhLulus: '2026-08-11' },
      { id: uid('r'), bil: 2, tarikhPermohonan: '', nama: 'Muhammad Faris Musawwi bin Ruslan', noMatrik: 'P125395', penyelia: 'Dr. Abdul Rahman Mohmad', program: 'MSc', semester: '', tarikhTerima: '', tarikhTerimaPencalonan: '', tarikhLulus: '' },
      { id: uid('r'), bil: 3, tarikhPermohonan: 'belum serah notis', nama: 'Rawhan Haque', noMatrik: 'P127643', penyelia: 'Dr. Ooi Poh Choon', program: 'PhD', semester: '', tarikhTerima: '', tarikhTerimaPencalonan: '', tarikhLulus: '' }
    ],
    serahTesis: [
      { id: uid('r'), bil: 1, tarikhPermohonan: '', nama: 'Syazwani Izrah bin Badrudin', noMatrik: '130001', penyelia: 'Dr. Rhonira Latif', semester: '', tarikhHantar: '', laporanPL: '', laporanPD: '', tarikhUpdate: '' },
      { id: uid('r'), bil: 2, tarikhPermohonan: '', nama: 'Muhammad Faris Musawwi bin Ruslan', noMatrik: 'P125395', penyelia: 'Dr. Abdul Rahman Mohmad', semester: '', tarikhHantar: '2026-03-25', laporanPL: '', laporanPD: '', tarikhUpdate: '2026-08-10' },
      { id: uid('r'), bil: 3, tarikhPermohonan: '2026-07-15', nama: 'Muhammad Khairul Ashraf bin Azmi', noMatrik: 'P154124', penyelia: 'Dr. Ahmad Razif Muhammad\nProf. Madya Ir. Dr. Abang Annuar Ehsan', semester: '', tarikhHantar: '2026-07-31', laporanPL: '2026-09-03', laporanPD: '', tarikhUpdate: '' }
    ],
    jps: [
      { id: uid('r'), bil: 1, perkara: 'MSYT JPS BIL. 1/2026', tarikhMsyt: '', tarikhMinit: '', serahKP: '', serahTP: '', emelJkuasa: '' },
      { id: uid('r'), bil: 2, perkara: 'MSYT JPS BIL. 2/2026', tarikhMsyt: '', tarikhMinit: '', serahKP: '', serahTP: '', emelJkuasa: '' },
      { id: uid('r'), bil: 3, perkara: 'MSYT JPS BIL. 6/2026', tarikhMsyt: '2026-08-11', tarikhMinit: '', serahKP: '', serahTP: '', emelJkuasa: '' }
    ],
    tangguh: [
      { id: uid('r'), bil: 1, tarikhPermohonan: '2026-02-04', nama: 'Hakim Izani', noMatrik: 'P117864', penyelia: 'Dr. Ahmad Razif Muhammad', alasan: 'Kewangan', semester: '2/2025-2026', tarikhLulus: '2026-02-12' },
      { id: uid('r'), bil: 2, tarikhPermohonan: '', nama: 'Syahirah Tombel', noMatrik: '', penyelia: 'Dr. Mohd Zulhakimi Ab Razak', alasan: 'Kewangan', semester: '2/2025-2026', tarikhLulus: '' }
    ],
    kpi: [
      { id: uid('r'), bil: 1, tarikh: '', semester: 'S1/2026-2027', nama: 'Vasanthan A/L Subramaniam', cadanganPenyelia: 'Dr. Maria binti Abu Bakar', noPelajar: '', program: 'Doktor Falsafah', tajuk: '', bentuk: 'Separuh masa', tarikhTerima: '', tarikhLulus: '2026-04-15', tarikhKemaskini: '' },
      { id: uid('r'), bil: 2, tarikh: '', semester: 'S1/2026-2027', nama: 'Mohamad Shafiq bin Mohamad Radzuan', cadanganPenyelia: 'Dr. Atiqah binti Mohd Afdzaluddin', noPelajar: '860927565833', program: 'Sarjana Sains Kejuruteraan Mikro dan Nanoelektronik', tajuk: 'Fundamental Study and Powder Metallurgy Manufacturing of Cu6Sn5 Intermetallic Powder for Electronic Interconnection Applications.', bentuk: 'Sepenuh masa', tarikhTerima: '2026-05-26', tarikhLulus: '2026-06-10', tarikhKemaskini: '2026-06-10' }
    ],
    tambahMasa: [
      { id: uid('r'), bil: 1, tarikhPermohonan: '', nama: 'Nur Ariena Hanis binti Mohd Nor', noMatrik: 'P127641', penyelia: 'Prof. Madya Dr. Siow Kim Shyong', alasan: 'memerlukan tempoh masa yang lebih bagi menyiapkan tesis dan menerbitkan artikel jurnal', semester: 'Semester 2 Sesi 2025/2026', tarikhTindakan: '', tarikhLulus: '' },
      { id: uid('r'), bil: 2, tarikhPermohonan: '', nama: 'Syahirah Hinayadullah', noMatrik: '', penyelia: 'Dr. Abdul Rahman Mohmad', alasan: '', semester: '', tarikhTindakan: '', tarikhLulus: '' }
    ],
    pdpl: [
      { id: uid('r'), bil: 1, tarikhPermohonan: '', nama: 'Muhammad Faris Musawwi bin Ruslan', noMatrik: 'P125395', penyelia: 'Dr. Abdul Rahman Mohmad', semester: '', tarikhTerima: '2026-02-12', tarikhLulus: '', pemeriksaLuar: 'Prof. Dr. Goh Boon Tong (UM)', pemeriksaDalam: 'Dr. Dilla Duryha Berhanuddin', tarikhHantar: '', tarikhUpdate: '2026-03-05' },
      { id: uid('r'), bil: 2, tarikhPermohonan: '', nama: 'Mohd Erwan bin Basiron', noMatrik: 'P130002', penyelia: 'Prof. Dr. Azman bin Jalar', semester: '', tarikhTerima: '', tarikhLulus: '', pemeriksaLuar: 'Prof. Madya Dr. Abdullah Aziz bin Saad (USM)', pemeriksaDalam: 'Dr. Muhammad Aniq Shazni bin Mohammad Haniff', tarikhHantar: '', tarikhUpdate: '' }
    ],
    senat: [
      { id: uid('r'), bil: 1, nama: 'Nur Nazhifah binti Yusoff (P100212)', program: 'Sarjana Sains', penyelia: 'Penyelia Utama\nDr. Norhayati binti Abu Bakar', semester: '13', terimaTesis: '', permohonanPTSL: '', pengesahanPTSL: '', permohonanJPS: '', kelulusanJPS: '', hantarSenat: '', tarikhUpdate: '' },
      { id: uid('r'), bil: 2, nama: 'Nur Adliha binti Abdullah (P94705)', program: 'Doktor Falsafah', penyelia: 'Pengerusi JK Penyeliaan\nYM Tengku Hasnan Tengku Abdul Aziz', semester: '13', terimaTesis: '', permohonanPTSL: '', pengesahanPTSL: '', permohonanJPS: '', kelulusanJPS: '', hantarSenat: '', tarikhUpdate: '' }
    ],
    graduan: [
      { id: uid('r'), bil: 1, nama: 'Nur Nazhifah binti Yusoff (P100212)', program: 'Sarjana Sains', penyelia: 'Penyelia Utama\nDr. Norhayati binti Abu Bakar', semester: '13', catatan: 'Senat 527' },
      { id: uid('r'), bil: 2, nama: 'Nur Adliha binti Abdullah (P94705)', program: 'Doktor Falsafah', penyelia: 'Pengerusi JK Penyeliaan\nYM Tengku Hasnan Tengku Abdul Aziz', semester: '13', catatan: 'Senat 527' }
    ],
    honorarium: [
      { id: uid('r'), bil: 1, nama: '', noMatrik: '', tarikhViva: '', pdpl: '', jumlah: '', tarikhSedia: '', tarikhPos: '' },
      { id: uid('r'), bil: 2, nama: '', noMatrik: '', tarikhViva: '', pdpl: '', jumlah: '', tarikhSedia: '', tarikhPos: '' }
    ],
    penyeliaPelajar: [
      { id: uid('r'), bil: 1, namaPenyelia: 'PROF. DR. AZMAN JALAR @JALIL (K007353)', pelajar: [
        { id: uid('r'), bil: 1, nama: 'LIM EE MAY', program: 'Sarjana Sains', noPelajar: 'P137670', penyeliaBersama: 'DR. MARIA BINTI ABU BAKAR', statusKhas: '', tarikhStatus: '', semesterPengajian: 6, sejarahSemester: [
          { sesi: '2/2022-2023', semesterPengajian: '6', catatan: '' },
          { sesi: '1/2023-2024', semesterPengajian: '', catatan: 'akan hantar notis dalam masa terdekat' },
          { sesi: '2/2023-2024', semesterPengajian: '6', catatan: 'Dalam tempoh pembetulan tesis.' }
        ] },
        { id: uid('r'), bil: 2, nama: 'BALOGUN BASHIR TEMITOPE', program: 'Doktor Falsafah', noPelajar: 'P117630', penyeliaBersama: 'DR. MARIA BINTI ABU BAKAR\nDR. ATIQAH BINTI MOHD AFDZALUDDIN', statusKhas: '', tarikhStatus: '', semesterPengajian: 7, sejarahSemester: [] },
        { id: uid('r'), bil: 3, nama: 'MOHD ERWAN BIN BASIRON', program: 'Sarjana Sains', noPelajar: 'P130002', penyeliaBersama: 'DR. MARIA BINTI ABU BAKAR', statusKhas: '', tarikhStatus: '', semesterPengajian: 6, sejarahSemester: [] }
      ] },
      { id: uid('r'), bil: 2, namaPenyelia: 'PROF. DR. AZRUL AZLAN HAMZAH (K014762)', pelajar: [
        { id: uid('r'), bil: 1, nama: 'ROHARSYAFINAZ BINTI ROSLAN', program: 'Sarjana Sains', noPelajar: 'P121441', penyeliaBersama: 'DR. AHMAD GHADAFI ISMAIL, PROF. MADYA DR. P. SUSTHITHA MENON', statusKhas: 'graduasi', tarikhStatus: '2024-11-15', semesterPengajian: 8, sejarahSemester: [] },
        { id: uid('r'), bil: 2, nama: 'MANAL BINTI AMMAR', program: 'Sarjana Sains', noPelajar: 'P152989', penyeliaBersama: 'TIADA', statusKhas: '', tarikhStatus: '', semesterPengajian: 4, sejarahSemester: [] },
        { id: uid('r'), bil: 3, nama: 'ARIFAH SYAHIRAH BINTI ABDUL RAHMAN', program: 'Doktor Falsafah', noPelajar: 'P153583', penyeliaBersama: 'TIADA', statusKhas: '', tarikhStatus: '', semesterPengajian: 3, sejarahSemester: [] }
      ] },
      { id: uid('r'), bil: 3, namaPenyelia: 'PROF. DR. DEE CHANG FU (K013525)', pelajar: [
        { id: uid('r'), bil: 1, nama: 'MOHAMAD NIZAR HADI BIN MOHAMAD NASSIR', program: 'Doktor Falsafah', noPelajar: 'P109012', penyeliaBersama: 'PROF. DR. AZRUL AZLAN BIN HAMZAH\nDR. AHMAD GHADAFI BIN ISMAIL', statusKhas: '', tarikhStatus: '', semesterPengajian: 12, sejarahSemester: [] },
        { id: uid('r'), bil: 2, nama: 'MUHAMAD ARIF BIN SHAHARIAH', program: 'Doktor Falsafah', noPelajar: 'P154207', penyeliaBersama: 'DR. NG PEI YUEN', statusKhas: 'diberhentikan', tarikhStatus: '2025-09-01', semesterPengajian: 4, sejarahSemester: [] }
      ] }
    ]
  };
  state.reminders = [
    { id: uid('rm'), wsId: 'viva', recordId: null, title: 'Susulan laporan pemeriksa', ref: 'Mohd Faris Musawwi bin Ruslan', date: todayISO(), priority: 'med', note: 'Hubungi pemeriksa untuk susulan laporan', done: false, demo: true }
  ];
  state.history = [];
  state.trash = [];
  state.lastVisit = null;
  state.seq = 1;
  state.settings = defaultSettings();
  save();
}

/* ---------- ROUTER ---------- */
function go(viewId) {
  ui.view = viewId;
  ui.editingRow = null;
  ui.draft = null;
  if (viewId !== 'dashboard' && viewId !== 'admin') { state.lastVisit = viewId; save(); }
  render();
  closeSidebar();
  window.scrollTo(0, 0);
}

/* ---------- SIDEBAR ---------- */
const CATEGORY_META = {
  'KEMASUKAN': { icon: 'userPlus', grad: 'g-blue' },
  'PENGURUSAN PENGAJIAN': { icon: 'clock', grad: 'g-purple' },
  'TESIS & PEPERIKSAAN': { icon: 'fileText', grad: 'g-teal' },
  'PENGIJAZAHAN': { icon: 'award', grad: 'g-amber' },
  'PENTADBIRAN': { icon: 'wallet', grad: 'g-slate' },
  'PENYELIA-PELAJAR': { icon: 'userCheck', grad: 'g-indigo' }
};

function renderSidebar() {
  const nav = document.getElementById('sidebarNav');
  let html = '';
  const dashActive = ui.view === 'dashboard' ? ' is-active' : '';
  html += '<button class="nav-item' + dashActive + '" data-go="dashboard">' + icon('grid', 17) + '<span class="nav-item__label">Dashboard</span></button>';

  GROUP_ORDER.forEach(function (group) {
    const items = WORKSPACES.filter(function (w) { return w.group === group; });
    if (!items.length) return;
    const meta = CATEGORY_META[group] || { icon: 'grid', grad: 'g-blue' };
    const catCount = items.reduce(function (s, w) { return s + (state.records[w.id] || []).length; }, 0);
    const catAttention = items.reduce(function (s, w) { return s + countAttention(w.id); }, 0);

    html += '<div class="nav-cat">';
    html += '<div class="nav-cat__head">';
    html += '<span class="nav-cat__badge ' + meta.grad + '">' + icon(meta.icon, 16) + '</span>';
    html += '<div class="nav-cat__meta"><div class="nav-cat__name">' + esc(group) + '</div>';
    html += '<div class="nav-cat__stat">' + (catAttention > 0 ? '<b>' + catAttention + '</b> perlu perhatian' : 'Tiada perhatian') + '</div></div>';
    html += '<span class="nav-cat__count">' + catCount + '</span>';
    html += '</div>';

    items.forEach(function (w) {
      const cnt = (state.records[w.id] || []).length;
      const active = ui.view === w.id ? ' is-active' : '';
      const att = countAttention(w.id);
      html += '<button class="nav-item' + active + '" data-go="' + w.id + '">' +
        '<span class="nav-item__icon">' + icon(w.icon, 17) + '</span>' +
        '<span class="nav-item__label">' + esc(w.name) + '</span>' +
        (att > 0 ? '<span class="nav-item__dot warn" title="' + att + ' perlu perhatian"></span>' : '') +
        '<span class="nav-item__count">' + cnt + '</span></button>';
    });
    html += '</div>';
  });

  /* Sistem — tetapan admin */
  const adminActive = ui.view === 'admin' ? ' is-active' : '';
  html += '<div class="nav-cat">';
  html += '<div class="nav-cat__head">';
  html += '<span class="nav-cat__badge g-slate">' + icon('settings', 16) + '</span>';
  html += '<div class="nav-cat__meta"><div class="nav-cat__name">SISTEM</div>';
  html += '<div class="nav-cat__stat">Konfigurasi aplikasi</div></div>';
  html += '</div>';
  const trashCount = (state.trash || []).length;
  html += '<button class="nav-item' + adminActive + '" data-go="admin">' +
    '<span class="nav-item__icon">' + icon('settings', 17) + '</span>' +
    '<span class="nav-item__label">Admin & Tetapan</span>' +
    (trashCount > 0 ? '<span class="nav-item__count">' + trashCount + '</span>' : '') + '</button>';
  html += '</div>';

  nav.innerHTML = html;
  nav.querySelectorAll('[data-go]').forEach(function (b) { b.addEventListener('click', function () { go(b.dataset.go); }); });
}

/* Count per-record attention bagi satu workspace (mengikut tetapan peraturan admin) */
function countAttention(wsId) {
  const st = ensureSettings();
  let n = 0;
  const recs = state.records[wsId] || [];
  const soon = st.attentionThresholds.vivaSoonDays;
  if (wsId === 'serahTesis' && st.attentionRules.serahTesis) recs.forEach(function (r) { if (r.tarikhHantar && !r.laporanPD) n++; if (r.tarikhHantar && !r.laporanPL) n++; });
  else if (wsId === 'viva' && st.attentionRules.viva) recs.forEach(function (r) { const d = daysFromNow(r.tarikhViva); if (d != null && d >= 0 && d <= soon && !r.tarikhUpdate) n++; });
  else if (wsId === 'notis' && st.attentionRules.notis) recs.forEach(function (r) { if (r.tarikhTerimaPencalonan && !r.tarikhLulus) n++; });
  else if (wsId === 'jps' && st.attentionRules.jps) recs.forEach(function (r) { const d = daysFromNow(r.tarikhMsyt); if (d != null && d < 0 && !r.tarikhMinit) n++; });
  else if (wsId === 'senat' && st.attentionRules.senat) recs.forEach(function (r) { if (r.hantarSenat && !r.tarikhUpdate) n++; });
  return n;
}

/* ---------- MAIN RENDER ---------- */
function render() {
  ensureSettings();
  renderSidebar();
  applySettingsToChrome();
  const c = document.getElementById('content');
  if (ui.view === 'dashboard') c.innerHTML = renderDashboard();
  else if (ui.view === 'admin') c.innerHTML = renderAdmin();
  else if (ui.view === 'penyeliaPelajar') c.innerHTML = renderPenyeliaWorkspace();
  else c.innerHTML = renderWorkspace(ws(ui.view));
  bindContent();
}

/* Kemas kini elemen chrome (topbar/brand) mengikut tetapan */
function applySettingsToChrome() {
  const s = ensureSettings();
  const nameEl = document.querySelector('.sidebar__title');
  const subEl = document.querySelector('.sidebar__subtitle');
  const opName = document.querySelector('.topbar__user-name');
  const opRole = document.querySelector('.topbar__user-role');
  const av = document.querySelector('.avatar');
  const badge = document.querySelector('.demo-badge');
  if (nameEl) nameEl.textContent = s.orgName;
  if (subEl) subEl.textContent = s.orgSub;
  if (opName) opName.textContent = s.operatorName;
  if (opRole) opRole.textContent = s.operatorRole;
  if (av) av.textContent = s.operatorInitials;
  if (badge) badge.style.display = s.showDemoBadge ? '' : 'none';
}

/* ---------- DASHBOARD ---------- */
function renderDashboard() {
  const attention = computeAttention();
  const reminders = state.reminders.filter(function (r) { return !r.done; }).sort(function (a, b) { return (parseDate(a.date) || 0) - (parseDate(b.date) || 0); });
  const upcoming = computeUpcoming();

  let totalRecords = 0;
  Object.keys(state.records).forEach(function (k) { totalRecords += state.records[k].length; });

  let html = '';
  html += '<div class="page-head">';
  html += '<div><h1 class="page-head__title">Dashboard</h1>';
  html += '<p class="page-head__desc">Ringkasan kerja harian, perkara memerlukan perhatian dan countdown acara</p></div>';
  if (state.lastVisit) html += '<button class="btn btn--ghost" data-go="' + state.lastVisit + '">' + icon('history', 16) + '<span class="btn-label">Sambung: ' + esc(ws(state.lastVisit).name) + '</span></button>';
  html += '</div>';

  const attCls = attention.length > 0 ? 'g-alert' : 'g-blue';
  html += '<div class="stat-grid">';
  html += '<button class="stat-card ' + attCls + '" data-scroll="attention"><div class="stat-card__top"><span class="stat-card__label">' + icon('alert', 15) + ' Perlu Perhatian</span><span class="stat-card__icon">' + icon('alert', 19) + '</span></div><div><div class="stat-card__value">' + attention.length + '</div><div class="stat-card__sub">Rekod dengan langkah belum lengkap</div></div></button>';
  html += '<button class="stat-card g-purple" data-scroll="reminders"><div class="stat-card__top"><span class="stat-card__label">' + icon('bell', 15) + ' Peringatan Aktif</span><span class="stat-card__icon">' + icon('bell', 19) + '</span></div><div><div class="stat-card__value">' + reminders.length + '</div><div class="stat-card__sub">Peringatan manual belum selesai</div></div></button>';
  html += '<button class="stat-card g-teal" data-scroll="upcoming"><div class="stat-card__top"><span class="stat-card__label">' + icon('calendar', 15) + ' Acara 60 Hari</span><span class="stat-card__icon">' + icon('calendar', 19) + '</span></div><div><div class="stat-card__value">' + upcoming.length + '</div><div class="stat-card__sub">Viva & mesyuarat akan datang</div></div></button>';
  html += '</div>';

  html += '<div class="mini-grid">';
  html += '<div class="mini-card"><span class="mini-card__icon g-blue">' + icon('grid', 17) + '</span><div><div class="mini-card__value">' + totalRecords + '</div><div class="mini-card__label">Jumlah Rekod</div></div></div>';
  html += '<div class="mini-card"><span class="mini-card__icon g-amber">' + icon('clock', 17) + '</span><div><div class="mini-card__value">' + reminders.length + '</div><div class="mini-card__label">Menunggu Tindakan</div></div></div>';
  html += '</div>';

  html += '<div class="dash-grid"><div class="dash-col">';
  html += '<section class="panel" id="sec-attention"><div class="panel__head"><h2 class="section-title" style="margin:0">' + icon('alert', 17) + ' Perkara Memerlukan Perhatian</h2><span class="badge badge--warn">' + attention.length + ' item</span></div><div class="panel__body">';
  html += attention.length ? attention.slice(0, 12).map(renderAttentionItem).join('') : emptyState('Tiada perkara memerlukan perhatian', 'Semua rekod kelihatan lengkap setakat ini.');
  html += '</div></section></div>';

  html += '<div class="dash-col">';
  html += '<section class="panel" id="sec-reminders"><div class="panel__head"><h2 class="section-title" style="margin:0">' + icon('bell', 17) + ' Peringatan</h2><button class="btn btn--subtle btn--sm" data-add-reminder>' + icon('plus', 14) + ' Tambah</button></div><div class="panel__body">';
  html += reminders.length ? reminders.slice(0, 10).map(renderReminderItem).join('') : emptyState('Tiada peringatan aktif', 'Tambah peringatan untuk acara atau susulan penting.');
  html += '</div></section>';
  html += '<section class="panel" id="sec-upcoming"><div class="panel__head"><h2 class="section-title" style="margin:0">' + icon('calendar', 17) + ' Acara Akan Datang</h2></div><div class="panel__body">';
  html += upcoming.length ? upcoming.slice(0, 10).map(renderUpcomingItem).join('') : emptyState('Tiada acara dalam 60 hari', 'Tarikh viva atau mesyuarat akan dipaparkan di sini.');
  html += '</div></section></div></div>';

  html += '<div class="legend" style="margin-top:20px"><span><i style="background:var(--destructive)"></i> Keutamaan tinggi</span><span><i style="background:var(--warning)"></i> Keutamaan sederhana</span><span><i style="background:var(--primary)"></i> Keutamaan rendah</span><span>' + icon('info', 13) + ' "Belum direkodkan" bermaksud data belum diisi, bukan semestinya kerja belum dibuat.</span></div>';
  return html;
}

function renderAttentionItem(a) {
  return '<button class="attention-item" data-go="' + a.wsId + '" data-focus="' + (a.recordId || '') + '">' +
    '<span class="attention-item__pri p-' + a.prio + '"></span>' +
    '<span class="attention-item__main"><span class="attention-item__title">' + esc(a.title) + '</span>' +
    '<span class="attention-item__meta">' + esc(a.meta) + '</span></span>' +
    '<span class="attention-item__tag">' + esc(a.tag) + '</span></button>';
}

function fmtCount(d) {
  if (d == null) return { cls: '', num: '—', unit: '' };
  if (d < 0) return { cls: 'is-overdue', num: Math.abs(d), unit: 'lewat' };
  if (d === 0) return { cls: 'is-today', num: '0', unit: 'hari ini' };
  return { cls: '', num: d, unit: 'hari lagi' };
}

function renderReminderItem(r) {
  const c = fmtCount(daysFromNow(r.date));
  return '<div class="reminder-item">' +
    '<span class="reminder-item__count ' + c.cls + '"><span class="reminder-item__num">' + c.num + '</span><span class="reminder-item__unit">' + c.unit + '</span></span>' +
    '<span class="attention-item__main"><span class="reminder-item__title">' + esc(r.title) + '</span>' +
    '<span class="reminder-item__sub">' + esc(r.ref || '') + (r.ref ? ' · ' : '') + esc(fmtDate(r.date)) + (r.demo ? ' · contoh' : '') + '</span></span>' +
    '<button class="row-action" data-reminder-done="' + r.id + '" title="Tandakan selesai">' + icon('checkCircle', 15) + '</button></div>';
}

function renderUpcomingItem(u) {
  const c = fmtCount(daysFromNow(u.date));
  return '<button class="reminder-item" data-go="' + u.wsId + '" data-focus="' + (u.recordId || '') + '">' +
    '<span class="reminder-item__count ' + c.cls + '"><span class="reminder-item__num">' + c.num + '</span><span class="reminder-item__unit">' + c.unit + '</span></span>' +
    '<span class="attention-item__main"><span class="reminder-item__title">' + esc(u.title) + '</span>' +
    '<span class="reminder-item__sub">' + esc(u.label) + ' · ' + esc(fmtDate(u.date)) + '</span></span></button>';
}

function emptyState(title, sub) {
  return '<div class="empty-state">' + icon('checkCircle', 34) + '<div style="font-weight:600;color:var(--fg)">' + esc(title) + '</div><div>' + esc(sub) + '</div></div>';
}

/* ---------- ADMIN & TETAPAN ---------- */
function renderAdmin() {
  const s = ensureSettings();
  const totalRec = Object.keys(state.records).reduce(function (n, k) { return n + state.records[k].length; }, 0);
  const trashCount = (state.trash || []).length;

  let html = '<div class="page-head"><div>';
  html += '<h1 class="page-head__title">Admin & Tetapan</h1>';
  html += '<p class="page-head__desc">Konfigurasi sistem, pengendali dan arkib padaman</p></div></div>';

  /* Mini stats */
  html += '<div class="mini-grid" style="grid-template-columns:repeat(3,1fr)">';
  html += '<div class="mini-card"><span class="mini-card__icon g-blue">' + icon('database', 17) + '</span><div><div class="mini-card__value">' + totalRec + '</div><div class="mini-card__label">Jumlah Rekod</div></div></div>';
  html += '<div class="mini-card"><span class="mini-card__icon g-purple">' + icon('history', 17) + '</span><div><div class="mini-card__value">' + (state.history || []).length + '</div><div class="mini-card__label">Entri Sejarah</div></div></div>';
  html += '<div class="mini-card"><span class="mini-card__icon g-amber">' + icon('trash', 17) + '</span><div><div class="mini-card__value">' + trashCount + '</div><div class="mini-card__label">Dalam Arkib Padaman</div></div></div>';
  html += '</div>';

  html += '<div class="admin-grid">';

  /* Kad 1: Identiti Organisasi */
  html += '<section class="panel"><div class="panel__head"><h2 class="section-title" style="margin:0">' + icon('sliders', 17) + ' Identiti Sistem</h2></div><div class="panel__body panel__body--pad">';
  html += adminField('set-orgName', 'Nama Sistem', s.orgName, 'text');
  html += adminField('set-orgSub', 'Subtajuk (Organisasi)', s.orgSub, 'text');
  html += '<div class="field" style="margin-top:12px"><label>Tunjuk Label Demo</label><select id="set-showDemoBadge"><option value="1"' + (s.showDemoBadge ? ' selected' : '') + '>Ya</option><option value="0"' + (!s.showDemoBadge ? ' selected' : '') + '>Tidak</option></select></div>';
  html += '</div></section>';

  /* Kad 2: Pengendali */
  html += '<section class="panel"><div class="panel__head"><h2 class="section-title" style="margin:0">' + icon('users', 17) + ' Pengendali</h2></div><div class="panel__body panel__body--pad">';
  html += adminField('set-operatorName', 'Nama Pengendali', s.operatorName, 'text');
  html += adminField('set-operatorRole', 'Peranan', s.operatorRole, 'text');
  html += adminField('set-operatorInitials', 'Inisial Avatar (2 huruf)', s.operatorInitials, 'text', 2);
  html += '</div></section>';

  /* Kad 3: Paparan & Jadual */
  html += '<section class="panel"><div class="panel__head"><h2 class="section-title" style="margin:0">' + icon('grid', 17) + ' Paparan & Jadual</h2></div><div class="panel__body panel__body--pad">';
  html += adminField('set-pageSize', 'Bilangan Baris Sehalaman', s.pageSize, 'number');
  html += adminField('set-upcomingDays', 'Horizon Countdown (hari)', s.attentionThresholds.upcomingDays, 'number');
  html += adminField('set-vivaSoonDays', 'Ambang "Viva Terdekat" (hari)', s.attentionThresholds.vivaSoonDays, 'number');
  html += '<div class="field" style="margin-top:12px"><label>Sesi Aktif (Semester Semasa)</label><select id="set-sesiAktif">' + selectOptions(generateSemesterOptions(), s.semesterAktif.sesi) + '</select><span class="field__hint">Digunakan untuk label "SEM. PENGAJIAN" dalam Senarai Pelajar Mengikut Penyelia.</span></div>';
  html += '</div></section>';

  /* Kad 4: Peraturan Perhatian */
  html += '<section class="panel"><div class="panel__head"><h2 class="section-title" style="margin:0">' + icon('alert', 17) + ' Peraturan "Perlu Perhatian"</h2></div><div class="panel__body panel__body--pad">';
  const ruleLabels = {
    serahTesis: 'Serah Tesis — laporan PD/PL belum direkodkan',
    viva: 'Peperiksaan Lisan — update SMP belum direkodkan',
    notis: 'Notis Serah Tesis — tarikh lulus belum direkodkan',
    jps: 'Mesyuarat JPS — minit belum siap',
    senat: 'Pengesahan Senat — update SMP belum direkodkan'
  };
  Object.keys(ruleLabels).forEach(function (k) {
    const on = s.attentionRules[k];
    html += '<label class="switch-row"><span>' + esc(ruleLabels[k]) + '</span><input type="checkbox" data-rule="' + k + '"' + (on ? ' checked' : '') + '></label>';
  });
  html += '</div></section>';

  /* Kad 5: Padaman & Arkib */
  html += '<section class="panel"><div class="panel__head"><h2 class="section-title" style="margin:0">' + icon('shield', 17) + ' Padaman & Arkib</h2></div><div class="panel__body panel__body--pad">';
  html += adminField('set-trashRetentionDays', 'Tempoh Simpan Arkib (hari)', s.trashRetentionDays, 'number');
  html += '<div class="field" style="margin-top:12px"><label>Sahkan Sebelum Padam</label><select id="set-confirmDelete"><option value="1"' + (s.confirmDelete ? ' selected' : '') + '>Ya (disyorkan)</option><option value="0"' + (!s.confirmDelete ? ' selected' : '') + '>Tidak</option></select></div>';
  html += '<div class="field" style="margin-top:12px"><label>Rujukan Padaman</label><p class="field__hint">Rekod yang dipadamkan disimpan dalam Arkib Padaman dan boleh dipulihkan pada bila-bila masa. Hanya jadual data sebenar yang berubah — tiada maklumat hilang semasa demo.</p></div>';
  html += '</div></section>';

  /* Kad 6: Data & Selenggara */
  html += '<section class="panel"><div class="panel__head"><h2 class="section-title" style="margin:0">' + icon('database', 17) + ' Data & Selenggara</h2></div><div class="panel__body panel__body--pad">';
  html += '<div class="admin-actions">';
  html += '<button class="btn btn--ghost" data-export-all>' + icon('download', 16) + ' Eksport Semua (JSON)</button>';
  html += '<button class="btn btn--ghost" data-reset-demo>' + icon('history', 16) + ' Reset Data Demo</button>';
  html += '</div>';
  html += '<p class="field__hint" style="margin-top:10px">Eksport JSON berguna sebagai sandaran penuh sebelum reset atau sebelum pindah ke backend sebenar.</p>';
  html += '</div></section>';

  html += '</div>'; /* admin-grid */

  /* Arkib Padaman */
  html += '<section class="panel" style="margin-top:18px"><div class="panel__head"><h2 class="section-title" style="margin:0">' + icon('trash', 17) + ' Arkib Padaman</h2>';
  html += '<div style="display:flex;gap:8px;align-items:center"><span class="badge badge--warn">' + trashCount + ' rekod</span>';
  if (trashCount) html += '<button class="btn btn--danger btn--sm" data-empty-trash>' + icon('trash', 14) + ' Kosongkan</button>';
  html += '</div></div><div class="panel__body">';
  if (!trashCount) {
    html += emptyState('Arkib kosong', 'Rekod yang dipadamkan akan muncul di sini dan boleh dipulihkan.');
  } else {
    (state.trash || []).forEach(function (t) {
      const w = ws(t.wsId);
      const label = t.record.nama || t.record.perkara || '(rekod)';
      html += '<div class="trash-item">';
      html += '<span class="attention-item__main"><span class="reminder-item__title">' + esc(label) + '</span>';
      html += '<span class="reminder-item__sub">' + esc(w ? w.name : t.wsId) + ' · dipadam ' + esc(new Date(t.at).toLocaleString('ms-MY')) + '</span></span>';
      html += '<div class="cell-actions"><button class="row-action" data-restore="' + t.id + '" title="Pulihkan" style="color:var(--accent);border-color:var(--accent)">' + icon('restore', 15) + '</button>';
      html += '<button class="row-action is-danger" data-purge="' + t.id + '" title="Padam kekal">' + icon('trash', 15) + '</button></div>';
      html += '</div>';
    });
  }
  html += '</div></section>';

  /* Footer credit */
  html += '<div class="admin-credit">Developed by: <strong>Hab Digital IMEN</strong></div>';

  return html;
}

function adminField(id, label, value, type, maxlength) {
  return '<div class="field" style="margin-top:12px"><label>' + esc(label) + '</label>' +
    '<input id="' + id + '" type="' + type + '" value="' + esc(value) + '"' + (maxlength ? ' maxlength="' + maxlength + '"' : '') + '></div>';
}

/* ---------- WORKSPACE PENYELIA-PELAJAR (bersarang) ---------- */
/* Penjana senarai sesi semester (dropdown) — 3 ke belakang, 1 ke hadapan */
function generateSemesterOptions() {
  const now = new Date();
  const y = now.getFullYear();
  const out = [];
  for (let sy = y - 3; sy <= y + 1; sy++) {
    out.push('1/' + sy + '-' + (sy + 1));
    out.push('2/' + sy + '-' + (sy + 1));
  }
  return out;
}

function renderPenyeliaWorkspace() {
  const w = ws('penyeliaPelajar');
  const st = ensureSettings();
  const showAll = ui.filter['penyeliaPelajar'] === 'semua';
  const recs = state.records.penyeliaPelajar || [];
  const q = (ui.query['penyeliaPelajar'] || '').toLowerCase().trim();

  function filterPelajar(plist) {
    let list = plist || [];
    if (q) list = list.filter(function (p) {
      return String(p.nama || '').toLowerCase().indexOf(q) !== -1 || String(p.noPelajar || '').toLowerCase().indexOf(q) !== -1;
    });
    return list;
  }

  let html = '<div class="page-head">';
  html += '<div><h1 class="page-head__title">' + esc(w.name) + '</h1><p class="page-head__desc">' + esc(w.desc || '') + '</p></div>';
  html += '<button class="btn btn--primary" data-add-penyelia>' + icon('plus', 16) + '<span class="btn-label">Tambah Penyelia</span></button>';
  html += '</div>';

  /* Toolbar: carian + toggle Aktif/Semua */
  html += '<div class="ws-toolbar">';
  html += '<div class="ws-search' + (q ? ' has-value' : '') + '"><svg class="ws-search__icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="M21 21l-4.3-4.3"/></svg>';
  html += '<input type="search" data-ws-search="penyeliaPelajar" value="' + esc(ui.query['penyeliaPelajar'] || '') + '" placeholder="Cari pelajar (nama / no. pelajar)…" aria-label="Cari pelajar">';
  html += '<button class="ws-search__clear" data-ws-clear="penyeliaPelajar" aria-label="Kosongkan carian">' + icon('x', 15) + '</button></div>';
  html += '<div class="seg-toggle">';
  html += '<button class="seg' + (!showAll ? ' is-active' : '') + '" data-seg="aktif">Aktif</button>';
  html += '<button class="seg' + (showAll ? ' is-active' : '') + '" data-seg="semua">Semua</button>';
  html += '</div>';
  html += '</div>';

  const sesiAktif = st.semesterAktif.sesi;

  if (!recs.length) {
    html += '<div class="panel"><div class="panel__body">' + emptyState('Tiada penyelia', 'Klik "Tambah Penyelia" untuk mula.') + '</div></div>';
    return html;
  }

  /* Bina panel penyelia */
  recs.forEach(function (p) {
    const allPelajar = filterPelajar(p.pelajar);
    const aktif = allPelajar.filter(function (x) { return !x.statusKhas; });
    const lepas = allPelajar.filter(function (x) { return x.statusKhas; });
    const open = (ui.penyeliaOpen && ui.penyeliaOpen[p.id]) !== false; // default terbuka

    html += '<section class="penyelia-panel' + (open ? ' is-open' : '') + '" data-penyelia="' + p.id + '">';
    html += '<button class="penyelia-panel__head" data-toggle-penyelia="' + p.id + '">';
    html += '<span class="penyelia-panel__chev">' + icon('chevronRight', 16) + '</span>';
    html += '<span class="penyelia-panel__icon">' + icon('userCheck', 15) + '</span>';
    html += '<span class="penyelia-panel__title">' + esc(p.namaPenyelia) + '</span>';
    html += '<span class="badge badge--info">' + aktif.length + ' aktif</span>';
    if (lepas.length) html += '<span class="badge badge--warn">' + lepas.length + ' lepas</span>';
    html += '</button>';

    if (open) {
      html += '<div class="penyelia-panel__body">';
      html += renderPenyeliaTable(aktif, sesiAktif, p.id);
      if (lepas.length && (showAll || true)) {
        html += renderLepasPanel(p.id, lepas, sesiAktif, showAll);
      }
      html += '</div>';
    }
    html += '</section>';
  });

  return html;
}

function renderPenyeliaTable(pelajar, sesiAktif, penyeliaId) {
  const cols = ws('penyeliaPelajar').columns;
  if (!pelajar.length) {
    return '<div class="empty-state" style="padding:16px">' + icon('userCheck', 24) + '<div>Tiada pelajar aktif.</div></div>';
  }
  let html = '<div class="table-scroll"><table class="grid penyelia-table"><thead><tr>';
  cols.forEach(function (c, i) {
    html += '<th class="' + colClass(c, i) + '" title="' + esc(c.header) + '">' + esc(c.header) + '</th>';
  });
  html += '<th class="col-actions">TINDAKAN</th></tr></thead><tbody>';
  pelajar.forEach(function (p) {
    html += renderPenyeliaRow(p, penyeliaId);
  });
  html += '</tbody></table></div>';
  return html;
}

function renderPenyeliaRow(p, penyeliaId) {
  const cols = ws('penyeliaPelajar').columns;
  const cells = cols.map(function (c, i) {
    const cls = colClass(c, i);
    let val = p[c.key] || '';
    if (c.key === 'status') val = statusPelajarLabel(p.statusKhas);
    if (c.key === 'semesterPengajian') val = p.semesterPengajian || '';
    if (c.key === 'penyeliaBersama') val = (p.penyeliaBersama || '').replace(/\n/g, ', ');
    return '<td class="' + cls + '">' + esc(val) + '</td>';
  }).join('');

  const statusBadge = p.statusKhas
    ? '<span class="badge badge--warn">' + esc(statusPelajarLabel(p.statusKhas)) + '</span>'
    : '';

  return '<tr data-pelajar="' + p.id + '">' + cells +
    '<td class="col-actions"><div class="cell-actions">' +
    '<button class="row-action" data-edit-pelajar="' + penyeliaId + '|' + p.id + '" title="Edit / Status">' + icon('edit', 15) + '</button>' +
    '<button class="row-action" data-status-pelajar="' + penyeliaId + '|' + p.id + '" title="Ubah Status">' + icon('checkCircle', 15) + '</button>' +
    '</div>' + statusBadge + '</td></tr>';
}

function renderLepasPanel(penyeliaId, lepas, sesiAktif, showAll) {
  const open = ui.lepasOpen && ui.lepasOpen[penyeliaId];
  let html = '<section class="lepas-panel' + (open ? ' is-open' : '') + '">';
  html += '<button class="lepas-panel__head" data-toggle-lepas="' + penyeliaId + '">';
  html += '<span class="locked-panel__chev">' + icon('chevronRight', 15) + '</span>';
  html += '<span class="locked-panel__title">Rekod Lepas</span>';
  html += '<span class="badge badge--warn">' + lepas.length + '</span>';
  html += '<span class="locked-panel__hint">Graduasi / Menarik Diri / Diberhentikan</span>';
  html += '</button>';
  if (open) {
    html += '<div class="lepas-panel__body">';
    html += renderPenyeliaTable(lepas, sesiAktif, penyeliaId);
    html += '</div>';
  }
  html += '</section>';
  return html;
}

/* ---------- PENYELIA: drawer edit pelajar ---------- */
function openPelajarDrawer(penyeliaId, pelajarId) {
  const p = (state.records.penyeliaPelajar || []).find(function (x) { return x.id === penyeliaId; });
  if (!p) return;
  const pelajar = pelajarId ? p.pelajar.find(function (x) { return x.id === pelajarId; }) : null;
  const isNew = !pelajar;
  const rec = pelajar || { id: '', bil: (p.pelajar.length + 1), nama: '', program: '', noPelajar: '', penyeliaBersama: '', statusKhas: '', tarikhStatus: '', semesterPengajian: '', sejarahSemester: [] };

  let body = '<input type="hidden" id="dPenyeliaId" value="' + penyeliaId + '">';
  body += '<input type="hidden" id="dPelajarId" value="' + (isNew ? '' : pelajarId) + '">';
  body += '<div class="field"><label>Nama Pelajar <span class="req">*</span></label><input id="dNama" value="' + esc(rec.nama) + '"></div>';
  body += '<div class="field" style="margin-top:12px"><label>Program</label><select id="dProgram">' + selectOptions(PROGRAM_OPTIONS, rec.program) + '</select></div>';
  body += '<div class="field" style="margin-top:12px"><label>No. Pelajar</label><input id="dNoPelajar" value="' + esc(rec.noPelajar) + '"></div>';
  body += '<div class="field" style="margin-top:12px"><label>Penyelia Bersama</label><textarea id="dPenyeliaBersama">' + esc(rec.penyeliaBersama) + '</textarea></div>';
  body += '<div class="field" style="margin-top:12px"><label>Semester Pengajian (semasa)</label><input id="dSemesterPengajian" value="' + esc(rec.semesterPengajian) + '"></div>';
  body += '<div class="field" style="margin-top:12px"><label>Status Pelajar</label><select id="dStatus">' + selectOptions(STATUS_PELAJAR_OPTIONS, rec.statusKhas) + '</select></div>';
  body += '<div class="field" style="margin-top:12px"><label>Tarikh Status</label><input type="date" id="dTarikhStatus" value="' + esc(rec.tarikhStatus) + '"><span class="field__hint">Wajib untuk graduasi; optional untuk menarik diri / diberhentikan.</span></div>';

  /* Sejarah semester */
  body += '<div class="field" style="margin-top:16px"><label>Sejarah Semester</label><div id="dSejarah" class="sejarah-list">';
  (rec.sejarahSemester || []).forEach(function (s, idx) {
    body += sejarahRow(s.sesi, s.semesterPengajian, s.catatan, idx);
  });
  body += '</div><button class="penyelia-add" data-add-sejarah>' + icon('plus', 13) + ' Tambah Semester</button></div>';

  openDrawer({
    title: (isNew ? 'Tambah Pelajar' : 'Edit Pelajar'),
    sub: esc(p.namaPenyelia),
    body: body,
    foot: '<button class="btn btn--ghost" data-close-drawer>Batal</button><button class="btn btn--primary" id="savePelajar">' + icon('save', 16) + ' Simpan</button>',
    onMount: function (root) {
      root.querySelectorAll('[data-add-sejarah]').forEach(function (b) {
        b.addEventListener('click', function () {
          const list = root.querySelector('#dSejarah');
          const div = document.createElement('div');
          div.innerHTML = sejarahRow('', '', '', Date.now());
          list.appendChild(div);
        });
      });
      root.querySelectorAll('[data-remove-sejarah]').forEach(function (b) {
        b.addEventListener('click', function () { b.closest('.sejarah-row').remove(); });
      });
      root.querySelector('#dStatus').addEventListener('change', function () {
        const tgl = root.querySelector('#dTarikhStatus');
        tgl.required = this.value === 'graduasi';
      });
      root.querySelector('#savePelajar').addEventListener('click', function () {
        savePelajarFromDrawer(root);
      });
    }
  });
}

function sejarahRow(sesi, sp, catatan, idx) {
  return '<div class="sejarah-row">' +
    '<select class="sejarah-sesi">' + selectOptions(generateSemesterOptions(), sesi) + '</select>' +
    '<input class="sejarah-sp" placeholder="Sem. pengajian" value="' + esc(sp || '') + '">' +
    '<input class="sejarah-catatan" placeholder="Catatan" value="' + esc(catatan || '') + '">' +
    '<button type="button" class="penyelia-remove" data-remove-sejarah title="Buang">' + icon('x', 13) + '</button>' +
    '</div>';
}

function selectOptions(opts, selected) {
  return opts.map(function (o) {
    const v = typeof o === 'object' ? o.value : o;
    const l = typeof o === 'object' ? o.label : o;
    return '<option value="' + esc(v) + '"' + (v === selected ? ' selected' : '') + '>' + esc(l) + '</option>';
  }).join('');
}

function savePelajarFromDrawer(root) {
  const penyeliaId = root.querySelector('#dPenyeliaId').value;
  const pelajarId = root.querySelector('#dPelajarId').value;
  const p = (state.records.penyeliaPelajar || []).find(function (x) { return x.id === penyeliaId; });
  if (!p) return;

  const nama = root.querySelector('#dNama').value.trim();
  if (!nama) { toast('Nama pelajar wajib diisi', 'error'); return; }

  const statusKhas = root.querySelector('#dStatus').value;
  const tarikhStatus = root.querySelector('#dTarikhStatus').value;
  if (statusKhas === 'graduasi' && !tarikhStatus) { toast('Tarikh wajib untuk status graduasi', 'error'); return; }

  const sejarahSemester = [];
  root.querySelectorAll('#dSejarah .sejarah-row').forEach(function (row) {
    const sesi = row.querySelector('.sejarah-sesi').value;
    const sp = row.querySelector('.sejarah-sp').value;
    const catatan = row.querySelector('.sejarah-catatan').value;
    if (sesi || sp || catatan) sejarahSemester.push({ sesi: sesi, semesterPengajian: sp, catatan: catatan });
  });

  const data = {
    nama: nama,
    program: root.querySelector('#dProgram').value,
    noPelajar: root.querySelector('#dNoPelajar').value,
    penyeliaBersama: root.querySelector('#dPenyeliaBersama').value,
    semesterPengajian: root.querySelector('#dSemesterPengajian').value,
    statusKhas: statusKhas,
    tarikhStatus: tarikhStatus,
    sejarahSemester: sejarahSemester
  };

  if (pelajarId) {
    const existing = p.pelajar.find(function (x) { return x.id === pelajarId; });
    const oldStatus = existing ? existing.statusKhas : '';
    Object.assign(existing, data);
    syncGraduan(existing, oldStatus);
    toast('Pelajar dikemas kini', 'success');
  } else {
    const newPelajar = Object.assign({ id: uid('r'), bil: p.pelajar.length + 1 }, data);
    p.pelajar.push(newPelajar);
    syncGraduan(newPelajar, '');
    toast('Pelajar ditambah', 'success');
  }
  save();
  closeDrawer();
  render();
}

/* ---------- Status cepat (dropdown di baris) ---------- */
function openStatusQuick(penyeliaId, pelajarId) {
  const p = (state.records.penyeliaPelajar || []).find(function (x) { return x.id === penyeliaId; });
  if (!p) return;
  const pelajar = p.pelajar.find(function (x) { return x.id === pelajarId; });
  if (!pelajar) return;

  let body = '<div class="field"><label>Status Pelajar</label><select id="qStatus">' + selectOptions(STATUS_PELAJAR_OPTIONS, pelajar.statusKhas) + '</select></div>';
  body += '<div class="field" style="margin-top:12px"><label>Tarikh Status</label><input type="date" id="qTarikh" value="' + esc(pelajar.tarikhStatus || '') + '"><span class="field__hint">Wajib untuk graduasi; optional untuk lain.</span></div>';

  openModal({
    title: 'Ubah Status',
    sub: esc(pelajar.nama),
    body: body,
    foot: '<button class="btn btn--ghost" data-close-modal>Batal</button><button class="btn btn--primary" id="saveStatus">Simpan</button>',
    onMount: function (root) {
      root.querySelector('#saveStatus').addEventListener('click', function () {
        const status = root.querySelector('#qStatus').value;
        const tarikh = root.querySelector('#qTarikh').value;
        if (status === 'graduasi' && !tarikh) { toast('Tarikh wajib untuk graduasi', 'error'); return; }
        const old = pelajar.statusKhas;
        pelajar.statusKhas = status;
        pelajar.tarikhStatus = tarikh;
        syncGraduan(pelajar, old);
        save();
        closeModal();
        render();
        toast('Status dikemas kini', 'success');
      });
    }
  });
}

/* ---------- AUTO-SYNC GRADUAN ---------- */
function syncGraduan(pelajar, oldStatus) {
  const baru = pelajar.statusKhas;
  if (baru === 'graduasi') {
    /* tambah / kemas kini rekod graduan */
    const existing = (state.records.graduan || []).find(function (g) {
      return g.noPelajar && pelajar.noPelajar && g.noPelajar === pelajar.noPelajar;
    }) || (state.records.graduan || []).find(function (g) { return g._syncPelajarId === pelajar.id; });

    const entry = {
      id: existing ? existing.id : uid('r'),
      bil: existing ? existing.bil : nextId('graduan'),
      nama: pelajar.nama + (pelajar.noPelajar ? ' (' + pelajar.noPelajar + ')' : ''),
      program: pelajar.program || '',
      penyelia: 'Penyelia Utama\n' + (pelajar.penyeliaBersama || ''),
      semester: pelajar.semesterPengajian || '',
      catatan: 'Graduasi ' + (pelajar.tarikhStatus ? fmtDate(pelajar.tarikhStatus) : ''),
      noPelajar: pelajar.noPelajar,
      _syncPelajarId: pelajar.id
    };
    if (existing) {
      Object.assign(existing, entry);
    } else {
      state.records.graduan = state.records.graduan || [];
      state.records.graduan.push(entry);
    }
  } else if (oldStatus === 'graduasi') {
    /* buang rekod graduan yang dijana auto */
    state.records.graduan = (state.records.graduan || []).filter(function (g) {
      return g._syncPelajarId !== pelajar.id && g.noPelajar !== pelajar.noPelajar;
    });
  }
}

/* ---------- Tambah penyelia ---------- */
function addPenyelia() {
  const nama = prompt('Nama penyelia:');
  if (!nama || !nama.trim()) return;
  state.records.penyeliaPelajar = state.records.penyeliaPelajar || [];
  state.records.penyeliaPelajar.push({ id: uid('r'), bil: state.records.penyeliaPelajar.length + 1, namaPenyelia: nama.trim(), pelajar: [] });
  save();
  render();
  toast('Penyelia ditambah', 'success');
}

/* ---------- ATTENTION ---------- */
function computeAttention() {
  const st = ensureSettings();
  const soon = st.attentionThresholds.vivaSoonDays;
  const rules = st.attentionRules;
  const out = [];
  function push(wsId, recordId, title, meta, tag, prio) { out.push({ wsId: wsId, recordId: recordId, title: title, meta: meta, tag: tag, prio: prio }); }

  if (rules.serahTesis) (state.records.serahTesis || []).forEach(function (r) {
    if (r.tarikhHantar && !r.laporanPD) push('serahTesis', r.id, 'Laporan PD belum direkodkan', r.nama, 'Semakan', 'med');
    if (r.tarikhHantar && !r.laporanPL) push('serahTesis', r.id, 'Laporan PL belum direkodkan', r.nama, 'Semakan', 'med');
  });
  if (rules.viva) (state.records.viva || []).forEach(function (r) {
    const d = daysFromNow(r.tarikhViva);
    if (d != null && d >= 0 && d <= soon && !r.tarikhUpdate) push('viva', r.id, 'Update SMP belum direkodkan', r.nama + ' · viva dalam ' + d + ' hari', 'Segera', 'high');
  });
  if (rules.notis) (state.records.notis || []).forEach(function (r) {
    if (r.tarikhTerimaPencalonan && !r.tarikhLulus) push('notis', r.id, 'Tarikh lulus belum direkodkan', r.nama, 'Semakan', 'low');
  });
  if (rules.jps) (state.records.jps || []).forEach(function (r) {
    const d = daysFromNow(r.tarikhMsyt);
    if (d != null && d < 0 && !r.tarikhMinit) push('jps', r.id, 'Minit belum siap', r.perkara, 'Susulan', 'med');
  });
  if (rules.senat) (state.records.senat || []).forEach(function (r) {
    if (r.hantarSenat && !r.tarikhUpdate) push('senat', r.id, 'Update SMP belum direkodkan', r.nama, 'Semakan', 'low');
  });
  const rank = { high: 0, med: 1, low: 2 };
  return out.sort(function (a, b) { return rank[a.prio] - rank[b.prio]; });
}

function computeUpcoming() {
  const st = ensureSettings();
  const horizon = st.attentionThresholds.upcomingDays;
  const out = [];
  (state.records.viva || []).forEach(function (r) {
    if (r.tarikhViva) out.push({ wsId: 'viva', recordId: r.id, title: 'Viva: ' + r.nama, label: r.program || '', date: r.tarikhViva });
  });
  (state.records.jps || []).forEach(function (r) {
    if (r.tarikhMsyt) out.push({ wsId: 'jps', recordId: r.id, title: r.perkara, label: 'Mesyuarat JPS', date: r.tarikhMsyt });
  });
  return out.filter(function (u) { const d = daysFromNow(u.date); return d != null && d >= 0 && d <= horizon; })
    .sort(function (a, b) { return (parseDate(a.date) || 0) - (parseDate(b.date) || 0); });
}

/* ---------- WORKSPACE ---------- */
function renderWorkspace(w) {
  const cols = w.columns;
  const allRows = (state.records[w.id] || []).slice();
  const q = (ui.query[w.id] || '').toLowerCase().trim();
  const fk = ui.filter[w.id] || 'all';
  const s = ui.sort[w.id];

  function matchQuery(rows) {
    if (!q) return rows;
    return rows.filter(function (r) { return cols.some(function (c) { return String(r[c.key] || '').toLowerCase().indexOf(q) !== -1; }); });
  }
  function doSort(rows) {
    if (!s) return rows;
    return rows.slice().sort(function (a, b) {
      const av = a[s.key] == null ? '' : String(a[s.key]);
      const bv = b[s.key] == null ? '' : String(b[s.key]);
      const ad = parseDate(av), bd = parseDate(bv);
      let cmp;
      if (ad && bd) cmp = ad - bd;
      else cmp = av.localeCompare(bv, 'ms', { numeric: true });
      return s.dir === 'desc' ? -cmp : cmp;
    });
  }

  /* Pisah rekod aktif vs terkunci */
  const activeRows = allRows.filter(function (r) { return !r.locked; });
  const lockedRows = allRows.filter(function (r) { return r.locked; });

  let rows = matchQuery(applyChipFilter(w, activeRows, fk));
  rows = doSort(rows);

  const lockedSearch = doSort(matchQuery(lockedRows));

  const total = rows.length;
  const page = ui.page[w.id] || 1;
  const pageCount = Math.max(1, Math.ceil(total / ui.pageSize));
  const curPage = Math.min(page, pageCount);
  const start = (curPage - 1) * ui.pageSize;
  const pageRows = rows.slice(start, start + ui.pageSize);
  const chips = chipDefs(w);
  const lockedCount = lockedRows.length;

  let html = '<div class="page-head">';
  html += '<div><h1 class="page-head__title">' + esc(w.name) + '</h1><p class="page-head__desc">' + esc(w.desc || '') + '</p></div>';
  html += '<button class="btn btn--primary" data-add-row="' + w.id + '">' + icon('plus', 16) + '<span class="btn-label">Tambah Rekod</span></button>';
  html += '</div>';

  html += '<div class="ws-toolbar">';
  html += '<div class="ws-search' + (q ? ' has-value' : '') + '"><svg class="ws-search__icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="M21 21l-4.3-4.3"/></svg>';
  html += '<input type="search" data-ws-search="' + w.id + '" value="' + esc(ui.query[w.id] || '') + '" placeholder="Cari dalam ' + esc(w.name) + '…" aria-label="Cari dalam workspace">';
  html += '<button class="ws-search__clear" data-ws-clear="' + w.id + '" aria-label="Kosongkan carian">' + icon('x', 15) + '</button></div>';
  html += '<button class="btn btn--ghost" data-export="' + w.id + '">' + icon('download', 16) + '<span class="btn-label">Eksport</span></button>';
  html += '<span class="badge badge--info">' + total + ' rekod aktif' + (lockedCount ? ' · ' + lockedCount + ' dikunci' : '') + ((q || fk !== 'all') ? ' · ditapis' : '') + '</span></div>';

  if (chips.length > 1) {
    html += '<div class="chip-row">' + chips.map(function (ch) {
      return '<button class="chip' + (fk === ch.key ? ' is-active' : '') + '" data-chip="' + w.id + '|' + ch.key + '">' + esc(ch.label) + ' <span class="chip__count">' + ch.count + '</span></button>';
    }).join('') + '</div>';
  }

  html += '<div class="table-wrap"><div class="table-scroll"><table class="grid"><thead><tr>';
  html += cols.map(function (c, i) {
    return '<th class="' + colClass(c, i) + '" data-sortable="' + w.id + '|' + c.key + '" title="' + esc(c.header) + '">' + esc(c.header) + renderSortInd(w.id, c.key) + '</th>';
  }).join('');
  html += '<th class="col-actions">TINDAKAN</th></tr></thead><tbody>';
  if (pageRows.length) html += pageRows.map(function (r) { return renderRow(w, r); }).join('');
  else html += '<tr><td colspan="' + (cols.length + 1) + '">' + emptyState('Tiada rekod aktif', (q || fk !== 'all') ? 'Cuba tukar carian atau penapis.' : 'Klik Tambah Rekod untuk mula.') + '</td></tr>';
  html += '</tbody></table></div>';

  html += '<div class="table-foot"><span>' + (total ? (start + 1) + '–' + Math.min(start + ui.pageSize, total) + ' daripada ' + total + ' rekod aktif' : 'Tiada rekod aktif') + '</span>';
  if (pageCount > 1) {
    html += '<div class="pagination">';
    html += '<button class="page-btn" data-page="' + w.id + '|' + (curPage - 1) + '" ' + (curPage <= 1 ? 'disabled' : '') + '>' + icon('chevronLeft', 14) + '</button>';
    html += pageButtons(w.id, curPage, pageCount);
    html += '<button class="page-btn" data-page="' + w.id + '|' + (curPage + 1) + '" ' + (curPage >= pageCount ? 'disabled' : '') + '>' + icon('chevronRight', 14) + '</button></div>';
  }
  html += '</div></div>';

  /* Panel Rekod Dikunci (collapsible) */
  if (lockedCount) html += renderLockedPanel(w, lockedSearch);

  return html;
}

/* Panel collapsible untuk rekod dikunci */
function renderLockedPanel(w, lockedRows) {
  const open = ui.lockedOpen && ui.lockedOpen[w.id];
  let html = '<section class="locked-panel' + (open ? ' is-open' : '') + '" id="lockedPanel">';
  html += '<button class="locked-panel__head" data-toggle-locked="' + w.id + '">';
  html += '<span class="locked-panel__chev">' + icon('chevronRight', 16) + '</span>';
  html += '<span class="locked-panel__icon">' + icon('lock', 15) + '</span>';
  html += '<span class="locked-panel__title">Rekod Dikunci</span>';
  html += '<span class="badge badge--warn">' + lockedRows.length + '</span>';
  html += '<span class="locked-panel__hint">Rekod muktamad — buka kunci untuk edit</span>';
  html += '</button>';

  if (open) {
    html += '<div class="locked-panel__body"><div class="table-scroll"><table class="grid locked-table"><thead><tr>';
    html += w.columns.map(function (c, i) {
      return '<th class="' + colClass(c, i) + '" title="' + esc(c.header) + '">' + esc(c.header) + '</th>';
    }).join('');
    html += '<th class="col-actions">TINDAKAN</th></tr></thead><tbody>';
    html += lockedRows.map(function (r) { return renderRow(w, r); }).join('');
    html += '</tbody></table></div></div>';
  }
  html += '</section>';
  return html;
}

function renderSortInd(wsId, key) {
  const s = ui.sort[wsId];
  if (!s || s.key !== key) return '';
  return '<span class="sort-ind">' + (s.dir === 'asc' ? '▲' : '▼') + '</span>';
}

/* Kelas kolum: BIL dijadikan kecil & padat, kolum pertama sticky */
function colClass(c, i) {
  const parts = [];
  if (i === 0) parts.push('col-sticky');
  if (c.key === 'bil') parts.push('col-bil');
  return parts.join(' ');
}

function pageButtons(wsId, cur, count) {
  let html = '';
  const nums = [];
  for (let i = 1; i <= count; i++) {
    if (i === 1 || i === count || Math.abs(i - cur) <= 1) nums.push(i);
    else if (nums[nums.length - 1] !== '…') nums.push('…');
  }
  nums.forEach(function (n) {
    if (n === '…') html += '<span style="color:var(--muted-fg);padding:0 4px">…</span>';
    else html += '<button class="page-btn' + (n === cur ? ' is-active' : '') + '" data-page="' + wsId + '|' + n + '">' + n + '</button>';
  });
  return html;
}

/* ---------- CHIP FILTERS ---------- */
function chipDefs(w) {
  const recs = state.records[w.id] || [];
  const defs = [{ key: 'all', label: 'Semua', count: recs.length }];
  function add(key, label, fn) { defs.push({ key: key, label: label, count: recs.filter(fn).length }); }

  if (w.id === 'serahTesis') add('pending', 'Laporan belum lengkap', function (r) { return r.tarikhHantar && (!r.laporanPD || !r.laporanPL); });
  else if (w.id === 'viva') add('upcoming', 'Viva akan datang', function (r) { const d = daysFromNow(r.tarikhViva); return d != null && d >= 0; });
  else if (w.id === 'jps') add('minit', 'Minit belum siap', function (r) { return r.tarikhMsyt && !r.tarikhMinit; });
  else if (w.id === 'notis') add('pending', 'Lulus belum rekod', function (r) { return r.tarikhTerimaPencalonan && !r.tarikhLulus; });
  else if (w.id === 'senat') add('smppending', 'Update SMP belum', function (r) { return r.hantarSenat && !r.tarikhUpdate; });
  return defs;
}

function applyChipFilter(w, rows, fk) {
  if (!fk || fk === 'all') return rows;
  const key = w.id + ':' + fk;
  if (key === 'serahTesis:pending') return rows.filter(function (r) { return r.tarikhHantar && (!r.laporanPD || !r.laporanPL); });
  if (key === 'viva:upcoming') return rows.filter(function (r) { const d = daysFromNow(r.tarikhViva); return d != null && d >= 0; });
  if (key === 'jps:minit') return rows.filter(function (r) { return r.tarikhMsyt && !r.tarikhMinit; });
  if (key === 'notis:pending') return rows.filter(function (r) { return r.tarikhTerimaPencalonan && !r.tarikhLulus; });
  if (key === 'senat:smppending') return rows.filter(function (r) { return r.hantarSenat && !r.tarikhUpdate; });
  return rows;
}

/* ---------- ROW ---------- */
function renderRow(w, r) {
  const editing = ui.editingRow === r.id;
  const cols = w.columns;
  const cells = cols.map(function (c, i) {
    const val = (editing && ui.draft) ? (ui.draft[c.key] || '') : (r[c.key] || '');
    const cls = colClass(c, i);
    if (editing && ui.draft && !c.noEdit) return '<td class="' + cls + '">' + renderCellInput(c, val) + '</td>';
    if (c.type === 'date' && val) return '<td class="' + cls + '" title="' + esc(fmtDateLong(val)) + '">' + esc(fmtDate(val)) + '</td>';
    return '<td class="' + cls + '">' + esc(val) + '</td>';
  }).join('');

  const lockIcon = r.locked ? '<span class="row-lock-mark" title="Rekod dikunci">' + icon('lock', 12) + '</span>' : '';

  const actions = editing
    ? '<div class="cell-actions"><button class="row-action" data-save-row="' + w.id + '|' + r.id + '" title="Semak & Simpan" style="color:var(--accent);border-color:var(--accent)">' + icon('save', 15) + '</button>' +
      '<button class="row-action is-danger" data-cancel-row="' + r.id + '" title="Batal">' + icon('x', 15) + '</button></div>'
    : '<div class="cell-actions"><button class="row-action" data-edit-row="' + w.id + '|' + r.id + '" title="Edit">' + icon('edit', 15) + '</button>' +
      '<button class="row-action" data-row-menu="' + w.id + '|' + r.id + '" title="Lagi tindakan">' + icon('more', 15) + '</button></div>';

  const rowCls = [editing ? 'is-editing' : '', r.locked ? 'is-locked' : ''].filter(Boolean).join(' ');
  return '<tr class="' + rowCls + '" data-record="' + r.id + '">' + cells + '<td class="col-actions">' + lockIcon + actions + '</td></tr>';
}

function renderCellInput(c, val) {
  if (c.multi) return renderMultiInput(c, val);
  if (c.type === 'select') {
    let opts = '<option value=""></option>';
    (c.options || []).forEach(function (o) { opts += '<option value="' + esc(o) + '"' + (o === val ? ' selected' : '') + '>' + esc(o) + '</option>'; });
    if (val && (c.options || []).indexOf(val) === -1) opts += '<option value="' + esc(val) + '" selected>' + esc(val) + '</option>';
    return '<select class="cell-select" data-field="' + c.key + '">' + opts + '</select>';
  }
  if (c.type === 'textarea') return '<textarea class="cell-textarea" data-field="' + c.key + '">' + esc(val) + '</textarea>';
  if (c.type === 'date') {
    const d = parseDate(val);
    const iso = d ? (d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0')) : '';
    return '<input type="date" class="cell-input" data-field="' + c.key + '" value="' + esc(iso) + '">';
  }
  if (c.type === 'number') return '<input type="text" inputmode="numeric" class="cell-input" data-field="' + c.key + '" value="' + esc(val) + '">';
  return '<input type="text" class="cell-input" data-field="' + c.key + '" value="' + esc(val) + '">';
}

/* Input senarai boleh ulang untuk medan penyelia (1 hingga banyak) */
function renderMultiInput(c, val) {
  let names = '';
  if (c.autoLabel) {
    names = extractPenyeliaNames(val);
  } else {
    names = String(val || '').split('\n').map(function (s) { return s.trim(); }).filter(Boolean);
  }
  if (!names.length) names = [''];

  let rows = names.map(function (n) {
    return '<div class="penyelia-row">' +
      '<input type="text" class="cell-input penyelia-input" data-field="' + c.key + '" data-multi="1" value="' + esc(n) + '" placeholder="Nama penyelia…">' +
      '<button type="button" class="penyelia-remove" data-remove-penyelia title="Buang">' + icon('x', 13) + '</button>' +
      '</div>';
  }).join('');

  return '<div class="penyelia-list" data-multi-key="' + c.key + '">' +
    '<div class="penyelia-rows">' + rows + '</div>' +
    '<button type="button" class="penyelia-add" data-add-penyelia="' + c.key + '">' + icon('plus', 13) + ' Tambah Penyelia</button>' +
    (c.autoLabel ? '<span class="penyelia-hint">Label (Penyelia Utama / Bersama / Pengerusi) dijana automatik.</span>' : '') +
    '</div>';
}

/* ---------- EVENT BINDING ---------- */
function bindContent() {
  const c = document.getElementById('content');

  c.querySelectorAll('[data-go]').forEach(function (b) {
    b.addEventListener('click', function () {
      const focus = b.dataset.focus;
      go(b.dataset.go);
      if (focus) setTimeout(function () { focusRecord(focus); }, 80);
    });
  });

  c.querySelectorAll('[data-scroll]').forEach(function (b) {
    b.addEventListener('click', function () {
      const el = document.getElementById('sec-' + b.dataset.scroll);
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
  });

  c.querySelectorAll('[data-add-reminder]').forEach(function (b) { b.addEventListener('click', function () { openReminderModal(); }); });

  c.querySelectorAll('[data-reminder-done]').forEach(function (b) {
    b.addEventListener('click', function () {
      const r = state.reminders.find(function (x) { return x.id === b.dataset.reminderDone; });
      if (r) { r.done = true; save(); render(); toast('Peringatan ditandakan selesai', 'success'); }
    });
  });

  c.querySelectorAll('[data-add-row]').forEach(function (b) { b.addEventListener('click', function () { addRow(b.dataset.addRow); }); });
  c.querySelectorAll('[data-edit-row]').forEach(function (b) { b.addEventListener('click', function () { startEdit(b.dataset.editRow); }); });
  c.querySelectorAll('[data-row-menu]').forEach(function (b) { b.addEventListener('click', function (e) { e.stopPropagation(); openRowMenu(b, b.dataset.rowMenu); }); });
  c.querySelectorAll('[data-save-row]').forEach(function (b) { b.addEventListener('click', function () { requestSave(b.dataset.saveRow); }); });
  c.querySelectorAll('[data-cancel-row]').forEach(function (b) { b.addEventListener('click', function () { cancelEdit(); }); });

  c.querySelectorAll('[data-sortable]').forEach(function (th) {
    th.addEventListener('click', function () {
      const parts = th.dataset.sortable.split('|');
      const wsId = parts[0], key = parts[1];
      const cur = ui.sort[wsId];
      ui.sort[wsId] = { key: key, dir: cur && cur.key === key && cur.dir === 'asc' ? 'desc' : 'asc' };
      render();
    });
  });

  c.querySelectorAll('[data-chip]').forEach(function (b) {
    b.addEventListener('click', function () {
      const parts = b.dataset.chip.split('|');
      ui.filter[parts[0]] = parts[1];
      ui.page[parts[0]] = 1;
      render();
    });
  });

  c.querySelectorAll('[data-page]').forEach(function (b) {
    b.addEventListener('click', function () {
      const parts = b.dataset.page.split('|');
      ui.page[parts[0]] = parseInt(parts[1], 10);
      render();
      window.scrollTo(0, 0);
    });
  });

  c.querySelectorAll('[data-export]').forEach(function (b) { b.addEventListener('click', function () { exportWorkspace(b.dataset.export); }); });

  /* Senarai penyelia: tambah / buang baris (dalam mod edit) */
  c.querySelectorAll('[data-add-penyelia]').forEach(function (b) {
    b.addEventListener('click', function () {
      const wrap = b.closest('.penyelia-list');
      if (!wrap) return;
      const rows = wrap.querySelector('.penyelia-rows');
      const div = document.createElement('div');
      div.className = 'penyelia-row';
      div.innerHTML = '<input type="text" class="cell-input penyelia-input" data-field="' + b.dataset.addPenyelia + '" data-multi="1" value="" placeholder="Nama penyelia…">' +
        '<button type="button" class="penyelia-remove" data-remove-penyelia title="Buang">' + icon('x', 13) + '</button>';
      rows.appendChild(div);
      bindPenyeliaRemove(div.querySelector('[data-remove-penyelia]'));
      div.querySelector('input').focus();
    });
  });
  function bindPenyeliaRemove(btn) {
    if (!btn) return;
    btn.addEventListener('click', function () {
      const row = btn.closest('.penyelia-row');
      const wrap = btn.closest('.penyelia-list');
      if (!wrap) return;
      const rows = wrap.querySelectorAll('.penyelia-row');
      if (rows.length <= 1) { row.querySelector('input').value = ''; return; }
      row.remove();
    });
  }
  c.querySelectorAll('[data-remove-penyelia]').forEach(bindPenyeliaRemove);

  /* Panel Rekod Dikunci: buka/tutup */
  c.querySelectorAll('[data-toggle-locked]').forEach(function (b) {
    b.addEventListener('click', function () {
      const id = b.dataset.toggleLocked;
      ui.lockedOpen[id] = !ui.lockedOpen[id];
      render();
    });
  });

  const searchInput = c.querySelector('[data-ws-search]');
  if (searchInput) {
    let t = null;
    searchInput.addEventListener('input', function () {
      const id = searchInput.dataset.wsSearch;
      clearTimeout(t);
      t = setTimeout(function () {
        ui.query[id] = searchInput.value;
        ui.page[id] = 1;
        const pos = searchInput.selectionStart;
        render();
        const ni = document.querySelector('[data-ws-search="' + id + '"]');
        if (ni) { ni.focus(); ni.setSelectionRange(pos, pos); }
      }, 200);
    });
  }
  c.querySelectorAll('[data-ws-clear]').forEach(function (b) {
    b.addEventListener('click', function () { ui.query[b.dataset.wsClear] = ''; ui.page[b.dataset.wsClear] = 1; render(); });
  });

  /* ---------- ADMIN binding ---------- */
  const adminBind = {
    'set-orgName': function (v) { ensureSettings().orgName = v; },
    'set-orgSub': function (v) { ensureSettings().orgSub = v; },
    'set-showDemoBadge': function (v) { ensureSettings().showDemoBadge = v === '1'; },
    'set-operatorName': function (v) { ensureSettings().operatorName = v; },
    'set-operatorRole': function (v) { ensureSettings().operatorRole = v; },
    'set-operatorInitials': function (v) { ensureSettings().operatorInitials = (v || 'PN').slice(0, 2).toUpperCase(); },
    'set-pageSize': function (v) { const n = Math.max(5, Math.min(200, parseInt(v, 10) || 25)); ensureSettings().pageSize = n; ui.pageSize = n; },
    'set-upcomingDays': function (v) { ensureSettings().attentionThresholds.upcomingDays = Math.max(1, parseInt(v, 10) || 60); },
    'set-vivaSoonDays': function (v) { ensureSettings().attentionThresholds.vivaSoonDays = Math.max(1, parseInt(v, 10) || 14); },
    'set-trashRetentionDays': function (v) { ensureSettings().trashRetentionDays = Math.max(1, parseInt(v, 10) || 30); },
    'set-confirmDelete': function (v) { ensureSettings().confirmDelete = v === '1'; },
    'set-sesiAktif': function (v) { ensureSettings().semesterAktif.sesi = v; }
  };
  Object.keys(adminBind).forEach(function (id) {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('change', function () {
      adminBind[id](el.value);
      save();
      applySettingsToChrome();
      renderSidebar();
      toast('Tetapan disimpan', 'success');
    });
  });

  c.querySelectorAll('[data-rule]').forEach(function (chk) {
    chk.addEventListener('change', function () {
      ensureSettings().attentionRules[chk.dataset.rule] = chk.checked;
      save();
      renderSidebar();
      toast('Peraturan perhatian dikemas kini', 'success');
    });
  });

  c.querySelectorAll('[data-restore]').forEach(function (b) { b.addEventListener('click', function () { restoreTrash(b.dataset.restore); }); });
  c.querySelectorAll('[data-purge]').forEach(function (b) { b.addEventListener('click', function () { purgeTrash(b.dataset.purge); }); });
  c.querySelectorAll('[data-empty-trash]').forEach(function (b) { b.addEventListener('click', function () { emptyTrash(); }); });
  c.querySelectorAll('[data-reset-demo]').forEach(function (b) { b.addEventListener('click', function () { resetDemo(); }); });
  c.querySelectorAll('[data-export-all]').forEach(function (b) { b.addEventListener('click', function () { exportAllJson(); }); });

  /* ---------- PENYELIA-PELAJAR binding ---------- */
  c.querySelectorAll('[data-toggle-penyelia]').forEach(function (b) {
    b.addEventListener('click', function () {
      const id = b.dataset.togglePenyelia;
      ui.penyeliaOpen[id] = (ui.penyeliaOpen[id] === undefined) ? false : !ui.penyeliaOpen[id];
      render();
    });
  });
  c.querySelectorAll('[data-toggle-lepas]').forEach(function (b) {
    b.addEventListener('click', function () {
      const id = b.dataset.toggleLepas;
      ui.lepasOpen[id] = !ui.lepasOpen[id];
      render();
    });
  });
  c.querySelectorAll('[data-seg]').forEach(function (b) {
    b.addEventListener('click', function () {
      ui.filter['penyeliaPelajar'] = b.dataset.seg === 'semua' ? 'semua' : 'aktif';
      render();
    });
  });
  c.querySelectorAll('[data-add-penyelia]').forEach(function (b) { b.addEventListener('click', function () { addPenyelia(); }); });
  c.querySelectorAll('[data-edit-pelajar]').forEach(function (b) {
    b.addEventListener('click', function () {
      const parts = b.dataset.editPelajar.split('|');
      openPelajarDrawer(parts[0], parts[1]);
    });
  });
  c.querySelectorAll('[data-status-pelajar]').forEach(function (b) {
    b.addEventListener('click', function () {
      const parts = b.dataset.statusPelajar.split('|');
      openStatusQuick(parts[0], parts[1]);
    });
  });
}

/* ---------- EDIT FLOW ---------- */
function startEdit(encoded) {
  const parts = encoded.split('|');
  const wsId = parts[0], id = parts[1];
  const rec = (state.records[wsId] || []).find(function (r) { return r.id === id; });
  if (!rec) return;
  if (rec.locked) {
    /* Rekod dikunci: minta pengesahan buka kunci dahulu, kemudian terus masuk edit (pilihan a) */
    requestUnlock(wsId, rec, function () {
      ui.editingRow = id;
      ui.draft = JSON.parse(JSON.stringify(rec));
      render();
      focusRecord(id);
    });
    return;
  }
  ui.editingRow = id;
  ui.draft = JSON.parse(JSON.stringify(rec));
  render();
}

/* Kunci rekod */
function lockRecord(wsId, id) {
  const rec = (state.records[wsId] || []).find(function (r) { return r.id === id; });
  if (!rec) return;
  rec.locked = true;
  rec.lockedAt = new Date().toISOString();
  save();
  render();
  toast('Rekod dikunci — dipindahkan ke senarai Rekod Dikunci', 'success');
}

/* Minta pengesahan untuk buka kunci. onDone dipanggil selepas berjaya. */
function requestUnlock(wsId, rec, onDone) {
  openModal({
    title: 'Buka Kunci Rekod',
    sub: esc(rec.nama || rec.perkara || 'rekod ini'),
    body: '<div class="toast-inline" style="color:#b45309">' + icon('alert', 16) + '<span>Rekod ini dikunci. Buka kunci hanya jika anda benar-benar perlu mengeditnya.</span></div>',
    foot: '<button class="btn btn--ghost" data-close-modal>Batal</button><button class="btn btn--primary" id="confirmUnlock">' + icon('lockOpen', 16) + ' Buka Kunci</button>',
    onMount: function (root) {
      root.querySelector('#confirmUnlock').addEventListener('click', function () {
        closeModal();
        rec.locked = false;
        delete rec.lockedAt;
        save();
        if (typeof onDone === 'function') onDone();
        else { render(); toast('Rekod dibuka kunci', 'success'); }
      });
    }
  });
}

function cancelEdit() {
  if (ui.draft) {
    const changed = hasChanges();
    if (changed && !confirm('Ada perubahan belum disimpan. Batal dan buang perubahan?')) return;
  }
  ui.editingRow = null;
  ui.draft = null;
  render();
  toast('Perubahan dibatalkan');
}

function hasChanges() {
  if (!ui.draft) return false;
  const w = ws(ui.view);
  const rec = (state.records[w.id] || []).find(function (r) { return r.id === ui.editingRow; });
  if (!rec) return false;
  return w.columns.some(function (c) { return !c.noEdit && String(ui.draft[c.key] || '') !== String(rec[c.key] || ''); });
}

function collectDraftEdits() {
  const c = document.getElementById('content');
  if (!ui.draft) return;
  const w = ws(ui.view);
  const multiKeys = {};
  (w.columns || []).forEach(function (col) { if (col.multi) multiKeys[col.key] = col; });

  c.querySelectorAll('tr.is-editing [data-field]').forEach(function (inp) {
    const key = inp.dataset.field;
    if (multiKeys[key]) return; /* dikendalikan di bawah */
    ui.draft[key] = inp.value;
  });

  /* Kolum penyelia: baca senarai, jana label jika autoLabel */
  Object.keys(multiKeys).forEach(function (key) {
    const names = readMultiValue(key);
    if (names == null) return;
    const col = multiKeys[key];
    ui.draft[key] = col.autoLabel ? formatPenyelia(names) : names.join('\n');
  });
}

function requestSave(encoded) {
  collectDraftEdits();
  const parts = encoded.split('|');
  const wsId = parts[0], id = parts[1];
  const w = ws(wsId);
  const rec = (state.records[wsId] || []).find(function (r) { return r.id === id; });
  if (!rec || !ui.draft) return;

  const errors = [];
  w.columns.forEach(function (col) {
    if (col.required && !String(ui.draft[col.key] || '').trim()) errors.push(col.header + ' wajib diisi');
  });

  const diffs = [];
  w.columns.forEach(function (col) {
    if (col.noEdit) return;
    const oldV = String(rec[col.key] || '');
    const newV = String(ui.draft[col.key] || '');
    if (oldV !== newV) diffs.push({ key: col.key, header: col.header, oldV: oldV, newV: newV, important: isImportantField(col) });
  });

  if (errors.length) {
    openModal({
      title: 'Tidak boleh disimpan',
      body: '<div class="toast-inline">' + icon('alert', 16) + '<span>' + esc(errors.join(' · ')) + '</span></div>',
      foot: '<button class="btn btn--ghost" data-close-modal>Kembali Edit</button>'
    });
    return;
  }

  if (!diffs.length) { ui.editingRow = null; ui.draft = null; render(); toast('Tiada perubahan'); return; }

  const important = diffs.some(function (d) { return d.important; });
  if (important) { openDiffReview(wsId, id, rec, diffs); }
  else commitEdits(wsId, id, rec, diffs);
}

function isImportantField(col) {
  const h = col.header.toLowerCase();
  return h.indexOf('tarikh') !== -1 || h.indexOf('nama') !== -1 || h.indexOf('pelajar') !== -1 || h.indexOf('jumlah') !== -1 || h.indexOf('matrik') !== -1;
}

function openDiffReview(wsId, id, rec, diffs) {
  let body = '<p class="field__hint" style="margin-top:0">Sila semak perubahan berikut sebelum menyimpan.</p><div class="diff-list">';
  diffs.forEach(function (d) {
    body += '<div class="diff-row"><div class="diff-row__label">' + esc(d.header) + '</div><div class="diff-row__vals">';
    body += '<span class="diff-old">' + (esc(d.oldV) || '<em style="opacity:.6">kosong</em>') + '</span>';
    body += '<span class="diff-arrow">→</span>';
    body += '<span class="diff-new">' + (esc(d.newV) || '<em style="opacity:.6">kosong</em>') + '</span>';
    body += '</div></div>';
  });
  body += '</div>';

  openModal({
    title: 'Semak Perubahan',
    sub: 'Medan penting telah diubah',
    wide: true,
    body: body,
    foot: '<button class="btn btn--ghost" data-close-modal>Kembali Edit</button><button class="btn btn--primary" id="confirmSave">' + icon('save', 16) + ' Sahkan & Simpan</button>',
    onMount: function (root) {
      root.querySelector('#confirmSave').addEventListener('click', function () {
        closeModal();
        commitEdits(wsId, id, rec, diffs);
      });
    }
  });
}

function commitEdits(wsId, id, rec, diffs) {
  diffs.forEach(function (d) {
    state.history.unshift({ id: uid('h'), wsId: wsId, recordId: id, field: d.header, oldVal: d.oldV, newVal: d.newV, at: new Date().toISOString() });
    rec[d.key] = d.newV;
  });
  if (state.history.length > 300) state.history.length = 300;
  save();
  ui.editingRow = null;
  ui.draft = null;
  render();
  toast(diffs.length + ' perubahan disimpan', 'success');
}

/* ---------- ADD ROW ---------- */
function addRow(wsId) {
  const w = ws(wsId);
  const rec = { id: uid('r'), bil: nextId(wsId) };
  w.columns.forEach(function (c) { if (!c.noEdit && !(c.key in rec)) rec[c.key] = ''; });
  state.records[wsId] = state.records[wsId] || [];
  state.records[wsId].push(rec);
  save();
  ui.editingRow = rec.id;
  ui.draft = JSON.parse(JSON.stringify(rec));
  /* Navigate ke halaman terakhir supaya rekod baharu kelihatan */
  const total = state.records[wsId].length;
  ui.page[wsId] = Math.max(1, Math.ceil(total / ui.pageSize));
  render();
  setTimeout(function () { focusRecord(rec.id); }, 120);
  toast('Rekod baharu ditambah — sila isi dan simpan');
}

/* ---------- DELETE (BOLEH PULIH) ---------- */
function deleteRow(wsId, id) {
  const st = ensureSettings();
  const rec = (state.records[wsId] || []).find(function (r) { return r.id === id; });
  if (!rec) return;
  const label = rec.nama || rec.perkara || 'rekod ini';
  const retain = st.trashRetentionDays;

  if (st.confirmDelete) {
    openModal({
      title: 'Padamkan Rekod',
      sub: esc(label),
      body: '<div class="toast-inline" style="color:#b45309">' + icon('alert', 16) + '<span>Rekod akan dipadamkan daripada senarai, tetapi <strong>disimpan dalam Arkib Padaman selama ' + retain + ' hari</strong> dan boleh dipulihkan.</span></div>' +
        '<p class="field__hint" style="margin:12px 0 0">Anda boleh rujuk semula di <strong>Admin &amp; Tetapan → Arkib Padaman</strong>.</p>',
      foot: '<button class="btn btn--ghost" data-close-modal>Batal</button><button class="btn btn--danger" id="confirmDelete">' + icon('trash', 16) + ' Padamkan</button>',
      onMount: function (root) {
        root.querySelector('#confirmDelete').addEventListener('click', function () {
          closeModal();
          doDelete(wsId, id, rec);
        });
      }
    });
  } else {
    doDelete(wsId, id, rec);
  }
}

function doDelete(wsId, id, rec) {
  state.trash.unshift({ id: uid('t'), wsId: wsId, record: rec, at: new Date().toISOString() });
  state.records[wsId] = state.records[wsId].filter(function (r) { return r.id !== id; });
  save();
  render();
  toast('Rekod dipadamkan — boleh dipulihkan di Arkib Padaman', 'warn');
}

/* Pulihkan rekod yang dipadam */
function restoreTrash(trashId) {
  const item = (state.trash || []).find(function (t) { return t.id === trashId; });
  if (!item) return;
  state.records[item.wsId] = state.records[item.wsId] || [];
  state.records[item.wsId].push(item.record);
  state.trash = state.trash.filter(function (t) { return t.id !== trashId; });
  save();
  render();
  toast('Rekod dipulihkan', 'success');
}

/* Padamkan kekal (buang dari arkib) */
function purgeTrash(trashId) {
  if (!confirm('Padamkan kekal? Tindakan ini tidak boleh dibatalkan.')) return;
  state.trash = (state.trash || []).filter(function (t) { return t.id !== trashId; });
  save();
  render();
  toast('Rekod dipadamkan kekal', 'warn');
}

function emptyTrash() {
  if (!(state.trash || []).length) return;
  if (!confirm('Kosongkan seluruh Arkib Padaman? Semua rekod dalam arkib akan dipadamkan kekal.')) return;
  state.trash = [];
  save();
  render();
  toast('Arkib Padaman dikosongkan', 'warn');
}


/* ---------- ROW MENU ---------- */
function openRowMenu(btn, encoded) {
  closeRowMenu();
  const parts = encoded.split('|');
  const wsId = parts[0], id = parts[1];
  const rec = (state.records[wsId] || []).find(function (r) { return r.id === id; });
  if (!rec) return;

  const menu = document.createElement('div');
  menu.className = 'row-menu';
  menu.id = 'rowMenu';
  menu.innerHTML =
    (rec.locked
      ? '<button data-act="unlock" style="color:var(--warning)">' + icon('lockOpen', 15) + ' Buka Kunci</button>'
      : '<button data-act="lock">' + icon('lock', 15) + ' Kunci Rekod</button>') +
    '<button data-act="copy">' + icon('copy', 15) + ' Salin ke Workspace</button>' +
    '<button data-act="reminder">' + icon('bell', 15) + ' Tetapkan Peringatan</button>' +
    '<button data-act="history">' + icon('history', 15) + ' Sejarah Perubahan</button>' +
    '<button data-act="delete" style="color:var(--destructive)">' + icon('trash', 15) + ' Padamkan</button>';

  document.body.appendChild(menu);
  const r = btn.getBoundingClientRect();
  const mw = menu.offsetWidth;
  let left = Math.min(r.left, window.innerWidth - mw - 12);
  let top = r.bottom + 6;
  if (top + menu.offsetHeight > window.innerHeight - 12) top = r.top - menu.offsetHeight - 6;
  menu.style.left = Math.max(12, left) + 'px';
  menu.style.top = Math.max(12, top) + 'px';

  menu.querySelectorAll('button').forEach(function (b) {
    b.addEventListener('click', function () {
      closeRowMenu();
      const act = b.dataset.act;
      if (act === 'lock') lockRecord(wsId, id);
      else if (act === 'unlock') requestUnlock(wsId, rec);
      else if (act === 'copy') openCopyModal(wsId, rec);
      else if (act === 'reminder') openReminderModal(wsId, rec);
      else if (act === 'history') openHistoryModal(wsId, id);
      else if (act === 'delete') deleteRow(wsId, id);
    });
  });

  setTimeout(function () { document.addEventListener('click', closeRowMenu, { once: true }); }, 0);
}

function closeRowMenu() {
  const m = document.getElementById('rowMenu');
  if (m) m.remove();
}

/* ---------- COPY TO WORKSPACE ---------- */
function openCopyModal(wsId, rec) {
  const targets = WORKSPACES.filter(function (w) { return w.id !== wsId; });
  let body = '<p class="field__hint" style="margin-top:0">Pilih workspace sasaran. Hanya medan pengenalan (nama, no. matrik, program, penyelia, semester) akan dicadangkan untuk disalin.</p>';
  body += '<div class="field"><label>Workspace Sasaran</label><select id="copyTarget">';
  targets.forEach(function (t) { body += '<option value="' + t.id + '">' + esc(t.name) + '</option>'; });
  body += '</select></div>';

  openModal({
    title: 'Salin ke Workspace',
    sub: esc(rec.nama || rec.pelajar || rec.perkara || 'Rekod'),
    wide: true,
    body: body,
    foot: '<button class="btn btn--ghost" data-close-modal>Batal</button><button class="btn btn--primary" id="doCopy">' + icon('copy', 16) + ' Salin Rekod</button>',
    onMount: function (root) {
      root.querySelector('#doCopy').addEventListener('click', function () {
        const targetId = root.querySelector('#copyTarget').value;
        performCopy(wsId, targetId, rec);
      });
    }
  });
}

function performCopy(srcWsId, targetId, rec) {
  const srcW = ws(srcWsId), tgtW = ws(targetId);
  const newRec = { id: uid('r'), bil: nextId(targetId) };
  tgtW.columns.forEach(function (c) { if (!c.noEdit) newRec[c.key] = ''; });

  const copied = [];
  srcW.columns.filter(function (c) { return c.identity; }).forEach(function (sc) {
    const val = String(rec[sc.key] || '').trim();
    if (!val) return;
    const match = tgtW.columns.find(function (tc) {
      return tc.identity && (tc.header.toLowerCase() === sc.header.toLowerCase() || tc.key === sc.key || sameFieldName(tc.key, sc.key));
    });
    if (match) { newRec[match.key] = val; copied.push(match.header); }
  });

  const dup = (state.records[targetId] || []).find(function (r) {
    const srcId = idVal(rec), tgtId2 = idVal(r);
    return srcId && tgtId2 && srcId.toLowerCase() === tgtId2.toLowerCase();
  });

  const body = '<p>Medan yang akan disalin: <strong>' + (copied.length ? esc(copied.join(', ')) : 'tiada medan sepadan') + '</strong></p>' +
    (dup ? '<div class="toast-inline">' + icon('alert', 16) + '<span>Pelajar ini mungkin sudah ada dalam workspace sasaran. Tambah sebagai rekod baharu?</span></div>' : '') +
    '<p class="field__hint">Tarikh proses, kelulusan dan status tidak disalin secara lalai.</p>';

  closeModal();
  openModal({
    title: 'Sahkan Salinan',
    sub: esc(tgtW.name),
    body: body,
    foot: '<button class="btn btn--ghost" data-close-modal>Batal</button><button class="btn btn--primary" id="confirmCopy">Tambah Rekod</button>',
    onMount: function (root) {
      root.querySelector('#confirmCopy').addEventListener('click', function () {
        state.records[targetId] = state.records[targetId] || [];
        state.records[targetId].push(newRec);
        save();
        closeModal();
        go(targetId);
        setTimeout(function () { focusRecord(newRec.id); }, 200);
        toast('Rekod disalin ke ' + tgtW.name, 'success');
      });
    }
  });
}

function idVal(rec) { return rec.noMatrik || rec.noPelajar || ''; }

function sameFieldName(a, b) {
  const map = { nama: 1, noMatrik: 1, noPelajar: 1, program: 1, penyelia: 1, cadanganPenyelia: 1, semester: 1 };
  const na = a.toLowerCase().replace(/[^a-z]/g, '');
  const nb = b.toLowerCase().replace(/[^a-z]/g, '');
  return na === nb && map[a];
}

/* ---------- REMINDER MODAL ---------- */
function openReminderModal(wsId, rec) {
  const body = '<div class="field"><label>Tajuk Peringatan</label><input id="remTitle" placeholder="Contoh: Susulan laporan pemeriksa" value="' + esc(rec ? 'Susulan: ' + (rec.nama || rec.perkara || '') : '') + '"></div>' +
    '<div class="field" style="margin-top:12px"><label>Tarikh</label><input type="date" id="remDate" value="' + todayISO() + '"></div>' +
    '<div class="field" style="margin-top:12px"><label>Keutamaan</label><select id="remPri"><option value="low">Rendah</option><option value="med" selected>Sederhana</option><option value="high">Tinggi</option></select></div>' +
    '<div class="field" style="margin-top:12px"><label>Catatan</label><textarea id="remNote" placeholder="Catatan tindakan…"></textarea></div>';

  openModal({
    title: 'Tetapkan Peringatan',
    body: body,
    foot: '<button class="btn btn--ghost" data-close-modal>Batal</button><button class="btn btn--primary" id="saveRem">' + icon('save', 16) + ' Simpan</button>',
    onMount: function (root) {
      root.querySelector('#saveRem').addEventListener('click', function () {
        const title = root.querySelector('#remTitle').value.trim();
        if (!title) { toast('Tajuk peringatan diperlukan', 'error'); return; }
        state.reminders.unshift({
          id: uid('rm'), wsId: wsId || null, recordId: rec ? rec.id : null,
          title: title, ref: rec ? (rec.nama || rec.perkara || '') : '', date: root.querySelector('#remDate').value,
          priority: root.querySelector('#remPri').value, note: root.querySelector('#remNote').value, done: false
        });
        save();
        closeModal();
        if (ui.view === 'dashboard') render();
        toast('Peringatan ditetapkan', 'success');
      });
    }
  });
}

/* ---------- HISTORY MODAL ---------- */
function openHistoryModal(wsId, id) {
  const items = state.history.filter(function (h) { return h.wsId === wsId && h.recordId === id; });
  let body;
  if (!items.length) body = '<div class="empty-state">' + icon('history', 32) + '<div>Tiada sejarah perubahan bagi rekod ini.</div></div>';
  else {
    body = '<div class="timeline">';
    items.forEach(function (h) {
      body += '<div class="timeline__item"><div class="timeline__time">' + esc(new Date(h.at).toLocaleString('ms-MY')) + '</div>' +
        '<div class="timeline__field">' + esc(h.field) + '</div>' +
        '<div class="timeline__change"><del>' + (esc(h.oldVal) || 'kosong') + '</del> → <ins>' + (esc(h.newVal) || 'kosong') + '</ins></div></div>';
    });
    body += '</div>';
  }
  openModal({ title: 'Sejarah Perubahan', wide: true, body: body, foot: '<button class="btn btn--ghost" data-close-modal>Tutup</button>' });
}

/* ---------- MODAL ---------- */
function openModal(opts) {
  closeModal();
  const overlay = document.createElement('div');
  overlay.className = 'overlay';
  overlay.id = 'overlay';
  overlay.innerHTML =
    '<div class="modal' + (opts.wide ? ' modal--wide' : '') + '" role="dialog" aria-modal="true">' +
    '<div class="modal__head"><div><h3 class="modal__title">' + esc(opts.title || '') + '</h3>' +
    (opts.sub ? '<p class="modal__sub">' + opts.sub + '</p>' : '') + '</div>' +
    '<button class="modal__close" data-close-modal aria-label="Tutup">' + icon('x', 18) + '</button></div>' +
    '<div class="modal__body">' + opts.body + '</div>' +
    (opts.foot ? '<div class="modal__foot">' + opts.foot + '</div>' : '') +
    '</div>';
  document.body.appendChild(overlay);

  overlay.querySelectorAll('[data-close-modal]').forEach(function (b) { b.addEventListener('click', closeModal); });
  overlay.addEventListener('click', function (e) { if (e.target === overlay) closeModal(); });
  document.addEventListener('keydown', escClose);
  if (opts.onMount) opts.onMount(overlay);
}

function escClose(e) { if (e.key === 'Escape') closeModal(); }

function closeModal() {
  const o = document.getElementById('overlay');
  if (o) o.remove();
  document.removeEventListener('keydown', escClose);
}

/* ---------- DRAWER ---------- */
function openDrawer(opts) {
  closeDrawer();
  const overlay = document.createElement('div');
  overlay.className = 'overlay';
  overlay.id = 'drawerOverlay';
  overlay.innerHTML =
    '<div class="drawer" role="dialog" aria-modal="true">' +
    '<div class="drawer__head"><div><h3 class="modal__title">' + esc(opts.title || '') + '</h3>' +
    (opts.sub ? '<p class="modal__sub">' + esc(opts.sub) + '</p>' : '') + '</div>' +
    '<button class="modal__close" data-close-drawer aria-label="Tutup">' + icon('x', 18) + '</button></div>' +
    '<div class="drawer__body">' + opts.body + '</div>' +
    (opts.foot ? '<div class="drawer__foot">' + opts.foot + '</div>' : '') +
    '</div>';
  document.body.appendChild(overlay);

  overlay.querySelectorAll('[data-close-drawer]').forEach(function (b) { b.addEventListener('click', closeDrawer); });
  overlay.addEventListener('click', function (e) { if (e.target === overlay) closeDrawer(); });
  document.addEventListener('keydown', escCloseDrawer);
  if (opts.onMount) opts.onMount(overlay);
}

function escCloseDrawer(e) { if (e.key === 'Escape') closeDrawer(); }

function closeDrawer() {
  const o = document.getElementById('drawerOverlay');
  if (o) o.remove();
  document.removeEventListener('keydown', escCloseDrawer);
}

/* ---------- EXPORT ---------- */
function exportWorkspace(wsId) {
  const w = ws(wsId);
  const rows = state.records[wsId] || [];
  const head = w.columns.map(function (c) { return '"' + c.header.replace(/"/g, '""') + '"'; }).join(',');
  const body = rows.map(function (r) {
    return w.columns.map(function (c) {
      const v = String(r[c.key] || '').replace(/"/g, '""');
      return '"' + v + '"';
    }).join(',');
  }).join('\n');
  const csv = '\uFEFF' + head + '\n' + body;
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = wsId + '_demo.csv';
  a.click();
  URL.revokeObjectURL(a.href);
  toast('Eksport ' + w.name + ' dimuat turun', 'success');
}

/* Sandaran penuh semua workspace + tetapan (JSON) */
function exportAllJson() {
  const payload = {
    exportedAt: new Date().toISOString(),
    app: 'Sistem Pengurusan Siswazah (Demo)',
    settings: ensureSettings(),
    records: state.records,
    reminders: state.reminders,
    history: state.history,
    trash: state.trash
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'sps_sandaran_penuh_' + todayISO() + '.json';
  a.click();
  URL.revokeObjectURL(a.href);
  toast('Sandaran penuh dimuat turun', 'success');
}

/* ---------- FOCUS RECORD ---------- */
function focusRecord(id) {
  const el = document.querySelector('tr[data-record="' + id + '"]');
  if (el) {
    el.classList.add('is-target');
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    setTimeout(function () { el.classList.remove('is-target'); }, 1900);
  } else {
    toast('Rekod dijumpai tetapi tidak dipaparkan — tukar carian/penapis', 'warn');
  }
}

/* ---------- GLOBAL SEARCH ---------- */
function setupGlobalSearch() {
  const input = document.getElementById('globalSearch');
  const res = document.getElementById('globalResults');
  let t = null;

  input.addEventListener('input', function () {
    clearTimeout(t);
    t = setTimeout(function () { runGlobalSearch(input.value.trim()); }, 200);
  });

  document.addEventListener('click', function (e) {
    if (!document.getElementById('globalSearchWrap').contains(e.target)) {
      res.classList.add('hidden');
    }
  });

  function runGlobalSearch(q) {
    if (!q || q.length < 2) { res.classList.add('hidden'); res.innerHTML = ''; return; }
    const ql = q.toLowerCase();
    const hits = [];
    WORKSPACES.forEach(function (w) {
      (state.records[w.id] || []).forEach(function (r) {
        const name = String(r.nama || '').toLowerCase();
        const idv = String(idVal(r) || '').toLowerCase();
        if (name.indexOf(ql) !== -1 || idv.indexOf(ql) !== -1) {
          hits.push({ wsId: w.id, wsName: w.name, id: r.id, name: r.nama || '(tiada nama)', meta: idVal(r) || '' });
        }
      });
    });
    if (!hits.length) {
      res.innerHTML = '<div class="search-result" style="cursor:default">Tiada padanan dijumpai</div>';
      res.classList.remove('hidden');
      return;
    }
    res.innerHTML = hits.slice(0, 30).map(function (h) {
      return '<button class="search-result" data-gresult="' + h.wsId + '|' + h.id + '"><div class="search-result__ws">' + esc(h.wsName) + '</div>' +
        '<div class="search-result__name">' + esc(h.name) + '</div>' +
        (h.meta ? '<div class="search-result__meta">' + esc(h.meta) + '</div>' : '') + '</button>';
    }).join('');
    res.classList.remove('hidden');
    res.querySelectorAll('[data-gresult]').forEach(function (b) {
      b.addEventListener('click', function () {
        const parts = b.dataset.gresult.split('|');
        ui.query[parts[0]] = '';
        ui.filter[parts[0]] = 'all';
        res.classList.add('hidden');
        input.value = '';
        go(parts[0]);
        setTimeout(function () { focusRecord(parts[1]); }, 100);
      });
    });
  }
}

/* ---------- SIDEBAR TOGGLE ---------- */
function openSidebar() { document.getElementById('sidebar').classList.add('is-open'); document.getElementById('scrim').classList.add('is-open'); }
function closeSidebar() { document.getElementById('sidebar').classList.remove('is-open'); document.getElementById('scrim').classList.remove('is-open'); }

/* ---------- RESET ---------- */
function resetDemo() {
  if (!confirm('Reset semua data demo kepada keadaan asal? Perubahan rekod akan dibuang. Tetapan sistem akan dikekalkan.')) return;
  const keep = ensureSettings();
  localStorage.removeItem(STORAGE_KEY);
  state = { records: {}, reminders: [], history: [], trash: [], lastVisit: null, seq: 1, settings: keep };
  seedData();
  state.settings = keep;
  ui.view = 'dashboard'; ui.editingRow = null; ui.draft = null;
  ui.sort = {}; ui.filter = {}; ui.query = {}; ui.page = {};
  ui.pageSize = keep.pageSize;
  save();
  render();
  toast('Data demo telah direset', 'success');
}

/* ---------- INIT ---------- */
function init() {
  const had = load();
  if (!had) seedData();
  ensureSettings();
  ui.pageSize = state.settings.pageSize || 25;

  document.getElementById('menuBtn').addEventListener('click', openSidebar);
  document.getElementById('scrim').addEventListener('click', closeSidebar);
  document.getElementById('resetDemoBtn').addEventListener('click', resetDemo);

  setupGlobalSearch();
  render();
}

document.addEventListener('DOMContentLoaded', init);
