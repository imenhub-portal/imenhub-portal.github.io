/* Nested workspace, non-destructive CSV merge and student identity helpers.
   Classic script intentionally supports both file:// and GitHub Pages. */
function identityKey(w) { return w.columns.some(c => c.key === 'noMatrik') ? 'noMatrik' : 'noPelajar'; }
function visibleColumns(w) { return w.columns.filter(c => !['noMatrik', 'noPelajar'].includes(c.key)); }
function studentIdentity(r) {
  const match = String(r.nama || '').match(/\s*\(\s*(P\d+)\s*\)\s*$/i);
  const id = r.noMatrik || r.noPelajar || (match ? match[1] : '');
  return { nama: match ? r.nama.slice(0, match.index).trim() : r.nama || '', id: String(id).trim() };
}
function studentIdentityHTML(r) {
  const value = studentIdentity(r);
  return '<span class="student-name">' + esc(value.nama) + '</span><span class="student-id">' + esc(value.id) + '</span>';
}
function normalizedMatric(r) {
  const value = studentIdentity(r).id.toUpperCase().replace(/\s/g, '');
  return /^P\d+$/.test(value) ? value : /^\d{5,6}$/.test(value) ? 'P' + value : ''; // Legacy short matric; never a 12-digit IC.
}

/* ---------- REGISTRI PELAJAR (konsistensi merentas worksheet) ----------
   Sumber kebenaran: workspace penyeliaPelajar (nama, program, penyelia).
   Kunci: normalizedMatric. Digunakan untuk autofill + amaran konflik. */
function buildStudentRegistry() {
  const reg = {};
  (state.records.penyeliaPelajar || []).forEach(function (g) {
    (g.pelajar || []).forEach(function (p) {
      const key = normalizedMatric(p);
      if (!key) return;
      const nama = studentIdentity(p).nama;
      if (!reg[key]) reg[key] = { nama: nama, program: p.program || '', penyelia: g.namaPenyelia, penyeliaBersama: p.penyeliaBersama || '', namaAliases: {} };
      if (reg[key].nama && nama && reg[key].nama !== nama) reg[key].namaAliases[nama] = true;
      if (!reg[key].program && p.program) reg[key].program = p.program;
    });
  });
  return reg;
}

/* Cari rekod kanonikal untuk satu matrik (dari teks). */
function lookupStudent(matric) {
  const key = normalizedMatric({ noMatrik: matric });
  if (!key) return null;
  return buildStudentRegistry()[key] || null;
}

/* Nilai kanonikal untuk cadangan (nama/program) bagi matrik tertentu. */
function registrySuggestion(matric) {
  const hit = lookupStudent(matric);
  return hit ? { nama: hit.nama, program: hit.program, penyeliaBersama: hit.penyeliaBersama } : null;
}

/* Semak konflik: matrik sama tetapi nama/program berbeza daripada registri. */
function registryConflict(matric, nama, program) {
  const hit = lookupStudent(matric);
  if (!hit) return null;
  const conflicts = [];
  const cleanNama = String(nama || '').trim();
  const cleanProg = String(program || '').trim();
  if (hit.nama && cleanNama && normalizeName(hit.nama) !== normalizeName(cleanNama)) {
    conflicts.push({ field: 'NAMA', canonical: hit.nama, entered: cleanNama });
  }
  if (hit.program && cleanProg && hit.program !== cleanProg) {
    conflicts.push({ field: 'PROGRAM', canonical: hit.program, entered: cleanProg });
  }
  return conflicts.length ? { ref: hit, conflicts: conflicts } : null;
}

function normalizeName(s) {
  return String(s || '').toUpperCase().replace(/\s+/g, ' ').replace(/[^A-Z0-9 ]/g, '').trim();
}

/* Pancarkan (sync) nilai kanonikal ke SEMUA rekod pelajar sama dalam semua worksheet.
   Hanya mengemas kini bila matrik padan dan medan berkenaan wujud. Pulangkan bilangan dikemas kini. */
function propagateCanonical(matric, canonical) {
  const key = normalizedMatric({ noMatrik: matric });
  if (!key) return 0;
  let changed = 0;
  WORKSPACES.forEach(function (w) {
    if (w.id === 'penyeliaPelajar') return;
    const hasNama = w.columns.some(function (c) { return c.key === 'nama'; });
    const hasProgram = w.columns.some(function (c) { return c.key === 'program'; });
    (state.records[w.id] || []).forEach(function (r) {
      if (normalizedMatric(r) !== key) return;
      if (hasNama && canonical.nama && studentIdentity(r).nama !== canonical.nama) {
        r._legacyNama = r._legacyNama || r.nama; r.nama = canonical.nama; changed++;
      }
      if (hasProgram && canonical.program && r.program !== canonical.program) {
        r._legacyProgram = r._legacyProgram || r.program; r.program = canonical.program; changed++;
      }
    });
  });
  return changed;
}

/* Amaran penuh: kumpul konflik dalam semua worksheet terhadap registri. */
function registryAudit() {
  const reg = buildStudentRegistry();
  const issues = [];
  WORKSPACES.forEach(function (w) {
    if (w.id === 'penyeliaPelajar') return;
    const hasNama = w.columns.some(function (c) { return c.key === 'nama'; });
    const hasProgram = w.columns.some(function (c) { return c.key === 'program'; });
    (state.records[w.id] || []).forEach(function (r) {
      const key = normalizedMatric(r);
      if (!key || !reg[key]) return;
      if (hasNama) {
        const entered = studentIdentity(r).nama;
        if (reg[key].nama && entered && normalizeName(reg[key].nama) !== normalizeName(entered)) {
          issues.push({ wsId: w.id, wsName: w.name, recordId: r.id, field: 'NAMA', entered: entered, canonical: reg[key].nama, matric: key });
        }
      }
      if (hasProgram && r.program && reg[key].program && r.program !== reg[key].program) {
        issues.push({ wsId: w.id, wsName: w.name, recordId: r.id, field: 'PROGRAM', entered: r.program, canonical: reg[key].program, matric: key });
      }
    });
  });
  return issues;
}
function searchableRecords(w) {
  return w.nested ? (state.records[w.id] || []).flatMap(g => [{id:g.id,nama:g.namaPenyelia}, ...(g.pelajar || [])]) : state.records[w.id] || [];
}
function cloneData(value) { return JSON.parse(JSON.stringify(value)); }
function programValue(value) {
  const v = String(value || '').trim();
  if (!v) return '';
  if (/^(SARJANA|SARJANA SAINS|MSc|Master|Sarjana)$/i.test(v) || /sarjana/i.test(v)) return PROGRAM_OPTIONS[1];
  if (/^(KEDOKTORAN|PhD|Doktor Falsafah)$/i.test(v) || /doktor|phd/i.test(v)) return PROGRAM_OPTIONS[0];
  return v;
}
WORKSPACES.forEach(w => {
  if (w.columns.some(c => c.key === 'nama') && !w.columns.some(c => ['noMatrik','noPelajar'].includes(c.key))) {
    w.columns.splice(2, 0, {key:'noPelajar',header:'NO. PELAJAR',type:'text',identity:true});
  }
  w.columns.forEach(c => {
    if (['penyelia','cadanganPenyelia','penyeliaBersama'].includes(c.key)) c.autoLabel = true;
    if (c.key === 'program') { c.type = 'select'; c.options = PROGRAM_OPTIONS; }
  });
});

function prepareStudentData(had) {
  WORKSPACES.forEach(w => {
    state.records[w.id] = state.records[w.id] || [];
    searchableRecords(w).forEach(r => {
      if (!r.nama) return;
      const identity = studentIdentity(r);
      if (identity.nama !== r.nama) {
        r._legacyNama = r._legacyNama || r.nama;
        r.nama = identity.nama;
        if (!r[identityKey(w)]) r[identityKey(w)] = identity.id;
      }
      if (r.program) {
        const mapped = programValue(r.program);
        if (mapped !== r.program) { r._legacyProgram = r._legacyProgram || r.program; r.program = mapped; }
      }
    });
  });
  if (!had && window.PENYELIA_SEED) {
    state.records.penyeliaPelajar = cloneData(window.PENYELIA_SEED.groups);
    state.penyeliaImportVersion = window.PENYELIA_SEED.version;
    state.penyeliaSourceManifest = cloneData(window.PENYELIA_SEED.sources);
  }
  /* Migrasi label: medan penyeliaBersama mesti sentiasa 'Penyelia Bersama', bukan 'Penyelia Utama'. */
  (state.records.penyeliaPelajar || []).forEach(g => (g.pelajar || []).forEach(p => {
    if (!p.penyeliaBersama) return;
    const fixed = formatPenyeliaBersama(extractPenyeliaNames(p.penyeliaBersama));
    if (fixed && fixed !== p.penyeliaBersama) {
      p._legacyPenyeliaBersama = p._legacyPenyeliaBersama || p.penyeliaBersama;
      p.penyeliaBersama = fixed;
    }
  }));
  save();
}

function supervisorKey(g) {
  const staff = String(g.namaPenyelia || '').match(/\(K\d+\)/i);
  return staff ? staff[0].toUpperCase() : String(g.namaPenyelia || '').toUpperCase().replace(/\s+/g, ' ').trim();
}
function importPlan() {
  const groups = state.records.penyeliaPelajar || [];
  const plan = { groups: [], students: [], matches: [], archived: 0 };
  const archivedRecords = (state.trash || []).filter(t => t.wsId === 'penyeliaPelajar').flatMap(t => [t.record, ...(t.record.pelajar || [])]);
  const archived = new Set(archivedRecords.map(r => r.id));
  const archivedMatric = new Set(archivedRecords.map(normalizedMatric).filter(Boolean));
  const archivedParents = new Set(archivedRecords.filter(r => Array.isArray(r.pelajar)).map(supervisorKey));
  (window.PENYELIA_SEED?.groups || []).forEach(source => {
    if (archived.has(source.id) || archivedParents.has(supervisorKey(source))) { plan.archived++; return; }
    const parent = groups.find(g => g.id === source.id || supervisorKey(g) === supervisorKey(source));
    if (!parent) plan.groups.push(source);
    source.pelajar.forEach(incoming => {
      if (archived.has(incoming.id) || archivedMatric.has(normalizedMatric(incoming))) { plan.archived++; return; }
      const matches = groups.flatMap(g => (g.pelajar || []).map(p => ({parent:g,record:p})))
        .filter(x => x.record.id === incoming.id || (normalizedMatric(incoming) && normalizedMatric(x.record) === normalizedMatric(incoming)));
      if (matches.length) plan.matches.push({incoming, source, matches});
      else plan.students.push({incoming, source, parent});
    });
  });
  return plan;
}
function previewPenyeliaImport() {
  const plan = importPlan();
  const body = '<p>' + plan.groups.length + ' penyelia baharu · ' + plan.students.length + ' pelajar baharu · ' + plan.matches.length + ' padanan sedia ada · ' + plan.archived + ' dalam arkib.</p>' +
    '<p>Rekod sedia ada dan suntingan dikekalkan. Salinan sumber CSV disimpan bersama padanan untuk semakan; nilai sedia ada tidak diganti.</p>' +
    '<details><summary>Laporan gabungan tiga sumber</summary><pre>' + esc(JSON.stringify(window.PENYELIA_SEED.report,null,2)) + '</pre></details>' +
    plan.matches.map((x,index) => '<details><summary>' + esc(x.incoming.noPelajar + ' — ' + x.incoming.nama) + ' · Cadangan status: ' + esc(statusPelajarLabel(x.incoming.statusKhas)) + '</summary>' +
      '<p>Penyelia sumber: ' + esc(x.source.namaPenyelia) + '</p>' +
      (x.matches.length === 1 ? '<label><input type="checkbox" data-adopt-import="' + index + '"> Gunakan versi gabungan CSV dan penyelia sumber untuk rekod ini (versi lama disimpan)</label>' : '<p>Padanan berganda: semak secara manual; tiada rekod diganti.</p>') +
      '<pre>' + esc(JSON.stringify({sediaAda:x.matches.map(m => ({penyelia:m.parent.namaPenyelia,record:m.record})), sumberCSV:x.incoming}, null, 2)) + '</pre></details>').join('');
  openDrawer({title:'Semak Import CSV', body, foot:'<button class="btn btn--ghost" data-close-drawer>Batal</button><button class="btn btn--primary" id="mergeImport">Gabung — kekalkan suntingan</button>',
    onMount(root) { root.querySelector('#mergeImport').onclick = () => { plan.adopt = [...root.querySelectorAll('[data-adopt-import]:checked')].map(b => Number(b.dataset.adoptImport)); mergePenyeliaImport(plan); closeDrawer(); render(); }; }});
}
function mergePenyeliaImport(plan) {
  const groups = state.records.penyeliaPelajar;
  plan.groups.forEach(g => groups.push({...cloneData(g), pelajar:[]}));
  plan.students.forEach(x => {
    const parent = x.parent || groups.find(g => g.id === x.source.id);
    parent.pelajar.push(cloneData(x.incoming));
  });
  plan.matches.forEach((x,index) => x.matches.forEach(m => {
    m.record.csvSources = m.record.csvSources || {};
    m.record.csvSources[window.PENYELIA_SEED.version] = {supervisor: x.source.namaPenyelia, record:cloneData(x.incoming)};
    if ((plan.adopt || []).includes(index) && x.matches.length === 1) {
      const previous = cloneData(m.record);
      delete previous.csvSources;
      delete previous.importBackups;
      const id = m.record.id;
      const backups = m.record.importBackups || [];
      backups.push({at:new Date().toISOString(),parentId:m.parent.id,record:previous});
      Object.assign(m.record,cloneData(x.incoming),{id,importBackups:backups});
      const parent = groups.find(g => supervisorKey(g) === supervisorKey(x.source));
      if (parent && parent.id !== m.parent.id) {
        m.parent.pelajar = m.parent.pelajar.filter(p => p.id !== id);
        parent.pelajar.push(m.record);
      }
      // Import does not delete or overwrite any graduation record.
    }
  }));
  state.penyeliaSourceManifest = cloneData(window.PENYELIA_SEED.sources);
  state.penyeliaImportVersion = window.PENYELIA_SEED.version;
  save();
  toast('Import digabung; rekod dan suntingan lama dikekalkan', 'success');
}

function semesterIndex(s) {
  const m = String(s || '').match(/^([12])\/(\d{4})-\d{4}$/);
  return m ? Number(m[2]) * 2 + Number(m[1]) - 1 : null;
}
function semesterAt(index) { const y = Math.floor(index / 2); return (index % 2 + 1) + '/' + y + '-' + (y + 1); }
function selectedSession() { return ui.penyeliaSession || ensureSettings().semesterAktif.sesi; }
function recentSessions() {
  const index = semesterIndex(selectedSession());
  const recorded = (state.records.penyeliaPelajar || []).flatMap(g => (g.pelajar || []).flatMap(p => (p.sejarahSemester || []).map(s => s.sesi)));
  return [...new Set([...recorded, selectedSession()])].filter(s => semesterIndex(s) !== null && (index === null || semesterIndex(s) <= index))
    .sort((a,b) => semesterIndex(a)-semesterIndex(b)).slice(-3);
}
function allSessions() {
  const activeIndex = semesterIndex(ensureSettings().semesterAktif.sesi);
  const list = [...new Set([...generateSemesterOptions(), ensureSettings().semesterAktif.sesi,
    ...(state.records.penyeliaPelajar || []).flatMap(g => (g.pelajar || []).flatMap(p => (p.sejarahSemester || []).map(s => s.sesi)))])].filter(Boolean)
    .filter(s => semesterIndex(s) !== null && (activeIndex === null || semesterIndex(s) <= activeIndex));
  return list.sort((a,b) => semesterIndex(a) - semesterIndex(b));
}
function snapshot(p, sesi) {
  const found = (p.sejarahSemester || []).find(s => s.sesi === sesi);
  if (found) return found;
  if (p.sesiSemesterPengajian === sesi) return {semesterPengajian:p.semesterPengajian,catatan:''};
  return null;
}
const originalSelectOptions = selectOptions;
selectOptions = function(opts, selected) {
  const list = opts.slice();
  if (!list.some(o => String(typeof o === 'object' ? o.value : o) === String(selected ?? ''))) list.unshift(selected ?? '');
  return originalSelectOptions(list, selected ?? '');
};

renderPenyeliaWorkspace = function() {
  const w = ws('penyeliaPelajar'), groups = state.records.penyeliaPelajar || [];
  const showAll = ui.filter.penyeliaPelajar === 'semua';
  const q = (ui.query.penyeliaPelajar || '').toLowerCase().trim();
  let html = '<div class="page-head"><div><h1 class="page-head__title">' + esc(w.name) + '</h1><p class="page-head__desc">' + esc(w.desc) + '</p></div><button class="btn btn--primary" data-add-penyelia>Tambah Penyelia</button></div>';
  html += '<div class="ws-toolbar"><div class="ws-search"><input type="search" data-ws-search="penyeliaPelajar" aria-label="Cari pelajar atau penyelia" value="' + esc(q) + '" placeholder="Cari pelajar (nama / no. pelajar)…"></div>';
  html += '<div class="seg-toggle"><button class="seg' + (!showAll?' is-active':'') + '" data-seg="aktif">Aktif</button><button class="seg' + (showAll?' is-active':'') + '" data-seg="semua">Semua</button></div>';
  html += '<label>Papar sesi <select id="viewSession">' + selectOptions(allSessions(), selectedSession()) + '</select></label><button class="btn btn--ghost" data-export="penyeliaPelajar">Eksport</button><button class="btn btn--ghost" id="previewImport">Semak Import CSV</button></div>';
  if (window.PENYELIA_SEED && state.penyeliaImportVersion !== window.PENYELIA_SEED.version) html += '<p class="toast-inline">Dataset CSV tersedia. Semak Import CSV untuk gabung dengan rekod sedia ada.</p>';
  groups.forEach(g => {
    const list = (g.pelajar || []).filter(p => !q || (g.namaPenyelia + ' ' + p.nama + ' ' + studentIdentity(p).id).toLowerCase().includes(q));
    if (q && !list.length && !g.namaPenyelia.toLowerCase().includes(q)) return;
    const active = list.filter(p => !p.statusKhas), past = list.filter(p => p.statusKhas);
    const open = ui.penyeliaOpen[g.id] ?? active.length > 0;
    html += '<section class="penyelia-panel' + (open?' is-open':'') + '" data-penyelia="' + esc(g.id) + '"><button class="penyelia-panel__head" data-toggle-penyelia="' + esc(g.id) + '" aria-expanded="' + open + '"><span class="penyelia-panel__chev">' + icon('chevronRight',16) + '</span><span class="penyelia-panel__title">' + esc(g.namaPenyelia) + '</span><span class="badge badge--info">' + active.length + ' aktif</span><span class="badge badge--warn">' + past.length + ' lepas</span></button>';
    if (open) {
      html += '<div class="penyelia-panel__body"><div class="nested-actions"><button class="btn btn--ghost" data-supervisor-edit="' + esc(g.id) + '">Edit Penyelia</button><button class="btn btn--ghost" data-student-add="' + esc(g.id) + '">Tambah Pelajar</button></div>';
      html += renderPenyeliaTable(active, selectedSession(), g.id);
      if (showAll && past.length) html += renderLepasPanel(g.id, past, selectedSession(), true);
      html += '</div>';
    }
    html += '</section>';
  });
  return html;
};
renderPenyeliaTable = function(students, sesi, parentId) {
  const sessions = recentSessions();
  let html = '<div class="table-scroll"><table class="grid penyelia-table"><thead><tr><th class="col-bil">BIL</th><th>PELAJAR</th><th>PROGRAM</th><th>PENYELIA BERSAMA</th>';
  sessions.forEach(s => { html += '<th>SEM. PENGAJIAN<br>' + esc(s) + '</th>'; });
  html += '<th>STATUS TERKINI</th><th>STATUS</th><th>TINDAKAN</th></tr></thead><tbody>';
  students.forEach(p => {
    html += '<tr data-pelajar="' + esc(p.id) + '" data-record="' + esc(p.id) + '"><td>' + esc(p.bil) + '</td><td class="student-identity">' + studentIdentityHTML(p) + '</td><td>' + esc(p.program) + '</td><td>' + esc(formatPenyeliaBersama(extractPenyeliaNames(p.penyeliaBersama))) + '</td>';
    sessions.forEach(s => { const snap = snapshot(p,s); html += '<td>' + (snap ? esc(snap.semesterPengajian) + (snap.catatan ? '<div class="student-id">' + esc(snap.catatan) + '</div>' : '') : '—') + '</td>'; });
    html += '<td>' + esc(p.statusTerkini || '') + '</td><td>' + esc(statusPelajarLabel(p.statusKhas)) + '<div class="student-id">' + esc(fmtDate(p.tarikhStatus)) + (p.needsDateReview ? 'Tahun ' + esc(p.tahunGraduasi || 'tidak diketahui') + ' — tarikh perlu semakan' : '') + '</div></td>';
    html += '<td class="col-actions"><div class="cell-actions">' +
      '<button class="row-btn" data-edit-pelajar="' + esc(parentId + '|' + p.id) + '" title="Edit pelajar">' + icon('edit', 14) + ' Edit</button>' +
      '<button class="row-btn" data-status-pelajar="' + esc(parentId + '|' + p.id) + '" title="Ubah status">' + icon('checkCircle', 14) + ' Status</button>' +
      '</div></td></tr>';
  });
  if (!students.length) html += '<tr><td colspan="' + (sessions.length + 7) + '" class="empty-cell">Tiada pelajar aktif.</td></tr>';
  return html + '</tbody></table></div>';
};

const originalBindContent = bindContent;
bindContent = function() {
  originalBindContent();
  const root = document.getElementById('content');
  root.querySelector('#previewImport')?.addEventListener('click', previewPenyeliaImport);
  root.querySelector('#viewSession')?.addEventListener('change', e => { ui.penyeliaSession = e.target.value; render(); });
  root.querySelectorAll('[data-supervisor-edit]').forEach(b => b.onclick = () => openSupervisorDrawer(b.dataset.supervisorEdit));
  root.querySelectorAll('[data-student-add]').forEach(b => b.onclick = () => openPelajarDrawer(b.dataset.studentAdd));
  // Replace the old default-open toggle listener with a state-aware button.
  root.querySelectorAll('[data-toggle-penyelia]').forEach(old => {
    const b = old.cloneNode(true); old.replaceWith(b);
    b.onclick = () => { ui.penyeliaOpen[b.dataset.togglePenyelia] = b.getAttribute('aria-expanded') !== 'true'; render(); };
  });
};
addPenyelia = function() { openSupervisorDrawer(); };
function openSupervisorDrawer(id) {
  const existing = (state.records.penyeliaPelajar || []).find(g => g.id === id);
  openDrawer({title: existing ? 'Edit Penyelia' : 'Tambah Penyelia',
    body:'<div class="field"><label for="supervisorName">NAMA PENYELIA</label><input id="supervisorName" value="' + esc(existing?.namaPenyelia || '') + '"></div>',
    foot:(existing ? '<button class="btn btn--danger" id="deleteSupervisor">Padamkan</button>':'') + '<button class="btn btn--ghost" data-close-drawer>Batal</button><button class="btn btn--primary" id="saveSupervisor">Simpan</button>',
    onMount(root) {
      root.querySelector('#saveSupervisor').onclick = () => {
        const name = root.querySelector('#supervisorName').value.trim();
        if (!name) { toast('Nama penyelia wajib diisi','error'); return; }
        if (existing) existing.namaPenyelia = name;
        else state.records.penyeliaPelajar.push({id:uid('supervisor'),bil:nextId('penyeliaPelajar'),namaPenyelia:name,pelajar:[]});
        save(); closeDrawer(); render();
      };
      root.querySelector('#deleteSupervisor')?.addEventListener('click', () => confirmNestedDelete(existing));
    }});
}
function confirmNestedDelete(parent, student) {
  openModal({title:'Padamkan Rekod',body:'<p>' + esc(student?.nama || parent.namaPenyelia) + '</p><p>Rekod disimpan dalam Arkib Padaman dan boleh dipulihkan.</p>',
    foot:'<button class="btn btn--ghost" data-close-modal>Batal</button><button class="btn btn--danger" id="confirmNestedDelete">Padamkan</button>',
    onMount(root) { root.querySelector('#confirmNestedDelete').onclick = () => {
      state.trash.unshift({id:uid('t'),wsId:'penyeliaPelajar',kind:student?'student':'supervisor',parentId:parent.id,record:cloneData(student || parent),at:new Date().toISOString()});
      if (student) parent.pelajar = parent.pelajar.filter(p => p.id !== student.id);
      else state.records.penyeliaPelajar = state.records.penyeliaPelajar.filter(g => g.id !== parent.id);
      save(); closeModal(); closeDrawer(); render();
    }; }});
}
const originalRestoreTrash = restoreTrash;
restoreTrash = function(id) {
  const item = state.trash.find(t => t.id === id);
  if (!item || item.wsId !== 'penyeliaPelajar') return originalRestoreTrash(id);
  let target = state.records.penyeliaPelajar;
  if (item.kind === 'student') {
    const parent = target.find(g => g.id === item.parentId);
    if (!parent) { toast('Pulihkan penyelia dahulu sebelum pelajar.', 'warn'); return; }
    target = parent.pelajar;
  } else if (!Array.isArray(item.record.pelajar)) { toast('Jenis rekod arkib tidak dikenal pasti.', 'error'); return; }
  if (target.some(r => r.id === item.record.id)) { toast('Rekod sudah wujud.', 'warn'); return; }
  target.push(item.record);
  state.trash = state.trash.filter(t => t.id !== id);
  save(); render();
};

const originalOpenPelajarDrawer = openPelajarDrawer;
openPelajarDrawer = function(parentId, studentId) {
  originalOpenPelajarDrawer(parentId, studentId);
  const root = document.getElementById('drawerOverlay');
  if (!root) return;
  const parent = state.records.penyeliaPelajar.find(g => g.id === parentId);
  const student = parent.pelajar.find(p => p.id === studentId);
  const body = root.querySelector('.drawer__body');
  const note = document.createElement('div'); note.className = 'field';
  note.innerHTML = '<label for="dStatusTerkini">STATUS TERKINI</label><textarea id="dStatusTerkini">' + esc(student?.statusTerkini || '') + '</textarea><label for="dSemesterSesi">Sesi untuk Semester Pengajian (semasa)</label><select id="dSemesterSesi">' + selectOptions(allSessions(), student?.sesiSemesterPengajian || '') + '</select>';
  body.appendChild(note);
  if (student?.csvSources || student?.sourceVersions) {
    const source = document.createElement('details');
    source.innerHTML = '<summary>Sumber CSV (semakan konflik)</summary><pre>' + esc(JSON.stringify({versions:student.sourceVersions,conflicts:student.conflicts,statusEvidence:student.statusEvidence,csvSources:student.csvSources},null,2)) + '</pre>';
    body.appendChild(source);
  }
  if (student) {
    const button = document.createElement('button'); button.className = 'btn btn--danger'; button.textContent = 'Padamkan';
    root.querySelector('.drawer__foot').prepend(button);
    button.onclick = () => confirmNestedDelete(parent,student);
  }
  // Delegation also handles semester rows added after opening the drawer.
  root.addEventListener('click', e => { const b = e.target.closest('[data-remove-sejarah]'); if (b) b.closest('.sejarah-row')?.remove(); });
  root.querySelector('#savePelajar').replaceWith(root.querySelector('#savePelajar').cloneNode(true));
  root.querySelector('#savePelajar').onclick = () => saveNestedStudent(root,parent,student);
};
function saveNestedStudent(root,parent,student) {
  const value = id => root.querySelector('#' + id).value;
  const name = value('dNama').trim(), status = value('dStatus'), date = value('dTarikhStatus');
  if (!name) return toast('Nama pelajar wajib diisi','error');
  if (status === 'graduasi' && !date && !(student?.statusKhas === 'graduasi' && student?.needsDateReview)) return toast('Tarikh wajib untuk status graduasi','error');
  const history = [...root.querySelectorAll('.sejarah-row')].map(row => {
    const sesi = row.querySelector('.sejarah-sesi').value;
    return {...(student?.sejarahSemester || []).find(s => s.sesi === sesi), sesi,
      semesterPengajian:row.querySelector('.sejarah-sp').value, catatan:row.querySelector('.sejarah-catatan').value};
  });
  const nonempty = history.filter(s => s.sesi || s.semesterPengajian || s.catatan);
  if (nonempty.some(s => !s.sesi) || new Set(nonempty.map(s => s.sesi)).size !== nonempty.length) return toast('Pilih sesi unik bagi setiap sejarah semester.','error');
  const session = value('dSemesterSesi'), count = value('dSemesterPengajian');
  if (session) {
    const existing = nonempty.find(s => s.sesi === session);
    // Only explicit changes to the current count replace a snapshot.
    if (!student || count !== String(student.semesterPengajian || '') || session !== student.sesiSemesterPengajian) {
      if (existing) existing.semesterPengajian = count;
      else nonempty.push({sesi:session,semesterPengajian:count,catatan:''});
    }
  }
  const record = student || {id:uid('student'),bil:Math.max(0,...parent.pelajar.map(p => Number(p.bil)||0))+1};
  const pbNames = [...root.querySelectorAll('#dPenyeliaBersamaList .penyelia-input')].map(i => i.value.trim()).filter(Boolean);
  Object.assign(record,{nama:name,noPelajar:value('dNoPelajar').trim(),program:value('dProgram'),
    penyeliaBersama:formatPenyeliaBersama(pbNames),
    semesterPengajian:count,sesiSemesterPengajian:session,sejarahSemester:nonempty,
    statusTerkini:value('dStatusTerkini'),statusKhas:status,tarikhStatus:date});
  record.needsDateReview = status === 'graduasi' && !date;
  if (!student) parent.pelajar.push(record);
  syncGraduan(record);
  /* Pancarkan identiti kanonikal ke semua worksheet (elak typo/percanggahan). */
  const propagat = propagateCanonical(studentIdentity(record).id, { nama: studentIdentity(record).nama, program: record.program });
  save(); closeDrawer(); render();
  if (propagat) toast('Identiti diselaraskan ke ' + propagat + ' medan di worksheet lain', 'success');
}

// Ownership is explicit. Legacy _syncPelajarId alone is not proof of creation.
// Matching manual/legacy records are left untouched and never adopted or deleted.
syncGraduan = function(p) {
  const groups = state.records.penyeliaPelajar || [];
  const parent = groups.find(g => (g.pelajar || []).some(s => s.id === p.id));
  const list = state.records.graduan = state.records.graduan || [];
  const owned = list.find(g => g._syncOrigin === 'penyeliaPelajar:v2' && g._syncPelajarId === p.id);
  if (p.statusKhas !== 'graduasi') {
    state.records.graduan = list.filter(g => !(g._syncOrigin === 'penyeliaPelajar:v2' && g._syncPelajarId === p.id));
    return;
  }
  if (!parent || !p.tarikhStatus) return; // Imported historic graduates stay in past; no invented date or sync.
  const matric = normalizedMatric(p);
  if (!owned && list.some(g => g._syncPelajarId === p.id || (matric && normalizedMatric(g) === matric))) return;
  const entry = owned || {id:'graduan-' + p.id,bil:nextId('graduan'),_syncOrigin:'penyeliaPelajar:v2',_syncPelajarId:p.id,_syncParentId:parent.id};
  Object.assign(entry,{nama:studentIdentity(p).nama,noPelajar:studentIdentity(p).id,program:p.program,
    penyelia:formatPenyelia([parent.namaPenyelia,...extractPenyeliaNames(p.penyeliaBersama).filter(n => n.toUpperCase() !== 'TIADA')]),
    semester:p.semesterPengajian || '',catatan:'Graduasi ' + fmtDate(p.tarikhStatus),tarikhGraduasi:p.tarikhStatus});
  entry._syncSourceIds = (p.sourceVersions || []).map(v => v.source + ':' + v.row);
  if (!owned) list.push(entry);
};
const originalFocusRecord = focusRecord;
focusRecord = function(id) {
  if (ui.view === 'penyeliaPelajar') {
    const parent = state.records.penyeliaPelajar.find(g => g.id === id || (g.pelajar || []).some(p => p.id === id));
    if (parent) {
      ui.query.penyeliaPelajar = ''; ui.filter.penyeliaPelajar = 'semua';
      ui.penyeliaOpen[parent.id] = true; ui.lepasOpen[parent.id] = true; render();
      const node = [...document.querySelectorAll('[data-record],[data-penyelia]')].find(el => el.dataset.record === id || el.dataset.penyelia === id);
      node?.scrollIntoView({block:'center'}); node?.classList.add('is-target'); return;
    }
  } else {
    const rows = state.records[ui.view] || [], index = rows.filter(r => !r.locked).findIndex(r => r.id === id);
    ui.query[ui.view] = ''; ui.filter[ui.view] = 'all'; delete ui.sort[ui.view];
    if (index >= 0) ui.page[ui.view] = Math.floor(index/ui.pageSize)+1;
    if (rows.find(r => r.id === id)?.locked) ui.lockedOpen[ui.view] = true;
    render();
  }
  originalFocusRecord(id);
};
const originalExportWorkspace = exportWorkspace;
exportWorkspace = function(id) {
  if (id !== 'penyeliaPelajar') return originalExportWorkspace(id);
  const sessions = allSessions();
  const rows = [['BIL','NAMA PENYELIA','PELAJAR','PROGRAM','NO. PELAJAR','PENYELIA BERSAMA',...sessions,'STATUS TERKINI','STATUS','TARIKH STATUS','SEJARAH SEMESTER']];
  state.records.penyeliaPelajar.forEach(g => {
    if (!g.pelajar.length) rows.push([g.bil,g.namaPenyelia]);
    g.pelajar.forEach(p => rows.push([g.bil,g.namaPenyelia,p.nama,p.program,p.noPelajar,p.penyeliaBersama,
      ...sessions.map(s => { const snap = snapshot(p,s); return snap ? [snap.semesterPengajian,snap.catatan].filter(Boolean).join('\n') : ''; }),
      p.statusTerkini,statusPelajarLabel(p.statusKhas),p.tarikhStatus,JSON.stringify(p.sejarahSemester || [])]));
  });
  const csv = '\uFEFF' + rows.map(row => row.map(v => '"' + String(v ?? '').replace(/"/g,'""') + '"').join(',')).join('\r\n');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'})); a.download = 'penyeliaPelajar.csv'; a.click(); URL.revokeObjectURL(a.href);
};
