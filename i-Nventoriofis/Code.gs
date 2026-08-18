// ============================================================
//  i-NVENTORI OFIS — Alat Tulis & Aset Alih, IMEN
//
//  Two categories, and only two:
//    Alat Tulis (consumable) — kertas, pen, dakwat. Tracked by quantity.
//    Aset Alih  (fixed_asset) — laptop, projektor, mikrofon. Tracked one by
//                               one, borrowed and returned.
//  The enum keys stay `consumable` / `fixed_asset` because they are internal
//  and live rows already carry them; only the labels are Malay.
//  Backend: Google Apps Script + Google Sheets
//
//  Paste this ENTIRE file into the bound Apps Script project's Code.gs,
//  then Deploy -> Manage deployments -> Edit -> New version -> Deploy.
//  NEVER click "New deployment" — it mints a new /exec URL and breaks
//  the google.script.run shim in index.html.
//
//  Architecture note: this file is the Apps Script equivalent of a
//  Laravel app. The mapping is deliberate and documented in CLAUDE.md:
//    migrations       -> ensureSheets_()      (idempotent, header auto-heal)
//    Eloquent + casts -> _readTable_()
//    query scopes     -> SCOPES
//    DB::transaction  -> _txn_()              (LockService)
//    FormRequest      -> _validate_()
//    routes + auth    -> doPost switch + ADMIN_ONLY_ACTIONS
//    Scheduler        -> checkDates() + installTriggers()
// ============================================================

// ── Deployment constants ────────────────────────────────────
// SPREADSHEET_ID / FOLDER_ID are not secrets (access is controlled by
// Drive sharing, not obscurity) so they may live in this public file.
// The admin password may NOT — see _adminPassword_().
const SPREADSHEET_ID = '1Xk1aKMmWR3AFTvWlMJDKUmv-5Ik_GSvyZeqXX30j_6E';
const FOLDER_ID      = '1NQUV2EHXpQhlWzcCm2aaIEOUvZ4CGOE7';
const ADMIN_EMAIL    = 'imenmakmal@gmail.com';
const APP_NAME       = 'i-Nventori Ofis';

// Where doGet fetches the live UI from (single source of truth: the repo).
const UI_RAW_URL   = 'https://raw.githubusercontent.com/imenhub-portal/imenhub-portal.github.io/main/i-Nventoriofis/index.html';
const PAGES_ORIGIN = 'https://imenhub-portal.github.io';

// Admin password. Script Properties ONLY — this file is committed to a
// PUBLIC repo, so a literal here would be world-readable. Fails CLOSED:
// if the property is missing, every gated action is denied.
function _adminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
}

// The Pentadbir Inventori: the person who approves requests and receives
// every notification. Held in the Config sheet so it can be changed from
// Tetapan without a redeploy. ADMIN_EMAIL above is the FALLBACK only — if
// Config is empty (a sheet predating this, or a fresh deployment) mail must
// still reach somebody rather than silently vanish.
const CONFIG_DEFAULTS = {
  admin_name:  'Pentadbir Inventori',
  admin_email: 'hazde@ukm.edu.my',
  admin_cc:    ''
};

function _config_() {
  try {
    const rows = _readTable_('Config');
    const out = {};
    Object.keys(CONFIG_DEFAULTS).forEach(function (k) { out[k] = CONFIG_DEFAULTS[k]; });
    rows.forEach(function (r) {
      if (r.key && String(r.value || '') !== '') out[String(r.key)] = String(r.value);
    });
    return out;
  } catch (e) {
    return CONFIG_DEFAULTS;   // sheet missing entirely — never throw here
  }
}

function _adminName_()  { return _config_().admin_name || 'Pentadbir Inventori'; }
function _adminEmail_() { return _config_().admin_email || ADMIN_EMAIL; }

// Extra addresses, comma or semicolon separated. Returns '' when unset so it
// can be handed straight to MailApp without a conditional.
function _adminCc_() {
  const raw = String(_config_().admin_cc || '').trim();
  if (!raw) return '';
  return raw.split(/[,;]+/).map(function (x) { return x.trim(); })
            .filter(Boolean).join(',');
}

// ============================================================
//  SCHEMA  ("migrations")
//
//  Column names are the specification's, verbatim, so the sheet stays
//  reviewable against the original data model. Order matters only for
//  first creation — every read/write goes through a header index, so
//  columns may be reordered in the sheet without breaking anything.
// ============================================================
const SCHEMA = {
  Categories: ['id', 'name', 'type', 'created_at', 'updated_at'],

  Locations: ['id', 'building', 'floor', 'room_number', 'created_at', 'updated_at'],

  // Key/value rather than one row of columns, so a future setting never
  // needs a migration — the same shape i-print's Config sheet uses.
  Config: ['key', 'value', 'updated_at'],

  Custodians: ['id', 'employee_id', 'name', 'email', 'department', 'created_at', 'updated_at'],

  Items: [
    'id', 'asset_tag', 'name', 'category_id', 'location_id', 'serial_number',
    'item_type', 'quantity_total', 'quantity_available', 'min_stock_alert',
    'status',
    'date_acquired', 'date_last_maintained',
    'date_next_inspection', 'date_last_audited', 'date_decommissioned',
    'custodian_id', 'photo_url', 'receipt_url', 'notes',
    // Return-proof state for the CURRENT loan only. All three are cleared at
    // check-in, which is what kills the token.
    'return_token', 'return_photo_url', 'return_claimed_at',
    'created_at', 'updated_at', 'deleted_at'
  ],

  // Requests raised by staff from the public page. Unlike Transactions
  // this IS mutable — a request moves through pending -> approved/rejected
  // — but the resulting handover still lands in the immutable ledger.
  // One row per requested LINE, with group_id shared across the lines of a
  // single submission. A header/detail pair of tabs would be more normalised,
  // but this keeps svcDecideRequest working per row, makes partial approval
  // natural, and needs no join.
  Requests: [
    'id', 'group_id', 'item_id', 'request_type', 'requester_name', 'requester_email',
    'quantity', 'purpose', 'needed_date',
    'status', 'admin_notes', 'custodian_id',
    'created_at', 'actioned_at'
  ],

  // IMMUTABLE LEDGER. Nothing in this file ever calls setValue/deleteRow on
  // this tab — only appendRow (enforced by the guard in _update_). A
  // check-in appends a NEW row; it does not edit the original check_out row.
  Transactions: [
    'id', 'item_id', 'user_id', 'custodian_id', 'action_type', 'quantity',
    'transaction_date', 'expected_return_date', 'actual_return_date',
    // Two named columns rather than one shared 'photo' field, so whose
    // evidence a row holds is never ambiguous. Only check_in rows use them.
    'photo_borrower', 'photo_admin',
    // Where a restock came from. Structured rather than buried in
    // reason_notes, so "how much paper did we get from Kedai Ali" is
    // answerable. Only stock_add rows carry it.
    'source',
    'reason_notes', 'created_at'
  ]
};

// Columns cast to Date on read. Everything date-shaped is named date_*
// except these, which are spelled out.
const DATE_COLS = {
  transaction_date: 1, expected_return_date: 1, actual_return_date: 1,
  needed_date: 1, actioned_at: 1, return_claimed_at: 1,
  created_at: 1, updated_at: 1, deleted_at: 1
};
const NUM_COLS = {
  id: 1, category_id: 1, location_id: 1, custodian_id: 1, item_id: 1,
  quantity_total: 1, quantity_available: 1, min_stock_alert: 1, quantity: 1
};

const ITEM_TYPES   = ['fixed_asset', 'consumable'];
const ITEM_STATUS  = ['available', 'assigned', 'maintenance', 'decommissioned', 'disposed'];
const ACTION_TYPES = ['stock_add', 'stock_remove', 'check_out', 'check_in', 'decommission', 'audit_adjust'];

// Removal reasons offered by the decommission drawer (spec 5.2).
const REMOVAL_REASONS = ['Rosak', 'E-waste', 'Hilang', 'Dijual', 'Dipindah'];

const REQUEST_STATUS = ['pending', 'approved', 'rejected', 'cancelled'];

// ============================================================
//  MIGRATIONS — ensureSheets_()
//
//  Idempotent: creates missing tabs, and AUTO-HEALS missing columns by
//  appending them to the header row rather than throwing. This is the
//  same technique i-nstrumen uses for its late-added `unit` column, and
//  it is what makes deploying a schema change safe: paste the new
//  Code.gs, and the next request repairs the sheet in place. Existing
//  rows simply read blank for the new column.
// ============================================================
function ensureSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(SCHEMA).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]);
      sh.setFrozenRows(1);
      return;
    }
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var header  = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    var missing = SCHEMA[name].filter(function (c) { return header.indexOf(c) === -1; });
    if (missing.length) {
      sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    }
  });
  return true;
}

// ============================================================
//  DATA ACCESS  ("Eloquent")
// ============================================================

function _sheet_(name) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  if (!sh) throw new Error('Sheet tidak dijumpai: ' + name + ' (jalankan ensureSheets_)');
  return sh;
}

function _headerIndex_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const header  = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const idx = {};
  header.forEach(function (h, i) { if (h) idx[h] = i; });
  return idx;
}

function _cast_(col, v) {
  if (v === '' || v === null || v === undefined) return null;
  if (col.indexOf('date_') === 0 || DATE_COLS[col]) {
    const d = (v instanceof Date) ? v : new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
  if (NUM_COLS[col]) {
    const n = Number(v);
    return isNaN(n) ? null : n;
  }
  return String(v);
}

// Reads a whole tab into plain objects, applying casts. Soft-deleted rows
// are excluded unless withTrashed is true (Laravel's default scope).
// `_row` is the physical sheet row, needed by _update_.
function _readTable_(name, withTrashed) {
  const sh = _sheet_(name);
  if (sh.getLastRow() < 2) return [];
  const idx  = _headerIndex_(sh);
  const cols = Object.keys(idx);
  // Blank-row detection keys off `id` where there is one. Config is
  // key/value and has no id column, so fall back to the first column —
  // otherwise every row in such a tab is silently treated as blank and
  // discarded, which is exactly what happened when Config was added.
  const presence = (idx.id !== undefined) ? idx.id : idx[cols[0]];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const out  = [];
  for (var r = 0; r < vals.length; r++) {
    var row = vals[r];
    if (String(row[presence] === undefined ? '' : row[presence]).trim() === '') continue; // blank row
    var o = { _row: r + 2 };
    for (var c = 0; c < cols.length; c++) o[cols[c]] = _cast_(cols[c], row[idx[cols[c]]]);
    if (!withTrashed && o.deleted_at) continue;
    out.push(o);
  }
  return out;
}

// Next surrogate id. Only ever called inside _txn_ — the lock is what
// makes read-max-then-write-max+1 safe against concurrent callers.
function _nextId_(name) {
  const sh = _sheet_(name);
  if (sh.getLastRow() < 2) return 1;
  const idx  = _headerIndex_(sh);
  const vals = sh.getRange(2, idx.id + 1, sh.getLastRow() - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < vals.length; i++) {
    var n = Number(vals[i][0]);
    if (!isNaN(n) && n > max) max = n;
  }
  return max + 1;
}

// Appends one record, mapping object keys onto the LIVE header order
// (not SCHEMA order) so a reordered or auto-healed sheet still works.
function _insert_(name, obj) {
  const sh    = _sheet_(name);
  const idx   = _headerIndex_(sh);
  const width = sh.getLastColumn();
  const row   = [];
  for (var i = 0; i < width; i++) row.push('');
  Object.keys(obj).forEach(function (k) {
    if (idx[k] === undefined) return;
    row[idx[k]] = (obj[k] === null || obj[k] === undefined) ? '' : obj[k];
  });
  sh.appendRow(row);
  return obj;
}

// The ONLY path by which a ledger row ever leaves the sheet.
//
// The ledger is otherwise append-only and _update_ refuses the tab outright.
// This exists for one narrow case: an item registered by mistake — the same
// pen entered twice — whose ledger rows describe stock that never arrived.
// Keeping those rows does not preserve history, it fabricates it. Callers
// must establish that the item is a mis-entry before calling this.
//
// Rows are deleted bottom-up because deleting a row shifts every row below
// it up by one, which would corrupt the remaining indices.
function _purgeLedgerRows_(itemId) {
  const sh = _sheet_('Transactions');
  const rows = _readTable_('Transactions').filter(function (t) {
    return Number(t.item_id) === Number(itemId);
  });
  rows.sort(function (a, b) { return b._row - a._row; });
  rows.forEach(function (r) { sh.deleteRow(r._row); });
  return rows.length;
}

// Patches specific columns of one row, addressed by the `_row` that
// _readTable_ attached. Refuses to touch the ledger — that guard is the
// mechanical enforcement of "immutable" in the spec, not just a comment.
function _update_(name, rowNum, patch) {
  if (name === 'Transactions') {
    throw new Error('Ledger Transactions tidak boleh diubah (append-only).');
  }
  const sh  = _sheet_(name);
  const idx = _headerIndex_(sh);
  Object.keys(patch).forEach(function (k) {
    if (idx[k] === undefined) return;
    var v = patch[k];
    sh.getRange(rowNum, idx[k] + 1).setValue(v === null || v === undefined ? '' : v);
  });
  return true;
}

// ============================================================
//  TRANSACTIONS  ("DB::transaction")
//
//  Apps Script has no real DB transaction, so atomicity comes from a
//  script-wide lock: every read-modify-write-plus-ledger-append runs
//  alone. Without this, two concurrent check-outs can both read
//  quantity_available = 1 and both decrement it to 0, handing the same
//  unit to two people. The lock is always released in `finally`.
// ============================================================
function _txn_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, error: 'Pelayan sibuk. Sila cuba sebentar lagi.' };
  }
  try {
    var result = fn();
    _bustCache_();
    return { success: true, result: result };
  } catch (e) {
    return { success: false, error: (e && e.message) || String(e) };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  VALIDATION  ("FormRequest")
//
//  Declarative schema per action. Messages are Bahasa Malaysia because
//  they surface directly in the UI.
//    { field: { required, type, in, min, max, def, label } }
// ============================================================
function _validate_(schema, payload) {
  const data = payload || {};
  const out  = {};
  const errs = [];

  Object.keys(schema).forEach(function (field) {
    var rule  = schema[field];
    var label = rule.label || field;
    var v     = data[field];

    var empty = (v === undefined || v === null || String(v).trim() === '');
    if (empty) {
      if (rule.required) errs.push(label + ' wajib diisi.');
      out[field] = (rule.def === undefined ? null : rule.def);
      return;
    }

    if (rule.type === 'number') {
      var n = Number(v);
      if (isNaN(n)) { errs.push(label + ' mesti nombor.'); return; }
      if (rule.min !== undefined && n < rule.min) errs.push(label + ' mesti sekurang-kurangnya ' + rule.min + '.');
      if (rule.max !== undefined && n > rule.max) errs.push(label + ' tidak boleh melebihi ' + rule.max + '.');
      v = n;
    } else if (rule.type === 'date') {
      var d = (v instanceof Date) ? v : new Date(v);
      if (isNaN(d.getTime())) { errs.push(label + ' bukan tarikh yang sah.'); return; }
      v = d;
    } else if (rule.type === 'email') {
      v = String(v).trim();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) { errs.push(label + ' bukan emel yang sah.'); return; }
    } else {
      v = String(v).trim();
      if (rule.max !== undefined && v.length > rule.max) {
        errs.push(label + ' terlalu panjang (maks ' + rule.max + ' aksara).');
      }
    }

    if (rule.in && rule.in.indexOf(v) === -1) {
      errs.push(label + ' tidak sah. Pilihan: ' + rule.in.join(', ') + '.');
      return;
    }
    out[field] = v;
  });

  if (errs.length) throw new Error(errs.join(' '));
  return out;
}

// ============================================================
//  DATE HELPERS
//
//  All comparisons are day-granular. A warranty expiring "today" and one
//  expiring at 23:59 today must behave identically, so every date is
//  flattened to local midnight before any arithmetic.
// ============================================================
function _startOfDay_(d) {
  const x = (d instanceof Date) ? new Date(d.getTime()) : new Date(d);
  if (isNaN(x.getTime())) return null;
  x.setHours(0, 0, 0, 0);
  return x;
}

function _daysBetween_(from, to) {
  const a = _startOfDay_(from), b = _startOfDay_(to);
  if (!a || !b) return null;
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

function _today_() { return _startOfDay_(new Date()); }

// ============================================================
//  QUERY SCOPES  ("scopeExpiringWarranty", "scopeOverdue", ...)
//
//  Pure predicates, shared by the read API (dashboard badges) and by the
//  daily checkDates() job (alert emails). Sharing them is the point: a
//  second copy of this logic is how the bell and the email end up
//  disagreeing about what is overdue.
// ============================================================
const AUDIT_STALE_DAYS = 365;

const SCOPES = {
  // Active = not soft-deleted, not decommissioned/disposed.
  active: function (item) {
    return !item.deleted_at && item.status !== 'decommissioned' && item.status !== 'disposed';
  },

  // Inspection due today or earlier.
  inspectionDue: function (item, today) {
    if (!SCOPES.active(item) || !item.date_next_inspection) return false;
    var d = _daysBetween_(today || _today_(), item.date_next_inspection);
    return d !== null && d <= 0;
  },

  // Never audited, or last audited more than 365 days ago.
  auditDue: function (item, today) {
    if (!SCOPES.active(item)) return false;
    if (!item.date_last_audited) return true;
    var d = _daysBetween_(item.date_last_audited, today || _today_());
    return d !== null && d > AUDIT_STALE_DAYS;
  },

  // Consumable at or below its reorder threshold.
  lowStock: function (item) {
    if (!SCOPES.active(item) || item.item_type !== 'consumable') return false;
    var min = Number(item.min_stock_alert || 0);
    return Number(item.quantity_available || 0) <= min;
  },

  // Ledger scope: an open loan whose expected return date has passed.
  // Operates on a check_out row that has no matching later check_in.
  overdue: function (loan, today) {
    if (!loan || !loan.expected_return_date) return false;
    if (loan.actual_return_date) return false;
    var d = _daysBetween_(today || _today_(), loan.expected_return_date);
    return d !== null && d < 0;
  }
};

// Open loans = check_out rows with no subsequent check_in for the same item.
// Derived from the ledger rather than stored, so it cannot drift from it.
function _openLoans_(transactions) {
  const byItem = {};
  const sorted = (transactions || []).slice().sort(function (a, b) {
    return (a.id || 0) - (b.id || 0);
  });
  sorted.forEach(function (t) {
    if (t.action_type === 'check_out') byItem[t.item_id] = t;
    else if (t.action_type === 'check_in') delete byItem[t.item_id];
    else if (t.action_type === 'decommission') delete byItem[t.item_id];
  });
  return Object.keys(byItem).map(function (k) { return byItem[k]; });
}

// ============================================================
//  SERVER CACHE
//
//  getInitialData() re-reads five tabs on every call. The finished
//  payload is cached briefly and busted by every write (see _txn_), so a
//  user's own action always invalidates before their next load, and a
//  cold client is never staler than the cache TTL.
// ============================================================
const CACHE_KEY = 'inv_initial_v1';
const CACHE_TTL = 45; // seconds

function _bustCache_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (e) { /* best-effort */ }
}

// ============================================================
//  READ API
// ============================================================

// Full app payload. `admin` gates the personal-data tables: custodian
// names and emails are PDPA-relevant, so an anonymous caller gets the
// inventory but never a bulk directory dump. (The i-nstrumen
// GetAllIMenianUsers incident is exactly this mistake.)
function getInitialData(adminPass) {
  ensureSheets_();
  const isAdmin = _isAdmin_(adminPass);

  var payload = null;
  try {
    var hit = CacheService.getScriptCache().get(CACHE_KEY);
    if (hit) payload = JSON.parse(hit);
  } catch (e) { payload = null; }

  if (!payload) {
    payload = _buildPayload_();
    try {
      CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(payload), CACHE_TTL);
    } catch (e) { /* payload too large for cache — serve it live */ }
  }

  if (!isAdmin) {
    payload.custodians   = [];
    payload.transactions = [];
    payload.requests     = [];
    payload.config       = {};   // the officer's address is contact detail, not public
    payload.alerts       = { overdue: [], inspection: [], audit: [], low_stock: [] };
  }
  payload.is_admin  = isAdmin;
  payload.server_ts = new Date().toISOString();
  return payload;
}

function _buildPayload_() {
  const today  = _today_();
  const items  = _readTable_('Items');
  const cats   = _readTable_('Categories');
  const locs   = _readTable_('Locations');
  const custs  = _readTable_('Custodians');
  const txns   = _readTable_('Transactions');
  const reqs   = _readTable_('Requests');
  const loans  = _openLoans_(txns);

  // Attach the open loan to its item so the grid can show who holds what
  // without the client having to replay the ledger itself.
  const loanByItem = {};
  loans.forEach(function (l) { loanByItem[l.item_id] = l; });
  items.forEach(function (it) {
    var l = loanByItem[it.id];
    it.open_loan = l ? {
      transaction_id:       l.id,
      custodian_id:         l.custodian_id,
      expected_return_date: l.expected_return_date,
      transaction_date:     l.transaction_date,
      is_overdue:           SCOPES.overdue(l, today)
    } : null;
  });

  return {
    items:        items,
    categories:   cats,
    locations:    locs,
    custodians:   custs,
    transactions: txns,
    requests: reqs.slice().sort(function (a, b) { return (b.id || 0) - (a.id || 0); }),
    config: _config_(),
    alerts: {
      inspection: items.filter(function (i) { return SCOPES.inspectionDue(i, today); }).map(_alertRef_),
      audit:      items.filter(function (i) { return SCOPES.auditDue(i, today); }).map(_alertRef_),
      low_stock:  items.filter(function (i) { return SCOPES.lowStock(i); }).map(_alertRef_),
      overdue:    loans.filter(function (l) { return SCOPES.overdue(l, today); }).map(function (l) {
        var it = _findById_(items, l.item_id);
        return {
          id:        l.item_id,
          asset_tag: it ? it.asset_tag : '?',
          name:      it ? it.name : '(item dipadam)',
          custodian_id: l.custodian_id,
          expected_return_date: l.expected_return_date,
          days_overdue: Math.abs(_daysBetween_(today, l.expected_return_date) || 0)
        };
      })
    },
    meta: {
      item_types:      ITEM_TYPES,
      item_status:     ITEM_STATUS,
      action_types:    ACTION_TYPES,
      removal_reasons: REMOVAL_REASONS
    }
  };
}

function _alertRef_(i) {
  return {
    id: i.id, asset_tag: i.asset_tag, name: i.name, status: i.status,
    date_next_inspection: i.date_next_inspection,
    date_last_audited:    i.date_last_audited,
    quantity_available:   i.quantity_available,
    min_stock_alert:      i.min_stock_alert
  };
}

function _findById_(rows, id) {
  for (var i = 0; i < rows.length; i++) if (Number(rows[i].id) === Number(id)) return rows[i];
  return null;
}

// Full history for one item — the audit-trail drawer.
function getItemHistory(itemId, adminPass) {
  _requireAdmin_(adminPass);
  const txns = _readTable_('Transactions').filter(function (t) {
    return Number(t.item_id) === Number(itemId);
  });
  txns.sort(function (a, b) { return (b.id || 0) - (a.id || 0); });
  return txns;
}

// ============================================================
//  PUBLIC CATALOG  (no password — this is what staff browse)
//
//  Deliberately a separate, narrower payload than getInitialData: it
//  carries only what a requester needs to decide whether to ask for
//  something. No ledger, no custodian directory, and critically NO name
//  of whoever currently holds an asset — an unavailable item reports its
//  expected return date and nothing more. Staff names are personal data
//  and this endpoint is reachable by anyone with the URL.
// ============================================================
function getPublicCatalog() {
  ensureSheets_();
  const today = _today_();
  const items = _readTable_('Items');
  const locs  = _readTable_('Locations');
  const loans = _openLoans_(_readTable_('Transactions'));

  const loanByItem = {};
  loans.forEach(function (l) { loanByItem[l.item_id] = l; });

  const locName = {};
  locs.forEach(function (l) {
    locName[l.id] = [l.building, l.floor, l.room_number].filter(String).join(' · ');
  });

  const out = [];
  items.forEach(function (i) {
    if (i.status === 'decommissioned' || i.status === 'disposed') return;
    var loan = loanByItem[i.id];
    var isAsset = i.item_type === 'fixed_asset';

    out.push({
      id: i.id,
      asset_tag: i.asset_tag || '',
      name: i.name,
      item_type: i.item_type,
      category: isAsset ? 'Aset Alih' : 'Alat Tulis',
      location: locName[i.location_id] || '',
      quantity_available: Number(i.quantity_available || 0),
      available: isAsset ? (!loan && Number(i.quantity_available || 0) > 0)
                         : Number(i.quantity_available || 0) > 0,
      // Date only — never the holder's name.
      expected_return_date: loan ? loan.expected_return_date : null,
      is_overdue: loan ? SCOPES.overdue(loan, today) : false
    });
  });

  out.sort(function (a, b) {
    if (a.item_type !== b.item_type) return a.item_type === 'fixed_asset' ? -1 : 1;
    return String(a.name).localeCompare(String(b.name));
  });
  return { items: out, server_ts: new Date().toISOString() };
}

// Public. Creates nothing but pending rows, so the worst case is an admin
// rejecting some noise.
//
// Every line is validated BEFORE anything is written: a request for "2 rim
// kertas and 4 pen" where the pens are out of stock is rejected whole rather
// than half-recorded, which would leave the requester thinking they asked for
// both and the admin seeing only one.
function svcSubmitRequest(p) {
  const v = _validate_({
    requester_name:  { required: true, max: 120, label: 'Nama' },
    requester_email: { required: true, type: 'email', max: 120, label: 'Emel' },
    purpose:         { required: true, max: 500, label: 'Tujuan' },
    needed_date:     { type: 'date', label: 'Tarikh diperlukan' }
  }, p);

  const rawLines = (p && p.lines) || [];
  if (!rawLines.length) throw new Error('Sila pilih sekurang-kurangnya satu barang.');
  if (rawLines.length > 20) throw new Error('Terlalu banyak barang dalam satu permohonan (maks 20).');

  const items   = _readTable_('Items');
  const loans   = _openLoans_(_readTable_('Transactions'));
  const pending = _readTable_('Requests').filter(function (r) { return r.status === 'pending'; });
  const email   = String(v.requester_email).toLowerCase();

  const seen  = {};
  const lines = [];

  rawLines.forEach(function (raw) {
    var itemId = Number(raw && raw.item_id);
    var qty    = Math.max(1, Number((raw && raw.quantity) || 1));

    var item = _findById_(items, itemId);
    if (!item) throw new Error('Barang tidak dijumpai (#' + itemId + ').');
    if (item.status === 'decommissioned' || item.status === 'disposed') {
      throw new Error(item.name + ' telah dilupuskan.');
    }
    if (seen[itemId]) throw new Error(item.name + ' dimasukkan lebih daripada sekali.');
    seen[itemId] = true;

    var isAsset = item.item_type === 'fixed_asset';

    // Availability is re-checked here, never trusted from the browser — the
    // catalog it chose from can be minutes stale.
    if (isAsset) {
      var open = loans.filter(function (l) { return Number(l.item_id) === Number(item.id); });
      if (open.length) {
        throw new Error(item.name + ' sedang dipinjam. Sila cuba semula selepas ia dipulangkan.');
      }
      qty = 1;   // one physical thing
    } else if (Number(item.quantity_available || 0) < qty) {
      throw new Error(item.name + ': baki tidak mencukupi (baki ' + item.quantity_available + ').');
    }

    // One pending line per person per item — stops a double-tapped button
    // filling the queue with duplicates.
    var dup = pending.filter(function (r) {
      return Number(r.item_id) === Number(item.id) &&
             String(r.requester_email).toLowerCase() === email;
    });
    if (dup.length) {
      throw new Error('Anda sudah memohon ' + item.name + ' dan ia masih menunggu kelulusan.');
    }

    lines.push({ item: item, quantity: qty, isAsset: isAsset });
  });

  // Everything validated — now write.
  const now     = new Date();
  const groupId = 'REQ-' + now.getTime() + '-' + Math.floor(Math.random() * 1000);
  const ids     = [];
  var nextId    = _nextId_('Requests');

  lines.forEach(function (ln) {
    _insert_('Requests', {
      id: nextId,
      group_id: groupId,
      item_id: ln.item.id,
      request_type: ln.isAsset ? 'asset' : 'inventory',
      requester_name: v.requester_name,
      requester_email: email,
      quantity: ln.quantity,
      purpose: v.purpose,
      needed_date: v.needed_date,
      status: 'pending',
      admin_notes: '',
      custodian_id: null,
      created_at: now,
      actioned_at: null
    });
    ids.push(nextId);
    nextId++;
  });

  // ONE acknowledgement listing every line — four pens must not mean four
  // emails.
  var mailed = true;
  try { _mailRequestReceived_(lines, v, groupId); } catch (e) { mailed = false; }
  return { group_id: groupId, ids: ids, count: ids.length, status: 'pending', mailed: mailed };
}

// Decides every pending line of one submission in a single transaction.
// Delegates each line to svcDecideRequest so the ledger row, recipient
// resolution and handover emails are identical to an admin-initiated
// handover — there is no second implementation to drift.
function svcDecideGroup(p, decision) {
  const groupId = String((p && p.group_id) || '').trim();
  if (!groupId) throw new Error('Kumpulan permohonan tidak dinyatakan.');

  const rows = _readTable_('Requests').filter(function (r) {
    return String(r.group_id) === groupId && r.status === 'pending';
  });
  if (!rows.length) throw new Error('Tiada permohonan menunggu dalam kumpulan ini.');

  const out = [];
  rows.forEach(function (r) {
    out.push(svcDecideRequest({
      id: r.id,
      admin_notes: p.admin_notes,
      expected_return_date: p.expected_return_date
    }, decision));
  });
  return { group_id: groupId, decided: out.length, status: decision, lines: out };
}

// Admin decision. Approving an ASSET hands it over for real: the same
// svcCheckOut path runs, so the ledger and the notification emails are
// identical to a handover the admin started themselves. Approving an
// INVENTORY request issues the stock the same way Keluar Stok does.
function svcDecideRequest(p, decision) {
  const v = _validate_({
    id:                   { required: true, type: 'number', label: 'Permohonan' },
    admin_notes:          { max: 500, label: 'Catatan' },
    expected_return_date: { type: 'date', label: 'Tarikh jangka pulang' }
  }, p);

  const rows = _readTable_('Requests');
  const req  = _findById_(rows, v.id);
  if (!req) throw new Error('Permohonan tidak dijumpai.');
  if (req.status !== 'pending') {
    throw new Error('Permohonan ini telah pun diuruskan (' + req.status + ').');
  }

  const item = _findById_(_readTable_('Items'), req.item_id);
  if (!item) throw new Error('Item tidak lagi wujud.');

  const now = new Date();

  if (decision === 'rejected') {
    _update_('Requests', req._row, {
      status: 'rejected', admin_notes: v.admin_notes || '', actioned_at: now
    });
    var rMailed = true;
    try { _mailRequestDecided_(item, req, 'rejected', v.admin_notes); } catch (e) { rMailed = false; }
    return { id: req.id, status: 'rejected', mailed: rMailed };
  }

  // ── Approved ──
  if (req.request_type === 'asset') {
    if (!v.expected_return_date) throw new Error('Sila tetapkan tarikh jangka pulang.');
    // Reuses the real handover path — same validation, same ledger row,
    // same emails. There is no second implementation to drift.
    var out = svcCheckOut({
      id: item.id,
      recipient_name: req.requester_name,
      recipient_email: req.requester_email,
      expected_return_date: v.expected_return_date,
      reason_notes: 'Permohonan #' + req.id + ': ' + (req.purpose || '')
    });
    _update_('Requests', req._row, {
      status: 'approved', admin_notes: v.admin_notes || '',
      custodian_id: out.custodian_id, actioned_at: now
    });
    return { id: req.id, status: 'approved', mailed: out.mailed, custodian_id: out.custodian_id };
  }

  // Inventory: issue the stock to the requester, creating them if needed.
  var recipient = _resolveRecipient_({
    recipient_name: req.requester_name,
    recipient_email: req.requester_email,
  });
  svcStockChange({
    id: item.id, quantity: req.quantity, custodian_id: recipient.id,
    reason_notes: 'Permohonan #' + req.id + ': ' + (req.purpose || '')
  }, 'stock_remove');

  _update_('Requests', req._row, {
    status: 'approved', admin_notes: v.admin_notes || '',
    custodian_id: recipient.id, actioned_at: now
  });
  var iMailed = true;
  try { _mailRequestDecided_(item, req, 'approved', v.admin_notes); } catch (e) { iMailed = false; }
  return { id: req.id, status: 'approved', mailed: iMailed };
}

// ============================================================
//  AUTH  ("auth middleware")
//
//  One shared admin password, held in Script Properties. The check is
//  SERVER-side on every gated action — a client-side boolean is never
//  trusted, because doPost is callable anonymously from anywhere.
//  Fails closed: no property set means nobody is admin.
// ============================================================
function _isAdmin_(pass) {
  const real = _adminPassword_();
  if (!real) return false;
  return String(pass || '') === real;
}

function _requireAdmin_(pass) {
  if (!_isAdmin_(pass)) throw new Error('Akses ditolak. Sila log masuk semula.');
  return true;
}

function adminLogin(pass) {
  // Deliberately returns only a boolean — no token, no echo of the
  // password. The client re-sends the password on each gated call.
  //
  // `unconfigured` distinguishes "no ADMIN_PASSWORD property set yet" from
  // "wrong password". Without it, a fresh deployment looks like a typo and
  // sends you hunting for the wrong problem.
  if (!_adminPassword_()) return { ok: false, unconfigured: true };
  return { ok: _isAdmin_(pass) };
}

// Every action name the frontend may invoke, and whether it needs admin.
// Anything not listed here is rejected by handleAction().
const ACTIONS_PUBLIC = ['Ping', 'SubmitRequest', 'SubmitReturnProof'];
const ACTIONS_ADMIN  = [
  'AddItem', 'UpdateItem', 'DeleteItem',
  'StockAdd', 'StockRemove',
  'CheckOut', 'CheckIn',
  'Decommission', 'AuditAdjust',
  'SaveCategory', 'DeleteCategory',
  'SaveLocation', 'DeleteLocation',
  'SaveCustodian', 'DeleteCustodian',
  'RunDateCheck', 'SeedReference', 'SaveConfig',
  'ApproveRequest', 'RejectRequest',
  'ApproveGroup', 'RejectGroup'
];

// ============================================================
//  ACTION DISPATCH  ("controllers")
//
//  One entry point, one explicit switch. Auth is enforced here — before
//  the switch — so no individual handler can forget it.
// ============================================================
function handleAction(actionType, payload) {
  const p = payload || {};
  if (ACTIONS_ADMIN.indexOf(actionType) !== -1) {
    if (!_isAdmin_(p.__pass)) {
      return { success: false, error: 'Akses ditolak. Sila log masuk semula.' };
    }
  } else if (ACTIONS_PUBLIC.indexOf(actionType) === -1) {
    return { success: false, error: 'Tindakan tidak dikenali: ' + actionType };
  }

  switch (actionType) {
    case 'Ping':            return { success: true, result: 'pong' };
    case 'SubmitRequest':   return _txn_(function () { return svcSubmitRequest(p); });
    case 'SubmitReturnProof': return _txn_(function () { return svcSubmitReturnProof(p); });
    case 'ApproveRequest':  return _txn_(function () { return svcDecideRequest(p, 'approved'); });
    case 'RejectRequest':   return _txn_(function () { return svcDecideRequest(p, 'rejected'); });
    case 'ApproveGroup':    return _txn_(function () { return svcDecideGroup(p, 'approved'); });
    case 'RejectGroup':     return _txn_(function () { return svcDecideGroup(p, 'rejected'); });
    case 'AddItem':         return _txn_(function () { return svcAddItem(p); });
    case 'UpdateItem':      return _txn_(function () { return svcUpdateItem(p); });
    case 'DeleteItem':      return _txn_(function () { return svcDeleteItem(p); });
    case 'StockAdd':        return _txn_(function () { return svcStockChange(p, 'stock_add'); });
    case 'StockRemove':     return _txn_(function () { return svcStockChange(p, 'stock_remove'); });
    case 'CheckOut':        return _txn_(function () { return svcCheckOut(p); });
    case 'CheckIn':         return _txn_(function () { return svcCheckIn(p); });
    case 'Decommission':    return _txn_(function () { return svcDecommission(p); });
    case 'AuditAdjust':     return _txn_(function () { return svcAuditAdjust(p); });
    case 'SaveCategory':    return _txn_(function () { return svcSaveRef('Categories', p); });
    case 'DeleteCategory':  return _txn_(function () { return svcDeleteRef('Categories', p); });
    case 'SaveLocation':    return _txn_(function () { return svcSaveRef('Locations', p); });
    case 'DeleteLocation':  return _txn_(function () { return svcDeleteRef('Locations', p); });
    case 'SaveCustodian':   return _txn_(function () { return svcSaveRef('Custodians', p); });
    case 'DeleteCustodian': return _txn_(function () { return svcDeleteRef('Custodians', p); });
    case 'RunDateCheck':    return { success: true, result: checkDates(true) };
    case 'SeedReference':   return _txn_(function () { return svcSeedReference(); });
    case 'SaveConfig':      return _txn_(function () { return svcSaveConfig(p); });
    default:                return { success: false, error: 'Tindakan tidak dikenali: ' + actionType };
  }
}

// ============================================================
//  INVENTORY SERVICE  ("App\Services\InventoryService")
//
//  Every function below runs INSIDE _txn_ (see handleAction) and every
//  one that changes stock or lifecycle appends exactly one ledger row.
//  None of them may be called directly from doPost.
// ============================================================

// Writes one immutable ledger row. The single place Transactions is
// ever written; there is no update path at all.
function _ledger_(fields) {
  const now = new Date();
  return _insert_('Transactions', {
    id:                   _nextId_('Transactions'),
    item_id:              fields.item_id,
    user_id:              fields.user_id || 'admin',
    custodian_id:         fields.custodian_id === undefined ? null : fields.custodian_id,
    action_type:          fields.action_type,
    quantity:             fields.quantity,
    transaction_date:     fields.transaction_date || now,
    expected_return_date: fields.expected_return_date || null,
    actual_return_date:   fields.actual_return_date || null,
    photo_borrower:       fields.photo_borrower || '',
    photo_admin:          fields.photo_admin || '',
    source:               fields.source || '',
    reason_notes:         fields.reason_notes || '',
    created_at:           now
  });
}

// A capability token, not a password. Handed to one borrower for one loan,
// it grants exactly three things: read that item's public details, one Drive
// upload, and attach a photo URL to that loan. Cleared at check-in.
function _token_() {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';   // no I/O/0/1 — these get read aloud
  var out = '';
  for (var i = 0; i < 24; i++) out += abc.charAt(Math.floor(Math.random() * abc.length));
  return out;
}

// ── Asset tag: AST-YYYY-NNNN ────────────────────────────────
// Sequence is per acquisition year and derived from the tags already in
// the sheet, so it survives row deletion and manual edits. Called only
// inside the lock, which is what stops two simultaneous adds from
// claiming the same number.
function _nextAssetTag_(year) {
  const y     = String(year || new Date().getFullYear());
  const items = _readTable_('Items', true); // include soft-deleted: their tags are still taken
  const re    = new RegExp('^AST-' + y + '-(\\d{4,})$');
  var max = 0;
  items.forEach(function (it) {
    var m = re.exec(String(it.asset_tag || ''));
    if (m) { var n = parseInt(m[1], 10); if (n > max) max = n; }
  });
  const next = max + 1;
  // Past 9999 the field simply widens to 5 digits rather than wrapping
  // and colliding with AST-YYYY-0001.
  const padded = next < 10000 ? ('000' + next).slice(-4) : String(next);
  return 'AST-' + y + '-' + padded;
}

const ITEM_SCHEMA = {
  name:                 { required: true,  max: 120, label: 'Nama item' },
  location_id:          { required: true,  type: 'number', label: 'Lokasi' },
  item_type:            { required: true,  in: ITEM_TYPES, label: 'Jenis item' },
  serial_number:        { max: 80,  label: 'Nombor siri' },
  quantity_total:       { type: 'number', min: 0, def: 1, label: 'Kuantiti' },
  min_stock_alert:      { type: 'number', min: 0, def: 0, label: 'Amaran stok minimum' },
  date_acquired:        { required: true, type: 'date', label: 'Tarikh perolehan' },
  date_next_inspection: { type: 'date', label: 'Tarikh pemeriksaan seterusnya' },
  date_last_maintained: { type: 'date', label: 'Tarikh penyelenggaraan terakhir' },
  date_last_audited:    { type: 'date', label: 'Tarikh audit terakhir' },
  photo_url:            { max: 500, label: 'Foto aset' },
  receipt_url:          { max: 500, label: 'Resit/Invois' },
  notes:                { max: 2000, label: 'Catatan' }
};

function svcAddItem(p) {
  const v = _validate_(ITEM_SCHEMA, p);

  // Referential integrity — Sheets has no foreign keys, so it is checked here.
  if (!_findById_(_readTable_('Locations'), v.location_id)) throw new Error('Lokasi tidak wujud.');

  // A fixed asset is one physical thing; quantity is meaningless above 1.
  const qty = (v.item_type === 'fixed_asset') ? 1 : Math.max(0, Number(v.quantity_total || 0));
  const now = new Date();
  const id  = _nextId_('Items');
  // Only movable assets carry a tag — an asset tag identifies ONE physical
  // thing you can point at and label. Inventory is a quantity of
  // interchangeable units, so it is identified by name instead.
  const tag = (v.item_type === 'fixed_asset') ? _nextAssetTag_(v.date_acquired.getFullYear()) : '';

  const item = {
    id: id,
    asset_tag: tag,
    name: v.name,
    location_id: v.location_id,
    serial_number: v.serial_number || '',
    item_type: v.item_type,
    quantity_total: qty,
    quantity_available: qty,
    min_stock_alert: v.min_stock_alert || 0,
    status: qty > 0 ? 'available' : 'maintenance',
    date_acquired: v.date_acquired,
    date_last_maintained: v.date_last_maintained,
    date_next_inspection: v.date_next_inspection,
    date_last_audited: v.date_last_audited,
    date_decommissioned: null,
    custodian_id: null,
    photo_url: v.photo_url || '',
    receipt_url: v.receipt_url || '',
    notes: v.notes || '',
    created_at: now,
    updated_at: now,
    deleted_at: null
  };
  _insert_('Items', item);

  _ledger_({
    item_id: id, action_type: 'stock_add', quantity: qty,
    transaction_date: v.date_acquired,
    reason_notes: 'Pendaftaran item baharu (' + tag + ')'
  });

  return { id: id, asset_tag: tag };
}

// Edits descriptive/lifecycle fields. Deliberately cannot change stock
// levels or status — those only move through the ledgered actions, so
// the ledger can never disagree with the item row.
function svcUpdateItem(p) {
  const items = _readTable_('Items');
  const item  = _findById_(items, p.id);
  if (!item) throw new Error('Item tidak dijumpai.');

  const v = _validate_(ITEM_SCHEMA, p);
  if (!_findById_(_readTable_('Locations'), v.location_id)) throw new Error('Lokasi tidak wujud.');

  _update_('Items', item._row, {
    name: v.name,
    location_id: v.location_id,
    serial_number: v.serial_number || '',
    item_type: v.item_type,
    min_stock_alert: v.min_stock_alert || 0,
    date_acquired: v.date_acquired,
    date_last_maintained: v.date_last_maintained,
    date_next_inspection: v.date_next_inspection,
    date_last_audited: v.date_last_audited,
    photo_url: v.photo_url || '',
    receipt_url: v.receipt_url || '',
    notes: v.notes || '',
    updated_at: new Date()
  });
  return { id: item.id };
}

// Soft delete. The row and its whole ledger history stay in the sheet.
// Removes an item that should never have been registered — a duplicate, or
// a typo. This is NOT svcDecommission: that records the disposal of
// something which genuinely existed and keeps it in the register.
//
// This is a HARD delete. The item row and its ledger rows are removed
// outright, because a mis-entry's ledger rows describe stock that never
// arrived; leaving them behind would keep a fiction in the audit trail.
//
// The dangerous case is an item that had a real life — stock issued to
// staff — whose ledger rows record things that DID happen. Those are only
// purged when the caller explicitly acknowledges it (`confirm_history`), so
// it can never happen by accident.
function svcDeleteItem(p) {
  const v = _validate_({
    id:     { required: true, type: 'number', label: 'Item' },
    reason: { required: true, max: 200, label: 'Sebab' }
  }, p);

  const item = _findById_(_readTable_('Items'), v.id);
  if (!item) throw new Error('Item tidak dijumpai.');

  // Somebody is physically holding this. Deleting it would strand the loan
  // with nothing to return to.
  if (_openLoans_(_readTable_('Transactions')).some(function (l) {
    return Number(l.item_id) === Number(item.id);
  })) {
    throw new Error('Item masih dipinjam. Daftar masuk dahulu sebelum memadam.');
  }

  // Anything beyond the registration stock_add means this item was actually
  // used, and its ledger rows record real events.
  const all = _readTable_('Transactions').filter(function (t) {
    return Number(t.item_id) === Number(item.id);
  });
  const real = all.filter(function (t) { return t.action_type !== 'stock_add'; }).length;

  if (real > 0 && !p.confirm_history) {
    throw new Error('Item ini ada ' + real + ' pergerakan sebenar direkodkan. ' +
      'Memadamnya akan membuang rekod tersebut secara kekal. ' +
      'Sahkan sekali lagi, atau gunakan Lupuskan untuk mengekalkan rekodnya.');
  }

  const purged = _purgeLedgerRows_(item.id);
  _sheet_('Items').deleteRow(item._row);

  return {
    id: item.id, name: item.name, asset_tag: item.asset_tag || '',
    ledger_removed: purged, real_movements: real, reason: v.reason
  };
}

// ── Consumable stock in/out ─────────────────────────────────
// `custodian_id` is optional and only meaningful on a removal: it records
// WHO the stock was issued to. Without it, "who took the pens" could only
// ever be free text in reason_notes — unsearchable and impossible to total
// per person. Removals for damage or loss legitimately have no recipient,
// which is why it is optional rather than required.
function svcStockChange(p, action) {
  const v = _validate_({
    id:            { required: true, type: 'number', label: 'Item' },
    quantity:      { required: true, type: 'number', min: 1, label: 'Kuantiti' },
    custodian_id:  { type: 'number', label: 'Penerima' },
    // Topup only. Stock often arrives days before anyone records it, so the
    // received date is captured rather than assumed to be today.
    received_date: { type: 'date', label: 'Tarikh terima' },
    source:        { max: 120, label: 'Sumber' },
    reason_notes:  { required: true, max: 500, label: 'Sebab' }
  }, p);

  const item = _findById_(_readTable_('Items'), v.id);
  if (!item) throw new Error('Item tidak dijumpai.');
  if (item.status === 'decommissioned' || item.status === 'disposed') {
    throw new Error('Item telah dilupuskan — stok tidak boleh diubah.');
  }

  var recipient = null;
  if (v.custodian_id) {
    recipient = _findById_(_readTable_('Custodians'), v.custodian_id);
    if (!recipient) throw new Error('Penerima tidak wujud.');
    if (action === 'stock_add') {
      throw new Error('Penerima hanya untuk pengeluaran stok, bukan penambahan.');
    }
  }

  const isAdd = (action === 'stock_add');

  // Backdating a delivery is normal. Post-dating one is not: the stock is not
  // physically there yet, so counting it would make the balance a lie.
  var when = new Date();
  var source = '';
  if (isAdd) {
    if (v.received_date) {
      if (_daysBetween_(_today_(), v.received_date) > 0) {
        throw new Error('Tarikh terima tidak boleh pada masa hadapan — stok belum sampai.');
      }
      when = v.received_date;
    }
    source = v.source || '';
  } else if (v.received_date || v.source) {
    // An issue happens at the counter, in the moment. Allowing either field
    // here would let a withdrawal be quietly backdated or attributed to a
    // supplier, neither of which means anything.
    throw new Error('Tarikh terima dan sumber hanya untuk penambahan stok.');
  }

  const delta = isAdd ? v.quantity : -v.quantity;
  const avail = Number(item.quantity_available || 0) + delta;
  const total = Number(item.quantity_total || 0) + delta;
  if (avail < 0) {
    throw new Error('Stok tidak mencukupi. Baki semasa: ' + item.quantity_available + '.');
  }

  _update_('Items', item._row, {
    quantity_available: avail,
    quantity_total:     Math.max(0, total),
    status:             avail > 0 ? (item.status === 'assigned' ? 'assigned' : 'available') : 'maintenance',
    updated_at:         new Date()
  });

  _ledger_({
    item_id: item.id, action_type: action, quantity: delta,
    custodian_id: recipient ? recipient.id : null,
    transaction_date: when,
    source: source,
    reason_notes: v.reason_notes
  });

  return {
    id: item.id, quantity_available: avail,
    custodian_id: recipient ? recipient.id : null,
    transaction_date: when, source: source
  };
}

// ── Check-out ───────────────────────────────────────────────
function svcCheckOut(p) {
  const v = _validate_({
    id:                   { required: true, type: 'number', label: 'Item' },
    custodian_id:         { type: 'number', label: 'Penerima' },
    recipient_name:       { max: 120, label: 'Nama penerima' },
    recipient_email:      { type: 'email', max: 120, label: 'Emel penerima' },
    recipient_department: { max: 80, label: 'Jabatan' },
    expected_return_date: { required: true, type: 'date', label: 'Tarikh jangka pulang' },
    quantity:             { type: 'number', min: 1, def: 1, label: 'Kuantiti' },
    reason_notes:         { max: 500, label: 'Catatan' }
  }, p);

  const item = _findById_(_readTable_('Items'), v.id);
  if (!item) throw new Error('Item tidak dijumpai.');
  if (item.status === 'decommissioned' || item.status === 'disposed') {
    throw new Error('Item telah dilupuskan dan tidak boleh dipinjam.');
  }

  const custodian = _resolveRecipient_(v);

  // A fixed asset already on loan has no second unit to hand out. This is
  // checked BEFORE the stock check: an asset that is out on loan also has
  // quantity_available 0, and "stok tidak mencukupi" would be a true but
  // useless answer to "why can't I borrow this?".
  if (item.item_type === 'fixed_asset') {
    var open = _openLoans_(_readTable_('Transactions')).filter(function (l) {
      return Number(l.item_id) === Number(item.id);
    });
    if (open.length) throw new Error('Aset ini sudah dipinjam dan belum dipulangkan.');
  }

  const qty = (item.item_type === 'fixed_asset') ? 1 : v.quantity;
  const avail = Number(item.quantity_available || 0) - qty;
  if (avail < 0) throw new Error('Stok tidak mencukupi untuk dipinjam. Baki: ' + item.quantity_available + '.');

  if (_daysBetween_(_today_(), v.expected_return_date) < 0) {
    throw new Error('Tarikh jangka pulang tidak boleh sebelum hari ini.');
  }

  const returnToken = _token_();
  _update_('Items', item._row, {
    quantity_available: avail,
    status:             'assigned',
    custodian_id:       custodian.id,
    return_token:       returnToken,
    return_photo_url:   '',
    return_claimed_at:  null,
    updated_at:         new Date()
  });

  _ledger_({
    item_id: item.id, custodian_id: custodian.id,
    action_type: 'check_out', quantity: -qty,
    expected_return_date: v.expected_return_date,
    reason_notes: v.reason_notes || ''
  });

  const txnId = _nextId_('Transactions') - 1;   // the row just appended
  const cat = (item.item_type === 'fixed_asset') ? 'Aset Alih' : 'Alat Tulis';
  const locRow = _findById_(_readTable_('Locations'), item.location_id);
  const loc = locRow ? [locRow.building, locRow.floor, locRow.room_number].filter(String).join(' · ') : '';

  // Notification is best-effort: a mail quota failure must not roll back a
  // check-out that already happened and is already in the ledger.
  var mailed = true;
  try { _mailCheckOut_(item, custodian, v.expected_return_date, qty, txnId, loc, cat, returnToken); }
  catch (e) { mailed = false; }

  return {
    id: item.id, custodian_id: custodian.id, custodian_name: custodian.name,
    custodian_email: custodian.email, quantity_available: avail, mailed: mailed
  };
}

// ── Check-in ────────────────────────────────────────────────
function svcCheckIn(p) {
  const v = _validate_({
    id:                 { required: true, type: 'number', label: 'Item' },
    actual_return_date: { type: 'date', label: 'Tarikh pulang' },
    photo_admin:        { max: 500, label: 'Foto admin' },
    reason_notes:       { max: 500, label: 'Catatan' }
  }, p);

  const item = _findById_(_readTable_('Items'), v.id);
  if (!item) throw new Error('Item tidak dijumpai.');

  const loans = _openLoans_(_readTable_('Transactions')).filter(function (l) {
    return Number(l.item_id) === Number(item.id);
  });
  if (!loans.length) throw new Error('Tiada rekod pinjaman terbuka untuk item ini.');

  const loan = loans[0];
  const qty  = Math.abs(Number(loan.quantity || 1));
  const returned = v.actual_return_date || new Date();

  // Clearing return_token is what expires the borrower's link: the loan it
  // referred to no longer exists.
  _update_('Items', item._row, {
    quantity_available: Number(item.quantity_available || 0) + qty,
    status:             'available',
    custodian_id:       null,
    return_token:       '',
    return_photo_url:   '',
    return_claimed_at:  null,
    updated_at:         new Date()
  });

  // A NEW ledger row — the original check_out row is never edited.
  _ledger_({
    item_id: item.id, custodian_id: loan.custodian_id,
    action_type: 'check_in', quantity: qty,
    expected_return_date: loan.expected_return_date,
    actual_return_date: returned,
    photo_borrower: item.return_photo_url || '',
    photo_admin: v.photo_admin || '',
    reason_notes: v.reason_notes || ''
  });

  const txnId = _nextId_('Transactions') - 1;
  const custodian = _findById_(_readTable_('Custodians'), loan.custodian_id);
  const cat = (item.item_type === 'fixed_asset') ? 'Aset Alih' : 'Alat Tulis';
  const locRow = _findById_(_readTable_('Locations'), item.location_id);
  const loc = locRow ? [locRow.building, locRow.floor, locRow.room_number].filter(String).join(' · ') : '';

  // Closes the handover in writing on both ends — the person who took the
  // asset gets written proof they gave it back.
  var mailed = true;
  try { _mailCheckIn_(item, custodian, loan, returned, txnId, loc, cat); }
  catch (e) { mailed = false; }

  return {
    id: item.id, mailed: mailed,
    custodian_name: custodian ? custodian.name : null,
    custodian_email: custodian ? custodian.email : null,
    returned_late: SCOPES.overdue({
      expected_return_date: loan.expected_return_date, actual_return_date: null
    }, _startOfDay_(returned))
  };
}

// ── Decommission / disposal (spec 5.2) ──────────────────────
function svcDecommission(p) {
  const v = _validate_({
    id:                  { required: true, type: 'number', label: 'Item' },
    reason:              { required: true, in: REMOVAL_REASONS, label: 'Sebab pelupusan' },
    date_decommissioned: { required: true, type: 'date', label: 'Tarikh pelupusan' },
    status:              { in: ['decommissioned', 'disposed'], def: 'decommissioned', label: 'Status' },
    reason_notes:        { max: 500, label: 'Catatan' }
  }, p);

  const item = _findById_(_readTable_('Items'), v.id);
  if (!item) throw new Error('Item tidak dijumpai.');
  if (item.status === 'decommissioned' || item.status === 'disposed') {
    throw new Error('Item ini telah pun dilupuskan.');
  }

  const openLoan = _openLoans_(_readTable_('Transactions')).filter(function (l) {
    return Number(l.item_id) === Number(item.id);
  });
  if (openLoan.length) throw new Error('Item masih dipinjam. Daftar masuk dahulu sebelum melupuskan.');

  const qty = Number(item.quantity_available || 0);

  _update_('Items', item._row, {
    status:              v.status,
    quantity_available:  0,
    date_decommissioned: v.date_decommissioned,
    custodian_id:        null,
    updated_at:          new Date()
  });

  _ledger_({
    item_id: item.id, action_type: 'decommission', quantity: -qty,
    transaction_date: v.date_decommissioned,
    reason_notes: v.reason + (v.reason_notes ? ' | ' + v.reason_notes : '')
  });

  return { id: item.id, status: v.status };
}

// ── Audit adjustment ────────────────────────────────────────
function svcAuditAdjust(p) {
  const v = _validate_({
    id:            { required: true, type: 'number', label: 'Item' },
    counted:       { required: true, type: 'number', min: 0, label: 'Kuantiti dikira' },
    date_audited:  { type: 'date', label: 'Tarikh audit' },
    reason_notes:  { max: 500, label: 'Catatan' }
  }, p);

  const item = _findById_(_readTable_('Items'), v.id);
  if (!item) throw new Error('Item tidak dijumpai.');

  const before   = Number(item.quantity_available || 0);
  const variance = v.counted - before;
  const audited  = v.date_audited || new Date();

  _update_('Items', item._row, {
    quantity_available: v.counted,
    quantity_total:     Math.max(Number(item.quantity_total || 0) + variance, v.counted),
    date_last_audited:  audited,
    updated_at:         new Date()
  });

  _ledger_({
    item_id: item.id, action_type: 'audit_adjust', quantity: variance,
    transaction_date: audited,
    reason_notes: 'Audit: dijangka ' + before + ', dikira ' + v.counted +
                  ' (varians ' + (variance >= 0 ? '+' : '') + variance + ')' +
                  (v.reason_notes ? ' | ' + v.reason_notes : '')
  });

  return { id: item.id, variance: variance };
}

// ── Reference tables (Categories / Locations / Custodians) ──
const REF_SCHEMAS = {
  Categories: {
    name: { required: true, max: 80, label: 'Nama kategori' },
    type: { required: true, in: ITEM_TYPES, label: 'Jenis' }
  },
  Locations: {
    building:    { required: true, max: 80, label: 'Bangunan' },
    floor:       { max: 40, label: 'Tingkat' },
    room_number: { max: 40, label: 'Nombor bilik' }
  },
  Custodians: {
    employee_id: { required: true, max: 40, label: 'ID pekerja' },
    name:        { required: true, max: 120, label: 'Nama' },
    email:       { required: true, type: 'email', max: 120, label: 'Emel' },
    department:  { max: 80, label: 'Jabatan' }
  }
};

function svcSaveRef(table, p) {
  const v   = _validate_(REF_SCHEMAS[table], p);
  const now = new Date();
  const rows = _readTable_(table);

  // Uniqueness the spec declares but Sheets cannot enforce.
  if (table === 'Custodians') {
    var clash = rows.filter(function (r) {
      if (p.id && Number(r.id) === Number(p.id)) return false;
      return String(r.employee_id) === v.employee_id ||
             String(r.email).toLowerCase() === v.email.toLowerCase();
    });
    if (clash.length) throw new Error('ID pekerja atau emel tersebut telah digunakan.');
  }

  if (p.id) {
    var existing = _findById_(rows, p.id);
    if (!existing) throw new Error('Rekod tidak dijumpai.');
    v.updated_at = now;
    _update_(table, existing._row, v);
    return { id: existing.id };
  }

  const id = _nextId_(table);
  v.id = id; v.created_at = now; v.updated_at = now;
  _insert_(table, v);
  return { id: id };
}

function svcDeleteRef(table, p) {
  const row = _findById_(_readTable_(table), p.id);
  if (!row) throw new Error('Rekod tidak dijumpai.');

  // Block deletes that would orphan an item — Sheets has no ON DELETE.
  const items = _readTable_('Items');
  const fk = { Locations: 'location_id', Custodians: 'custodian_id' }[table];
  if (!fk) { _sheet_(table).deleteRow(row._row); return { id: row.id }; }
  const used = items.filter(function (it) { return Number(it[fk]) === Number(row.id); });
  if (used.length) {
    throw new Error('Tidak boleh dipadam — masih digunakan oleh ' + used.length + ' item.');
  }

  _sheet_(table).deleteRow(row._row);
  return { id: row.id };
}

// Upserts by key, so adding a setting later needs no schema change.
function svcSaveConfig(p) {
  const v = _validate_({
    admin_name:  { required: true, max: 120, label: 'Nama pentadbir' },
    admin_email: { required: true, type: 'email', max: 120, label: 'Emel pentadbir' },
    admin_cc:    { max: 240, label: 'Emel salinan' }
  }, p);

  // Each CC address is checked individually — one bad entry in a list would
  // otherwise make MailApp reject the whole send at notification time,
  // silently, long after this form was saved.
  if (v.admin_cc) {
    const bad = String(v.admin_cc).split(/[,;]+/)
      .map(function (x) { return x.trim(); }).filter(Boolean)
      .filter(function (x) { return !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x); });
    if (bad.length) throw new Error('Emel salinan tidak sah: ' + bad.join(', '));
  }

  const now  = new Date();
  const rows = _readTable_('Config');
  Object.keys(v).forEach(function (k) {
    var existing = null;
    for (var i = 0; i < rows.length; i++) if (String(rows[i].key) === k) existing = rows[i];
    if (existing) _update_('Config', existing._row, { value: v[k] === null ? '' : v[k], updated_at: now });
    else _insert_('Config', { key: k, value: v[k] === null ? '' : v[k], updated_at: now });
  });
  return { saved: true };
}

// First-run reference data, so a fresh sheet is usable immediately.
// Idempotent: skips any table that already has rows.
function svcSeedReference() {
  const now = new Date();
  const added = { locations: 0 };

  if (!_readTable_('Locations').length) {
    [['IMEN', 'Aras 1', 'Pejabat Am'], ['IMEN', 'Aras 1', 'Stor'],
     ['IMEN', 'Aras 2', 'Makmal 1'], ['IMEN', 'Aras 2', 'Bilik Mesyuarat']
    ].forEach(function (l, i) {
      _insert_('Locations', {
        id: i + 1, building: l[0], floor: l[1], room_number: l[2],
        created_at: now, updated_at: now
      });
      added.locations++;
    });
  }

  return added;
}

// ============================================================
//  EMAIL  (helpers ported from i-print/_backend/Code.gs)
//
//  _esc_ is not optional: item names, notes and custodian names are
//  user-supplied and would otherwise be injected raw into HTML mail.
// ============================================================
function _esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function _emailLayout_(title, inner) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;">' +
    '<div style="background:#0f172a;color:#ffffff;padding:16px 20px;font-size:15px;font-weight:bold;">i-Nventori Ofis — ' + _esc_(title) + '</div>' +
    '<div style="padding:20px;color:#1e293b;font-size:14px;line-height:1.6;">' + inner + '</div>' +
    '<div style="padding:12px 20px;background:#f8fafc;color:#94a3b8;font-size:11px;">i-Nventori Ofis &middot; IMEN, UKM &middot; Emel automatik, jangan balas.</div>' +
  '</div>';
}

function _kvRows_(pairs) {
  return '<table style="width:100%;border-collapse:collapse;font-size:13px;margin:12px 0;">' +
    pairs.map(function (p) {
      return '<tr>' +
        '<td style="padding:6px 8px;color:#64748b;width:42%;vertical-align:top;border-bottom:1px solid #f1f5f9;">' + _esc_(p[0]) + '</td>' +
        '<td style="padding:6px 8px;font-weight:bold;color:#0f172a;border-bottom:1px solid #f1f5f9;">' + _esc_(p[1]) + '</td>' +
      '</tr>';
    }).join('') + '</table>';
}

function _fmtDate_(d) {
  if (!d) return '—';
  const x = (d instanceof Date) ? d : new Date(d);
  if (isNaN(x.getTime())) return '—';
  const dd = ('0' + x.getDate()).slice(-2);
  const mm = ('0' + (x.getMonth() + 1)).slice(-2);
  return dd + '/' + mm + '/' + x.getFullYear();
}

function _fmtDateTime_(d) {
  if (!d) return '—';
  const x = (d instanceof Date) ? d : new Date(d);
  if (isNaN(x.getTime())) return '—';
  return _fmtDate_(x) + ', ' + ('0' + x.getHours()).slice(-2) + ':' + ('0' + x.getMinutes()).slice(-2);
}

// Sent to the recipient on check-out (spec 5.3).
function _mailCheckOut_(item, custodian, expected, qty, txnId, loc, cat, returnToken) {
  if (!custodian || !custodian.email) return;
  const now = new Date();
  const body = _emailLayout_('Aset Didaftar Keluar',
    '<p>Salam <b>' + _esc_(custodian.name) + '</b>,</p>' +
    '<p>Aset berikut telah <b>didaftarkan keluar kepada anda</b> dan kini di bawah jagaan anda:</p>' +
    _kvRows_([
      ['Rujukan', 'TRX_' + ('00000' + txnId).slice(-6)],
      ['Tag Aset', item.asset_tag],
      ['Nama Aset', item.name],
      ['Kategori', cat || '—'],
      ['Lokasi Asal', loc || '—'],
      ['Nombor Siri', item.serial_number || '—'],
      ['Kuantiti', String(qty)],
      ['Tarikh & Masa Didaftar Keluar', _fmtDateTime_(now)],
      ['Tarikh Jangka Pulang', _fmtDate_(expected)],
      ['Direkod Oleh', _adminName_()]
    ]) +
    '<p style="color:#b45309;margin-top:14px;">Sila pulangkan sebelum atau pada <b>' + _fmtDate_(expected) + '</b>. ' +
    'Peringatan automatik akan dihantar jika aset belum dipulangkan selepas tarikh tersebut.</p>' +
    (returnToken
      ? '<div style="margin:18px 0;padding:16px;background:#fff7ed;border:1px solid #fed7aa;border-radius:10px;">' +
        '<p style="margin:0 0 10px;font-size:13px;color:#9a3412;"><b>Bila anda pulangkan aset ini</b>, tekan butang ' +
        'di bawah pada telefon anda. Kamera akan terbuka — snap keadaan aset sebagai bukti hantar. ' +
        'Tiada log masuk diperlukan.</p>' +
        '<a href="' + PAGES_ORIGIN + '/i-Nventoriofis/?pulang=' + returnToken + '" ' +
        'style="display:inline-block;background:#f4511e;color:#ffffff;text-decoration:none;' +
        'padding:11px 20px;border-radius:999px;font-size:14px;font-weight:bold;">Pulangkan &amp; Snap Bukti</a>' +
        '</div>'
      : '') +
    '<p style="font-size:12px;color:#64748b;">Simpan emel ini sebagai rekod penerimaan anda.</p>'
  );
  MailApp.sendEmail({
    to: custodian.email, cc: [_adminEmail_(), _adminCc_()].filter(Boolean).join(','),
    subject: '[i-Nventori Ofis] Aset ' + item.asset_tag + ' didaftar keluar kepada anda',
    htmlBody: body
  });
}

// Sent to the same person on return, so the handover is closed off in
// writing on both ends — the recipient gets proof they returned it.
function _mailCheckIn_(item, custodian, loan, returned, txnId, loc, cat) {
  if (!custodian || !custodian.email) return;
  const out = loan && loan.transaction_date ? _fmtDateTime_(loan.transaction_date) : '—';
  const days = loan && loan.transaction_date ? _daysBetween_(loan.transaction_date, returned) : null;
  const late = loan && loan.expected_return_date
    ? _daysBetween_(loan.expected_return_date, returned) : null;
  const status = (late === null) ? 'Selesai'
    : (late > 0 ? 'Lewat ' + late + ' hari' : 'Dipulangkan tepat pada masa');

  const body = _emailLayout_('Aset Telah Dipulangkan',
    '<p>Salam <b>' + _esc_(custodian.name) + '</b>,</p>' +
    '<p>Terima kasih. Pemulangan aset berikut telah <b>direkodkan</b> dan anda tidak lagi bertanggungjawab ke atasnya:</p>' +
    _kvRows_([
      ['Rujukan', 'TRX_' + ('00000' + txnId).slice(-6)],
      ['Tag Aset', item.asset_tag],
      ['Nama Aset', item.name],
      ['Kategori', cat || '—'],
      ['Nombor Siri', item.serial_number || '—'],
      ['Tarikh & Masa Didaftar Keluar', out],
      ['Tarikh & Masa Dipulangkan', _fmtDateTime_(returned)],
      ['Tempoh Dipinjam', (days === null ? '—' : days + ' hari')],
      ['Status Pemulangan', status],
      ['Dikembalikan Ke', loc || '—'],
      ['Direkod Oleh', _adminName_()]
    ]) +
    '<p style="color:' + (late > 0 ? '#b45309' : '#047857') + ';margin-top:14px;">' +
    (late > 0
      ? 'Aset dipulangkan <b>' + late + ' hari selepas</b> tarikh jangka pulang.'
      : 'Aset dipulangkan dalam tempoh yang ditetapkan. Terima kasih.') + '</p>' +
    '<p style="font-size:12px;color:#64748b;">Simpan emel ini sebagai bukti pemulangan anda.</p>'
  );
  MailApp.sendEmail({
    to: custodian.email, cc: [_adminEmail_(), _adminCc_()].filter(Boolean).join(','),
    subject: '[i-Nventori Ofis] Aset ' + item.asset_tag + ' telah dipulangkan',
    htmlBody: body
  });
}

// Resolves the person an asset is being handed to. Either an existing
// custodian is picked, or a name + email is typed straight into the
// check-out form. In the latter case the person is created (or matched by
// email) so the ledger keeps a real foreign key and the same person is
// reusable next time — rather than the name living as loose text.
function _resolveRecipient_(v) {
  if (v.custodian_id) {
    const found = _findById_(_readTable_('Custodians'), v.custodian_id);
    if (!found) throw new Error('Penerima tidak wujud.');
    return found;
  }
  if (!v.recipient_name || !v.recipient_email) {
    throw new Error('Sila pilih penerima, atau masukkan nama dan emel penerima.');
  }
  const email = String(v.recipient_email).trim().toLowerCase();
  const rows  = _readTable_('Custodians');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].email || '').toLowerCase() === email) return rows[i];
  }
  const now = new Date();
  const id  = _nextId_('Custodians');
  const rec = {
    id: id,
    employee_id: 'ADH-' + ('000' + id).slice(-3),   // ad-hoc, still unique
    name: String(v.recipient_name).trim(),
    email: email,
    department: v.recipient_department || '',
    created_at: now, updated_at: now
  };
  _insert_('Custodians', rec);
  return rec;
}

// Acknowledgement to the requester, and a heads-up to the officer. Both
// list every line of the submission, so one request means one email
// regardless of how many items it carries.
function _mailRequestReceived_(lines, v, groupId) {
  const now = new Date();
  const itemRows = lines.map(function (ln) {
    return [ln.item.name + (ln.item.asset_tag ? ' (' + ln.item.asset_tag + ')' : ''),
            (ln.isAsset ? 'Aset Alih' : 'Alat Tulis') + ' × ' + ln.quantity];
  });
  const meta = [
    ['Rujukan', groupId],
    ['Bilangan Barang', String(lines.length)],
    ['Tujuan', v.purpose],
    ['Tarikh Diperlukan', v.needed_date ? _fmtDate_(v.needed_date) : '—'],
    ['Dihantar Pada', _fmtDateTime_(now)]
  ];

  MailApp.sendEmail({
    to: v.requester_email,
    subject: '[i-Nventori Ofis] Permohonan diterima — ' + lines.length + ' barang',
    htmlBody: _emailLayout_('Permohonan Diterima',
      '<p>Salam <b>' + _esc_(v.requester_name) + '</b>,</p>' +
      '<p>Permohonan anda telah diterima dan sedang <b>menunggu kelulusan</b>.</p>' +
      '<h3 style="font-size:14px;margin:16px 0 4px;">Barang Dipohon</h3>' +
      _kvRows_(itemRows) +
      _kvRows_(meta) +
      '<p style="font-size:12px;color:#64748b;">Anda akan menerima emel sebaik permohonan ini diluluskan atau ditolak.</p>')
  });

  MailApp.sendEmail({
    to: _adminEmail_(), cc: _adminCc_(),
    subject: '[i-Nventori Ofis] Permohonan baharu — ' + _esc_(v.requester_name) + ' (' + lines.length + ' barang)',
    htmlBody: _emailLayout_('Permohonan Baharu Menunggu',
      '<p>Permohonan baharu daripada <b>' + _esc_(v.requester_name) + '</b> (' +
        _esc_(v.requester_email) + '):</p>' +
      _kvRows_(itemRows) +
      _kvRows_(meta) +
      '<p>Buka i-Nventori Ofis &rarr; Permohonan untuk meluluskan atau menolak.</p>')
  });
}

function _mailRequestDecided_(item, req, decision, notes) {
  const approved = decision === 'approved';
  MailApp.sendEmail({
    to: req.requester_email,
    cc: [_adminEmail_(), _adminCc_()].filter(Boolean).join(','),
    subject: '[i-Nventori Ofis] Permohonan ' + (approved ? 'DILULUSKAN' : 'DITOLAK') + ' — ' + item.name,
    htmlBody: _emailLayout_(approved ? 'Permohonan Diluluskan' : 'Permohonan Ditolak',
      '<p>Salam <b>' + _esc_(req.requester_name) + '</b>,</p>' +
      '<p>Permohonan anda telah <b>' + (approved ? 'diluluskan' : 'ditolak') + '</b>.</p>' +
      _kvRows_([
        ['Rujukan', 'REQ_' + ('00000' + req.id).slice(-6)],
        ['Item', item.name],
        ['Kuantiti', String(req.quantity || 1)],
        ['Keputusan', approved ? 'Diluluskan' : 'Ditolak'],
        ['Catatan Admin', notes || '—'],
        ['Tarikh & Masa Keputusan', _fmtDateTime_(new Date())]
      ]) +
      (approved
        ? '<p style="color:#047857;">Sila hubungi pejabat untuk mengambil item. Emel berasingan mengandungi butiran penuh serahan.</p>'
        : '<p style="color:#b45309;">Sila hubungi admin jika anda memerlukan penjelasan lanjut.</p>'))
  });
}

// ============================================================
//  DATE AUTOMATION ENGINE  ("inventory:check-dates", spec 5.4)
//
//  Runs daily from a time-driven trigger, and on demand from the
//  dashboard's "Semak Tarikh" button. Uses SCOPES for all four checks so
//  the digest can never disagree with the badges the UI shows.
//
//  Returns its counts whether or not mail was sent, so the on-demand
//  call can render the same numbers.
// ============================================================
function checkDates(skipMail) {
  ensureSheets_();
  const today = _today_();
  const items = _readTable_('Items');
  const txns  = _readTable_('Transactions');
  const custs = _readTable_('Custodians');
  const loans = _openLoans_(txns);

  const inspection = items.filter(function (i) { return SCOPES.inspectionDue(i, today); });
  const audit      = items.filter(function (i) { return SCOPES.auditDue(i, today); });
  const lowStock   = items.filter(function (i) { return SCOPES.lowStock(i); });
  const overdue    = loans.filter(function (l) { return SCOPES.overdue(l, today); });

  const summary = {
    checked_at: new Date().toISOString(),
    inspection: inspection.length,
    audit:      audit.length,
    low_stock:  lowStock.length,
    overdue:    overdue.length
  };

  const total = inspection.length + audit.length + lowStock.length + overdue.length;
  if (skipMail || total === 0) return summary;

  var html = '<p>Ringkasan amaran inventori bagi ' + _fmtDate_(today) + ':</p>';

  if (overdue.length) {
    html += '<h3 style="font-size:14px;color:#be123c;margin:16px 0 4px;">Pinjaman lewat (' + overdue.length + ')</h3>' +
      _kvRows_(overdue.map(function (l) {
        var it = _findById_(items, l.item_id);
        var cu = _findById_(custs, l.custodian_id);
        return [
          (it ? it.asset_tag + ' — ' + it.name : 'Item #' + l.item_id) +
            (cu ? ' (' + cu.name + ')' : ''),
          'Sepatutnya pulang ' + _fmtDate_(l.expected_return_date)
        ];
      }));
  }
  if (inspection.length) {
    html += '<h3 style="font-size:14px;color:#b45309;margin:16px 0 4px;">Pemeriksaan tertunggak (' + inspection.length + ')</h3>' +
      _kvRows_(inspection.map(function (i) {
        return [i.asset_tag + ' — ' + i.name, _fmtDate_(i.date_next_inspection)];
      }));
  }
  if (lowStock.length) {
    html += '<h3 style="font-size:14px;color:#b45309;margin:16px 0 4px;">Stok rendah (' + lowStock.length + ')</h3>' +
      _kvRows_(lowStock.map(function (i) {
        return [i.asset_tag + ' — ' + i.name, 'Baki ' + i.quantity_available + ' (min ' + i.min_stock_alert + ')'];
      }));
  }
  if (audit.length) {
    html += '<h3 style="font-size:14px;color:#64748b;margin:16px 0 4px;">Audit tertunggak &gt;365 hari (' + audit.length + ')</h3>' +
      _kvRows_(audit.slice(0, 25).map(function (i) {
        return [i.asset_tag + ' — ' + i.name, i.date_last_audited ? _fmtDate_(i.date_last_audited) : 'Belum pernah diaudit'];
      }));
    if (audit.length > 25) html += '<p style="color:#94a3b8;font-size:12px;">…dan ' + (audit.length - 25) + ' lagi.</p>';
  }

  try {
    MailApp.sendEmail({
      to: _adminEmail_(), cc: _adminCc_(),
      subject: '[i-Nventori Ofis] ' + total + ' amaran inventori — ' + _fmtDate_(today),
      htmlBody: _emailLayout_('Amaran Harian', html)
    });
    summary.mailed = true;
  } catch (e) {
    summary.mailed = false;
    summary.mail_error = String(e);
  }
  return summary;
}

// ============================================================
//  TRIGGER INSTALLER  ("Laravel Scheduler")
//
//  Run ONCE manually from the Apps Script editor after deploying.
//  Removes its own previous triggers first so running it twice does not
//  produce two daily emails.
// ============================================================
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkDates') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkDates').timeBased().everyDays(1).atHour(7).create();
  return 'Trigger harian checkDates dipasang (07:00).';
}

// ============================================================
//  DRIVE UPLOAD  (asset photo / receipt)
//
//  Ported from i-print. The bytes never pass through Apps Script: Drive
//  opens a resumable session and the browser PUTs the file straight to
//  Google's endpoint.
//
//  SECURITY: returns the session URI, never ScriptApp.getOAuthToken().
//  The URI authorises exactly one upload into one folder; a leaked OAuth
//  token would expose the owner's entire Drive.
// ============================================================
function _safeOrigin_(origin) {
  if (origin === PAGES_ORIGIN) return origin;
  if (origin === 'https://script.google.com') return origin;
  // Apps Script's sandbox origin is https://n-<random>-script.googleusercontent.com
  // — the hash differs per session, so it has to be matched by shape.
  if (/^https:\/\/[a-z0-9-]+\.googleusercontent\.com$/.test(origin || '')) return origin;
  return PAGES_ORIGIN; // unknown origin -> fall back, never echo it blindly
}

// The Drive session body, shared so the admin path and the borrower's
// token path cannot drift apart.
function _driveSession_(fileName, mimeType, origin) {
  const safeName = String(fileName || 'lampiran').replace(/[^\w.\- ]+/g, '_').slice(0, 80);
  const meta = {
    name:     Date.now() + '_' + safeName,
    parents:  [FOLDER_ID],
    mimeType: mimeType || 'application/octet-stream'
  };
  const res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true',
    {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        Origin: _safeOrigin_(origin)
      },
      payload: JSON.stringify(meta),
      muteHttpExceptions: true
    }
  );
  if (res.getResponseCode() !== 200) {
    return { ok: false, error: 'Drive init HTTP ' + res.getResponseCode() };
  }
  const headers    = res.getAllHeaders();
  const sessionUri = headers['Location'] || headers['location'];
  if (!sessionUri) return { ok: false, error: 'Drive tidak memulangkan session URI.' };
  return { ok: true, sessionUri: sessionUri };
}

function startResumableUpload(fileName, mimeType, origin, adminPass) {
  try {
    _requireAdmin_(adminPass);
    return _driveSession_(fileName, mimeType, origin);
  } catch (e) {
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

// ============================================================
//  RETURN PROOF  (borrower side, token-gated)
//
//  The borrower never logs in. The check-out email carries ?pulang=<token>,
//  and that token is the only credential — scoped to one open loan, and dead
//  the moment the loan is closed.
// ============================================================

// Resolves a token to its item, or throws. Every token path goes through
// here so the "open loan only" rule is stated once.
function _itemByReturnToken_(token) {
  const t = String(token || '').trim();
  if (!t || t.length < 12) throw new Error('Pautan tidak sah.');
  const rows = _readTable_('Items').filter(function (i) {
    return String(i.return_token || '') === t;
  });
  if (!rows.length) {
    throw new Error('Pautan tidak sah atau telah tamat tempoh. Aset ini mungkin sudah direkod sebagai dipulangkan.');
  }
  const item = rows[0];
  if (item.status === 'decommissioned' || item.status === 'disposed') {
    throw new Error('Aset ini telah dilupuskan.');
  }
  return item;
}

// Public read. Carries only what the borrower needs to recognise the item
// they are holding — never the holder's name, never the ledger.
function getReturnContext(token) {
  ensureSheets_();
  const item = _itemByReturnToken_(token);
  const loans = _openLoans_(_readTable_('Transactions')).filter(function (l) {
    return Number(l.item_id) === Number(item.id);
  });
  return {
    name: item.name,
    asset_tag: item.asset_tag || '',
    expected_return_date: loans.length ? loans[0].expected_return_date : null,
    already_claimed: !!item.return_claimed_at,
    claimed_at: item.return_claimed_at || null
  };
}

function startReturnUpload(token, fileName, mimeType, origin) {
  try {
    ensureSheets_();
    _itemByReturnToken_(token);          // throws unless the token is live
    return _driveSession_(fileName, mimeType, origin);
  } catch (e) {
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

// Records the borrower's evidence and tells the officer it arrived. This does
// NOT close the loan — the admin still confirms physical receipt.
function svcSubmitReturnProof(p) {
  const item = _itemByReturnToken_(p && p.token);
  const url = String((p && p.photo_url) || '').trim();
  if (!/^https:\/\//.test(url)) throw new Error('Foto tidak sah.');

  const now = new Date();
  _update_('Items', item._row, {
    return_photo_url: url,
    return_claimed_at: now,
    updated_at: now
  });

  var mailed = true;
  try { _mailReturnClaimed_(item, url, now); } catch (e) { mailed = false; }
  return { name: item.name, claimed_at: now, mailed: mailed };
}

function _mailReturnClaimed_(item, url, when) {
  MailApp.sendEmail({
    to: _adminEmail_(), cc: _adminCc_(),
    subject: '[i-Nventori Ofis] Bukti pemulangan diterima — ' + (item.asset_tag || item.name),
    htmlBody: _emailLayout_('Bukti Pemulangan Diterima',
      '<p>Peminjam telah menghantar bukti pemulangan:</p>' +
      _kvRows_([
        ['Tag Aset', item.asset_tag || '—'],
        ['Nama Aset', item.name],
        ['Dihantar Pada', _fmtDateTime_(when)]
      ]) +
      '<p><a href="' + _esc_(url) + '">Lihat foto bukti</a></p>' +
      '<p style="color:#b45309;">Pemulangan <b>belum</b> direkod. Sila sahkan penerimaan ' +
      'melalui Aset Alih &rarr; Pinjaman &rarr; Daftar Masuk.</p>')
  });
}

// ============================================================
//  HTTP ENTRY POINTS
// ============================================================

// Serves the UI. Fetches index.html live from the repo so GitHub Pages
// and the /exec URL always render the same build (single source of
// truth), falling back to a stub if GitHub is unreachable.
function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.format === 'json') {
    return ContentService
      .createTextOutput(JSON.stringify(getInitialData(p.pass)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var html;
  try {
    const res = UrlFetchApp.fetch(UI_RAW_URL, { muteHttpExceptions: true });
    html = (res.getResponseCode() === 200) ? res.getContentText() : null;
  } catch (err) { html = null; }

  if (!html) {
    html = '<!DOCTYPE html><html lang="ms"><body style="font-family:system-ui;padding:40px">' +
           '<h2>i-Nventori</h2><p>UI tidak dapat dimuatkan dari GitHub buat sementara waktu. ' +
           'Sila cuba <a href="' + PAGES_ORIGIN + '/i-Nventoriofis/">' + PAGES_ORIGIN + '/i-Nventoriofis/</a>.</p>' +
           '</body></html>';
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle('i-Nventori Ofis')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// JSON API for the GitHub-Pages-hosted UI. The frontend shim POSTs
// {fn, args} as text/plain — a CORS "simple request", because Apps
// Script cannot answer a preflight.
//
// SECURITY: dispatch is an explicit switch. Only these four entry points
// are reachable from outside; an arbitrary function name is rejected,
// and every privileged action is gated inside handleAction().
function doPost(e) {
  var payload;
  try {
    const req  = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const args = Array.isArray(req.args) ? req.args : [];
    var result;
    switch (req.fn) {
      case 'getInitialData':       result = getInitialData(args[0]); break;
      case 'getPublicCatalog':     result = getPublicCatalog(); break;
      case 'getReturnContext':     result = getReturnContext(args[0]); break;
      case 'startReturnUpload':    result = startReturnUpload(args[0], args[1], args[2], args[3]); break;
      case 'handleAction':         result = handleAction(args[0], args[1]); break;
      case 'adminLogin':           result = adminLogin(args[0]); break;
      case 'getItemHistory':       result = getItemHistory(args[0], args[1]); break;
      case 'startResumableUpload': result = startResumableUpload(args[0], args[1], args[2], args[3]); break;
      default: throw new Error('Unknown API function: ' + req.fn);
    }
    payload = { ok: true, result: result };
  } catch (err) {
    payload = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(payload))
                       .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  ONE-TIME SETUP — run from the Apps Script editor
// ============================================================
function setup() {
  ensureSheets_();
  const seeded = _txn_(function () { return svcSeedReference(); });
  const trig   = installTriggers();
  return 'Sheets sedia. Seed: ' + JSON.stringify(seeded) + '. ' + trig;
}
