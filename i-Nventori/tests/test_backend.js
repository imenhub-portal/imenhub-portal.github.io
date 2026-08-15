/**
 * i-Nventori — backend unit tests.
 *
 * There is no local Apps Script emulator, so this loads Code.gs into a vm
 * context with hand-written mocks for the Google services and drives the
 * real service functions against an in-memory spreadsheet. That makes the
 * tests exercise the actual write paths (locking, ledger appends, stock
 * arithmetic) rather than just the pure helpers.
 *
 * Run:  node tests/test_backend.js
 * This file is local-only tooling and is not deployed anywhere.
 */
'use strict';

const fs = require('fs');
const vm = require('vm');
const path = require('path');

const CODE_PATH = path.join(__dirname, '..', 'Code.gs');

// ── In-memory spreadsheet ─────────────────────────────────────────────
// Mirrors just enough of the Range/Sheet API for Code.gs: 1-indexed rows
// and columns, dense 2-D backing array, ragged rows padded on read.
function makeSheet(name) {
  const data = []; // array of arrays, row 0 = header
  const api = {
    _name: name,
    _data: data,
    getName: () => name,
    getLastRow: () => data.length,
    getLastColumn: () => data.reduce((m, r) => Math.max(m, r.length), 0),
    setFrozenRows: () => api,
    appendRow(row) { data.push(row.slice()); return api; },
    deleteRow(rowNum) { data.splice(rowNum - 1, 1); return api; },
    getRange(row, col, numRows, numCols) {
      numRows = numRows || 1;
      numCols = numCols || 1;
      return {
        getValues() {
          const out = [];
          for (let r = 0; r < numRows; r++) {
            const src = data[row - 1 + r] || [];
            const line = [];
            for (let c = 0; c < numCols; c++) {
              const v = src[col - 1 + c];
              line.push(v === undefined ? '' : v);
            }
            out.push(line);
          }
          return out;
        },
        setValues(vals) {
          for (let r = 0; r < vals.length; r++) {
            const ri = row - 1 + r;
            if (!data[ri]) data[ri] = [];
            for (let c = 0; c < vals[r].length; c++) data[ri][col - 1 + c] = vals[r][c];
          }
          return this;
        },
        setValue(v) { return this.setValues([[v]]); }
      };
    }
  };
  return api;
}

function makeSpreadsheet() {
  const sheets = {};
  return {
    _sheets: sheets,
    getSheetByName: (n) => sheets[n] || null,
    insertSheet(n) { sheets[n] = makeSheet(n); return sheets[n]; },
    getSheets: () => Object.keys(sheets).map((k) => sheets[k]),
    deleteSheet(sh) { delete sheets[sh._name]; }
  };
}

// ── Harness ───────────────────────────────────────────────────────────
function loadBackend(opts) {
  opts = opts || {};
  const ss = makeSpreadsheet();
  const sentMail = [];
  const props = { ADMIN_PASSWORD: opts.adminPassword === undefined ? 'rahsia' : opts.adminPassword };
  const cache = {};
  let lockHeld = 0;
  let lockMaxDepth = 0;

  // Every setValue/deleteRow that touches the ledger is recorded, so a
  // test can assert the ledger is genuinely append-only.
  const ledgerMutations = [];
  const origInsert = ss.insertSheet.bind(ss);
  ss.insertSheet = (n) => {
    const sh = origInsert(n);
    if (n === 'Transactions') {
      const realGetRange = sh.getRange.bind(sh);
      sh.getRange = (...a) => {
        const rng = realGetRange(...a);
        const wrap = { ...rng };
        wrap.setValue = (v) => { if (a[0] > 1) ledgerMutations.push(['setValue', a]); return rng.setValue(v); };
        wrap.setValues = (v) => { if (a[0] > 1) ledgerMutations.push(['setValues', a]); return rng.setValues(v); };
        wrap.getValues = rng.getValues;
        return wrap;
      };
      const realDelete = sh.deleteRow.bind(sh);
      sh.deleteRow = (r) => { ledgerMutations.push(['deleteRow', r]); return realDelete(r); };
    }
    return sh;
  };

  const sandbox = {
    console,
    SpreadsheetApp: { openById: () => ss },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (props[k] === undefined ? null : props[k]),
        setProperty: (k, v) => { props[k] = v; }
      })
    },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => (cache[k] === undefined ? null : cache[k]),
        put: (k, v) => { cache[k] = v; },
        remove: (k) => { delete cache[k]; }
      })
    },
    LockService: {
      getScriptLock: () => ({
        tryLock: () => {
          lockHeld++;
          lockMaxDepth = Math.max(lockMaxDepth, lockHeld);
          return true;
        },
        releaseLock: () => { lockHeld--; }
      })
    },
    MailApp: { sendEmail: (o) => { sentMail.push(o); } },
    UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getContentText: () => '', getAllHeaders: () => ({}) }) },
    ScriptApp: {
      getOAuthToken: () => 'token',
      getProjectTriggers: () => [],
      newTrigger: () => ({ timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create: () => {} }) }) }) }),
      deleteTrigger: () => {}
    },
    ContentService: { createTextOutput: (t) => ({ setMimeType: () => t }), MimeType: { JSON: 'json' } },
    HtmlService: {
      createHtmlOutput: (h) => ({ setTitle: function () { return this; }, addMetaTag: function () { return this; }, setXFrameOptionsMode: function () { return h; } }),
      XFrameOptionsMode: { ALLOWALL: 1 }
    },
    Utilities: { formatDate: (d) => String(d) }
  };

  vm.createContext(sandbox);
  // Top-level `function` declarations attach to the vm global, but `const`
  // ones stay in the script's lexical scope — so the constants Code.gs
  // declares have to be handed out explicitly, from inside the same script.
  const src = fs.readFileSync(CODE_PATH, 'utf8') +
    '\n;this.SCHEMA = SCHEMA; this.SCOPES = SCOPES; this.ITEM_TYPES = ITEM_TYPES;' +
    ' this.ITEM_STATUS = ITEM_STATUS; this.ACTION_TYPES = ACTION_TYPES;' +
    ' this.REMOVAL_REASONS = REMOVAL_REASONS; this.CACHE_KEY = CACHE_KEY;' +
    ' this.ACTIONS_ADMIN = ACTIONS_ADMIN; this.ACTIONS_PUBLIC = ACTIONS_PUBLIC;';
  new vm.Script(src, { filename: 'Code.gs' }).runInContext(sandbox);

  sandbox.__test = {
    ss, sentMail, cache, ledgerMutations,
    lockDepth: () => lockHeld,
    lockMaxDepth: () => lockMaxDepth,
    rows: (tab) => (ss.getSheetByName(tab) ? ss.getSheetByName(tab)._data : [])
  };
  return sandbox;
}

// ── Assertions ────────────────────────────────────────────────────────
let passed = 0;
const failures = [];

function ok(name, cond, detail) {
  if (cond) { passed++; return; }
  failures.push(name + (detail ? ' — ' + detail : ''));
}
function eq(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  ok(name, a === e, 'got ' + a + ', expected ' + e);
}
function throws(name, fn, re) {
  try { fn(); failures.push(name + ' — expected a throw, got none'); }
  catch (e) {
    if (re && !re.test(e.message)) failures.push(name + ' — wrong message: ' + e.message);
    else passed++;
  }
}
function section(t) { console.log('\n' + t); }

// Values built inside the vm come from a different realm, so `instanceof
// Date` is false here even for genuine Dates. Compare by tag instead.
function isDate(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

const day = 86400000;
function daysFromNow(n) {
  const d = new Date(); d.setHours(0, 0, 0, 0);
  return new Date(d.getTime() + n * day);
}
function iso(d) { return d.toISOString().slice(0, 10); }

// Bootstraps a context with sheets, seed data and one custodian.
function fresh() {
  const S = loadBackend();
  S.ensureSheets_();
  S._txn_(() => S.svcSeedReference());
  S._txn_(() => S.svcSaveRef('Custodians', {
    employee_id: 'E001', name: 'Aminah', email: 'aminah@ukm.edu.my', department: 'IMEN'
  }));
  return S;
}

// ══════════════════════════════════════════════════════════════════════
section('1. Depreciation (_bookValue_)');
{
  const S = loadBackend();
  const base = { unit_cost: 10000, salvage_value: 1000, useful_life_years: 5 };

  eq('brand new asset is worth full cost',
    S._bookValue_({ ...base, date_acquired: daysFromNow(0) }), 10000);

  // 1 year in: 10000 - ((10000-1000)/5 * 1) = 8200
  const oneYear = S._bookValue_({ ...base, date_acquired: daysFromNow(-365) });
  ok('one year of straight-line depreciation', Math.abs(oneYear - 8200) < 5, 'got ' + oneYear);

  // Half life: 10000 - (1800 * 2.5) = 5500
  const half = S._bookValue_({ ...base, date_acquired: daysFromNow(-Math.round(365.25 * 2.5)) });
  ok('mid-life value', Math.abs(half - 5500) < 10, 'got ' + half);

  eq('fully depreciated (exactly 5 years) clamps to salvage value',
    S._bookValue_({ ...base, date_acquired: daysFromNow(-Math.ceil(365.25 * 5)) }), 1000);
  ok('one day short of full life is still slightly above salvage',
    S._bookValue_({ ...base, date_acquired: daysFromNow(-(Math.ceil(365.25 * 5) - 2)) }) > 1000);

  // The clamp is the important one: without it this goes negative.
  eq('long past useful life never goes below salvage',
    S._bookValue_({ ...base, date_acquired: daysFromNow(-365 * 40) }), 1000);
  ok('never negative even with zero salvage',
    S._bookValue_({ unit_cost: 500, salvage_value: 0, useful_life_years: 2, date_acquired: daysFromNow(-365 * 30) }) === 0);

  eq('future acquisition date does not inflate value above cost',
    S._bookValue_({ ...base, date_acquired: daysFromNow(400) }), 10000);

  eq('no useful life recorded means no depreciation',
    S._bookValue_({ unit_cost: 750, date_acquired: daysFromNow(-3650) }), 750);

  eq('zero cost is zero', S._bookValue_({ unit_cost: 0, useful_life_years: 5, date_acquired: daysFromNow(-365) }), 0);
}

// ══════════════════════════════════════════════════════════════════════
section('2. SCOPES boundaries');
{
  const S = loadBackend();
  const today = S._today_();
  const item = (o) => ({ status: 'available', item_type: 'fixed_asset', ...o });

  // Warranty: inclusive 0..30, exclusive beyond.
  eq('warranty 29 days out is expiring', S.SCOPES.expiringWarranty(item({ date_warranty_expiry: daysFromNow(29) }), today), true);
  eq('warranty exactly 30 days out is expiring', S.SCOPES.expiringWarranty(item({ date_warranty_expiry: daysFromNow(30) }), today), true);
  eq('warranty 31 days out is NOT expiring', S.SCOPES.expiringWarranty(item({ date_warranty_expiry: daysFromNow(31) }), today), false);
  eq('warranty expiring today is expiring', S.SCOPES.expiringWarranty(item({ date_warranty_expiry: daysFromNow(0) }), today), true);
  eq('already-expired warranty is not "expiring"', S.SCOPES.expiringWarranty(item({ date_warranty_expiry: daysFromNow(-1) }), today), false);
  eq('already-expired warranty IS expired', S.SCOPES.expiredWarranty(item({ date_warranty_expiry: daysFromNow(-1) }), today), true);
  eq('decommissioned items are excluded from warranty alerts',
    S.SCOPES.expiringWarranty(item({ status: 'decommissioned', date_warranty_expiry: daysFromNow(5) }), today), false);

  // Inspection: due today counts.
  eq('inspection due today', S.SCOPES.inspectionDue(item({ date_next_inspection: daysFromNow(0) }), today), true);
  eq('inspection overdue yesterday', S.SCOPES.inspectionDue(item({ date_next_inspection: daysFromNow(-1) }), today), true);
  eq('inspection due tomorrow is not yet due', S.SCOPES.inspectionDue(item({ date_next_inspection: daysFromNow(1) }), today), false);

  // Audit: strictly greater than 365.
  eq('audited 364 days ago is fine', S.SCOPES.auditDue(item({ date_last_audited: daysFromNow(-364) }), today), false);
  eq('audited exactly 365 days ago is fine', S.SCOPES.auditDue(item({ date_last_audited: daysFromNow(-365) }), today), false);
  eq('audited 366 days ago is due', S.SCOPES.auditDue(item({ date_last_audited: daysFromNow(-366) }), today), true);
  eq('never audited is due', S.SCOPES.auditDue(item({ date_last_audited: null }), today), true);

  // Low stock: at threshold counts.
  const cons = (avail, min) => ({ status: 'available', item_type: 'consumable', quantity_available: avail, min_stock_alert: min });
  eq('stock above threshold is fine', S.SCOPES.lowStock(cons(6, 5)), false);
  eq('stock exactly at threshold is low', S.SCOPES.lowStock(cons(5, 5)), true);
  eq('stock below threshold is low', S.SCOPES.lowStock(cons(4, 5)), true);
  eq('fixed assets are never "low stock"', S.SCOPES.lowStock({ status: 'available', item_type: 'fixed_asset', quantity_available: 0, min_stock_alert: 5 }), false);

  // Overdue.
  eq('loan due tomorrow is not overdue', S.SCOPES.overdue({ expected_return_date: daysFromNow(1) }, today), false);
  eq('loan due today is not yet overdue', S.SCOPES.overdue({ expected_return_date: daysFromNow(0) }, today), false);
  eq('loan due yesterday is overdue', S.SCOPES.overdue({ expected_return_date: daysFromNow(-1) }, today), true);
  eq('returned loan is never overdue',
    S.SCOPES.overdue({ expected_return_date: daysFromNow(-10), actual_return_date: daysFromNow(-9) }, today), false);
}

// ══════════════════════════════════════════════════════════════════════
section('3. Asset tag generation');
{
  const S = fresh();
  const add = (over) => S._txn_(() => S.svcAddItem({
    name: 'Laptop', category_id: 1, location_id: 1, item_type: 'fixed_asset',
    date_acquired: '2026-03-01', unit_cost: 1000, ...over
  }));

  eq('first tag of the year', add().result.asset_tag, 'AST-2026-0001');
  eq('second tag increments', add().result.asset_tag, 'AST-2026-0002');
  eq('third tag increments', add().result.asset_tag, 'AST-2026-0003');
  eq('a different acquisition year restarts the sequence',
    add({ date_acquired: '2025-06-01' }).result.asset_tag, 'AST-2025-0001');
  eq('and the original year continues where it left off',
    add().result.asset_tag, 'AST-2026-0004');

  // Widening past 9999 rather than wrapping — wrapping would collide.
  const S2 = fresh();
  S2._txn_(() => {
    S2._insert_('Items', { id: 500, asset_tag: 'AST-2026-9999', name: 'x', item_type: 'fixed_asset', date_acquired: new Date('2026-01-01') });
  });
  eq('rolls over to 5 digits instead of colliding', S2._nextAssetTag_(2026), 'AST-2026-10000');

  // Soft-deleted items keep their tag reserved.
  const S3 = fresh();
  const first = S3._txn_(() => S3.svcAddItem({
    name: 'A', category_id: 1, location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-05'
  }));
  S3._txn_(() => S3.svcDeleteItem({ id: first.result.id }));
  const after = S3._txn_(() => S3.svcAddItem({
    name: 'B', category_id: 1, location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-05'
  }));
  eq('a soft-deleted item does not free its tag for reuse', after.result.asset_tag, 'AST-2026-0002');
}

// ══════════════════════════════════════════════════════════════════════
section('4. Check-out / check-in stock arithmetic');
{
  const S = fresh();
  const made = S._txn_(() => S.svcAddItem({
    name: 'Projektor', category_id: 1, location_id: 1, item_type: 'fixed_asset',
    date_acquired: '2026-01-10', unit_cost: 3000, useful_life_years: 5
  }));
  const id = made.result.id;
  const item = () => S._findById_(S._readTable_('Items'), id);

  eq('new fixed asset starts available with qty 1', [item().status, item().quantity_available], ['available', 1]);

  const out = S._txn_(() => S.svcCheckOut({ id, custodian_id: 1, expected_return_date: iso(daysFromNow(7)) }));
  ok('check-out succeeds', out.success, out.error);
  eq('check-out sets status assigned and decrements stock',
    [item().status, item().quantity_available, item().custodian_id], ['assigned', 0, 1]);

  const dbl = S._txn_(() => S.svcCheckOut({ id, custodian_id: 1, expected_return_date: iso(daysFromNow(7)) }));
  eq('a second check-out of the same asset is refused', dbl.success, false);
  ok('refusal explains why', /sudah dipinjam/.test(dbl.error), dbl.error);

  const back = S._txn_(() => S.svcCheckIn({ id }));
  ok('check-in succeeds', back.success, back.error);
  eq('check-in restores availability and CLEARS the custodian',
    [item().status, item().quantity_available, item().custodian_id], ['available', 1, null]);

  // The round-trip invariant.
  eq('stock returns exactly to its starting value', item().quantity_available, 1);

  const noLoan = S._txn_(() => S.svcCheckIn({ id }));
  eq('checking in something not on loan is refused', noLoan.success, false);

  // Backdated return dates are rejected up front.
  const past = S._txn_(() => S.svcCheckOut({ id, custodian_id: 1, expected_return_date: iso(daysFromNow(-1)) }));
  eq('a return date in the past is rejected', past.success, false);
}

// ══════════════════════════════════════════════════════════════════════
section('5. Consumable stock movements');
{
  const S = fresh();
  const made = S._txn_(() => S.svcAddItem({
    name: 'Kertas A4', category_id: 5, location_id: 2, item_type: 'consumable',
    quantity_total: 10, min_stock_alert: 3, unit_cost: 15, date_acquired: '2026-02-01'
  }));
  const id = made.result.id;
  const item = () => S._findById_(S._readTable_('Items'), id);

  eq('consumable honours its quantity', item().quantity_available, 10);

  S._txn_(() => S.svcStockChange({ id, quantity: 5, reason_notes: 'Pembelian' }, 'stock_add'));
  eq('stock_add increases availability', item().quantity_available, 15);

  S._txn_(() => S.svcStockChange({ id, quantity: 4, reason_notes: 'Guna' }, 'stock_remove'));
  eq('stock_remove decreases availability', item().quantity_available, 11);

  const over = S._txn_(() => S.svcStockChange({ id, quantity: 999, reason_notes: 'Guna' }, 'stock_remove'));
  eq('removing more than exists is refused', over.success, false);
  ok('refusal reports the real balance', /Baki semasa: 11/.test(over.error), over.error);
  eq('a refused removal leaves stock untouched', item().quantity_available, 11);

  const zero = S._txn_(() => S.svcStockChange({ id, quantity: 0, reason_notes: 'x' }, 'stock_add'));
  eq('a zero-quantity movement is rejected by validation', zero.success, false);
}

// ══════════════════════════════════════════════════════════════════════
section('6. Ledger immutability and completeness');
{
  const S = fresh();
  const made = S._txn_(() => S.svcAddItem({
    name: 'Monitor', category_id: 2, location_id: 1, item_type: 'fixed_asset',
    date_acquired: '2026-01-02', unit_cost: 800, useful_life_years: 4
  }));
  const id = made.result.id;

  S._txn_(() => S.svcCheckOut({ id, custodian_id: 1, expected_return_date: iso(daysFromNow(3)) }));
  S._txn_(() => S.svcCheckIn({ id }));
  S._txn_(() => S.svcAuditAdjust({ id, counted: 1 }));
  S._txn_(() => S.svcDecommission({ id, reason: 'Rosak', date_decommissioned: iso(daysFromNow(0)), writeoff_value: 50 }));

  const led = S._readTable_('Transactions').filter((t) => Number(t.item_id) === id);
  eq('the full lifecycle wrote five ledger rows', led.length, 5);
  eq('in the right order', led.map((t) => t.action_type),
    ['stock_add', 'check_out', 'check_in', 'audit_adjust', 'decommission']);

  eq('NO ledger row was ever edited or deleted', S.__test.ledgerMutations, []);

  throws('_update_ refuses the ledger outright',
    () => S._update_('Transactions', 2, { quantity: 999 }), /append-only/);

  const co = led.find((t) => t.action_type === 'check_out');
  const ci = led.find((t) => t.action_type === 'check_in');
  eq('check_out records a negative quantity', co.quantity, -1);
  eq('check_in records a positive quantity', ci.quantity, 1);
  ok('the original check_out row still has no actual_return_date', co.actual_return_date === null);
  ok('the check_in row carries the return date instead', isDate(ci.actual_return_date));

  const item = S._findById_(S._readTable_('Items'), id);
  eq('decommission sets the status and zeroes stock', [item.status, item.quantity_available], ['decommissioned', 0]);
  ok('decommission stamps the date', isDate(item.date_decommissioned));

  const again = S._txn_(() => S.svcDecommission({ id, reason: 'Rosak', date_decommissioned: iso(daysFromNow(0)) }));
  eq('an item cannot be decommissioned twice', again.success, false);
}

// ══════════════════════════════════════════════════════════════════════
section('7. Authorization');
{
  const S = fresh();
  const payload = {
    name: 'X', category_id: 1, location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-01'
  };

  eq('a gated action without a password is denied',
    S.handleAction('AddItem', payload).success, false);
  eq('a gated action with the wrong password is denied',
    S.handleAction('AddItem', { ...payload, __pass: 'salah' }).success, false);
  eq('a gated action with the right password proceeds',
    S.handleAction('AddItem', { ...payload, __pass: 'rahsia' }).success, true);

  eq('an unknown action name is rejected', S.handleAction('DropDatabase', { __pass: 'rahsia' }).success, false);
  eq('the public ping needs no password', S.handleAction('Ping', {}).success, true);

  eq('adminLogin accepts the right password', S.adminLogin('rahsia').ok, true);
  eq('adminLogin rejects the wrong one', S.adminLogin('salah').ok, false);

  // Fails closed when the property was never set.
  const S2 = loadBackend({ adminPassword: '' });
  S2.ensureSheets_();
  eq('with no ADMIN_PASSWORD set, nobody is admin', S2.adminLogin('').ok, false);
  eq('and gated actions stay denied', S2.handleAction('AddItem', { __pass: '' }).success, false);
  // A fresh deployment must not look like a typo.
  eq('an unset password reports itself as unconfigured', S2.adminLogin('apa-apa').unconfigured, true);
  ok('a configured password does not', S.adminLogin('salah').unconfigured === undefined);

  // Anonymous reads must not leak the custodian directory.
  const S3 = fresh();
  const anon = S3.getInitialData('salah');
  eq('an anonymous caller gets no custodian records', anon.custodians.length, 0);
  eq('and no ledger', anon.transactions.length, 0);
  eq('and is not marked admin', anon.is_admin, false);
  const admin = S3.getInitialData('rahsia');
  ok('an admin caller does get custodians', admin.custodians.length === 1);
}

// ══════════════════════════════════════════════════════════════════════
section('8. Validation and referential integrity');
{
  const S = fresh();
  const base = { name: 'X', category_id: 1, location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-01' };

  throws('a missing name is rejected', () => S.svcAddItem({ ...base, name: '' }), /Nama item wajib/);
  throws('an unknown item_type is rejected', () => S.svcAddItem({ ...base, item_type: 'kereta' }), /Jenis item tidak sah/);
  throws('a bad date is rejected', () => S.svcAddItem({ ...base, date_acquired: 'semalam' }), /bukan tarikh/);
  throws('a nonexistent category is rejected', () => S.svcAddItem({ ...base, category_id: 999 }), /Kategori tidak wujud/);
  throws('a nonexistent location is rejected', () => S.svcAddItem({ ...base, location_id: 999 }), /Lokasi tidak wujud/);
  throws('a negative cost is rejected', () => S.svcAddItem({ ...base, unit_cost: -5 }), /Kos seunit/);

  throws('a duplicate employee_id is rejected', () => S.svcSaveRef('Custodians', {
    employee_id: 'E001', name: 'Lain', email: 'lain@ukm.edu.my'
  }), /telah digunakan/);
  throws('a duplicate email is rejected', () => S.svcSaveRef('Custodians', {
    employee_id: 'E999', name: 'Lain', email: 'AMINAH@ukm.edu.my'
  }), /telah digunakan/);
  throws('a malformed email is rejected', () => S.svcSaveRef('Custodians', {
    employee_id: 'E002', name: 'Y', email: 'bukan-emel'
  }), /bukan emel/);

  // Deleting reference data still in use would orphan items.
  S._txn_(() => S.svcAddItem(base));
  throws('a category still in use cannot be deleted',
    () => S.svcDeleteRef('Categories', { id: 1 }), /masih digunakan oleh 1 item/);
  ok('an unused category can be deleted', S._txn_(() => S.svcDeleteRef('Categories', { id: 7 })).success);
}

// ══════════════════════════════════════════════════════════════════════
section('9. Migrations (ensureSheets_) and header auto-heal');
{
  const S = loadBackend();
  S.ensureSheets_();
  eq('all five tabs are created', Object.keys(S.__test.ss._sheets).sort(),
    ['Categories', 'Custodians', 'Items', 'Locations', 'Transactions']);
  eq('Items header matches the schema', S.__test.rows('Items')[0], S.SCHEMA.Items);

  // Simulate an older sheet missing a late-added column.
  const items = S.__test.ss.getSheetByName('Items');
  items._data[0] = S.SCHEMA.Items.filter((c) => c !== 'salvage_value' && c !== 'useful_life_years');
  items._data.push(['1', 'AST-2026-0001', 'Lama', '1', '1', '', 'fixed_asset', '1', '1', '0', '100', 'available',
    '2026-01-01', '', '', '', '', '', '', '', '', '', new Date(), new Date(), '']);
  const before = items._data[1].length;

  S.ensureSheets_();
  const header = items._data[0];
  ok('the missing columns are appended, not thrown on',
    header.includes('salvage_value') && header.includes('useful_life_years'));
  eq('running it again is idempotent', (S.ensureSheets_(), items._data[0].length), header.length);
  ok('the pre-existing row survives untouched', items._data[1].length === before);

  const read = S._readTable_('Items');
  eq('the healed row still reads', read.length, 1);
  eq('and the new column reads as empty rather than breaking', read[0].salvage_value, null);
  eq('so depreciation treats it as no-salvage', S._bookValue_(read[0]), 100);
}

// ══════════════════════════════════════════════════════════════════════
section('10. Locking, cache and the date engine');
{
  const S = fresh();

  eq('the lock is released after a successful transaction', S.__test.lockDepth(), 0);
  S._txn_(() => { throw new Error('boom'); });
  eq('the lock is released even when the body throws', S.__test.lockDepth(), 0);
  eq('a nested lock is never held twice at once', S.__test.lockMaxDepth(), 1);

  const failed = S._txn_(() => { throw new Error('kesalahan ujian'); });
  eq('a thrown error is reported, not swallowed', [failed.success, failed.error], [false, 'kesalahan ujian']);

  // A write must invalidate the cached payload.
  S.getInitialData('rahsia');
  ok('the payload gets cached', Object.keys(S.__test.cache).length === 1);
  S._txn_(() => S.svcSeedReference());
  eq('a write busts the cache', Object.keys(S.__test.cache).length, 0);

  // Date engine over a deliberately messy portfolio.
  const S2 = fresh();
  const mk = (over) => S2._txn_(() => S2.svcAddItem({
    name: 'Item', category_id: 1, location_id: 1, item_type: 'fixed_asset',
    date_acquired: '2026-01-01', unit_cost: 100, ...over
  })).result.id;

  mk({ date_warranty_expiry: iso(daysFromNow(10)) });                        // warranty
  mk({ date_next_inspection: iso(daysFromNow(-2)) });                        // inspection
  mk({ date_last_audited: iso(daysFromNow(-400)) });                         // audit
  const loanId = mk({});
  S2._txn_(() => S2.svcAddItem({
    name: 'Pen', category_id: 5, location_id: 1, item_type: 'consumable',
    quantity_total: 2, min_stock_alert: 5, date_acquired: '2026-01-01'
  }));                                                                        // low stock

  // An overdue loan, created by writing the ledger row directly with a
  // past due date (svcCheckOut correctly refuses to backdate one).
  S2._txn_(() => S2._ledger_({
    item_id: loanId, custodian_id: 1, action_type: 'check_out', quantity: -1,
    expected_return_date: daysFromNow(-5)
  }));

  const sum = S2.checkDates(true);
  eq('warranty alerts', sum.warranty, 1);
  eq('inspection alerts', sum.inspection, 1);
  eq('low stock alerts', sum.low_stock, 1);
  eq('overdue alerts', sum.overdue, 1);
  ok('audit alerts include the stale one and the never-audited ones', sum.audit >= 1, 'got ' + sum.audit);
  eq('skipMail sends nothing', S2.__test.sentMail.length, 0);

  // The digest and the dashboard must agree — same SCOPES, same numbers.
  const payload = S2.getInitialData('rahsia');
  eq('the dashboard warranty badge matches the digest', payload.alerts.warranty.length, sum.warranty);
  eq('the dashboard overdue badge matches the digest', payload.alerts.overdue.length, sum.overdue);
  eq('the dashboard low-stock badge matches the digest', payload.alerts.low_stock.length, sum.low_stock);
  eq('the dashboard audit badge matches the digest', payload.alerts.audit.length, sum.audit);

  ok('an overdue loan is flagged on the item itself', payload.items.find((i) => i.id === loanId).open_loan.is_overdue === true);
}

// ══════════════════════════════════════════════════════════════════════
section('11. Portfolio summary and notification');
{
  const S = fresh();
  S._txn_(() => S.svcAddItem({
    name: 'A', category_id: 1, location_id: 1, item_type: 'fixed_asset',
    date_acquired: iso(daysFromNow(-365)), unit_cost: 10000, salvage_value: 1000, useful_life_years: 5
  }));
  S._txn_(() => S.svcAddItem({
    name: 'B', category_id: 5, location_id: 1, item_type: 'consumable',
    date_acquired: iso(daysFromNow(0)), unit_cost: 20, quantity_total: 10
  }));

  const p = S.getInitialData('rahsia').portfolio;
  eq('acquisition cost counts consumables by quantity', p.acquisition_cost, 10000 + 200);
  ok('book value is below acquisition cost after a year', p.book_value < p.acquisition_cost);
  eq('depreciation is the difference', Math.round((p.acquisition_cost - p.book_value) * 100) / 100, p.depreciation);
  eq('asset count', p.asset_count, 2);

  // Decommissioned assets leave the portfolio.
  const id = S._readTable_('Items')[0].id;
  S._txn_(() => S.svcDecommission({ id, reason: 'Dijual', date_decommissioned: iso(daysFromNow(0)), writeoff_value: 4000 }));
  const p2 = S.getInitialData('rahsia').portfolio;
  eq('a decommissioned asset drops out of the portfolio', p2.asset_count, 1);

  // Check-out notification.
  const S2 = fresh();
  const nid = S2._txn_(() => S2.svcAddItem({
    name: 'Kamera', category_id: 1, location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-01'
  })).result.id;
  S2._txn_(() => S2.svcCheckOut({ id: nid, custodian_id: 1, expected_return_date: iso(daysFromNow(5)) }));
  eq('check-out emails the custodian', S2.__test.sentMail.length, 1);
  eq('addressed to them', S2.__test.sentMail[0].to, 'aminah@ukm.edu.my');
  ok('with the asset tag in the subject', /AST-2026-0001/.test(S2.__test.sentMail[0].subject));

  // HTML escaping — item names are user input.
  const S3 = fresh();
  const xid = S3._txn_(() => S3.svcAddItem({
    name: '<img src=x onerror=alert(1)>', category_id: 1, location_id: 1,
    item_type: 'fixed_asset', date_acquired: '2026-01-01'
  })).result.id;
  S3._txn_(() => S3.svcCheckOut({ id: xid, custodian_id: 1, expected_return_date: iso(daysFromNow(5)) }));
  const html = S3.__test.sentMail[0].htmlBody;
  ok('a script-ish item name is escaped in the email', !/<img src=x/.test(html) && /&lt;img/.test(html));
}

// ══════════════════════════════════════════════════════════════════════
console.log('\n' + '─'.repeat(60));
if (failures.length) {
  console.log(passed + ' passed, ' + failures.length + ' FAILED\n');
  failures.forEach((f) => console.log('  FAIL  ' + f));
  process.exit(1);
}
console.log(passed + ' assertions passed.');
