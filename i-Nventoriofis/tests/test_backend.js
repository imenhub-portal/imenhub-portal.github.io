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
    // Drive's resumable-session response is a 200 plus a Location header;
    // without the header _driveSession_ correctly reports failure, so the
    // mock has to supply it for the upload paths to be testable at all.
    UrlFetchApp: {
      fetch: () => ({
        getResponseCode: () => 200,
        getContentText: () => '',
        getAllHeaders: () => ({ Location: 'https://upload.googleapis.com/session/mock' })
      })
    },
    ScriptApp: {
      getOAuthToken: () => 'token',
      getProjectTriggers: () => [],
      newTrigger: () => ({ timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create: () => {} }) }) }) }),
      deleteTrigger: () => {}
    },
    // A faithful TextOutput: doPost's callers read getContent(), and the
    // frontend shim calls r.json() on the response, so the declared mime
    // type is part of the contract and worth being able to assert.
    ContentService: {
      createTextOutput: (t) => {
        const out = {
          __text: t, __mimeType: null,
          getContent: () => out.__text,
          setMimeType: (m) => { out.__mimeType = m; return out; }
        };
        return out;
      },
      MimeType: { JSON: 'JSON' }
    },
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
// A Date from inside the vm, rendered as a local calendar day. toISOString
// would shift across midnight in a positive timezone and make date
// comparisons fail for reasons that have nothing to do with the code.
function fmtISO(v) {
  if (!v) return '';
  const d = new Date(v.getTime ? v.getTime() : v);
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}

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
section('1. Query scopes (SCOPES)');
{
  const S = loadBackend();
  const today = S._today_();
  const item = (o) => ({ status: 'available', item_type: 'fixed_asset', ...o });

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

  // ── Issuing stock to a named member of staff ────────────────────────
  // The point of this is reporting: "how many pens did Aminah take" must
  // be answerable from structured data, not free text in the notes field.
  const issued = S._txn_(() => S.svcStockChange(
    { id, quantity: 3, custodian_id: 1, reason_notes: 'Bekalan pejabat' }, 'stock_remove'));
  ok('stock can be issued to a custodian', issued.success, issued.error);
  eq('and the recipient comes back', issued.result.custodian_id, 1);

  const led = S._readTable_('Transactions').filter((t) => Number(t.item_id) === id);
  const issueRow = led[led.length - 1];
  eq('the ledger row records the recipient', issueRow.custodian_id, 1);
  eq('as a removal', issueRow.action_type, 'stock_remove');
  eq('with a negative quantity', issueRow.quantity, -3);

  // Damage/loss has no recipient — the field must stay optional.
  const noOne = S._txn_(() => S.svcStockChange({ id, quantity: 1, reason_notes: 'Rosak' }, 'stock_remove'));
  ok('a removal with no recipient is still allowed', noOne.success, noOne.error);
  eq('and records none', S._readTable_('Transactions').slice(-1)[0].custodian_id, null);

  const ghost = S._txn_(() => S.svcStockChange(
    { id, quantity: 1, custodian_id: 999, reason_notes: 'x' }, 'stock_remove'));
  eq('an unknown recipient is rejected', ghost.success, false);
  ok('with a clear reason', /Penerima tidak wujud/.test(ghost.error), ghost.error);

  // A recipient on an ADDITION is nonsense — nobody "receives" a restock.
  const addTo = S._txn_(() => S.svcStockChange(
    { id, quantity: 1, custodian_id: 1, reason_notes: 'x' }, 'stock_add'));
  eq('a recipient on a stock addition is rejected', addTo.success, false);

  // Per-person totals must be derivable from the ledger alone.
  const perPerson = S._readTable_('Transactions')
    .filter((t) => t.action_type === 'stock_remove' && Number(t.custodian_id) === 1)
    .reduce((sum, t) => sum + Math.abs(Number(t.quantity || 0)), 0);
  eq('consumption per staff member totals correctly from the ledger', perPerson, 3);
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
  throws('a nonexistent location is rejected', () => S.svcAddItem({ ...base, location_id: 999 }), /Lokasi tidak wujud/);

  throws('a duplicate employee_id is rejected', () => S.svcSaveRef('Custodians', {
    employee_id: 'E001', name: 'Lain', email: 'lain@ukm.edu.my'
  }), /telah digunakan/);
  throws('a duplicate email is rejected', () => S.svcSaveRef('Custodians', {
    employee_id: 'E999', name: 'Lain', email: 'AMINAH@ukm.edu.my'
  }), /telah digunakan/);
  throws('a malformed email is rejected', () => S.svcSaveRef('Custodians', {
    employee_id: 'E002', name: 'Y', email: 'bukan-emel'
  }), /bukan emel/);

  // Deleting reference data still in use would orphan items. Categories are
  // gone (the item type is the category now), so Locations carry this rule.
  S._txn_(() => S.svcAddItem(base));
  throws('a location still in use cannot be deleted',
    () => S.svcDeleteRef('Locations', { id: 1 }), /masih digunakan oleh 1 item/);
  ok('an unused location can be deleted', S._txn_(() => S.svcDeleteRef('Locations', { id: 4 })).success);
}

// ══════════════════════════════════════════════════════════════════════
section('9. Migrations (ensureSheets_) and header auto-heal');
{
  const S = loadBackend();
  S.ensureSheets_();
  eq('all seven tabs are created', Object.keys(S.__test.ss._sheets).sort(),
    ['Categories', 'Config', 'Custodians', 'Items', 'Locations', 'Requests', 'Transactions']);
  eq('Items header matches the schema', S.__test.rows('Items')[0], S.SCHEMA.Items);

  // Simulate a sheet created before a column existed. Any column
  // demonstrates it; date_last_audited is used because nothing else in this
  // test depends on it.
  const items = S.__test.ss.getSheetByName('Items');
  const older = S.SCHEMA.Items.filter((c) => c !== 'date_last_audited');
  items._data[0] = older;
  const row = older.map((c) => {
    if (c === 'id') return '1';
    if (c === 'asset_tag') return 'AST-2026-0001';
    if (c === 'name') return 'Lama';
    if (c === 'category_id' || c === 'location_id') return '1';
    if (c === 'item_type') return 'fixed_asset';
    if (c === 'quantity_total' || c === 'quantity_available') return '1';
    if (c === 'status') return 'available';
    if (c === 'date_acquired') return '2026-01-01';
    if (c === 'created_at' || c === 'updated_at') return new Date();
    return '';
  });
  items._data.push(row);
  const before = items._data[1].length;

  S.ensureSheets_();
  const header = items._data[0];
  ok('the missing column is appended, not thrown on', header.includes('date_last_audited'));
  eq('running it again is idempotent', (S.ensureSheets_(), items._data[0].length), header.length);
  ok('the pre-existing row survives untouched', items._data[1].length === before);

  const read = S._readTable_('Items');
  eq('the healed row still reads', read.length, 1);
  eq('and the new column reads as empty rather than breaking', read[0].date_last_audited, null);

  // A row predating the column must remain fully usable.
  const outc = S._txn_(() => S.svcCheckOut({
    id: 1, recipient_name: 'A', recipient_email: 'a@ukm.edu.my',
    expected_return_date: iso(daysFromNow(3))
  }));
  ok('so a pre-existing asset can still be checked out', outc.success, outc.error);

  // A sheet with no Config tab at all is the real upgrade path.
  const S2 = loadBackend();
  S2.ensureSheets_();
  delete S2.__test.ss._sheets.Config;
  S2.ensureSheets_();
  ok('a missing Config tab is recreated', !!S2.__test.ss.getSheetByName('Config'));
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
  eq('inspection alerts', sum.inspection, 1);
  eq('low stock alerts', sum.low_stock, 1);
  eq('overdue alerts', sum.overdue, 1);
  ok('audit alerts include the stale one and the never-audited ones', sum.audit >= 1, 'got ' + sum.audit);
  eq('skipMail sends nothing', S2.__test.sentMail.length, 0);

  // The digest and the dashboard must agree — same SCOPES, same numbers.
  const payload = S2.getInitialData('rahsia');
  eq('the dashboard overdue badge matches the digest', payload.alerts.overdue.length, sum.overdue);
  eq('the dashboard low-stock badge matches the digest', payload.alerts.low_stock.length, sum.low_stock);
  eq('the dashboard audit badge matches the digest', payload.alerts.audit.length, sum.audit);

  ok('an overdue loan is flagged on the item itself', payload.items.find((i) => i.id === loanId).open_loan.is_overdue === true);
}

// ══════════════════════════════════════════════════════════════════════
section('11. Notifications');
{
  const S = fresh();
  const nid = S._txn_(() => S.svcAddItem({
    name: 'Kamera', category_id: 1, location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-01'
  })).result.id;
  S._txn_(() => S.svcCheckOut({ id: nid, custodian_id: 1, expected_return_date: iso(daysFromNow(5)) }));
  eq('check-out emails the custodian', S.__test.sentMail.length, 1);
  eq('addressed to them', S.__test.sentMail[0].to, 'aminah@ukm.edu.my');
  ok('with the asset tag in the subject', /AST-2026-0001/.test(S.__test.sentMail[0].subject));

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

section('11b. Asset tags belong to assets only');
{
  // A tag names one physical thing you can stick a label on. Inventory is
  // a quantity of interchangeable units, so it gets no tag at all.
  const S = fresh();
  const asset = S._txn_(() => S.svcAddItem({
    name: 'Projektor', category_id: 1, location_id: 1, item_type: 'fixed_asset',
    date_acquired: '2026-04-01'
  }));
  eq('an asset gets a tag', asset.result.asset_tag, 'AST-2026-0001');

  const inv = S._txn_(() => S.svcAddItem({
    name: 'Pen Biru', category_id: 5, location_id: 2, item_type: 'consumable',
    quantity_total: 100, min_stock_alert: 20, date_acquired: '2026-04-01'
  }));
  ok('inventory is accepted without one', inv.success, inv.error);
  eq('and its tag is blank', inv.result.asset_tag, '');

  const asset2 = S._txn_(() => S.svcAddItem({
    name: 'Kamera', category_id: 1, location_id: 1, item_type: 'fixed_asset',
    date_acquired: '2026-04-02'
  }));
  eq('inventory does not consume a number in the asset sequence',
    asset2.result.asset_tag, 'AST-2026-0002');

  const pen = S._readTable_('Items').filter((i) => i.name === 'Pen Biru')[0];
  eq('inventory remains fully usable without a tag', pen.quantity_available, 100);

  const moved = S._txn_(() => S.svcStockChange(
    { id: pen.id, quantity: 5, custodian_id: 1, reason_notes: 'Bekalan' }, 'stock_remove'));
  ok('including stock movements', moved.success, moved.error);
  eq('which still decrement correctly', moved.result.quantity_available, 95);
}

section('12. Handover: ad-hoc recipients and two-way email');
{
  const S = fresh();
  const mk = () => S._txn_(() => S.svcAddItem({
    name: 'Projektor', category_id: 1, location_id: 1, item_type: 'fixed_asset',
    date_acquired: '2026-01-10', unit_cost: 2200, serial_number: 'SN-EP-1'
  })).result.id;

  // ── Typing a name + email instead of picking an existing person ──
  const id = mk();
  const before = S._readTable_('Custodians').length;
  const out = S._txn_(() => S.svcCheckOut({
    id, recipient_name: 'Zulkifli bin Omar', recipient_email: 'zul@ukm.edu.my',
    recipient_department: 'Kejuruteraan', expected_return_date: iso(daysFromNow(7))
  }));
  ok('an asset can be handed to someone typed in on the spot', out.success, out.error);
  eq('the recipient is created as a real record', S._readTable_('Custodians').length, before + 1);
  eq('and returned to the caller', out.result.custodian_email, 'zul@ukm.edu.my');

  const created = S._readTable_('Custodians').filter((c) => c.email === 'zul@ukm.edu.my')[0];
  ok('with a generated employee id', /^ADH-[0-9]{3}$/.test(created.employee_id), created.employee_id);

  eq('a check-out email goes to them', S.__test.sentMail.length, 1);
  const coMail = S.__test.sentMail[0];
  eq('addressed correctly', coMail.to, 'zul@ukm.edu.my');
  ok('with the asset tag in the subject', /AST-2026-0001/.test(coMail.subject));
  ok('and a timestamp of the handover', /Masa Didaftar Keluar/.test(coMail.htmlBody));
  ok('and a reference number', /TRX_0000/.test(coMail.htmlBody));
  ok('and the expected return date', /Tarikh Jangka Pulang/.test(coMail.htmlBody));

  // ── The same email must not create a duplicate person ──
  const id2 = mk();
  const out2 = S._txn_(() => S.svcCheckOut({
    id: id2, recipient_name: 'Zulkifli B. Omar', recipient_email: 'ZUL@ukm.edu.my',
    expected_return_date: iso(daysFromNow(5))
  }));
  ok('a second handover to the same email succeeds', out2.success, out2.error);
  eq('and reuses the existing person rather than duplicating them',
    S._readTable_('Custodians').filter((c) => c.email === 'zul@ukm.edu.my').length, 1);
  eq('matching case-insensitively', out2.result.custodian_id, created.id);

  // ── Return closes the loop with a second email ──
  S.__test.sentMail.length = 0;
  const back = S._txn_(() => S.svcCheckIn({ id }));
  ok('check-in succeeds', back.success, back.error);
  eq('a return email is sent', S.__test.sentMail.length, 1);
  const ciMail = S.__test.sentMail[0];
  eq('to the person who held it', ciMail.to, 'zul@ukm.edu.my');
  ok('confirming the return in the subject', /dipulangkan/i.test(ciMail.subject), ciMail.subject);
  ok('with the return timestamp', /Masa Dipulangkan/.test(ciMail.htmlBody));
  ok('the duration held', /Tempoh Dipinjam/.test(ciMail.htmlBody));
  ok('and the on-time status', /tepat pada masa|Lewat/.test(ciMail.htmlBody));

  // ── Neither a person nor a name/email is an error ──
  const id3 = mk();
  const none = S._txn_(() => S.svcCheckOut({ id: id3, expected_return_date: iso(daysFromNow(3)) }));
  eq('a handover with no recipient at all is refused', none.success, false);
  ok('explaining both options', /pilih penerima|nama dan emel/i.test(none.error), none.error);

  const badMail = S._txn_(() => S.svcCheckOut({
    id: id3, recipient_name: 'X', recipient_email: 'bukan-emel', expected_return_date: iso(daysFromNow(3))
  }));
  eq('a malformed recipient email is refused', badMail.success, false);

  // Every asset is now "Aset Alih" — movable by definition — so there is no
  // non-portable case left to block. Any registered asset is loanable.
  const anyAsset = S._txn_(() => S.svcAddItem({
    name: 'Mikrofon', location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-10'
  }));
  ok('any asset can be registered without a portability flag', anyAsset.success, anyAsset.error);
  const micOut = S._txn_(() => S.svcCheckOut({
    id: anyAsset.result.id, recipient_name: 'A', recipient_email: 'a@ukm.edu.my',
    expected_return_date: iso(daysFromNow(3))
  }));
  ok('and is loanable straight away', micOut.success, micOut.error);
}

section('13. A notification failure must never lose the record');
{
  // Mail quota is a real limit in Apps Script. If sending throws, the
  // handover has already been written to the ledger and must stand.
  const S = loadBackend();
  S.MailApp.sendEmail = () => { throw new Error('quota exceeded'); };
  S.ensureSheets_();
  S._txn_(() => S.svcSeedReference());

  const id = S._txn_(() => S.svcAddItem({
    name: 'Kamera', category_id: 1, location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-01'
  })).result.id;
  const outc = S._txn_(() => S.svcCheckOut({
    id, recipient_name: 'B', recipient_email: 'b@ukm.edu.my', expected_return_date: iso(daysFromNow(4))
  }));
  ok('the handover still succeeds when email fails', outc.success, outc.error);
  eq('and reports that the notification did not go out', outc.result.mailed, false);
  eq('the ledger still recorded it',
    S._readTable_('Transactions').filter((t) => t.action_type === 'check_out').length, 1);

  const backc = S._txn_(() => S.svcCheckIn({ id }));
  ok('and the return likewise', backc.success, backc.error);
  eq('reporting the same', backc.result.mailed, false);
  eq('with the return in the ledger',
    S._readTable_('Transactions').filter((t) => t.action_type === 'check_in').length, 1);
}

section('14. Public catalog leaks nothing personal');
{
  const S = fresh();
  const projId = S._txn_(() => S.svcAddItem({
    name: 'Projektor', category_id: 1, location_id: 3, item_type: 'fixed_asset',
    date_acquired: '2026-01-10'
  })).result.id;
  S._txn_(() => S.svcAddItem({
    name: 'Pen Biru', category_id: 5, location_id: 2, item_type: 'consumable',
    quantity_total: 40, min_stock_alert: 10, date_acquired: '2026-01-10'
  }));

  const cat = S.getPublicCatalog();
  const names = cat.items.map((i) => i.name).sort();
  eq('both categories are offered', names, ['Pen Biru', 'Projektor']);

  const proj = cat.items.filter((i) => i.name === 'Projektor')[0];
  eq('an available asset says so', proj.available, true);
  ok('and reports its location', proj.location.length > 0, proj.location);

  // Hand it out, then re-read the public view.
  S._txn_(() => S.svcCheckOut({
    id: projId, custodian_id: 1, expected_return_date: iso(daysFromNow(9))
  }));
  const cat2 = S.getPublicCatalog();
  const proj2 = cat2.items.filter((i) => i.name === 'Projektor')[0];
  eq('once loaned it reads unavailable', proj2.available, false);
  ok('and gives the expected return date', isDate(proj2.expected_return_date));

  // The privacy boundary: the holder's identity must never appear.
  const blob = JSON.stringify(cat2);
  ok('the catalog never names who holds it', blob.indexOf('Aminah') === -1);
  ok('nor leaks any email address', blob.indexOf('@') === -1, blob.slice(0, 200));
  ok('nor carries a custodian id', blob.indexOf('custodian') === -1);
  ok('nor the ledger', blob.indexOf('transaction_date') === -1);
}

section('15. Request lifecycle (multi-line)');
{
  const S = fresh();
  const projId = S._txn_(() => S.svcAddItem({
    name: 'Projektor', location_id: 3, item_type: 'fixed_asset', date_acquired: '2026-01-10'
  })).result.id;
  const penId = S._txn_(() => S.svcAddItem({
    name: 'Pen Biru', location_id: 2, item_type: 'consumable',
    quantity_total: 40, min_stock_alert: 10, date_acquired: '2026-01-10'
  })).result.id;
  const paperId = S._txn_(() => S.svcAddItem({
    name: 'Kertas A4', location_id: 2, item_type: 'consumable',
    quantity_total: 20, min_stock_alert: 5, date_acquired: '2026-01-10'
  })).result.id;

  // ── One submission, many lines, no password ──
  S.__test.sentMail.length = 0;
  const sub = S.handleAction('SubmitRequest', {
    requester_name: 'Farah binti Ali', requester_email: 'farah@ukm.edu.my',
    purpose: 'Bengkel pelajar', needed_date: iso(daysFromNow(4)),
    lines: [{ item_id: penId, quantity: 4 }, { item_id: paperId, quantity: 2 }]
  });
  ok('a multi-line request needs no admin password', sub.success, sub.error);
  eq('it reports both lines', sub.result.count, 2);
  eq('and lands as pending', sub.result.status, 'pending');

  const rows = S._readTable_('Requests');
  eq('one row per line was written', rows.length, 2);
  eq('sharing a single group_id', new Set(rows.map((r) => r.group_id)).size, 1);
  eq('quantities are preserved per line',
    rows.map((r) => r.quantity).sort((a, b) => a - b), [2, 4]);

  // The whole point of grouping: four pens must not mean four emails.
  eq('exactly two emails go out, not one per line', S.__test.sentMail.length, 2);
  eq('one to the requester', S.__test.sentMail[0].to, 'farah@ukm.edu.my');
  ok('listing both items', /Pen Biru/.test(S.__test.sentMail[0].htmlBody) &&
                           /Kertas A4/.test(S.__test.sentMail[0].htmlBody));

  // ── A group is all-or-nothing ──
  const before = S._readTable_('Requests').length;
  const partial = S.handleAction('SubmitRequest', {
    requester_name: 'Zed', requester_email: 'zed@ukm.edu.my', purpose: 'Ujian',
    lines: [{ item_id: projId, quantity: 1 }, { item_id: penId, quantity: 9999 }]
  });
  eq('a group with one impossible line is refused', partial.success, false);
  ok('naming the offending item', /Pen Biru/.test(partial.error), partial.error);
  eq('and nothing at all was written', S._readTable_('Requests').length, before);

  // ── Guards ──
  eq('an empty request is refused', S.handleAction('SubmitRequest', {
    requester_name: 'A', requester_email: 'a@ukm.edu.my', purpose: 'X', lines: []
  }).success, false);
  eq('the same item twice in one request is refused', S.handleAction('SubmitRequest', {
    requester_name: 'A', requester_email: 'a@ukm.edu.my', purpose: 'X',
    lines: [{ item_id: penId, quantity: 1 }, { item_id: penId, quantity: 2 }]
  }).success, false);
  const dup = S.handleAction('SubmitRequest', {
    requester_name: 'Farah', requester_email: 'FARAH@ukm.edu.my', purpose: 'Lagi',
    lines: [{ item_id: penId, quantity: 1 }]
  });
  eq('a second pending request for the same item by the same person is refused', dup.success, false);
  ok('and says it is still waiting', /masih menunggu/i.test(dup.error), dup.error);

  // ── Deciding is admin-only ──
  const gid = rows[0].group_id;
  eq('approving a group without the password is denied',
    S.handleAction('ApproveGroup', { group_id: gid }).success, false);

  // ── Approving a group issues every line ──
  S.__test.sentMail.length = 0;
  const appr = S.handleAction('ApproveGroup', {
    __pass: 'rahsia', group_id: gid, admin_notes: 'Diluluskan'
  });
  ok('an admin can approve the whole group', appr.success, appr.error);
  eq('both lines were decided', appr.result.decided, 2);

  eq('pen stock was issued', S._findById_(S._readTable_('Items'), penId).quantity_available, 36);
  eq('paper stock was issued', S._findById_(S._readTable_('Items'), paperId).quantity_available, 18);
  const issues = S._readTable_('Transactions').filter((t) => t.action_type === 'stock_remove');
  eq('one ledger row per line', issues.length, 2);
  ok('each attributed to the requester', issues.every((t) => Number(t.custodian_id) > 0));
  eq('every line is now approved',
    S._readTable_('Requests').filter((r) => r.group_id === gid && r.status === 'approved').length, 2);

  // ── A mixed group: asset + stationery ──
  const mixed = S.handleAction('SubmitRequest', {
    requester_name: 'Rajesh', requester_email: 'rajesh@ukm.edu.my', purpose: 'Kelas',
    lines: [{ item_id: projId, quantity: 1 }, { item_id: paperId, quantity: 3 }]
  });
  ok('a mixed group can be submitted', mixed.success, mixed.error);
  const mApp = S.handleAction('ApproveGroup', {
    __pass: 'rahsia', group_id: mixed.result.group_id,
    expected_return_date: iso(daysFromNow(7))
  });
  ok('and approved with one return date for its assets', mApp.success, mApp.error);
  eq('the asset is checked out', S._findById_(S._readTable_('Items'), projId).status, 'assigned');
  eq('and the stationery deducted', S._findById_(S._readTable_('Items'), paperId).quantity_available, 15);
  eq('with a check_out ledger row',
    S._readTable_('Transactions').filter((t) => t.action_type === 'check_out').length, 1);

  // ── A loaned asset cannot be requested again ──
  const busy = S.handleAction('SubmitRequest', {
    requester_name: 'Lain', requester_email: 'lain@ukm.edu.my', purpose: 'Cuba',
    lines: [{ item_id: projId, quantity: 1 }]
  });
  eq('a loaned asset cannot be requested', busy.success, false);
  ok('and says to wait for its return', /sedang dipinjam/i.test(busy.error), busy.error);

  // ── Rejecting a group moves no stock ──
  const rej = S.handleAction('SubmitRequest', {
    requester_name: 'Siti', requester_email: 'siti@ukm.edu.my', purpose: 'Ujian',
    lines: [{ item_id: penId, quantity: 5 }]
  });
  const penBefore = S._findById_(S._readTable_('Items'), penId).quantity_available;
  S.__test.sentMail.length = 0;
  const rejected = S.handleAction('RejectGroup', {
    __pass: 'rahsia', group_id: rej.result.group_id, admin_notes: 'Stok perlu disimpan'
  });
  ok('a group can be rejected', rejected.success, rejected.error);
  eq('no stock moved', S._findById_(S._readTable_('Items'), penId).quantity_available, penBefore);
  ok('the requester is told', S.__test.sentMail.some((m) => m.to === 'siti@ukm.edu.my'));
  ok('with the reason', /Stok perlu disimpan/.test(S.__test.sentMail[0].htmlBody));

  // ── A decided group cannot be decided again ──
  eq('re-approving a decided group is refused',
    S.handleAction('ApproveGroup', { __pass: 'rahsia', group_id: gid }).success, false);

  // ── The queue stays admin-only ──
  eq('the request queue is admin-only', S.getInitialData('salah').requests.length, 0);
  ok('but an admin sees it', S.getInitialData('rahsia').requests.length >= 5);
}

section('16. Pentadbir Inventori (Config)');
{
  const S = fresh();

  // Defaults are readable before anything is saved.
  const c0 = S.getInitialData('rahsia').config;
  eq('the officer defaults to the seeded email', c0.admin_email, 'hazde@ukm.edu.my');
  eq('and a role name rather than a blank', c0.admin_name, 'Pentadbir Inventori');

  // Gated.
  eq('saving without the admin password is denied',
    S.handleAction('SaveConfig', { admin_name: 'X', admin_email: 'x@ukm.edu.my' }).success, false);

  const ok1 = S.handleAction('SaveConfig', {
    __pass: 'rahsia', admin_name: 'Hazde', admin_email: 'hazde@ukm.edu.my',
    admin_cc: 'pejabat@ukm.edu.my, ganti@ukm.edu.my'
  });
  ok('an admin can save the officer', ok1.success, ok1.error);

  const c1 = S.getInitialData('rahsia').config;
  eq('name persists', c1.admin_name, 'Hazde');
  eq('email persists', c1.admin_email, 'hazde@ukm.edu.my');
  ok('cc persists', /ganti@ukm\.edu\.my/.test(c1.admin_cc), c1.admin_cc);

  // Saving again updates rather than appending a second row per key.
  S.handleAction('SaveConfig', { __pass: 'rahsia', admin_name: 'Hazde B', admin_email: 'hazde@ukm.edu.my' });
  const keys = S._readTable_('Config').map((r) => r.key);
  eq('each key is stored once, not duplicated', keys.length, new Set(keys).size);
  eq('and the update took', S.getInitialData('rahsia').config.admin_name, 'Hazde B');

  // Validation.
  eq('a malformed officer email is rejected',
    S.handleAction('SaveConfig', { __pass: 'rahsia', admin_name: 'X', admin_email: 'bukan-emel' }).success, false);
  const badCc = S.handleAction('SaveConfig', {
    __pass: 'rahsia', admin_name: 'X', admin_email: 'x@ukm.edu.my', admin_cc: 'ok@ukm.edu.my, rosak'
  });
  eq('one bad address in the CC list is rejected', badCc.success, false);
  ok('naming the offender', /rosak/.test(badCc.error), badCc.error);
  ok('a blank CC is allowed', S.handleAction('SaveConfig', {
    __pass: 'rahsia', admin_name: 'X', admin_email: 'x@ukm.edu.my', admin_cc: ''
  }).success);

  // Privacy: the officer's address is contact detail, not public data.
  eq('an anonymous caller gets no config', Object.keys(S.getInitialData('salah').config).length, 0);
}

section('17. Notifications follow the configured officer');
{
  const S = fresh();
  const projId = S._txn_(() => S.svcAddItem({
    name: 'Projektor', location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-10'
  })).result.id;

  S.handleAction('SaveConfig', {
    __pass: 'rahsia', admin_name: 'Hazde', admin_email: 'hazde@ukm.edu.my',
    admin_cc: 'pejabat@ukm.edu.my'
  });

  S.__test.sentMail.length = 0;
  S.handleAction('SubmitRequest', {
    requester_name: 'Farah', requester_email: 'farah@ukm.edu.my',
    purpose: 'Bengkel', lines: [{ item_id: projId, quantity: 1 }]
  });
  const toAdmin = S.__test.sentMail.filter((m) => m.to === 'hazde@ukm.edu.my');
  eq('the officer is notified at the configured address', toAdmin.length, 1);
  ok('with the CC applied', /pejabat@ukm\.edu\.my/.test(toAdmin[0].cc || ''), toAdmin[0].cc);

  // Fallback: an empty Config must not send mail into the void.
  const S2 = fresh();
  S2.__test.ss._sheets.Config._data.length = 1;   // header only
  // The seeded default is itself a real address, so an empty Config still
  // resolves to somebody; ADMIN_EMAIL is the last resort behind it.
  ok('with Config empty the resolver still yields an address',
    /@/.test(S2._adminEmail_()), S2._adminEmail_());
  S2.__test.sentMail.length = 0;
  const pid = S2._txn_(() => S2.svcAddItem({
    name: 'Kamera', location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-10'
  })).result.id;
  S2.handleAction('SubmitRequest', {
    requester_name: 'A', requester_email: 'a@ukm.edu.my', purpose: 'X',
    lines: [{ item_id: pid, quantity: 1 }]
  });
  ok('and the notification still reaches somebody',
    S2.__test.sentMail.some((m) => /@/.test(m.to || '')));

  // The officer's name signs official handover mail.
  const S3 = fresh();
  S3.handleAction('SaveConfig', { __pass: 'rahsia', admin_name: 'Hazde', admin_email: 'hazde@ukm.edu.my' });
  const aid = S3._txn_(() => S3.svcAddItem({
    name: 'Mikrofon', location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-10'
  })).result.id;
  S3.__test.sentMail.length = 0;
  S3._txn_(() => S3.svcCheckOut({
    id: aid, recipient_name: 'B', recipient_email: 'b@ukm.edu.my',
    expected_return_date: iso(daysFromNow(5))
  }));
  ok('handover mail is signed by the officer, not a generic label',
    /Hazde/.test(S3.__test.sentMail[0].htmlBody) && !/Admin i-Nventori/.test(S3.__test.sentMail[0].htmlBody));
}

section('18. Return proof — the token is a capability, not a password');
{
  const S = fresh();
  const aid = S._txn_(() => S.svcAddItem({
    name: 'Projektor Epson', location_id: 3, item_type: 'fixed_asset',
    date_acquired: '2026-01-10', serial_number: 'SN-EP-1'
  })).result.id;

  // Before any loan there is no token to hold.
  throws('a garbage token is refused', () => S.getReturnContext('NOTATOKEN1234'), /tidak sah/);
  throws('an empty token is refused', () => S.getReturnContext(''), /tidak sah/);

  // ── Check-out mints one and mails the link ──
  S.__test.sentMail.length = 0;
  const out = S._txn_(() => S.svcCheckOut({
    id: aid, recipient_name: 'Farah', recipient_email: 'farah@ukm.edu.my',
    expected_return_date: iso(daysFromNow(5))
  }));
  ok('check-out succeeds', out.success, out.error);

  const item = S._findById_(S._readTable_('Items'), aid);
  const token = item.return_token;
  ok('a return token was minted', !!token && token.length === 24, String(token));
  ok('the check-out email carries the return link',
    /\?pulang=/.test(S.__test.sentMail[0].htmlBody));
  ok('and the link contains that token', S.__test.sentMail[0].htmlBody.indexOf(token) !== -1);
  ok('with a button label a borrower will understand',
    /Snap Bukti/.test(S.__test.sentMail[0].htmlBody));

  // ── What the token can read: the item, and nothing personal ──
  const ctx = S.getReturnContext(token);
  eq('the borrower sees which item it is', ctx.name, 'Projektor Epson');
  eq('and its tag', ctx.asset_tag, 'AST-2026-0001');
  ok('and the due date', isDate(ctx.expected_return_date));
  eq('no proof submitted yet', ctx.already_claimed, false);

  const blob = JSON.stringify(ctx);
  ok('the context never names the holder', blob.indexOf('Farah') === -1);
  ok('nor leaks any email address', blob.indexOf('@') === -1, blob);
  ok('nor any custodian reference', blob.indexOf('custodian') === -1);
  ok('nor the ledger', blob.indexOf('transaction') === -1);

  // ── A token belonging to nobody else's loan ──
  const bid = S._txn_(() => S.svcAddItem({
    name: 'Mikrofon', location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-01-10'
  })).result.id;
  S._txn_(() => S.svcCheckOut({
    id: bid, recipient_name: 'Rajesh', recipient_email: 'rajesh@ukm.edu.my',
    expected_return_date: iso(daysFromNow(5))
  }));
  const tokenB = S._findById_(S._readTable_('Items'), bid).return_token;
  ok('each loan gets its own token', token !== tokenB);
  eq('and a token resolves only to its own item', S.getReturnContext(tokenB).name, 'Mikrofon');

  // ── Submitting proof ──
  S.__test.sentMail.length = 0;
  const proof = S._txn_(() => S.svcSubmitReturnProof({
    token: token, photo_url: 'https://drive.google.com/file/d/abc123/view'
  }));
  ok('the borrower can submit proof with no password', proof.success, proof.error);
  const after = S._findById_(S._readTable_('Items'), aid);
  ok('the photo is recorded', /abc123/.test(after.return_photo_url), after.return_photo_url);
  ok('with a timestamp', isDate(after.return_claimed_at));
  eq('the officer is notified', S.__test.sentMail.length, 1);
  ok('at the configured address', /@/.test(S.__test.sentMail[0].to));
  ok('and told the return is NOT yet recorded',
    /belum/i.test(S.__test.sentMail[0].htmlBody));

  eq('the context now reports the claim', S.getReturnContext(token).already_claimed, true);

  // Submitting proof must not close the loan by itself — the admin confirms.
  eq('the asset is still on loan', S._findById_(S._readTable_('Items'), aid).status, 'assigned');
  eq('and no check_in row exists yet',
    S._readTable_('Transactions').filter((t) => t.action_type === 'check_in').length, 0);

  // A junk URL is refused rather than stored.
  throws('a non-https photo url is refused',
    () => S.svcSubmitReturnProof({ token: token, photo_url: 'javascript:alert(1)' }), /tidak sah/);

  // ── Check-in banks both photos and kills the token ──
  const ci = S._txn_(() => S.svcCheckIn({
    id: aid, photo_admin: 'https://drive.google.com/file/d/admin999/view'
  }));
  ok('check-in succeeds', ci.success, ci.error);

  const row = S._readTable_('Transactions').filter((t) => t.action_type === 'check_in')[0];
  ok('the borrower photo is on the ledger row', /abc123/.test(row.photo_borrower), row.photo_borrower);
  ok('and the admin photo too', /admin999/.test(row.photo_admin), row.photo_admin);
  eq('the two are kept apart, not merged', row.photo_borrower === row.photo_admin, false);

  const closed = S._findById_(S._readTable_('Items'), aid);
  eq('the token is cleared', closed.return_token, null);
  eq('the pending photo is cleared', closed.return_photo_url, null);
  eq('and the claim timestamp too', closed.return_claimed_at, null);

  throws('the old link no longer works', () => S.getReturnContext(token), /tidak sah|tamat tempoh/);
  const late = S._txn_(() => S.svcSubmitReturnProof({
    token: token, photo_url: 'https://drive.google.com/file/d/xyz/view'
  }));
  eq('and proof cannot be submitted against it', late.success, false);

  // The ledger stayed append-only through all of this.
  eq('no ledger row was ever edited or deleted', S.__test.ledgerMutations, []);

  // ── A fresh loan mints a fresh token ──
  const out2 = S._txn_(() => S.svcCheckOut({
    id: aid, recipient_name: 'Siti', recipient_email: 'siti@ukm.edu.my',
    expected_return_date: iso(daysFromNow(3))
  }));
  ok('the asset can be loaned again', out2.success, out2.error);
  const token2 = S._findById_(S._readTable_('Items'), aid).return_token;
  ok('with a different token', !!token2 && token2 !== token);

  // ── startReturnUpload is gated by the token, not by a password ──
  const up = S.startReturnUpload(token2, 'bukti.jpg', 'image/jpeg', 'https://imenhub-portal.github.io');
  eq('a live token opens an upload session', up.ok, true);
  const upBad = S.startReturnUpload('NOTATOKEN1234', 'x.jpg', 'image/jpeg', 'https://imenhub-portal.github.io');
  eq('a bogus token does not', upBad.ok, false);
}

section('19. Topping up Alat Tulis records when and from where');
{
  const S = fresh();
  const penId = S._txn_(() => S.svcAddItem({
    name: 'Pen Biru', location_id: 2, item_type: 'consumable',
    quantity_total: 10, min_stock_alert: 5, date_acquired: '2026-01-10'
  })).result.id;
  const bal = () => S._findById_(S._readTable_('Items'), penId).quantity_available;
  const lastAdd = () => S._readTable_('Transactions')
    .filter((t) => t.action_type === 'stock_add').slice(-1)[0];

  // ── A backdated delivery is filed on the day it arrived ──
  const back = fmtISO(daysFromNow(-5));
  const topup = S._txn_(() => S.svcStockChange({
    id: penId, quantity: 20, received_date: back,
    source: 'Kedai Ali', reason_notes: 'Pembelian bulanan'
  }, 'stock_add'));
  ok('a topup with details succeeds', topup.success, topup.error);
  eq('stock went up', bal(), 30);

  const row = lastAdd();
  eq('the source is stored as its own field', row.source, 'Kedai Ali');
  eq('and dated when it actually arrived, not today',
    fmtISO(row.transaction_date), back);
  ok('not today', fmtISO(row.transaction_date) !== fmtISO(daysFromNow(0)));

  // ── Both fields are optional ──
  const plain = S._txn_(() => S.svcStockChange({
    id: penId, quantity: 5, reason_notes: 'Tanpa butiran'
  }, 'stock_add'));
  ok('a topup with no date or source still works', plain.success, plain.error);
  eq('and defaults to today', fmtISO(lastAdd().transaction_date), fmtISO(daysFromNow(0)));
  eq('with a blank source', lastAdd().source, null);
  eq('stock still went up', bal(), 35);

  // ── A future date is refused ──
  const before = bal();
  const future = S._txn_(() => S.svcStockChange({
    id: penId, quantity: 100, received_date: fmtISO(daysFromNow(3)),
    source: 'Kedai Ali', reason_notes: 'Belum sampai'
  }, 'stock_add'));
  eq('a future received date is refused', future.success, false);
  ok('naming the reason', /masa hadapan|belum sampai/i.test(future.error), future.error);
  eq('and the balance is untouched', bal(), before);

  // Today itself is fine — the boundary must not be off by one.
  ok('today is accepted', S._txn_(() => S.svcStockChange({
    id: penId, quantity: 1, received_date: fmtISO(daysFromNow(0)),
    reason_notes: 'Sampai hari ini'
  }, 'stock_add')).success);

  // ── An issue cannot borrow these fields ──
  const sneakyDate = S._txn_(() => S.svcStockChange({
    id: penId, quantity: 1, received_date: fmtISO(daysFromNow(-10)), reason_notes: 'Guna'
  }, 'stock_remove'));
  eq('a withdrawal cannot be backdated', sneakyDate.success, false);
  const sneakySrc = S._txn_(() => S.svcStockChange({
    id: penId, quantity: 1, source: 'Kedai Ali', reason_notes: 'Guna'
  }, 'stock_remove'));
  eq('nor attributed to a supplier', sneakySrc.success, false);

  const issued = S._txn_(() => S.svcStockChange({
    id: penId, quantity: 2, custodian_id: 1, reason_notes: 'Bekalan pejabat'
  }, 'stock_remove'));
  ok('a plain withdrawal still works', issued.success, issued.error);
  const outRow = S._readTable_('Transactions').filter((t) => t.action_type === 'stock_remove')[0];
  eq('and carries no source', outRow.source, null);

  // ── Reporting: totals per source come out of the ledger ──
  S._txn_(() => S.svcStockChange({
    id: penId, quantity: 12, source: 'Kedai Ali', reason_notes: 'Tambahan'
  }, 'stock_add'));
  S._txn_(() => S.svcStockChange({
    id: penId, quantity: 8, source: 'Stor Pusat', reason_notes: 'Agihan'
  }, 'stock_add'));
  const fromAli = S._readTable_('Transactions')
    .filter((t) => t.action_type === 'stock_add' && t.source === 'Kedai Ali')
    .reduce((n, t) => n + Number(t.quantity || 0), 0);
  eq('how much came from one supplier is answerable', fromAli, 32);

  // ── The ledger is still append-only through all of this ──
  eq('no ledger row was edited or deleted', S.__test.ledgerMutations, []);
}

section('20. Padam — a mis-entry is removed outright, ledger and all');
{
  const S = fresh();
  const mkPen = (name) => S._txn_(() => S.svcAddItem({
    name: name, location_id: 2, item_type: 'consumable',
    quantity_total: 10, min_stock_alert: 2, date_acquired: '2026-01-10'
  })).result.id;

  // The scenario this exists for: the same thing entered twice.
  const realId = mkPen('Pen Biru');
  const dupId  = mkPen('Pen Biru');
  eq('both duplicates are in the list', S._readTable_('Items').length, 2);
  eq('each has its registration row', S._readTable_('Transactions').length, 2);

  // ── Gated ──
  eq('deleting without the admin password is denied',
    S.handleAction('DeleteItem', { id: dupId, reason: 'Berulang' }).success, false);

  // ── A reason is required ──
  const noReason = S.handleAction('DeleteItem', { __pass: 'rahsia', id: dupId });
  eq('a delete with no reason is refused', noReason.success, false);
  ok('naming the missing field', /Sebab/i.test(noReason.error), noReason.error);

  // ── The duplicate goes, and so does its fictional stock_add ──
  const del = S.handleAction('DeleteItem', {
    __pass: 'rahsia', id: dupId, reason: 'Tersalah masuk / berulang'
  });
  ok('an admin can delete it', del.success, del.error);
  eq('it reports what went', del.result.name, 'Pen Biru');
  eq('one ledger row was purged with it', del.result.ledger_removed, 1);
  eq('and it had no real movements', del.result.real_movements, 0);

  eq('the item is gone from the sheet entirely', S._readTable_('Items').length, 1);
  eq('not merely hidden', S._readTable_('Items', true).length, 1);
  eq('the surviving one is the original', S._readTable_('Items')[0].id, realId);

  // The survivor's own ledger row must be untouched — purging must not
  // take the neighbours with it.
  const left = S._readTable_('Transactions');
  eq('only the deleted item ledger row went', left.length, 1);
  eq('and the survivor keeps its own', Number(left[0].item_id), realId);

  eq('a deleted item cannot be deleted again',
    S.handleAction('DeleteItem', { __pass: 'rahsia', id: dupId, reason: 'Lagi' }).success, false);

  // ── An item someone is holding cannot be deleted ──
  const projId = S._txn_(() => S.svcAddItem({
    name: 'Projektor', location_id: 3, item_type: 'fixed_asset', date_acquired: '2026-01-10'
  })).result.id;
  S._txn_(() => S.svcCheckOut({
    id: projId, recipient_name: 'Farah', recipient_email: 'farah@ukm.edu.my',
    expected_return_date: iso(daysFromNow(5))
  }));
  const onLoan = S.handleAction('DeleteItem', {
    __pass: 'rahsia', id: projId, reason: 'Cuba padam'
  });
  eq('an item on loan is refused', onLoan.success, false);
  ok('telling the admin to check it in first', /Daftar masuk dahulu/i.test(onLoan.error), onLoan.error);

  // ── Real history is not destroyed by accident ──
  const usedId = mkPen('Pen Merah');
  S._txn_(() => S.svcStockChange(
    { id: usedId, quantity: 3, custodian_id: 1, reason_notes: 'Bekalan' }, 'stock_remove'));

  const guarded = S.handleAction('DeleteItem', {
    __pass: 'rahsia', id: usedId, reason: 'Rekod salah'
  });
  eq('an item with real movements is refused on the first attempt', guarded.success, false);
  ok('naming how many real events would be destroyed',
    /1 pergerakan sebenar/.test(guarded.error), guarded.error);
  ok('and offering Lupuskan as the alternative', /Lupuskan/.test(guarded.error), guarded.error);
  eq('nothing was removed', S._readTable_('Items').filter((i) => i.id === usedId).length, 1);

  // With the acknowledgement it proceeds — the admin is not blocked, only
  // stopped from doing it by accident.
  const forced = S.handleAction('DeleteItem', {
    __pass: 'rahsia', id: usedId, reason: 'Rekod salah', confirm_history: true
  });
  ok('an explicit confirmation lets it through', forced.success, forced.error);
  eq('and both its ledger rows go', forced.result.ledger_removed, 2);
  eq('reporting the real ones among them', forced.result.real_movements, 1);
  eq('the item is gone', S._readTable_('Items').filter((i) => i.id === usedId).length, 0);
  eq('and none of its ledger rows remain',
    S._readTable_('Transactions').filter((t) => Number(t.item_id) === usedId).length, 0);

  // ── The invariant that still holds ──
  // Rows are removed wholesale when their item is purged, but a ledger row
  // is NEVER edited in place — that would rewrite history rather than
  // retract it.
  const edits = S.__test.ledgerMutations.filter((m) => m[0] !== 'deleteRow');
  eq('no ledger row was ever EDITED', edits, []);
  ok('deletions only ever came from a purge',
    S.__test.ledgerMutations.every((m) => m[0] === 'deleteRow'));
  throws('_update_ still refuses the ledger outright',
    () => S._update_('Transactions', 2, { quantity: 999 }), /append-only/);

  // ── A purged tag is genuinely free again ──
  // Unlike the soft delete this replaced, nothing remains to reserve it.
  const S2 = fresh();
  const a1 = S2._txn_(() => S2.svcAddItem({
    name: 'Kamera', location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-03-01'
  }));
  eq('first asset tag', a1.result.asset_tag, 'AST-2026-0001');
  S2.handleAction('DeleteItem', { __pass: 'rahsia', id: a1.result.id, reason: 'Berulang' });
  const a2 = S2._txn_(() => S2.svcAddItem({
    name: 'Kamera Betul', location_id: 1, item_type: 'fixed_asset', date_acquired: '2026-03-02'
  }));
  eq('the tag is reissued, because nothing holds it any more',
    a2.result.asset_tag, 'AST-2026-0001');
  eq('and there is exactly one item', S2._readTable_('Items').length, 1);

  // ── Purged items leave the public catalog ──
  const cat = S2.getPublicCatalog();
  eq('the catalog shows only the survivor', cat.items.length, 1);
  eq('and it is the right one', cat.items[0].name, 'Kamera Betul');
}

// ══════════════════════════════════════════════════════════════════════
console.log('\n' + '─'.repeat(60));
section('21. Alat Tulis is consumed, not borrowed');
{
  const S = fresh();
  const pen = S._txn_(() => S.svcAddItem({
    name: 'Pen Biru', location_id: 2, item_type: 'consumable',
    quantity_total: 10, min_stock_alert: 2, date_acquired: '2026-01-10'
  })).result.id;

  // The UI no longer offers this, but the UI is not the guard. A loan on a
  // consumable could never be closed by a return — it would stay open for ever.
  const out = S.handleAction('CheckOut', {
    __pass: 'rahsia', id: pen, recipient_name: 'Farah',
    recipient_email: 'farah@ukm.edu.my', expected_return_date: iso(daysFromNow(3))
  });
  eq('borrowing an Alat Tulis is refused', out.success, false);
  ok('and it points to the right action instead', /Keluar Stok/.test(out.error), out.error);
  eq('no loan was opened', S._openLoans_(S._readTable_('Transactions')).length, 0);

  // The legitimate route still works and still leaves the balance right.
  const iss = S.handleAction('StockRemove', {
    __pass: 'rahsia', id: pen, quantity: 3, custodian_id: 1, reason_notes: 'Bekalan'
  });
  ok('issuing stock is unaffected', iss.success, iss.error);
  eq('and the balance drops', iss.result.quantity_available, 7);

  // Aset Alih keeps the borrow cycle untouched.
  const proj = S._txn_(() => S.svcAddItem({
    name: 'Projektor', location_id: 3, item_type: 'fixed_asset', date_acquired: '2026-01-10'
  })).result.id;
  ok('an Aset Alih can still be borrowed', S.handleAction('CheckOut', {
    __pass: 'rahsia', id: proj, recipient_name: 'Farah',
    recipient_email: 'farah@ukm.edu.my', expected_return_date: iso(daysFromNow(3))
  }).success);
}


section('22. Every notification identifies itself by name, not by address');
{
  const S = fresh();
  const proj = S._txn_(() => S.svcAddItem({
    name: 'Projektor', location_id: 3, item_type: 'fixed_asset', date_acquired: '2026-01-10'
  })).result.id;

  // Exercise the paths that actually send: hand-over, return, and a request
  // with its approval.
  S._txn_(() => S.svcCheckOut({
    id: proj, recipient_name: 'Farah', recipient_email: 'farah@ukm.edu.my',
    expected_return_date: iso(daysFromNow(3))
  }));
  S._txn_(() => S.svcCheckIn({ id: proj }));

  const pen = S._txn_(() => S.svcAddItem({
    name: 'Pen Biru', location_id: 2, item_type: 'consumable',
    quantity_total: 10, min_stock_alert: 2, date_acquired: '2026-01-10'
  })).result.id;
  const req = S.handleAction('SubmitRequest', {
    requester_name: 'Aiman', requester_email: 'aiman@ukm.edu.my',
    purpose: 'Bekalan pejabat', lines: [{ item_id: pen, quantity: 2 }]
  });
  ok('the request went in', req.success, req.error);

  const mail = S.__test.sentMail;
  ok('several notifications were sent', mail.length >= 4, 'sent ' + mail.length);

  const unnamed = mail.filter((m) => m.name !== 'i-Nventori Pejabat IMEN');
  eq('every one is sent under the system name', unnamed.map((m) => m.subject), []);

  // The point of the change: a recipient sees a name in the From line. The
  // address is still the owning account's — Apps Script cannot change that —
  // so this asserts the name is set, not that the address is hidden.
  ok('the name is a name, not an address',
    mail.every((m) => m.name && m.name.indexOf('@') === -1));

  // Subject and body agree with the From line, so nothing reads as a
  // different sender once the mail is open.
  ok('subjects carry the same name',
    mail.every((m) => m.subject.indexOf('i-Nventori Pejabat IMEN') !== -1),
    mail.map((m) => m.subject).join(' | '));
  ok('and so does the layout',
    mail.every((m) => m.htmlBody.indexOf('i-Nventori Pejabat IMEN') !== -1));

  // The guard that keeps this true: there is exactly one place that talks to
  // MailApp, so a notification added later cannot ship without the name.
  const src = fs.readFileSync(CODE_PATH, 'utf8');
  const raw = (src.match(/MailApp\.sendEmail/g) || []).length;
  eq('exactly one call site touches MailApp directly', raw, 1);
  ok('and it is inside the _sendMail_ helper',
    /function _sendMail_\(opts\) \{\s*opts\.name = MAIL_SENDER;\s*MailApp\.sendEmail\(opts\);/.test(src));
}


section('23. doPost — the entry point the whole frontend goes through');
{
  // This was never covered, and it is the one function whose absence takes
  // the entire app down while leaving doGet working, so the failure looks
  // like a network fault rather than a missing function.
  const S = fresh();
  const post = (body) => JSON.parse(
    S.doPost({ postData: { contents: JSON.stringify(body) } }).getContent());

  // Checked first and hard-stopped on, because every assertion below calls
  // it: without this the suite dies with a stack trace instead of naming
  // the one thing that is wrong.
  ok('doPost exists', typeof S.doPost === 'function');
  if (typeof S.doPost !== 'function') {
    throw new Error('doPost is missing from Code.gs — the whole API is unreachable');
  }

  const ping = post({ fn: 'handleAction', args: ['Ping', {}] });
  eq('a Ping round-trips', ping.ok, true);
  eq('and answers pong', ping.result.result, 'pong');

  // Every function the frontend can call must be routed. A name missing
  // from the switch fails only at run time, in production.
  const API_FNS = ['getInitialData', 'getPublicCatalog', 'getReturnContext',
    'startReturnUpload', 'handleAction', 'adminLogin', 'getItemHistory',
    'startResumableUpload'];
  const unrouted = API_FNS.filter((fn) =>
    /Unknown API function/.test(JSON.stringify(post({ fn: fn, args: [] }))));
  eq('every API function the page calls is routed', unrouted, []);

  // The frontend's shim reads {ok, result} / {ok, error}, so the envelope
  // matters as much as the payload.
  const bad = post({ fn: 'noSuchThing', args: [] });
  eq('an unknown function is reported, not thrown', bad.ok, false);
  ok('naming what was asked for', /noSuchThing/.test(bad.error), bad.error);

  // A thrown service error must come back as a normal envelope too,
  // otherwise the page sees a transport failure instead of the reason.
  const denied = post({ fn: 'handleAction', args: ['AddItem', { name: 'X' }] });
  eq('an unauthenticated write returns ok:true with success:false', denied.ok, true);
  eq('and the action itself reports the refusal', denied.result.success, false);

  // Malformed input must not take the endpoint down.
  const junk = JSON.parse(S.doPost({ postData: { contents: 'not json' } }).getContent());
  eq('malformed JSON is answered, not crashed', junk.ok, false);
  const empty = JSON.parse(S.doPost({}).getContent());
  eq('an empty request is answered too', empty.ok, false);

  // The response must be JSON — the shim calls r.json() on it.
  eq('the response is declared as JSON',
    S.doPost({ postData: { contents: '{"fn":"getPublicCatalog","args":[]}' } }).__mimeType,
    'JSON');
}


if (failures.length) {
  console.log(passed + ' passed, ' + failures.length + ' FAILED\n');
  failures.forEach((f) => console.log('  FAIL  ' + f));
  process.exit(1);
}
console.log(passed + ' assertions passed.');
