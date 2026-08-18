/**
 * Static checks on index.html, run before considering any frontend edit done.
 *
 *  1. Every inline <script> block parses as valid JS.
 *  2. Tags are balanced (the i-print "page won't load" class of bug).
 *  3. Every function called from an on*= handler actually exists.
 *  4. Every <use href="#i-..."> icon has a matching <g id> in the sprite.
 *  5. No CDN dependency loads at boot (they must be lazy-loaded).
 *
 * Run:  node tests/check_html.js
 */
'use strict';

const fs = require('fs');
const vm = require('vm');
const path = require('path');

const FILE = path.join(__dirname, '..', 'index.html');
const html = fs.readFileSync(FILE, 'utf8');
const problems = [];
let checks = 0;

function check(label, cond, detail) {
  checks++;
  if (!cond) problems.push(label + (detail ? ' — ' + detail : ''));
}

// ── 1. Inline script blocks parse ─────────────────────────────────────
const blocks = [];
const re = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi;
let m;
while ((m = re.exec(html)) !== null) {
  blocks.push({ code: m[1], line: html.slice(0, m.index).split('\n').length });
}
console.log('Found ' + blocks.length + ' inline <script> blocks.');

blocks.forEach((b, i) => {
  try {
    new vm.Script(b.code, { filename: 'block' + (i + 1) + '@line' + b.line });
    checks++;
  } catch (e) {
    problems.push('script block ' + (i + 1) + ' (line ~' + b.line + ') does not parse: ' + e.message);
  }
});

// ── 2. Tag balance ────────────────────────────────────────────────────
// Counts only the containers whose imbalance actually breaks layout.
['div', 'aside', 'table', 'tbody', 'thead', 'select', 'fieldset'].forEach((tag) => {
  const open = (html.match(new RegExp('<' + tag + '(?=[\\s>])', 'gi')) || []).length;
  const close = (html.match(new RegExp('</' + tag + '>', 'gi')) || []).length;
  check('<' + tag + '> balance', open === close, open + ' open vs ' + close + ' close');
});

// ── 3. Handler targets exist ──────────────────────────────────────────
// Every function referenced from an inline on*= attribute must be defined
// somewhere in the JS, or the button is silently dead at runtime.
const allJs = blocks.map((b) => b.code).join('\n');
const defined = new Set();
let d;
const defRe = /(?:^|\n)\s*function\s+([A-Za-z_$][\w$]*)/g;
while ((d = defRe.exec(allJs)) !== null) defined.add(d[1]);
const varFnRe = /(?:var|let|const)\s+([A-Za-z_$][\w$]*)\s*=\s*function/g;
while ((d = varFnRe.exec(allJs)) !== null) defined.add(d[1]);

// Handlers written directly in the HTML body...
const called = new Set();
// Bare calls only — a name preceded by "." is a method on some object and
// is not expected to be a top-level function here.
const callRe = /(^|[^.\w$])([A-Za-z_$][\w$]*)\s*\(/g;
function collectCalls(src) {
  let c;
  callRe.lastIndex = 0;
  while ((c = callRe.exec(src)) !== null) called.add(c[2]);
}
const onRe = /\son[a-z]+\s*=\s*"([^"]*)"/gi;
while ((d = onRe.exec(html)) !== null) collectCalls(d[1]);
// ...and handlers built inside generated HTML strings.
const strOnRe = /on[a-z]+="([^"]*?)(?:\\?'|")/gi;
while ((d = strOnRe.exec(allJs)) !== null) collectCalls(d[1]);

const BUILTINS = new Set([
  'if', 'for', 'while', 'switch', 'return', 'function', 'typeof', 'catch', 'confirm', 'alert',
  'parseInt', 'parseFloat', 'Number', 'String', 'Boolean', 'Array', 'Object', 'JSON', 'Math',
  'Date', 'RegExp', 'setTimeout', 'setInterval', 'encodeURIComponent', 'decodeURIComponent'
]);
const missing = [...called].filter((f) => !defined.has(f) && !BUILTINS.has(f));
check('all inline handlers resolve to a defined function', missing.length === 0, 'undefined: ' + missing.join(', '));

// ── 4. Icon sprite completeness ───────────────────────────────────────
const sprite = new Set();
let s;
const gRe = /<g id="(i-[\w-]+)"/g;
while ((s = gRe.exec(html)) !== null) sprite.add(s[1]);
const used = new Set();
const useRe = /href="#(i-[\w-]+)"/g;
while ((s = useRe.exec(html)) !== null) used.add(s[1]);
const missingIcons = [...used].filter((i) => !sprite.has(i));
check('every icon referenced exists in the sprite', missingIcons.length === 0, 'missing: ' + missingIcons.join(', '));
console.log('Icons: ' + sprite.size + ' defined, ' + used.size + ' used.');

// ── 4b. Every sprite <svg> declares a viewBox ─────────────────────────
// The sprite is authored on a 24x24 grid. An <svg> that uses it without a
// viewBox maps 1 user unit to 1 px, so any icon rendered below 24px is
// silently clipped to its top-left corner — which is exactly what shipped
// once. Cheap to assert, hard to spot by eye.
const spriteSvgs = html.match(/<svg[^>]*>\s*<use href="#i-/g) || [];
const missingVB = spriteSvgs.filter((t) => !/viewBox/.test(t));
check('every sprite <svg> has a viewBox', missingVB.length === 0,
  missingVB.length + ' of ' + spriteSvgs.length + ' missing');
// The JS helper builds icons for everything generated at runtime.
check('the icon() helper emits a viewBox', /function icon\([^)]*\)[\s\S]{0,320}?viewBox="0 0 24 24"/.test(allJs));

// ── 5. No blocking CDN dependency ─────────────────────────────────────
// Heavy libraries must be lazy-loaded, so nothing but fonts may appear as
// a <script src> or stylesheet in the document head.
const srcTags = html.match(/<script[^>]*\bsrc=["']([^"']+)["']/gi) || [];
check('no <script src> at boot (libs are lazy-loaded)', srcTags.length === 0, srcTags.join(', '));

const links = (html.match(/<link[^>]*href=["']([^"']+)["']/gi) || [])
  .filter((l) => !/fonts\.(googleapis|gstatic)\.com/.test(l));
check('no external stylesheet beyond Google Fonts', links.length === 0, links.join(', '));

// The lazy loaders themselves should still be present.
check('QR generator is lazy-loaded', /loadLib\('qr'/.test(allJs));
check('QR scanner is lazy-loaded', /loadLib\('scan'/.test(allJs));

// ── 5b. Row-menu handlers must survive being put in an attribute ──────
// mi(icon, label, fn) drops `fn` into an onclick attribute delimited by
// double quotes. A double quote inside `fn` therefore closes the attribute
// early and the browser keeps a truncated fragment — `openStock(1,` — so the
// menu entry becomes a silent no-op: no console error, nothing happens, until
// somebody clicks it. Tambah Stok and Keluar Stok shipped that way, because
// they are the only entries that pass a string argument.
//
// A static scan cannot catch this: the attribute template lives in mi() and
// the argument lives at the call site. So the call sites are evaluated here
// and the resulting attribute is checked the way a browser would read it.
{
  const broken = [];
  const call = /\bmi\(\s*'[^']*'\s*,\s*'[^']*'\s*,\s*([^;]+?)(?:,\s*true)?\s*\)\s*\)?;/g;
  let m, seen = 0;
  while ((m = call.exec(allJs)) !== null) {
    let fn;
    try { fn = new Function('id', 'return ' + m[1])(1); }
    catch (e) { continue; }            // not a literal we can evaluate
    if (typeof fn !== 'string') continue;
    seen++;
    // What the browser keeps: everything up to the first double quote.
    const kept = fn.indexOf('"') === -1 ? fn : fn.slice(0, fn.indexOf('"'));
    const opens = (kept.match(/\(/g) || []).length;
    const closes = (kept.match(/\)/g) || []).length;
    if (kept !== fn || opens !== closes) broken.push(m[1].trim() + '  →  kept: ' + kept);
  }
  check('every row-menu handler survives the attribute (' + seen + ' checked)',
    seen > 0 && broken.length === 0,
    broken.length ? broken.join('\n        ') : 'no mi() call sites found — regex is stale');
}

// ── 6. Deployment placeholders ────────────────────────────────────────
const hasPlaceholder = /PASTE_EXEC_URL_HERE/.test(html);
console.log('\nDeployment placeholder present: ' + (hasPlaceholder ? 'yes (still needs the /exec URL)' : 'no'));

// ── Report ────────────────────────────────────────────────────────────
console.log('\n' + '─'.repeat(60));
if (problems.length) {
  console.log(checks - problems.length + ' checks passed, ' + problems.length + ' FAILED\n');
  problems.forEach((p) => console.log('  FAIL  ' + p));
  process.exit(1);
}
console.log(checks + ' checks passed. ' +
  Math.round(html.length / 1024) + 'KB, ' + html.split('\n').length + ' lines.');
