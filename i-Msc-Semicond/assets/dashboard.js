/* ==========================================================
   IMEN · M.Sc. Semiconductor Engineering — analysis dashboard
   Full analysis of every survey answer + downloads (XLSX / CSV / JSON / PDF).
   English-only reporting; all language versions are combined.
   ========================================================== */
(function () {
  "use strict";

  const AUTH_KEY = "msc_dashboard_auth";
  const SEED_KEY = "msc_semicond_demo_seed_v1";
  const SCALE = ["", "Highly Unsuitable", "Unsuitable", "Moderate", "Suitable", "Highly Suitable"];
  const RC = ["", "#ef4444", "#f97316", "#eab308", "#14b8a6", "#2563eb"];
  const PALETTE = ["#2563eb", "#0ea5e9", "#d99a00", "#10b981", "#f472b6", "#f97316", "#8b5cf6", "#64748b", "#ef4444", "#14b8a6"];
  const SHORT_CAT = ["Semiconductor industry", "IC design house", "OSAT / Wafer fab", "Academia", "Government / agency", "GLC", "Private company", "Self-employed", "Student", "Others"];
  const FEE_MID = { "RM 10,000 - RM 14,999": 12500, "RM 15,000 - RM 19,999": 17500, "RM 20,000 - RM 25,000": 22500 };
  const STOP = new Set(("a an and are as at be been but by can could for from has have if in into is it its more most of on or our should so such than that the their them then there these they this to was we were will with would you your also add adding any all not no yes very well need needs more less too much many some other others course courses programme program subject subjects module modules include included including offer offered provide provided better good like want etc").split(" "));

  const state = { all: [], filtered: [], preview: false, filters: { cat: "", mode: "", interest: "", lang: "", days: 0 }, sort: "rank", page: 0, search: "", A: null };
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const f1 = (n) => (isFinite(n) ? n.toFixed(1) : "—");
  const f2 = (n) => (isFinite(n) && n > 0 ? n.toFixed(2) : "—");
  const pct = (n, d) => (d ? Math.round((n / d) * 100) : 0);
  const shortCat = (c) => {
    const i = CATEGORIES.indexOf(c);
    return i >= 0 ? SHORT_CAT[i] : c || "—";
  };

  /* ====================== auth ====================== */
  async function sha256(str) {
    const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(str));
    return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
  }
  const isDemo = () => APP_CONFIG.DEMO_MODE || !APP_CONFIG.SCRIPT_URL;

  async function unlock() {
    const pw = ($("password").value || "").trim();
    const err = $("gateError");
    const btn = $("unlockBtn");
    err.textContent = "";
    if (!pw) { err.textContent = "Enter the dashboard password."; return; }
    btn.disabled = true;
    btn.textContent = "Verifying…";
    try {
      if (isDemo()) {
        if (pw !== APP_CONFIG.DEMO_PASSWORD) throw new Error("Incorrect password.");
        enterApp();
      } else {
        const hash = await sha256(pw);
        const res = await fetch(APP_CONFIG.SCRIPT_URL + "?action=data&k=" + encodeURIComponent(hash));
        const data = await res.json();
        if (!data || !data.ok) throw new Error("Incorrect password.");
        state.all = (data.responses || []).map(norm);
        sessionStorage.setItem(AUTH_KEY, "1");
        enterApp(true);
      }
    } catch (e) {
      err.textContent = e.message || "Unable to sign in.";
      btn.disabled = false;
      btn.textContent = "Unlock dashboard";
      if (isDemo()) $("gateHint").classList.remove("hidden");
    }
  }
  function enterApp(haveData) {
    $("gate").classList.add("hidden");
    $("app").classList.remove("hidden");
    if (!haveData) {
      if (isDemo()) loadDemo();
    }
    $("badgeLive").classList.toggle("hidden", state.preview || isDemo());
    $("badgeSample").classList.toggle("hidden", !state.preview);
    $("sampleBanner").classList.toggle("hidden", !state.preview);
    buildFilterOptions();
    refresh();
  }
  function logout() {
    sessionStorage.removeItem(AUTH_KEY);
    location.href = location.pathname;
  }

  /* ====================== data ====================== */
  function igroup(s) {
    s = String(s || "");
    if (/enrolling personally|nominating|sponsoring/i.test(s)) return "Intent";
    if (/encourage/i.test(s)) return "Advocacy";
    if (/not interested/i.test(s)) return "Not interested";
    if (/unsure/i.test(s)) return "Unsure";
    return s ? "Other" : "";
  }
  function norm(r) {
    const ratings = {};
    COURSES.forEach((c) => {
      const v = Number(r[c.id]);
      if (v >= 1 && v <= 5) ratings[c.id] = v;
    });
    const ts = r.timestamp ? new Date(r.timestamp) : null;
    const fee = String(r.fee_range || "").trim();
    return {
      raw: r,
      ts: ts && !isNaN(ts) ? ts : null,
      cat: String(r.category || "").trim(),
      catOther: String(r.category_other || "").trim(),
      mode: String(r.study_mode || "").trim(),
      fee: fee,
      feeNum: fee === "Others" ? Number(r.fee_other) || 0 : FEE_MID[fee] || 0,
      interest: String(r.interest || "").trim(),
      ig: igroup(r.interest),
      lang: (String(r.language || "en").toLowerCase() === "ms" ? "Bahasa Melayu" : "English"),
      criteria: String(r.criteria || "").split("|").map((s) => s.trim()).filter(Boolean),
      feedback: String(r.feedback || "").trim(),
      ratings: ratings
    };
  }
  function loadDemo() {
    let actual = [];
    try { actual = JSON.parse(localStorage.getItem(APP_CONFIG.RESPONSE_KEY) || "[]"); } catch (e) {}
    state.all = actual.map(norm);
  }

  /* ---- deterministic sample data for the preview ---- */
  function rng(a) {
    return function () {
      a |= 0; a = (a + 0x6d2b79f5) | 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  function buildSample(n) {
    const r = rng(20261103);
    const pick = (a) => a[Math.floor(r() * a.length)];
    const wpick = (items, w) => { let t = r() * w.reduce((x, y) => x + y, 0); for (let i = 0; i < items.length; i++) { t -= w[i]; if (t <= 0) return items[i]; } return items[items.length - 1]; };
    const catW = [20, 12, 12, 14, 8, 6, 10, 3, 12, 3];
    const comments = [
      "Please add a short industrial attachment so students can work inside a fab.",
      "Consider more hands-on cleanroom hours for the microfabrication course.",
      "Strong programme. Add power semiconductor and wide bandgap devices as an elective.",
      "Weekend classes would help working engineers a lot.",
      "Include machine learning for yield analysis and test data.",
      "The fee should be supported by industry scholarships or HRD Corp claims.",
      "More focus on IC design EDA tools, Cadence or Synopsys licences for students.",
      "Good coverage of packaging. Please add chiplet and heterogeneous integration lab work.",
      "Project I and Project II should be linked to real company problems.",
      "Offer a certificate stack so working people can take single courses first.",
      "Add cybersecurity for hardware and supply chain traceability.",
      "Online delivery for theory, physical for lab sessions is the best mix.",
      "Photonics and silicon photonics are growing fast, keep this elective.",
      "Include soft skills such as technical writing and project leadership.",
      "Industry experts should co-teach selected modules.",
      "Would be useful to have a pathway to PhD for good graduates.",
      "Make the fee more affordable for local students and sponsorship friendly.",
      "Add reliability and failure analysis lab sessions with real equipment.",
      "Please include automotive semiconductor standards such as AEC-Q100.",
      "Excellent initiative, the industry urgently needs these engineers."
    ];
    const out = [];
    for (let i = 0; i < n; i++) {
      const catIdx = CATEGORIES.indexOf(wpick(CATEGORIES.filter((c) => c !== "Others").concat(["Others"]), catW));
      const cat = CATEGORIES[catIdx];
      const industry = catIdx <= 2 || catIdx === 6;
      const academic = catIdx === 3;
      const student = catIdx === 8;
      const day = Math.floor(Math.pow(r(), 0.8) * 30);
      const rec = {
        timestamp: new Date(Date.now() - day * 86400000 - Math.floor(r() * 20) * 3600000).toISOString(),
        email: "respondent" + (i + 1) + "@sample.example",
        name: "",
        language: r() > 0.58 ? "ms" : "en",
        category: cat,
        category_other: cat === "Others" ? "Consultant" : "",
        criteria: CRITERIA.filter((c, k) => r() < [0.62, 0.78, 0.66, 0.7][k] + (industry && k === 3 ? 0.1 : 0)).join(" | ") || CRITERIA[1],
        study_mode: wpick(STUDY_MODES, industry ? [14, 16, 38, 32] : student ? [46, 12, 28, 14] : [30, 14, 36, 20]),
        fee_range: wpick(FEE_RANGES, [38, 30, 14, 6]),
        fee_other: "",
        interest: wpick(INTEREST_OPTIONS, industry ? [22, 34, 22, 8, 14] : student ? [52, 6, 18, 8, 16] : academic ? [18, 20, 36, 10, 16] : [24, 18, 24, 12, 22]),
        feedback: r() < 0.38 ? pick(comments) : "",
        user_agent: "sample"
      };
      if (rec.fee_range === "Others") rec.fee_other = String(8000 + Math.floor(r() * 7) * 1000);
      COURSES.forEach((c) => {
        let base = c.type === "core" ? 4.15 : 3.85;
        const id = c.id;
        if (industry && ["core_2", "core_4", "core_5", "elective_1", "elective_4"].indexOf(id) >= 0) base += 0.35;
        if (academic && ["core_1", "core_3", "core_7", "core_8", "core_9"].indexOf(id) >= 0) base += 0.35;
        if (student && c.type === "elective") base += 0.25;
        if (id === "core_6") base -= 0.2;
        if (id === "elective_5") base -= 0.45;
        if (id === "core_7" && industry) base -= 0.3;
        let v = Math.round(base + (r() - 0.5) * 2.0);
        if (r() < 0.07) v -= 2;
        rec[id] = Math.max(1, Math.min(5, v));
      });
      out.push(rec);
    }
    return out;
  }

  /* ====================== filters ====================== */
  function applyFilters() {
    const f = state.filters;
    const since = f.days ? Date.now() - f.days * 86400000 : 0;
    state.filtered = state.all.filter((r) => {
      if (f.cat && r.cat !== f.cat) return false;
      if (f.mode && r.mode !== f.mode) return false;
      if (f.interest && r.interest !== f.interest) return false;
      if (f.lang && r.lang !== f.lang) return false;
      if (since && (!r.ts || r.ts.getTime() < since)) return false;
      return true;
    });
  }
  function opts(sel, label, values, current) {
    sel.innerHTML = '<option value="">' + esc(label) + "</option>" + values.map((v) => '<option value="' + esc(v) + '"' + (v === current ? " selected" : "") + ">" + esc(v) + "</option>").join("");
  }
  function buildFilterOptions() {
    const uniq = (k) => Array.from(new Set(state.all.map((r) => r[k]).filter(Boolean)));
    opts($("fCategory"), "All categories", CATEGORIES.filter((c) => uniq("cat").indexOf(c) >= 0), state.filters.cat);
    opts($("fMode"), "All study modes", STUDY_MODES.filter((c) => uniq("mode").indexOf(c) >= 0), state.filters.mode);
    opts($("fInterest"), "All interest levels", INTEREST_OPTIONS.filter((c) => uniq("interest").indexOf(c) >= 0), state.filters.interest);
    opts($("fLang"), "All languages", uniq("lang"), state.filters.lang);
    $("fPeriod").value = String(state.filters.days);
  }

  /* ====================== statistics ====================== */
  const mean = (a) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0);
  function sd(a) {
    if (a.length < 2) return 0;
    const m = mean(a);
    return Math.sqrt(a.reduce((s, v) => s + (v - m) * (v - m), 0) / (a.length - 1));
  }
  function median(a) {
    if (!a.length) return 0;
    const s = a.slice().sort((x, y) => x - y);
    const m = Math.floor(s.length / 2);
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }
  const byValue = (l) => l.slice().sort((a, b) => b.value - a.value);
  function tally(rows, keyFn, order) {
    const m = {};
    rows.forEach((r) => {
      const ks = [].concat(keyFn(r));
      ks.forEach((k) => { if (k) m[k] = (m[k] || 0) + 1; });
    });
    let keys = Object.keys(m);
    if (order) keys = order.filter((k) => m[k]).concat(keys.filter((k) => order.indexOf(k) < 0));
    else keys.sort((a, b) => m[b] - m[a]);
    return keys.map((k) => ({ label: k, value: m[k] }));
  }

  /* ====================== analysis engine ====================== */
  function compute() {
    const rows = state.filtered;
    const N = rows.length;
    const A = { N: N };

    // course statistics
    A.courses = COURSES.map((c) => {
      const v = rows.map((r) => r.ratings[c.id]).filter(Boolean);
      const n = v.length, m = mean(v), s = sd(v);
      const ci = n > 1 ? (1.96 * s) / Math.sqrt(n) : 0;
      const dist = [1, 2, 3, 4, 5].map((k) => v.filter((x) => x === k).length);
      const fav = pct(dist[3] + dist[4], n);
      const unf = pct(dist[0] + dist[1], n);
      let verdict = "Reconsider";
      if (m >= 4.2 && fav >= 80) verdict = "Strong";
      else if (m >= 3.8 && fav >= 65) verdict = "Good";
      else if (m >= 3.3) verdict = "Review";
      return { course: c, n: n, mean: m, sd: s, ci: ci, dist: dist, fav: fav, unf: unf, verdict: n ? verdict : "—" };
    });
    const ranked = A.courses.filter((c) => c.n).slice().sort((a, b) => b.mean - a.mean);
    ranked.forEach((c, i) => (c.rank = i + 1));
    A.ranked = ranked;
    const grp = (type) => {
      const list = A.courses.filter((c) => c.course.type === type && c.n);
      const all = [];
      rows.forEach((r) => COURSES.filter((c) => c.type === type).forEach((c) => r.ratings[c.id] && all.push(r.ratings[c.id])));
      return { type: type, mean: mean(all), sd: sd(all), n: all.length, fav: pct(all.filter((x) => x >= 4).length, all.length), best: list.slice().sort((a, b) => b.mean - a.mean)[0], worst: list.slice().sort((a, b) => a.mean - b.mean)[0] };
    };
    A.core = grp("core");
    A.elective = grp("elective");
    const allVals = [];
    rows.forEach((r) => COURSES.forEach((c) => r.ratings[c.id] && allVals.push(r.ratings[c.id])));
    A.overall = { mean: mean(allVals), fav: pct(allVals.filter((x) => x >= 4).length, allVals.length), n: allVals.length };

    // categories present
    A.categories = tally(rows, (r) => r.cat, CATEGORIES);
    A.catList = A.categories.map((c) => c.label);

    // course x category matrix
    A.segMatrix = A.courses.map((c) => ({
      course: c.course,
      cells: A.catList.map((cat) => {
        const v = rows.filter((r) => r.cat === cat).map((r) => r.ratings[c.course.id]).filter(Boolean);
        return { n: v.length, mean: mean(v) };
      })
    }));

    // demand
    A.interest = tally(rows, (r) => r.interest, INTEREST_OPTIONS);
    A.intent = pct(rows.filter((r) => r.ig === "Intent").length, N);
    A.advocacy = pct(rows.filter((r) => r.ig === "Advocacy").length, N);
    A.notInterested = pct(rows.filter((r) => r.ig === "Not interested").length, N);
    A.modes = tally(rows, (r) => r.mode, STUDY_MODES);
    A.fees = tally(rows, (r) => r.fee, FEE_RANGES);
    A.feeMedian = median(rows.map((r) => r.feeNum).filter(Boolean));
    A.interestByCat = A.catList.map((cat) => {
      const sub = rows.filter((r) => r.cat === cat);
      const g = { Intent: 0, Advocacy: 0, Unsure: 0, "Not interested": 0 };
      sub.forEach((r) => { if (g[r.ig] != null) g[r.ig]++; });
      return { cat: cat, n: sub.length, g: g };
    });
    A.modeByCat = A.catList.map((cat) => {
      const sub = rows.filter((r) => r.cat === cat);
      return { cat: cat, n: sub.length, counts: STUDY_MODES.map((m) => sub.filter((r) => r.mode === m).length) };
    });
    A.modeFee = STUDY_MODES.map((m) => ({ mode: m, counts: FEE_RANGES.map((f) => rows.filter((r) => r.mode === m && r.fee === f).length) }));

    // criteria
    A.criteria = tally(rows, (r) => r.criteria, CRITERIA);
    A.criteriaByCat = CRITERIA.map((cr) => ({
      criterion: cr,
      cells: A.catList.map((cat) => {
        const sub = rows.filter((r) => r.cat === cat);
        return { n: sub.length, v: sub.filter((r) => r.criteria.indexOf(cr) >= 0).length };
      })
    }));

    // respondents
    A.langs = tally(rows, (r) => r.lang);
    const days = {};
    rows.forEach((r) => { if (r.ts) { const k = r.ts.toISOString().slice(0, 10); days[k] = (days[k] || 0) + 1; } });
    const keys = Object.keys(days).sort();
    A.timeline = [];
    if (keys.length) {
      const start = new Date(keys[0] + "T00:00:00"), end = new Date(keys[keys.length - 1] + "T00:00:00");
      for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
        const k = d.toISOString().slice(0, 10);
        A.timeline.push({ day: k, n: days[k] || 0 });
      }
    }
    A.last7 = rows.filter((r) => r.ts && Date.now() - r.ts.getTime() < 7 * 86400000).length;

    // feedback + keywords
    A.feedback = rows.filter((r) => r.feedback).map((r) => ({ ts: r.ts, cat: r.cat, text: r.feedback }));
    const wc = {};
    A.feedback.forEach((f) => {
      const seen = new Set();
      f.text.toLowerCase().replace(/[^a-z0-9\s\-]/g, " ").split(/\s+/).forEach((w) => {
        if (w.length < 4 || STOP.has(w) || /^\d+$/.test(w) || seen.has(w)) return;
        seen.add(w);
        wc[w] = (wc[w] || 0) + 1;
      });
    });
    A.keywords = Object.keys(wc).filter((w) => wc[w] >= 2).map((w) => ({ label: w, value: wc[w] })).sort((a, b) => b.value - a.value).slice(0, 24);

    A.findings = findings(A);
    return A;
  }

  function findings(A) {
    const out = [];
    const N = A.N;
    if (!N) return out;
    const top = A.ranked[0], low = A.ranked[A.ranked.length - 1];
    const tc = A.categories[0];
    out.push("<b>" + N + " responses</b> analysed" + (tc ? ", the largest group being <b>" + esc(shortCat(tc.label)) + "</b> (" + pct(tc.value, N) + "%)" : "") + ". " + A.last7 + " arrived in the last 7 days.");
    if (A.overall.n) out.push("Across all 14 courses the average suitability is <b>" + f2(A.overall.mean) + " / 5</b> with <b>" + A.overall.fav + "%</b> favourable ratings (4–5). Core courses average <b>" + f2(A.core.mean) + "</b>; electives <b>" + f2(A.elective.mean) + "</b>.");
    if (top && low && top !== low) out.push("Highest rated: <b>" + esc(top.course.title) + "</b> (" + f2(top.mean) + "). Lowest rated: <b>" + esc(low.course.title) + "</b> (" + f2(low.mean) + (low.verdict === "Reconsider" || low.verdict === "Review" ? ", flagged “" + low.verdict + "”" : "") + ").");
    const strong = A.courses.filter((c) => c.verdict === "Strong").length, review = A.courses.filter((c) => c.verdict === "Review" || c.verdict === "Reconsider").length;
    out.push("<b>" + strong + "</b> of 14 courses are rated <b>Strong</b>" + (review ? "; <b>" + review + "</b> need a second look before the proposal is finalised." : " and none need review."));
    out.push("Demand: <b>" + A.intent + "%</b> would enrol or nominate/sponsor staff, a further <b>" + A.advocacy + "%</b> would encourage others, and <b>" + A.notInterested + "%</b> are not interested.");
    const m = byValue(A.modes), fe = byValue(A.fees), cr = byValue(A.criteria);
    if (m[0]) out.push("Preferred delivery: <b>" + esc(m[0].label) + "</b> (" + pct(m[0].value, N) + "%)" + (m[1] ? ", followed by " + esc(m[1].label) + " (" + pct(m[1].value, N) + "%)." : "."));
    if (fe[0]) out.push("Most chosen tuition band: <b>" + esc(fe[0].label) + "</b> (" + pct(fe[0].value, N) + "%)" + (A.feeMedian ? "; median expectation about <b>RM " + Math.round(A.feeMedian).toLocaleString("en-MY") + "</b>." : "."));
    if (cr[0]) out.push("Top market criterion: <b>" + esc(cr[0].label) + "</b> (" + pct(cr[0].value, N) + "% of respondents).");
    // segment gap
    let best = null;
    A.segMatrix.forEach((row) => row.cells.forEach((c, i) => {
      if (c.n >= 3) {
        const gap = c.mean - A.courses.find((x) => x.course === row.course).mean;
        if (!best || Math.abs(gap) > Math.abs(best.gap)) best = { gap: gap, course: row.course, cat: A.catList[i], mean: c.mean, n: c.n };
      }
    }));
    if (best && Math.abs(best.gap) >= 0.4) out.push("Biggest segment difference: <b>" + esc(shortCat(best.cat)) + "</b> rate <b>" + esc(best.course.title) + "</b> " + f2(best.mean) + " (" + (best.gap > 0 ? "+" : "") + f1(best.gap) + " versus all respondents, n=" + best.n + ").");
    if (N < 30) out.push("<span class='text-amber-700'>Note: fewer than 30 responses — treat means and intervals as indicative.</span>");
    return out;
  }

  /* ====================== rendering helpers ====================== */
  function bars(items, total, opt) {
    opt = opt || {};
    if (!items.length) return empty();
    const max = opt.max || Math.max.apply(null, items.map((i) => i.value)) || 1;
    return '<div class="space-y-3.5">' + items.map((it, i) => {
      const w = Math.max(2, (it.value / max) * 100);
      const color = opt.colors ? opt.colors[i % opt.colors.length] : PALETTE[i % PALETTE.length];
      return '<div><div class="flex items-end justify-between gap-3"><span class="text-[15px] font-semibold leading-snug text-slate-700">' + esc(it.label) + '</span><span class="shrink-0 text-[15px] font-extrabold text-slate-900">' + it.value + ' <span class="font-semibold text-slate-400">· ' + pct(it.value, total) + '%</span></span></div>' +
        '<div class="mt-1.5 h-3 w-full overflow-hidden rounded-full bg-slate-200"><div class="h-full rounded-full" style="width:' + w + '%;background:' + color + '"></div></div></div>';
    }).join("") + "</div>";
  }
  function empty() { return '<p class="rounded-xl border border-dashed border-slate-300 bg-slate-50 px-4 py-6 text-center text-[15px] font-semibold text-slate-400">No data for this filter yet.</p>'; }
  function stackBar(counts, colors, labels) {
    const t = counts.reduce((a, b) => a + b, 0);
    if (!t) return '<div class="h-4 rounded-full bg-slate-200"></div>';
    const seg = counts.map((c, i) => {
      if (!c) return "";
      const tip = (labels ? labels[i] : i + 1) + ": " + c + " (" + pct(c, t) + "%)";
      return '<div style="width:' + (c / t) * 100 + "%;background:" + colors[i] + '" title="' + esc(tip) + '"></div>';
    });
    return '<div class="flex h-4 w-full overflow-hidden rounded-full bg-slate-200">' + seg.join("") + "</div>";
  }
  function heatColor(v, lo, hi) {
    const t = Math.max(0, Math.min(1, (v - lo) / (hi - lo)));
    // red -> orange -> yellow -> teal -> blue
    const stops = [[239, 68, 68], [249, 115, 22], [234, 179, 8], [20, 184, 166], [37, 99, 235]];
    const p = t * (stops.length - 1), i = Math.min(stops.length - 2, Math.floor(p)), k = p - i;
    const c = stops[i].map((x, j) => Math.round(x + (stops[i + 1][j] - x) * k));
    return "rgb(" + c.join(",") + ")";
  }
  function donut(items, centerLabel) {
    if (!items.length) return empty();
    const total = items.reduce((a, b) => a + b.value, 0) || 1;
    let acc = 0;
    const stops = items.map((it, i) => { const s = (acc / total) * 100; acc += it.value; return PALETTE[i % PALETTE.length] + " " + s + "% " + (acc / total) * 100 + "%"; });
    return '<div class="flex flex-col items-center gap-4 sm:flex-row"><div class="relative h-36 w-36 shrink-0 rounded-full" style="background:conic-gradient(' + stops.join(",") + ')"><div class="absolute inset-[20%] flex flex-col items-center justify-center rounded-full bg-white"><span class="text-2xl font-extrabold">' + total + '</span><span class="text-[12px] font-bold uppercase tracking-wider text-slate-400">' + esc(centerLabel || "responses") + "</span></div></div>" +
      '<div class="w-full space-y-2">' + items.map((it, i) => '<div class="flex items-center gap-2.5"><span class="h-3 w-3 shrink-0 rounded-full" style="background:' + PALETTE[i % PALETTE.length] + '"></span><span class="flex-1 text-[15px] text-slate-600">' + esc(it.label) + '</span><span class="text-[15px] font-extrabold">' + it.value + ' · ' + pct(it.value, total) + "%</span></div>").join("") + "</div></div>";
  }
  function table(head, rowsHtml, cls) {
    return '<table class="tbl w-full ' + (cls || "") + '"><thead><tr>' + head.map((h) => "<th>" + esc(h) + "</th>").join("") + "</tr></thead><tbody>" + rowsHtml + "</tbody></table>";
  }

  /* ====================== render ====================== */
  function render() {
    const A = state.A;
    $("filterCount").textContent = A.N;
    $("totalCount").textContent = state.all.length;
    $("lastUpdated").textContent = new Date().toLocaleString("en-MY", { dateStyle: "medium", timeStyle: "short" });
    $("printMeta").textContent = "Generated " + new Date().toLocaleString("en-MY") + " · " + A.N + " responses" + (state.preview ? " · SAMPLE DATA" : "");
    renderActiveFilters();
    renderKpis(A);
    $("findings").innerHTML = A.N ? A.findings.map((f, i) => '<li class="flex gap-3"><span class="mt-0.5 flex h-7 w-7 shrink-0 items-center justify-center rounded-lg bg-royal/10 text-[14px] font-extrabold text-royal">' + (i + 1) + '</span><span class="text-[17px] leading-relaxed text-slate-700">' + f + "</span></li>").join("") : "<li>" + empty() + "</li>";
    renderCourses(A);
    renderSegments(A);
    renderDemand(A);
    renderCriteria(A);
    renderProfile(A);
    renderFeedback(A);
    state.page = 0;
    renderData();
  }

  function renderActiveFilters() {
    const f = state.filters, chips = [];
    const add = (label, key) => chips.push('<button data-clear="' + key + '" class="chip chip-gold !text-[14px]">' + esc(label) + ' ✕</button>');
    if (f.cat) add(shortCat(f.cat), "cat");
    if (f.mode) add(f.mode, "mode");
    if (f.interest) add(f.interest, "interest");
    if (f.lang) add(f.lang, "lang");
    if (f.days) add("Last " + f.days + " days", "days");
    $("activeFilters").innerHTML = chips.join("");
    $("activeFilters").querySelectorAll("[data-clear]").forEach((b) => b.addEventListener("click", () => { state.filters[b.dataset.clear] = b.dataset.clear === "days" ? 0 : ""; buildFilterOptions(); refresh(); }));
  }

  function kpi(label, value, sub, color, icon) {
    return '<div class="card print-block p-4 sm:p-5"><div class="flex items-center justify-between"><p class="text-[13px] font-extrabold uppercase tracking-[0.1em] text-slate-500">' + label + '</p><span class="flex h-8 w-8 items-center justify-center rounded-lg" style="background:' + color + '22;color:' + color + '">' + icon + '</span></div>' +
      '<p class="mt-2 text-[32px] font-extrabold leading-none tracking-tight" style="color:' + color + '">' + value + '</p><p class="mt-1.5 text-[14px] font-semibold leading-snug text-slate-500">' + sub + "</p></div>";
  }
  const IC = (p) => '<svg class="h-5 w-5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">' + p + "</svg>";
  function renderKpis(A) {
    const N = A.N;
    const modeTop = byValue(A.modes)[0], feeTop = byValue(A.fees)[0];
    $("kpis").innerHTML = [
      kpi("Responses", N, A.last7 + " in the last 7 days", "#2563eb", IC('<circle cx="9" cy="8" r="3.4"/><path d="M2.8 20c.6-3.6 3.1-5.4 6.2-5.4M16 6a3 3 0 010 5.5M18 15.2c1.7.5 2.7 1.9 3 4.3"/>')),
      kpi("Overall suitability", A.overall.n ? f2(A.overall.mean) + " / 5" : "—", A.overall.fav + "% favourable (4–5)", "#14b8a6", IC('<path d="M12 3.5l2.6 5.4 5.9.8-4.3 4.1 1 5.9L12 16.9l-5.2 2.8 1-5.9-4.3-4.1 5.9-.8z"/>')),
      kpi("Core vs elective", f2(A.core.mean) + " · " + f2(A.elective.mean), "Core average · elective average", "#d99a00", IC('<rect x="3" y="3" width="8" height="18" rx="1.5"/><rect x="13" y="9" width="8" height="12" rx="1.5"/>')),
      kpi("Demand (intent)", A.intent + "%", "would enrol or sponsor · +" + A.advocacy + "% advocate", "#10b981", IC('<path d="M5 13l4 4L19 7"/>')),
      kpi("Top-rated course", A.ranked[0] ? f2(A.ranked[0].mean) : "—", A.ranked[0] ? esc(A.ranked[0].course.title) : "no ratings yet", "#2563eb", IC('<path d="M8 21h8M12 17v4M7 4h10v5a5 5 0 01-10 0zM17 5h3v2a3 3 0 01-3 3M7 5H4v2a3 3 0 003 3"/>')),
      kpi("Lowest-rated course", A.ranked.length ? f2(A.ranked[A.ranked.length - 1].mean) : "—", A.ranked.length ? esc(A.ranked[A.ranked.length - 1].course.title) : "no ratings yet", "#ef4444", IC('<path d="M12 9v4M12 17h.01M10.3 3.9L2.4 18a2 2 0 001.7 3h15.8a2 2 0 001.7-3L13.7 3.9a2 2 0 00-3.4 0z"/>')),
      kpi("Preferred mode", modeTop ? pct(modeTop.value, N) + "%" : "—", modeTop ? esc(modeTop.label) : "—", "#8b5cf6", IC('<rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 20h8M12 16v4"/>')),
      kpi("Median fee expectation", A.feeMedian ? "RM " + (Math.round(A.feeMedian / 100) / 10) + "k" : "—", feeTop ? "Most chosen: " + esc(feeTop.label) : "—", "#f97316", IC('<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5v9M14.6 9.6c-.6-.8-1.5-1.2-2.6-1.2-1.4 0-2.4.7-2.4 1.8 0 2.6 5.2 1.2 5.2 3.8 0 1.1-1.1 1.9-2.8 1.9-1.2 0-2.2-.5-2.8-1.3"/>'))
    ].join("");
  }

  function verdictChip(v) {
    const c = { Strong: "bg-emerald-100 text-emerald-800", Good: "bg-sky-100 text-sky-800", Review: "bg-amber-100 text-amber-800", Reconsider: "bg-rose-100 text-rose-800" }[v] || "bg-slate-100 text-slate-500";
    return '<span class="inline-flex rounded-full px-2.5 py-1 text-[13px] font-extrabold ' + c + '">' + v + "</span>";
  }
  function renderCourses(A) {
    const sorts = [["rank", "Rank"], ["course", "Course order"], ["fav", "Favourable %"], ["n", "Responses"]];
    $("courseSort").innerHTML = '<span class="text-[14px] font-bold text-slate-500">Sort by</span>' + sorts.map((s) => '<button data-sort="' + s[0] + '" class="chip ' + (state.sort === s[0] ? "chip-blue" : "") + ' !text-[14px]">' + s[1] + "</button>").join("");
    $("courseSort").querySelectorAll("[data-sort]").forEach((b) => b.addEventListener("click", () => { state.sort = b.dataset.sort; renderCourses(state.A); }));
    let list = A.courses.slice();
    if (state.sort === "rank") list.sort((a, b) => (b.n ? b.mean : -1) - (a.n ? a.mean : -1));
    else if (state.sort === "fav") list.sort((a, b) => b.fav - a.fav);
    else if (state.sort === "n") list.sort((a, b) => b.n - a.n);
    $("courseTable").innerHTML = list.map((c, i) => {
      const lo = Math.max(1, c.mean - c.ci), hi = Math.min(5, c.mean + c.ci);
      return "<tr><td class='font-extrabold text-slate-400'>" + (c.rank || "—") + "</td>" +
        "<td><div class='font-extrabold text-slate-900'>" + esc(c.course.title) + "</div></td>" +
        "<td><span class='rounded-full px-2.5 py-1 text-[13px] font-extrabold " + (c.course.type === "core" ? "bg-amber-100 text-amber-800" : "bg-cyan-100 text-cyan-800") + "'>" + (c.course.type === "core" ? "Core" : "Elective") + "</span></td>" +
        "<td>" + c.n + "</td>" +
        "<td style='min-width:200px'><div class='flex items-center gap-2'><span class='w-12 text-[17px] font-extrabold' style='color:" + (c.n ? heatColor(c.mean, 2.5, 5) : "#94a3b8") + "'>" + f2(c.mean) + "</span>" +
        "<div class='relative h-3 flex-1 rounded-full bg-slate-200'><div class='absolute inset-y-0 rounded-full' style='left:0;width:" + (c.mean / 5) * 100 + "%;background:" + (c.n ? heatColor(c.mean, 2.5, 5) : "#cbd5e1") + "'></div>" +
        (c.n > 1 ? "<div class='absolute top-1/2 h-0.5 -translate-y-1/2 bg-slate-800/70' style='left:" + (lo / 5) * 100 + "%;width:" + ((hi - lo) / 5) * 100 + "%'></div>" : "") + "</div></div>" +
        "<div class='mt-0.5 text-[13px] font-semibold text-slate-400'>± " + f2(c.ci) + " (" + f2(lo) + "–" + f2(hi) + ")</div></td>" +
        "<td>" + f2(c.sd) + "</td><td class='font-extrabold'>" + c.fav + "%</td>" +
        "<td>" + stackBar(c.dist, RC.slice(1), SCALE.slice(1)) + "<div class='mt-1 flex justify-between text-[13px] font-semibold text-slate-400'>" + c.dist.map((d) => "<span>" + d + "</span>").join("") + "</div></td>" +
        "<td>" + verdictChip(c.verdict) + "</td></tr>";
    }).join("") || "<tr><td colspan='9'>" + empty() + "</td></tr>";
    $("courseLegend").innerHTML = '<span class="text-slate-500">Distribution of ratings:</span>' + [1, 2, 3, 4, 5].map((k) => '<span class="inline-flex items-center gap-1.5"><span class="h-3 w-3 rounded" style="background:' + RC[k] + '"></span>' + k + " " + SCALE[k] + "</span>").join("");
    const tc = (g, color, label) => '<div class="rounded-2xl border border-slate-300 bg-white p-4"><p class="text-[13px] font-extrabold uppercase tracking-[0.1em]" style="color:' + color + '">' + label + '</p><p class="mt-1 text-[28px] font-extrabold">' + f2(g.mean) + ' <span class="text-[15px] font-semibold text-slate-400">/ 5 · ' + g.fav + '% favourable</span></p>' +
      (g.best ? '<p class="mt-2 text-[15px] text-slate-600">Best: <b>' + esc(g.best.course.title) + "</b> (" + f2(g.best.mean) + ")</p><p class='text-[15px] text-slate-600'>Weakest: <b>" + esc(g.worst.course.title) + "</b> (" + f2(g.worst.mean) + ")</p>" : "<p class='mt-2 text-[15px] text-slate-400'>No ratings yet</p>") + "</div>";
    $("typeCompare").innerHTML = tc(A.core, "#d99a00", "Core courses (9)") + tc(A.elective, "#0ea5e9", "Electives (5)");
  }

  function renderSegments(A) {
    if (!A.catList.length) { $("segHeat").innerHTML = empty(); return; }
    let h = '<table class="tbl w-full min-w-[820px]"><thead><tr><th style="position:sticky;left:0;z-index:2">Course</th><th>All</th>' + A.catList.map((c, i) => "<th title='" + esc(c) + "'>" + esc(shortCat(c)) + "<br><span class='font-semibold normal-case text-slate-400'>n=" + A.categories[i].value + "</span></th>").join("") + "</tr></thead><tbody>";
    A.segMatrix.forEach((row) => {
      const all = A.courses.find((c) => c.course === row.course);
      h += "<tr><td style='position:sticky;left:0;background:#fff' class='font-bold'>" + esc(row.course.title) + "</td><td><span class='inline-flex w-14 justify-center rounded-md py-1 text-[15px] font-extrabold text-white' style='background:" + (all.n ? heatColor(all.mean, 2.5, 5) : "#94a3b8") + "'>" + f2(all.mean) + "</span></td>" +
        row.cells.map((c) => c.n ? "<td><span class='inline-flex w-14 justify-center rounded-md py-1 text-[15px] font-extrabold " + (c.n < 3 ? "opacity-40 " : "") + "text-white' style='background:" + heatColor(c.mean, 2.5, 5) + "' title='n=" + c.n + "'>" + f2(c.mean) + "</span></td>" : "<td class='text-slate-300'>—</td>").join("") + "</tr>";
    });
    $("segHeat").innerHTML = h + "</tbody></table>";
  }

  function renderDemand(A) {
    const N = A.N;
    $("interestChart").innerHTML = bars(A.interest, N, { colors: ["#10b981", "#14b8a6", "#0ea5e9", "#ef4444", "#94a3b8"] });
    $("modeChart").innerHTML = bars(A.modes, N, { colors: ["#2563eb", "#0ea5e9", "#8b5cf6", "#d99a00"] });
    $("feeChart").innerHTML = bars(A.fees, N, { colors: ["#14b8a6", "#2563eb", "#8b5cf6", "#94a3b8"] }) + (A.feeMedian ? '<p class="mt-4 rounded-xl bg-blue-50 px-4 py-3 text-[15px] font-semibold text-royal">Median fee expectation ≈ RM ' + Math.round(A.feeMedian).toLocaleString("en-MY") + "</p>" : "");
    const gcol = { Intent: "#10b981", Advocacy: "#0ea5e9", Unsure: "#94a3b8", "Not interested": "#ef4444" };
    $("interestByCat").innerHTML = A.interestByCat.length ? '<div class="space-y-3">' + A.interestByCat.map((r) => '<div><div class="flex justify-between text-[15px]"><span class="font-semibold text-slate-700">' + esc(shortCat(r.cat)) + '</span><span class="font-extrabold">' + pct(r.g.Intent + r.g.Advocacy, r.n) + '% <span class="font-semibold text-slate-400">of ' + r.n + "</span></span></div>" +
      '<div class="mt-1">' + stackBar(Object.keys(gcol).map((k) => r.g[k]), Object.values(gcol), Object.keys(gcol)) + "</div></div>").join("") + '<div class="flex flex-wrap gap-3 pt-1 text-[14px] font-semibold text-slate-600">' + Object.keys(gcol).map((k) => '<span class="inline-flex items-center gap-1.5"><span class="h-3 w-3 rounded" style="background:' + gcol[k] + '"></span>' + k + "</span>").join("") + "</div></div>" : empty();
    // mode x fee
    $("modeFee").innerHTML = A.N ? table(["Study mode"].concat(FEE_RANGES, ["Total"]), A.modeFee.map((r) => "<tr><td class='font-bold'>" + esc(r.mode) + "</td>" + r.counts.map((c) => "<td>" + c + "</td>").join("") + "<td class='font-extrabold'>" + r.counts.reduce((a, b) => a + b, 0) + "</td></tr>").join(""), "min-w-[640px]") : empty();
  }

  function renderCriteria(A) {
    $("criteriaChart").innerHTML = bars(A.criteria, A.N, { colors: ["#2563eb", "#0ea5e9", "#d99a00", "#10b981"] });
    if (!A.catList.length) { $("criteriaByCat").innerHTML = ""; return; }
    $("criteriaByCat").innerHTML = '<p class="mb-2 text-[14px] font-extrabold uppercase tracking-[0.1em] text-slate-500">Share selecting each criterion, by category</p>' +
      table(["Criterion"].concat(A.catList.map(shortCat)), A.criteriaByCat.map((r) => "<tr><td class='font-bold' style='min-width:260px'>" + esc(r.criterion) + "</td>" + r.cells.map((c) => "<td>" + (c.n ? pct(c.v, c.n) + "%" : "—") + "</td>").join("") + "</tr>").join(""), "min-w-[820px]");
  }

  function renderProfile(A) {
    $("catChart").innerHTML = bars(A.categories.map((c) => ({ label: shortCat(c.label), value: c.value })), A.N);
    $("langChart").innerHTML = donut(A.langs);
    if (!A.timeline.length) { $("timeline").innerHTML = empty(); return; }
    const max = Math.max.apply(null, A.timeline.map((d) => d.n)) || 1, W = 320, H = 110, bw = W / A.timeline.length;
    $("timeline").innerHTML = '<svg viewBox="0 0 ' + W + " " + (H + 18) + '" class="w-full">' + A.timeline.map((d, i) => '<rect x="' + (i * bw + 1) + '" y="' + (H - (d.n / max) * H) + '" width="' + Math.max(2, bw - 2) + '" height="' + (d.n / max) * H + '" rx="2" fill="#2563eb"><title>' + d.day + ": " + d.n + "</title></rect>").join("") +
      '<text x="0" y="' + (H + 14) + '" font-size="10" fill="#64748b">' + A.timeline[0].day + '</text><text x="' + W + '" y="' + (H + 14) + '" text-anchor="end" font-size="10" fill="#64748b">' + A.timeline[A.timeline.length - 1].day + "</text></svg>" +
      '<p class="mt-1 text-[14px] font-semibold text-slate-500">Busiest day: ' + max + " responses</p>";
  }

  function renderFeedback(A) {
    $("keywords").innerHTML = A.keywords.length ? A.keywords.map((k, i) => '<span class="inline-flex items-center gap-1.5 rounded-full border border-slate-300 bg-white px-3 py-1.5 font-bold text-slate-700" style="font-size:' + (14 + Math.min(6, k.value) ) + 'px">' + esc(k.label) + ' <span class="rounded-full bg-royal/10 px-1.5 text-[12px] font-extrabold text-royal">' + k.value + "</span></span>").join("") : '<span class="text-[15px] text-slate-400">Not enough comments for keyword themes yet.</span>';
    $("feedbackList").innerHTML = A.feedback.length ? A.feedback.slice().sort((a, b) => (b.ts || 0) - (a.ts || 0)).slice(0, 24).map((f) => '<blockquote class="rounded-xl border border-slate-300 bg-white p-4"><p class="text-[16px] leading-relaxed text-slate-800">“' + esc(f.text) + '”</p><p class="mt-2 text-[13px] font-bold text-slate-400">' + esc(shortCat(f.cat)) + (f.ts ? " · " + f.ts.toLocaleDateString("en-MY") : "") + "</p></blockquote>").join("") : "<p class='md:col-span-2'>" + empty() + "</p>";
  }

  function renderData() {
    const q = state.search.toLowerCase();
    let rows = state.filtered.slice().sort((a, b) => (b.ts || 0) - (a.ts || 0));
    if (q) rows = rows.filter((r) => (r.cat + " " + r.mode + " " + r.fee + " " + r.interest + " " + r.feedback + " " + r.lang).toLowerCase().indexOf(q) >= 0);
    const per = 12, pages = Math.max(1, Math.ceil(rows.length / per));
    state.page = Math.min(state.page, pages - 1);
    const slice = rows.slice(state.page * per, state.page * per + per);
    const avg = (r, type) => f2(mean(COURSES.filter((c) => c.type === type).map((c) => r.ratings[c.id]).filter(Boolean)));
    $("dataTable").innerHTML = slice.map((r) => "<tr><td class='whitespace-nowrap'>" + (r.ts ? r.ts.toLocaleDateString("en-MY") : "—") + "</td><td>" + esc(shortCat(r.cat)) + "</td><td>" + esc(r.mode) + "</td><td>" + esc(r.fee === "Others" ? "RM " + (r.feeNum || "?") : r.fee) + "</td><td>" + esc(r.interest) + "</td><td class='font-extrabold'>" + avg(r, "core") + "</td><td class='font-extrabold'>" + avg(r, "elective") + "</td><td>" + (r.lang === "English" ? "EN" : "BM") + "</td></tr>").join("") || "<tr><td colspan='8'>" + empty() + "</td></tr>";
    $("pageInfo").textContent = rows.length ? "Showing " + (state.page * per + 1) + "–" + Math.min(rows.length, state.page * per + per) + " of " + rows.length : "";
    $("prevPage").disabled = state.page === 0;
    $("nextPage").disabled = state.page >= pages - 1;
  }

  /* ====================== exports ====================== */
  function tables() {
    const A = state.A, N = A.N;
    const T = [];
    T.push({ name: "Summary", rows: [
      ["Metric", "Value"],
      ["Responses (filtered)", N], ["Responses (all)", state.all.length],
      ["Responses last 7 days", A.last7],
      ["Overall mean suitability (1-5)", +A.overall.mean.toFixed(3)], ["Overall favourable %", A.overall.fav],
      ["Core mean", +A.core.mean.toFixed(3)], ["Elective mean", +A.elective.mean.toFixed(3)],
      ["Intent % (enrol / nominate / sponsor)", A.intent], ["Advocacy %", A.advocacy], ["Not interested %", A.notInterested],
      ["Median fee expectation (RM)", Math.round(A.feeMedian)],
      ["Generated", new Date().toISOString()], ["Data", state.preview ? "SAMPLE DATA" : "Live responses"]
    ].concat([[""], ["Key findings"]].concat(A.findings.map((f) => [f.replace(/<[^>]+>/g, "")]))) });
    T.push({ name: "Course ratings", rows: [["Rank", "Course", "Type", "n", "Mean", "SD", "CI95 +/-", "Favourable %", "Unfavourable %", "Verdict"]].concat(A.courses.map((c) => [c.rank || "", c.course.title, c.course.type, c.n, +c.mean.toFixed(3), +c.sd.toFixed(3), +c.ci.toFixed(3), c.fav, c.unf, c.verdict])) });
    T.push({ name: "Rating distribution", rows: [["Course", "1 Highly Unsuitable", "2 Unsuitable", "3 Moderate", "4 Suitable", "5 Highly Suitable"]].concat(A.courses.map((c) => [c.course.title].concat(c.dist))) });
    T.push({ name: "Course x category", rows: [["Course"].concat(A.catList.map((c, i) => shortCat(c) + " (n=" + A.categories[i].value + ")"))].concat(A.segMatrix.map((r) => [r.course.title].concat(r.cells.map((c) => (c.n ? +c.mean.toFixed(2) : ""))))) });
    T.push({ name: "Interest", rows: [["Interest", "Count", "%"]].concat(A.interest.map((i) => [i.label, i.value, pct(i.value, N)])) });
    T.push({ name: "Interest by category", rows: [["Category", "n", "Intent", "Advocacy", "Unsure", "Not interested"]].concat(A.interestByCat.map((r) => [r.cat, r.n, r.g.Intent, r.g.Advocacy, r.g.Unsure, r.g["Not interested"]])) });
    T.push({ name: "Study mode", rows: [["Study mode", "Count", "%"]].concat(A.modes.map((i) => [i.label, i.value, pct(i.value, N)])) });
    T.push({ name: "Study mode by category", rows: [["Category", "n"].concat(STUDY_MODES)].concat(A.modeByCat.map((r) => [r.cat, r.n].concat(r.counts))) });
    T.push({ name: "Fee", rows: [["Fee band", "Count", "%"]].concat(A.fees.map((i) => [i.label, i.value, pct(i.value, N)])) });
    T.push({ name: "Mode x fee", rows: [["Study mode"].concat(FEE_RANGES)].concat(A.modeFee.map((r) => [r.mode].concat(r.counts))) });
    T.push({ name: "Criteria", rows: [["Criterion", "Count", "% of respondents"]].concat(A.criteria.map((i) => [i.label, i.value, pct(i.value, N)])) });
    T.push({ name: "Criteria by category", rows: [["Criterion"].concat(A.catList.map(shortCat))].concat(A.criteriaByCat.map((r) => [r.criterion].concat(r.cells.map((c) => (c.n ? pct(c.v, c.n) : ""))))) });
    T.push({ name: "Categories", rows: [["Category", "Count", "%"]].concat(A.categories.map((i) => [i.label, i.value, pct(i.value, N)])) });
    T.push({ name: "Language", rows: [["Language", "Count", "%"]].concat(A.langs.map((i) => [i.label, i.value, pct(i.value, N)])) });
    T.push({ name: "Timeline", rows: [["Date", "Responses"]].concat(A.timeline.map((d) => [d.day, d.n])) });
    T.push({ name: "Feedback keywords", rows: [["Keyword", "Comments mentioning"]].concat(A.keywords.map((k) => [k.label, k.value])) });
    T.push({ name: "Feedback", rows: [["Date", "Category", "Comment"]].concat(A.feedback.map((f) => [f.ts ? f.ts.toISOString().slice(0, 10) : "", f.cat, f.text])) });
    T.push({ name: "Raw responses", rows: rawRows() });
    return T;
  }
  function rawRows() {
    const head = ["timestamp", "category", "category_other", "language", "study_mode", "fee_range", "fee_other", "interest", "criteria", "feedback"].concat(COURSES.map((c) => c.id + " | " + c.title));
    return [head].concat(state.filtered.map((r) => [r.raw.timestamp || "", r.cat, r.catOther, r.lang, r.mode, r.fee, r.raw.fee_other || "", r.interest, r.criteria.join(" | "), r.feedback].concat(COURSES.map((c) => r.ratings[c.id] || ""))));
  }
  function csvCell(v) {
    const s = String(v == null ? "" : v);
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  const toCsv = (rows) => rows.map((r) => r.map(csvCell).join(",")).join("\r\n");
  function download(name, text, type) {
    const blob = text instanceof Blob ? text : new Blob(["﻿" + text], { type: type || "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  const stamp = () => new Date().toISOString().slice(0, 10) + (state.preview ? "-SAMPLE" : "");
  function exportXlsx() {
    if (!window.XLSX) { alert("The Excel library could not be loaded (offline?). Use the CSV download instead."); return; }
    const wb = XLSX.utils.book_new();
    tables().forEach((t) => {
      const ws = XLSX.utils.aoa_to_sheet(t.rows);
      const widths = [];
      t.rows.slice(0, 60).forEach((r) => r.forEach((c, i) => { widths[i] = Math.min(60, Math.max(widths[i] || 8, String(c == null ? "" : c).length + 2)); }));
      ws["!cols"] = widths.map((w) => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, ws, t.name.slice(0, 31));
    });
    XLSX.writeFile(wb, "IMEN-MSc-Semiconductor-analysis-" + stamp() + ".xlsx");
  }
  function exportSummaryCsv() {
    const out = [];
    tables().filter((t) => t.name !== "Raw responses").forEach((t) => { out.push([t.name.toUpperCase()]); t.rows.forEach((r) => out.push(r)); out.push([]); });
    download("IMEN-MSc-Semiconductor-analysis-summary-" + stamp() + ".csv", toCsv(out));
  }
  function exportRaw() { download("IMEN-MSc-Semiconductor-responses-" + stamp() + ".csv", toCsv(rawRows())); }
  function exportJson() {
    const obj = { generated: new Date().toISOString(), sample: state.preview, filters: state.filters, tables: {}, responses: state.filtered.map((r) => r.raw) };
    tables().forEach((t) => (obj.tables[t.name] = t.rows));
    download("IMEN-MSc-Semiconductor-analysis-" + stamp() + ".json", JSON.stringify(obj, null, 2), "application/json");
  }

  /* ====================== wiring ====================== */
  function refresh() {
    applyFilters();
    state.A = compute();
    render();
  }
  function boot() {
    const q = new URLSearchParams(location.search);
    $("unlockBtn").addEventListener("click", unlock);
    $("password").addEventListener("keydown", (e) => { if (e.key === "Enter") unlock(); });
    $("logoutBtn").addEventListener("click", logout);
    $("refreshBtn").addEventListener("click", () => { if (state.preview) { refresh(); } else if (isDemo()) { loadDemo(); buildFilterOptions(); refresh(); } else location.reload(); });
    $("clearFilters").addEventListener("click", () => { state.filters = { cat: "", mode: "", interest: "", lang: "", days: 0 }; buildFilterOptions(); refresh(); });
    [["fCategory", "cat"], ["fMode", "mode"], ["fInterest", "interest"], ["fLang", "lang"]].forEach((p) => $(p[0]).addEventListener("change", (e) => { state.filters[p[1]] = e.target.value; refresh(); }));
    $("fPeriod").addEventListener("change", (e) => { state.filters.days = Number(e.target.value); refresh(); });
    $("dataSearch").addEventListener("input", (e) => { state.search = e.target.value; state.page = 0; renderData(); });
    $("prevPage").addEventListener("click", () => { state.page--; renderData(); });
    $("nextPage").addEventListener("click", () => { state.page++; renderData(); });
    // download menu
    const menu = $("dlMenu");
    $("dlBtn").addEventListener("click", (e) => { e.stopPropagation(); menu.classList.toggle("open"); });
    document.addEventListener("click", () => menu.classList.remove("open"));
    menu.querySelectorAll("[data-dl]").forEach((b) => b.addEventListener("click", () => {
      const k = b.dataset.dl;
      if (k === "xlsx") exportXlsx();
      else if (k === "pdf") window.print();
      else if (k === "csv-raw") exportRaw();
      else if (k === "csv-summary") exportSummaryCsv();
      else if (k === "json") exportJson();
    }));

    if (isDemo()) $("gateHint").classList.remove("hidden");
    if (q.get("preview") === "1") {
      state.preview = true;
      state.all = buildSample(Number(q.get("n")) || 140).map(norm);
      enterApp(true);
    } else if (sessionStorage.getItem(AUTH_KEY) === "1" && isDemo()) {
      enterApp();
    }
  }
  document.addEventListener("DOMContentLoaded", boot);
})();
