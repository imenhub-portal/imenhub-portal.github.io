/* ==========================================================
   IMEN · M.Sc. Semiconductor Engineering — survey app
   ========================================================== */
const READ_KEY = APP_CONFIG.READ_KEY || "msc_semicond_read_v1";
const DRAFT_KEY = APP_CONFIG.DRAFT_KEY || "msc_semicond_draft_v1";
const DIE_KEY = "msc_semicond_die_v1";
const SKIM_CACHE = {}; // non-reactive memo for skimPoints()

function surveyApp() {
  return {
    /* ---------- state ---------- */
    lang: "en",
    started: false,
    resumeStep: 1,
    step: 1,
    deckIndex: 0,
    totalSteps: 3,
    tab: "core",
    loading: false,
    submitted: false,
    showAlert: false,
    alertTitle: "",
    alertMsg: "",
    diePlaced: false,
    circuitOn: false,
    dieLabel: { x: 0, y: 0, show: false },
    sheet: { open: false, course: null },
    viewer: null,
    deckViewer: null,
    readCourses: {},
    draftTimer: null,
    toast: "",

    form: {
      email: "",
      name: "",
      category: "",
      category_other: "",
      criteria: [],
      study_mode: "",
      fee_range: "",
      fee_other: "",
      interest: "",
      feedback: ""
    },

    ratings: {},

    /* ---------- init ---------- */
    init() {
      const saved = localStorage.getItem(APP_CONFIG.LANG_KEY);
      this.lang = saved === "ms" ? "ms" : "en";
      document.documentElement.lang = this.lang;

      COURSES.forEach((c) => {
        this.ratings[c.id] = 0;
      });

      // reading progress + in-progress answers
      try {
        const r = JSON.parse(localStorage.getItem(READ_KEY) || "{}");
        this.readCourses = r && typeof r === "object" ? r : {};
      } catch (e) {
        this.readCourses = {};
      }
      // the chip is a toy, not a gate: the "Drag me" label shows on every visit
      this.diePlaced = false;
      try {
        const d = JSON.parse(localStorage.getItem(DRAFT_KEY) || "null");
        if (d && d.form) {
          this.form = Object.assign({}, this.form, d.form);
          if (d.ratings) this.ratings = Object.assign({}, this.ratings, d.ratings);
          // always open on the landing page; Proceed resumes where they left off
          this.started = false;
          this.resumeStep = d.step || 1;
        }
      } catch (e) {}

      this.$watch("started", () => this.syncDeck());
      this.$watch("step", () => {
        this.syncDeck();
        this.resetScroll();
      });
      this.$watch("deckIndex", () => {
        this.syncDeck();
        this.resetScroll();
      });

      this.$nextTick(() => {
        const canvas = this.$refs.heroCanvas;
        if (canvas && window.IMEN) {
          try {
            this.hero = window.IMEN.initHeroScene(canvas);
            if (this.hero && this.hero.onPlace) {
              // circuit board powers up with the wafer pulse, not after it
              this.hero.onPlace(() => {
                this.circuitOn = true;
                this.onDiePlaced(); // hides the "Drag me" label straight away
              });
            }
            if (this.hero && this.hero.onUnlock) {
              this.hero.onUnlock(() => this.onDiePlaced());
            }
            if (this.hero && this.hero.dieLabelPos) {
              // pin the "Drag me" label to the chip straight from the render loop:
              // plain DOM writes, no reactive state and no second requestAnimationFrame loop
              const el = this.$refs.dieLabelEl;
              let lx = -999, ly = -999;
              const place = () => {
                if (this.diePlaced || !el) return;
                const p = this.hero.dieLabelPos();
                if (Math.abs(p.x - lx) < 0.5 && Math.abs(p.y - ly) < 0.5) return;
                lx = p.x;
                ly = p.y;
                el.style.left = p.x + "px";
                el.style.top = p.y + "px";
                el.style.visibility = "visible";
              };
              if (this.hero.onFrame) this.hero.onFrame(place);
              place();
            }
          } catch (e) {
            console.warn("Hero scene unavailable:", e);
          }
        }
        this.buildCircuit();
        const stage = document.querySelector(".hero-stage");
        if (stage && "IntersectionObserver" in window) {
          new IntersectionObserver((en) => stage.classList.toggle("paused", !en[0].isIntersecting)).observe(stage);
        }
        this.initReveal();
      });
    },

    /* ---------- hero circuit backdrop (decorative, seeded so it is stable) ---------- */
    buildCircuit() {
      const svg = this.$refs.circuitSvg;
      if (!svg) return;
      let seed = 11;
      const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
      const W = 1000, H = 560, cx = 500, cy = 292;
      const pick = (a) => a[Math.floor(rnd() * a.length)];
      const out = [];
      const parts = [];
      for (let i = 0; i < 46; i++) {
        const x = 30 + rnd() * (W - 60);
        const y = 30 + rnd() * (H - 60);
        if (Math.hypot(x - cx, (y - cy) * 1.4) < 170) continue;
        if (rnd() > 0.5) parts.push('<rect class="part" x="' + x.toFixed(0) + '" y="' + y.toFixed(0) + '" width="' + (16 + rnd() * 26).toFixed(0) + '" height="' + (9 + rnd() * 12).toFixed(0) + '" rx="2"/>');
        else parts.push('<circle class="part" cx="' + x.toFixed(0) + '" cy="' + y.toFixed(0) + '" r="' + (6 + rnd() * 6).toFixed(0) + '"/>');
      }
      for (let i = 0; i < 96; i++) {
        let base = Math.round(rnd() * 7) * (Math.PI / 4);
        const r0 = 90 + rnd() * 180;
        let x = cx + Math.cos(base) * r0 + (rnd() - 0.5) * 60;
        let y = cy + Math.sin(base) * r0 * 0.7 + (rnd() - 0.5) * 40;
        let d = "M" + x.toFixed(0) + " " + y.toFixed(0);
        let ang = base;
        for (let s = 0; s < 7; s++) {
          ang = base + pick([0, 0, Math.PI / 4, -Math.PI / 4]);
          const len = 45 + rnd() * 120;
          x += Math.cos(ang) * len;
          y += Math.sin(ang) * len * 0.75;
          d += " L" + x.toFixed(0) + " " + y.toFixed(0);
          if (x < -20 || x > W + 20 || y < -20 || y > H + 20) break;
        }
        const col = pick(["c1", "c2", "c3", "c1", "c2"]);
        const dur = (2.8 + rnd() * 3).toFixed(1);
        const delay = (-rnd() * 5).toFixed(1);
        out.push('<path class="trace" d="' + d + '"/>');
        if (i % 3 === 0) out.push('<path class="pulse ' + col + '" pathLength="100" d="' + d + '" style="animation-duration:' + dur + "s;animation-delay:" + delay + 's"/>');
        if (x > 8 && x < W - 8 && y > 8 && y < H - 8) out.push('<circle class="pad" cx="' + x.toFixed(0) + '" cy="' + y.toFixed(0) + '" r="4"/>');
      }
      svg.innerHTML = parts.join("") + out.join("");
    },

    onDiePlaced() {
      if (this.diePlaced) return;
      this.diePlaced = true;
    },

    initReveal() {
      // Motion is intentionally minimal: content is visible immediately.
      document.querySelectorAll(".reveal").forEach((el) => el.classList.add("in"));
    },

    /* ---------- i18n ---------- */
    t(key, vars) {
      return t(this.lang, key, vars);
    },

    /* ---------- reading progress ---------- */
    isRead(id) {
      return !!this.readCourses[id];
    },
    markRead(id) {
      if (this.readCourses[id]) return;
      this.readCourses = Object.assign({}, this.readCourses, { [id]: true });
      try {
        localStorage.setItem(READ_KEY, JSON.stringify(this.readCourses));
      } catch (e) {}
    },
    get readCount() {
      return Object.keys(this.readCourses).length;
    },
    get allRead() {
      return this.readCount >= COURSES.length;
    },
    get readProgress() {
      return Math.round((this.readCount / COURSES.length) * 100);
    },
    markAllRead() {
      const all = {};
      COURSES.forEach((c) => {
        all[c.id] = true;
      });
      this.readCourses = all;
      try {
        localStorage.setItem(READ_KEY, JSON.stringify(all));
      } catch (e) {}
    },
    setLang(l) {
      this.lang = l;
      localStorage.setItem(APP_CONFIG.LANG_KEY, l);
      document.documentElement.lang = l;
    },
    toggleLang() {
      this.setLang(this.lang === "en" ? "ms" : "en");
    },

    /* ---------- curriculum explorer: stages + deck ---------- */
    get stageGroups() {
      const order = [
        { key: "foundations", ids: ["core_1", "core_2", "core_3"] },
        { key: "fabrication", ids: ["core_4", "core_5"] },
        { key: "practice", ids: ["core_6", "core_7", "core_8", "core_9"] },
        { key: "specialise", ids: ["elective_1", "elective_2", "elective_3", "elective_4", "elective_5"] }
      ];
      return order.map((s, si) => {
        const courses = s.ids.map((id) => COURSES.find((c) => c.id === id)).filter(Boolean);
        return {
          key: s.key,
          index: si,
          courses: courses,
          read: courses.filter((c) => this.readCourses[c.id]).length
        };
      });
    },
    stageRead(group) {
      return group.courses.filter((c) => this.readCourses[c.id]).length;
    },
    courseIndex(id) {
      return COURSES.findIndex((c) => c.id === id);
    },
    courseNumber(id) {
      return this.courseIndex(id) + 1;
    },
    readSecs(course) {
      const words = ((course.summary || "") + " " + (course.overview || "")).split(/\s+/).length;
      return Math.max(15, Math.round(words / 4.2));
    },
    skimPoints(course) {
      if (!course || !course.overview) return [];
      // memoised: the template re-asks on every reactive change
      if (SKIM_CACHE[course.id]) return SKIM_CACHE[course.id];
      return (SKIM_CACHE[course.id] = this.computeSkim(course));
    },
    computeSkim(course) {
      const stop = new Set(["this", "that", "through", "students", "course", "their", "with", "from"]);
      const sents = course.overview
        .replace(/\s+/g, " ")
        .match(/[^.!?]+[.!?]/g) || [];
      const scored = sents
        .map((s) => s.trim())
        .filter((s) => s.length > 40 && s.length < 260)
        .map((s, i) => {
          let score = i === 0 ? 0.4 : 0;
          if (/covers|learn|emphasis|equip|develop|provides|focuses|explores/i.test(s)) score += 1;
          const words = s.toLowerCase().split(/\W+/);
          score += words.filter((w) => w.length > 7 && !stop.has(w)).length * 0.05;
          return { s: s.replace(/[.!?]+$/, ""), score: score };
        });
      return scored
        .sort((a, b) => b.score - a.score)
        .slice(0, 3)
        .sort((a, b) => scored.indexOf(a) - scored.indexOf(b))
        .map((x) => x.s);
    },
    nextUnread(fromId) {
      const i = this.courseIndex(fromId);
      for (let k = 1; k <= COURSES.length; k++) {
        const c = COURSES[(i + k) % COURSES.length];
        if (!this.readCourses[c.id]) return c;
      }
      return COURSES[(i + 1) % COURSES.length];
    },
    nextCourse() {
      if (!this.sheet.course) return;
      const n = this.nextUnread(this.sheet.course.id);
      this.openCourse(n);
    },
    prevCourse() {
      if (!this.sheet.course) return;
      const i = this.courseIndex(this.sheet.course.id);
      this.openCourse(COURSES[(i - 1 + COURSES.length) % COURSES.length]);
    },
    get sheetHasNext() {
      if (!this.sheet.course) return false;
      return !this.allRead;
    },

    /* ---------- survey step 1: one-by-one course deck ---------- */
    get deckCourse() {
      return COURSES[this.deckIndex] || COURSES[0];
    },
    get deckLast() {
      return this.deckIndex >= COURSES.length - 1;
    },
    syncDeck() {
      if (!(this.started && this.step === 1)) {
        if (this.deckViewer) {
          try {
            this.deckViewer.dispose();
          } catch (e) {}
          this.deckViewer = null;
        }
        return;
      }
      setTimeout(() => {
        const canvas = this.$refs.deckCanvas;
        if (!canvas || !window.IMEN) return;
        try {
          if (!this.deckViewer) {
            this.deckViewer = window.IMEN.createCourseViewer(canvas);
            if (this.deckViewer) this.deckViewer.start();
          }
          if (this.deckViewer) this.deckViewer.setCourse(this.deckCourse.scene);
        } catch (e) {
          console.warn("Course scene unavailable:", e);
        }
      }, 80);
    },
    /* ---------- icons: rating faces + preference options ---------- */
    rateColor(v) {
      return { 1: "#ef4444", 2: "#f97316", 3: "#eab308", 4: "#14b8a6", 5: "#2563eb" }[v];
    },
    rateIcon(v) {
      const f = {"1": "<circle cx=\"12\" cy=\"12\" r=\"9.5\"/><circle cx=\"9\" cy=\"11.6\" r=\".9\" fill=\"currentColor\" stroke=\"none\"/><circle cx=\"15\" cy=\"11.6\" r=\".9\" fill=\"currentColor\" stroke=\"none\"/><path d=\"M8 8.4l2.6 1M16 8.4l-2.6 1\"/><path d=\"M8 17.2c1.1-2.3 2.4-3.2 4-3.2s2.9.9 4 3.2\"/>", "2": "<circle cx=\"12\" cy=\"12\" r=\"9.5\"/><circle cx=\"9\" cy=\"10\" r=\".9\" fill=\"currentColor\" stroke=\"none\"/><circle cx=\"15\" cy=\"10\" r=\".9\" fill=\"currentColor\" stroke=\"none\"/><path d=\"M8.5 16.2c.9-1.5 2.2-2.2 3.5-2.2s2.6.7 3.5 2.2\"/>", "3": "<circle cx=\"12\" cy=\"12\" r=\"9.5\"/><circle cx=\"9\" cy=\"10\" r=\".9\" fill=\"currentColor\" stroke=\"none\"/><circle cx=\"15\" cy=\"10\" r=\".9\" fill=\"currentColor\" stroke=\"none\"/><path d=\"M8.5 15.2h7\"/>", "4": "<circle cx=\"12\" cy=\"12\" r=\"9.5\"/><circle cx=\"9\" cy=\"10\" r=\".9\" fill=\"currentColor\" stroke=\"none\"/><circle cx=\"15\" cy=\"10\" r=\".9\" fill=\"currentColor\" stroke=\"none\"/><path d=\"M8 14c1 1.8 2.4 2.6 4 2.6s3-.8 4-2.6\"/>", "5": "<circle cx=\"12\" cy=\"12\" r=\"9.5\"/><circle cx=\"9\" cy=\"9.6\" r=\".9\" fill=\"currentColor\" stroke=\"none\"/><circle cx=\"15\" cy=\"9.6\" r=\".9\" fill=\"currentColor\" stroke=\"none\"/><path d=\"M7.4 13.2h9.2a4.6 4.6 0 01-9.2 0z\" fill=\"currentColor\"/>"};
      return '<svg class="h-8 w-8" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">' + f[v] + "</svg>";
    },
    prefIcon(group, i) {
      const g = {"mode": ["<rect x=\"4\" y=\"3.5\" width=\"11\" height=\"17\" rx=\"1.5\"/><path d=\"M15 9.5h4a1.5 1.5 0 011.5 1.5v9.5H15M8 8h3M8 12h3M8 16h3\"/>", "<rect x=\"3\" y=\"4\" width=\"18\" height=\"12\" rx=\"2\"/><path d=\"M8 20h8M12 16v4\"/>", "<path d=\"M12 3l9 5-9 5-9-5z\"/><path d=\"M3 13l9 5 9-5\"/>", "<rect x=\"3.5\" y=\"5\" width=\"17\" height=\"15\" rx=\"2\"/><path d=\"M3.5 10h17M8 3v4M16 3v4\"/>"], "fee": ["<ellipse cx=\"12\" cy=\"17\" rx=\"7\" ry=\"2.6\"/><path d=\"M5 17v-0.5M19 17v-0.5\"/>", "<ellipse cx=\"12\" cy=\"17\" rx=\"7\" ry=\"2.6\"/><ellipse cx=\"12\" cy=\"13\" rx=\"7\" ry=\"2.6\"/><path d=\"M5 17v-0.5M19 17v-0.5\"/>", "<ellipse cx=\"12\" cy=\"17\" rx=\"7\" ry=\"2.6\"/><ellipse cx=\"12\" cy=\"13\" rx=\"7\" ry=\"2.6\"/><ellipse cx=\"12\" cy=\"9\" rx=\"7\" ry=\"2.6\"/><path d=\"M5 17v-0.5M19 17v-0.5\"/>", "<path d=\"M4 20h4L19 9a2.1 2.1 0 00-3-3L5 17z\"/>"], "interest": ["<circle cx=\"9\" cy=\"8\" r=\"3.4\"/><path d=\"M2.8 20c.6-3.6 3.1-5.4 6.2-5.4M15 16.5l2.2 2.2 4-4.4\"/>", "<rect x=\"3\" y=\"7\" width=\"18\" height=\"13\" rx=\"2\"/><path d=\"M8.5 7V5.5A1.5 1.5 0 0110 4h4a1.5 1.5 0 011.5 1.5V7M3 12.5h18\"/>", "<path d=\"M12 20s-7.5-4.6-7.5-10A4.2 4.2 0 0112 7.6 4.2 4.2 0 0119.5 10c0 5.4-7.5 10-7.5 10z\"/>", "<circle cx=\"12\" cy=\"12\" r=\"9\"/><path d=\"M9 9l6 6M15 9l-6 6\"/>", "<circle cx=\"12\" cy=\"12\" r=\"9\"/><path d=\"M9.6 9.4a2.5 2.5 0 114 2c-.9.6-1.6 1.1-1.6 2.3M12 16.8h.01\"/>"]};
      const p = (g[group] || [])[i];
      return p ? '<svg class="h-[18px] w-[18px]" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">' + p + "</svg>" : "";
    },
    // a new card / step always starts at the top, so nothing in the course text is missed
    resetScroll() {
      this.$nextTick(() => {
        document.querySelectorAll("#survey-card .overflow-y-auto").forEach((el) => {
          el.scrollTop = 0;
        });
      });
    },
    scaleLabel(v) {
      return t(this.lang, "scale_" + v);
    },
    deckNext() {
      if (!this.ratings[this.deckCourse.id]) return this.alert(this.t("err_title"), this.t("err_rate"));
      this.markRead(this.deckCourse.id);
      if (this.deckLast) return this.next();
      this.deckIndex++;
      this.$nextTick(() => this.scrollTop());
    },
    deckPrev() {
      if (this.deckIndex > 0) {
        this.deckIndex--;
        this.$nextTick(() => this.scrollTop());
      }
    },

    /* ---------- helpers ---------- */
    get progress() {
      if (!this.started) return 0;
      // 14 course cards + preferences + review
      const total = COURSES.length + 2;
      const done = this.step === 1 ? this.deckIndex : COURSES.length + (this.step - 2);
      return (done / total) * 100;
    },
    get coreCourses() {
      return COURSES.filter((c) => c.type === "core");
    },
    get electiveCourses() {
      return COURSES.filter((c) => c.type === "elective");
    },
    get visibleCourses() {
      return this.tab === "core" ? this.coreCourses : this.electiveCourses;
    },
    get ratedCount() {
      return COURSES.filter((c) => this.ratings[c.id] > 0).length;
    },
    get categories() { return CATEGORIES; },
    get criteriaOptions() { return CRITERIA; },
    get studyModes() { return STUDY_MODES; },
    get feeRanges() { return FEE_RANGES; },
    get interestOptions() { return INTEREST_OPTIONS; },
    get qCategory() { return CATEGORY_QUESTION; },
    get qCriteria() { return CRITERIA_QUESTION; },
    get qCore() { return CORE_QUESTION; },
    get qElective() { return ELECTIVE_QUESTION; },
    get qMode() { return STUDY_MODE_QUESTION; },
    get qFee() { return FEE_QUESTION; },
    get qInterest() { return INTEREST_QUESTION; },
    get qFeedback() { return FEEDBACK_QUESTION; },
    get courseCount() { return COURSES.length; },
    get programmeIntro() { return PROGRAMME.intro; },
    stepName(n) {
      return this.t("step_" + n);
    },

    scrollToSurvey() {
      this.scrollToCard();
    },
    scrollToHero() {
      window.scrollTo({ top: 0, behavior: "smooth" });
    },
    scrollTop() {
      this.scrollToCard();
    },
    scrollToCard() {
      const el = document.getElementById("survey-card");
      if (!el) return;
      const y = el.getBoundingClientRect().top + window.scrollY - 92;
      window.scrollTo({ top: Math.max(0, y), behavior: "smooth" });
    },

    alert(title, msg) {
      this.alertTitle = title;
      this.alertMsg = msg;
      this.showAlert = true;
    },

    /* ---------- course sheet ---------- */
    openCourse(course) {
      this.markRead(course.id);
      this.sheet.course = course;
      this.sheet.open = true;
      document.body.style.overflow = "hidden";
      setTimeout(() => {
        const canvas = this.$refs.courseCanvas;
        if (canvas && window.IMEN) {
          try {
            this.viewer = window.IMEN.createCourseViewer(canvas);
            this.viewer.setCourse(course.scene);
            this.viewer.start();
          } catch (e) {
            console.warn("Course scene unavailable:", e);
          }
        }
      }, 80);
    },
    closeCourse() {
      this.sheet.open = false;
      document.body.style.overflow = "";
      if (this.viewer) {
        try {
          this.viewer.dispose();
        } catch (e) {}
        this.viewer = null;
      }
      setTimeout(() => {
        this.sheet.course = null;
      }, 260);
    },
    rate(id, value) {
      this.ratings[id] = value;
      this.saveDraft();
    },

    /* ---------- profile gate ---------- */
    startSurvey() {
      const f = this.form;
      if (!this.emailOk(f.email)) return this.alert(this.t("err_title"), this.t("err_email"));
      if (!f.category) return this.alert(this.t("err_title"), this.t("err_required"));
      if (f.category === "Others" && !String(f.category_other || "").trim())
        return this.alert(this.t("err_title"), this.t("err_other"));
      if (!f.criteria.length) return this.alert(this.t("err_title"), this.t("err_criteria"));

      this.started = true;
      this.deckIndex = 0;
      this.step = this.resumeStep || 1;
      this.resumeStep = 1;
      this.saveDraft();
      this.$nextTick(() => window.scrollTo({ top: 0 }));
    },
    editProfile() {
      this.resumeStep = this.step;
      this.started = false;
      this.$nextTick(() => {
        window.dispatchEvent(new Event("resize")); // hero canvas was hidden
        this.scrollTop();
      });
    },

    saveDraft() {
      try {
        localStorage.setItem(DRAFT_KEY, JSON.stringify({ form: this.form, ratings: this.ratings, started: this.started, step: this.step }));
      } catch (e) {}
    },

    /* ---------- validation ---------- */
    emailOk(v) {
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || "").trim());
    },
    validateStep() {
      const f = this.form;
      if (this.step === 2) {
        if (!f.study_mode || !f.fee_range || !f.interest)
          return this.alert(this.t("err_title"), this.t("err_prefs"));
        if (f.fee_range === "Others" && !String(f.fee_other).trim())
          return this.alert(this.t("err_title"), this.t("err_other"));
        return true;
      }
      return true;
    },

    next() {
      if (!this.validateStep()) return;
      if (this.step < this.totalSteps) {
        this.step++;
        this.saveDraft();
        this.$nextTick(() => this.scrollTop());
      }
    },
    back() {
      if (this.step > 1) {
        this.step--;
        if (this.step === 1) this.deckIndex = COURSES.length - 1;
        this.saveDraft();
        this.$nextTick(() => this.scrollTop());
      }
    },
    goStep(n) {
      if (!this.started || n > this.step) return;
      this.step = n;
      if (n === 1) this.deckIndex = 0;
      this.saveDraft();
      this.$nextTick(() => this.scrollTop());
    },

    critIcon(i) {
      const p = ['<path d="M12 4v16M5 7h14M5 7l-2.5 6a3 3 0 005 0zM19 7l-2.5 6a3 3 0 005 0z"/>', '<rect x="6" y="6" width="12" height="12" rx="2"/><path d="M9 3v3M15 3v3M9 18v3M15 18v3M3 9h3M3 15h3M18 9h3M18 15h3"/>', '<rect x="3" y="7" width="18" height="13" rx="2"/><path d="M8.5 7V5.5A1.5 1.5 0 0110 4h4a1.5 1.5 0 011.5 1.5V7M3 12.5h18"/>', '<circle cx="9" cy="8.5" r="3"/><path d="M3 19.5c.6-3 2.8-4.5 6-4.5s5.4 1.5 6 4.5M16 6a3 3 0 010 5.5M18 15.2c1.7.5 2.7 1.9 3 4.3"/>'];
      return '<svg class="h-[18px] w-[18px]" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">' + (p[i] || p[0]) + "</svg>";
    },
    toggleCriterion(c) {
      const i = this.form.criteria.indexOf(c);
      if (i === -1) this.form.criteria.push(c);
      else this.form.criteria.splice(i, 1);
    },

    /* ---------- submit ---------- */
    async submit() {
      const f = this.form;
      if (!f.study_mode || !f.fee_range || !f.interest) {
        this.step = 2;
        this.$nextTick(() => this.scrollTop());
        return this.alert(this.t("err_title"), this.t("err_prefs"));
      }
      if (f.fee_range === "Others" && !String(f.fee_other).trim()) {
        this.step = 2;
        this.$nextTick(() => this.scrollTop());
        return this.alert(this.t("err_title"), this.t("err_other"));
      }
      this.loading = true;

      const payload = {
        timestamp: new Date().toISOString(),
        email: f.email.trim(),
        name: f.name.trim(),
        language: this.lang,
        category: f.category,
        category_other: f.category === "Others" ? f.category_other.trim() : "",
        criteria: f.criteria.join(" | "),
        study_mode: f.study_mode,
        fee_range: f.fee_range,
        fee_other: f.fee_range === "Others" ? String(f.fee_other).trim() : "",
        interest: f.interest,
        feedback: f.feedback.trim(),
        user_agent: navigator.userAgent
      };
      COURSES.forEach((c) => {
        payload[c.id] = this.ratings[c.id] || "";
      });

      try {
        if (APP_CONFIG.SCRIPT_URL && !APP_CONFIG.DEMO_MODE) {
          await fetch(APP_CONFIG.SCRIPT_URL, {
            method: "POST",
            mode: "no-cors",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify(payload)
          });
        } else {
          const key = APP_CONFIG.RESPONSE_KEY;
          const list = JSON.parse(localStorage.getItem(key) || "[]");
          list.push(payload);
          localStorage.setItem(key, JSON.stringify(list));
          await new Promise((r) => setTimeout(r, 700));
        }
        this.loading = false;
        this.submitted = true;
        try {
          localStorage.removeItem(DRAFT_KEY);
        } catch (e) {}
        window.scrollTo({ top: 0, behavior: "smooth" });
      } catch (e) {
        this.loading = false;
        this.alert(this.t("err_title"), this.t("err_submit"));
      }
    },

    reset() {
      this.submitted = false;
      this.started = false;
      this.resumeStep = 1;
      this.step = 1;
      this.form = {
        email: "",
        name: "",
        category: "",
        category_other: "",
        criteria: [],
        study_mode: "",
        fee_range: "",
        fee_other: "",
        interest: "",
        feedback: ""
      };
      COURSES.forEach((c) => {
        this.ratings[c.id] = 0;
      });
      try {
        localStorage.removeItem(DRAFT_KEY);
      } catch (e) {}
      this.$nextTick(() => window.scrollTo({ top: 0, behavior: "smooth" }));
    },

    /* ---------- review helpers ---------- */
    ratingWord(v) {
      return v ? v + " · " + SCALE_LABELS[v] : "—";
    },
    get reviewCore() {
      return this.coreCourses.map((c) => ({ title: c.title, value: this.ratings[c.id] }));
    },
    get reviewElective() {
      return this.electiveCourses.map((c) => ({ title: c.title, value: this.ratings[c.id] }));
    }
  };
}
