/* core_1 · Semiconductor Device Physics — cut-away n-MOSFET with inversion channel */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  IMEN.courseScenes = IMEN.courseScenes || {};
  IMEN.courseScenes.device = function (kit) {
    const T = THREE, M = kit.M;
    const g = new T.Group();
    const rnd = kit.rng(21);

    /* ---------- helpers ---------- */
    // cross-section texture: gradient + crystal-lattice dots
    function latticeTex(top, bot, dot, step, seed) {
      const r = kit.rng(seed);
      return kit.canvasTex(512, 256, function (c, W, H) {
        const gr = c.createLinearGradient(0, 0, 0, H);
        gr.addColorStop(0, top); gr.addColorStop(1, bot);
        c.fillStyle = gr; c.fillRect(0, 0, W, H);
        c.fillStyle = dot;
        for (let y = step / 2; y < H; y += step)
          for (let x = step / 2 + ((y / step) % 2 ? step / 2 : 0); x < W; x += step) {
            c.globalAlpha = 0.22 + r() * 0.2;
            c.beginPath(); c.arc(x, y, 2.2, 0, 6.283); c.fill();
          }
        c.globalAlpha = 1;
      });
    }
    function box(w, h, d, mat, x, y, z) {
      const m = new T.Mesh(new T.BoxGeometry(w, h, d), mat);
      m.position.set(x, y, z);
      g.add(m);
      return m;
    }

    /* ---------- p-type substrate (translucent crystal block) ---------- */
    const W = 4.8, D = 2.2, H = 1.5;
    const subMat = new T.MeshPhysicalMaterial({
      map: latticeTex("#a8733a", "#4a2f35", "#ffd9a0", 22, 3), color: 0xffffff, metalness: 0.2, roughness: 0.25,
      transparent: true, opacity: 0.7, clearcoat: 1, clearcoatRoughness: 0.08, depthWrite: false, envMapIntensity: 1.2, emissive: 0x3a2010, emissiveIntensity: 0.6
    });
    const sub = box(W, H, D, subMat, 0, -H / 2, 0);
    sub.renderOrder = 1;
    const edgeMat = new T.LineBasicMaterial({ color: 0xffd9a0, transparent: true, opacity: 0.55 });
    const subEdges = new T.LineSegments(new T.EdgesGeometry(sub.geometry), edgeMat);
    subEdges.position.copy(sub.position); g.add(subEdges);

    /* ---------- n+ wells ---------- */
    const wellMat = new T.MeshPhysicalMaterial({
      map: latticeTex("#2a6bff", "#0f2aa0", "#9fdcff", 22, 9), color: 0x7aa2ff, emissive: 0x1040ff, emissiveIntensity: 0.6,
      metalness: 0.1, roughness: 0.2, transparent: true, opacity: 0.72, clearcoat: 1, depthWrite: false
    });
    const wellW = 1.45, wellH = 0.5;
    [-1, 1].forEach(function (s) {
      const w = box(wellW, wellH, D + 0.03, wellMat, s * (0.9 + wellW / 2 + 0.02) , -wellH / 2 + 0.002, 0);
      w.renderOrder = 2;
      const e = new T.LineSegments(new T.EdgesGeometry(w.geometry), new T.LineBasicMaterial({ color: 0x8fd0ff, transparent: true, opacity: 0.8 }));
      e.position.copy(w.position); g.add(e);
    });

    /* ---------- depletion regions (glow around junctions) ---------- */
    const depMat = kit.M.additive(0xb15cff, 0.3);
    const depl = [];
    [-0.9, 0.9].forEach(function (x) {
      const m = new T.Mesh(new T.SphereGeometry(1, 32, 16, 0, Math.PI * 2, Math.PI / 2, Math.PI / 2), depMat.clone());
      m.scale.set(0.42, 0.78, 1.1);
      m.position.set(x, 0, 0);
      m.renderOrder = 3;
      g.add(m); depl.push(m);
    });

    /* ---------- inversion channel ---------- */
    const chanMat = kit.M.additive(0x33e6ff, 0.55);
    const chan = new T.Mesh(new T.BoxGeometry(1.8, 0.07, D - 0.04), chanMat);
    chan.position.set(0, -0.045, 0);
    chan.renderOrder = 4;
    g.add(chan);

    /* ---------- gate stack ---------- */
    const oxide = box(1.96, 0.07, D - 0.02, new T.MeshPhysicalMaterial({ color: 0xcfe6ff, roughness: 0.05, metalness: 0, transparent: true, opacity: 0.55, clearcoat: 1, emissive: 0x335577, emissiveIntensity: 0.4 }), 0, 0.035, 0);
    oxide.renderOrder = 5;
    const polyTex = kit.tex.brushed(256, 256, { repeat: [3, 1] });
    const polyMat = new T.MeshPhysicalMaterial({ color: 0xc4cbdc, metalness: 0.9, roughness: 0.28, clearcoat: 0.4 });
    const gate = new T.Mesh(kit.roundedBox(1.8, 0.36, D - 0.1, 0.04, 0.02), polyMat);
    gate.position.set(0, 0.07 + 0.18, 0);
    g.add(gate);
    // spacers
    const spMat = M.ceramic(0xd9e2f0);
    [-1, 1].forEach(function (s) { box(0.07, 0.3, D - 0.14, spMat, s * 0.935, 0.2, 0); });

    /* ---------- interlayer dielectric (glass) ---------- */
    const ildMat = M.glass(0x9cc4ff, 0.24);
    const ildW = 1.46;
    [-1, 1].forEach(function (s) {
      const m = box(ildW, 0.62, D, ildMat, s * (0.97 + ildW / 2), 0.31, 0);
      m.renderOrder = 6;
      const e = new T.LineSegments(new T.EdgesGeometry(m.geometry), new T.LineBasicMaterial({ color: 0xbfd8ff, transparent: true, opacity: 0.4 }));
      e.position.copy(m.position); g.add(e);
    });

    /* ---------- contacts, pads and leads ---------- */
    const cu = M.copper(), au = M.gold(); au.emissive = new T.Color(0x6a3c00); au.emissiveIntensity = 0.55; cu.emissive = new T.Color(0x4a1a08); cu.emissiveIntensity = 0.4;
    const viaGeo = new T.CylinderGeometry(0.085, 0.085, 1, 20);
    function via(x, z, y0, y1) {
      const m = new T.Mesh(viaGeo, cu);
      m.scale.y = y1 - y0; m.position.set(x, (y0 + y1) / 2, z); g.add(m);
    }
    function pad(x, y, w, d) {
      const m = new T.Mesh(kit.roundedBox(w, 0.12, d, 0.05, 0.03), au);
      m.position.set(x, y, 0); g.add(m);
    }
    function lead(x, y0, h, dx) {
      const pts = [new T.Vector3(x, y0, 0), new T.Vector3(x + dx * 0.15, y0 + h * 0.5, 0.05), new T.Vector3(x + dx, y0 + h, 0.1)];
      g.add(kit.tube(pts, 0.04, cu, 24));
      const ball = new T.Mesh(new T.SphereGeometry(0.075, 16, 12), au);
      ball.position.copy(pts[2]); g.add(ball);
    }
    [-1.62, 1.62].forEach(function (x) {
      via(x, -0.55, -0.12, 0.74); via(x, 0.55, -0.12, 0.74);
      pad(x, 0.74 + 0.06, 0.75, 1.8);
      lead(x, 0.86, 0.75, x < 0 ? -0.25 : 0.25);
    });
    via(0, -0.5, 0.4, 0.74); via(0, 0.5, 0.4, 0.74);
    pad(0, 0.8, 0.8, 1.7);
    lead(0, 0.86, 0.9, 0);
    // body contact
    const bodyPlate = new T.Mesh(kit.roundedBox(W - 0.1, 0.07, D - 0.1, 0.06, 0.03), M.aluminium());
    bodyPlate.position.set(0, -H - 0.05, 0); g.add(bodyPlate);
    lead(2.0, -H - 0.05, 0.5, 0.45);
    // pedestal glow strip
    const strip = new T.Mesh(new T.BoxGeometry(W - 0.1, 0.02, 0.02), M.glow(0x33aaff, 2.4));
    strip.position.set(0, -H - 0.12, D / 2 - 0.05); g.add(strip);

    /* ---------- electrons (instanced) ---------- */
    const NE = 200;
    const eGeo = new T.SphereGeometry(0.03, 8, 6);
    const eMat = new T.MeshBasicMaterial({ color: 0x9fffff, toneMapped: false });
    const elec = new T.InstancedMesh(eGeo, eMat, NE);
    elec.frustumCulled = false;
    g.add(elec);
    const eHalo = new T.InstancedMesh(new T.SphereGeometry(0.06, 8, 6), kit.M.additive(0x2bd7ff, 0.16), NE);
    eHalo.frustumCulled = false; eHalo.renderOrder = 7;
    g.add(eHalo);
    const eZ = new Float32Array(NE), eP = new Float32Array(NE), eS = new Float32Array(NE), eY = new Float32Array(NE);
    for (let i = 0; i < NE; i++) {
      eZ[i] = (rnd() - 0.5) * (D - 0.2);
      eP[i] = rnd();
      eS[i] = 0.7 + rnd() * 0.6;
      eY[i] = rnd();
    }
    // holes in p-substrate
    const NH = 70;
    const holes = new T.InstancedMesh(new T.SphereGeometry(0.04, 8, 6), new T.MeshBasicMaterial({ color: 0xffa24a, transparent: true, opacity: 0.8, toneMapped: false }), NH);
    holes.frustumCulled = false; g.add(holes);
    const hB = [];
    for (let i = 0; i < NH; i++) hB.push([(rnd() - 0.5) * 4.4, -0.55 - rnd() * 0.9, (rnd() - 0.5) * 1.9, rnd() * 6.28]);

    /* ---------- field lines ---------- */
    const fieldMat = kit.M.additive(0xc58cff, 0.35);
    const fields = [];
    [-0.7, -0.25, 0.25, 0.7].forEach(function (z, i) {
      const k = 1 + (i % 2) * 0.0;
      const c = new T.CatmullRomCurve3([
        new T.Vector3(0.95, -0.34, z), new T.Vector3(0.5, -0.62 * k, z), new T.Vector3(0, -0.74, z),
        new T.Vector3(-0.5, -0.62 * k, z), new T.Vector3(-0.95, -0.34, z)]);
      const m = new T.Mesh(new T.TubeGeometry(c, 40, 0.012, 6, false), fieldMat.clone());
      g.add(m); fields.push(m);
    });
    const vMat = kit.M.additive(0x7fe8ff, 0.4);
    const vlines = [];
    for (let i = 0; i < 6; i++) {
      const z = -0.85 + i * 0.34;
      for (let j = 0; j < 3; j++) {
        const x = -0.5 + j * 0.5;
        const m = new T.Mesh(new T.CylinderGeometry(0.008, 0.008, 0.07, 5), vMat);
        m.position.set(x, 0.03, z); g.add(m);
      }
    }

    /* ---------- labels ---------- */
    function lab(s, x, y, z, h, col) {
      const l = kit.label(s, h, { color: col || "#ffffff", size: 96 });
      l.position.set(x, y, z); g.add(l);
    }
    lab("S", -1.62, 1.75, 0.1, 0.34, "#ffd27a");
    lab("G", 0.3, 1.8, 0.1, 0.34, "#ffd27a");
    lab("D", 1.62, 1.75, 0.1, 0.34, "#ffd27a");
    lab("n+", -1.62, -0.25, D / 2 + 0.05, 0.28, "#d8f1ff");
    lab("n+", 1.62, -0.25, D / 2 + 0.05, 0.28, "#d8f1ff");
    lab("p-Si", 0, -1.05, D / 2 + 0.05, 0.26, "#ffe0b5");

    const sh = kit.contactShadow(3.6, 0.7);
    sh.position.y = -H - 0.16; g.add(sh);

    /* ---------- animation ---------- */
    const d = new T.Object3D();
    let lastT = 0;
    function update(t) {
      const dt = Math.min(0.1, Math.max(0, t - lastT)); lastT = t;
      const vg = 0.5 + 0.5 * Math.sin(t * 0.55 - 1.0);   // slow gate-voltage sweep
      const on = Math.min(1, Math.max(0, (vg - 0.12) / 0.3));
      chanMat.opacity = 0.08 + 0.6 * on;
      eMat.color.setRGB(0.6 + 0.4 * on, 1, 1);
      chan.scale.y = 0.5 + 1.5 * on;
      depl.forEach(function (m, i) {
        m.material.opacity = 0.22 + 0.1 * Math.sin(t * 1.3 + i) + 0.08 * (1 - on);
        m.scale.x = 0.42 - 0.08 * on;
      });
      fields.forEach(function (m, i) { m.material.opacity = 0.18 + 0.2 * (0.5 + 0.5 * Math.sin(t * 1.7 - i * 0.7)); });
      const speed = 0.04 + 0.2 * on;
      for (let i = 0; i < NE; i++) {
        eP[i] = (eP[i] + dt * speed * eS[i]) % 1;
        const x = -2.15 + eP[i] * 4.3;
        const inCh = Math.abs(x) < 0.95;
        const y = inCh ? -0.045 - eY[i] * 0.03 : -0.14 - eY[i] * 0.26;
        const share = (i / NE);
        // sources always carry some electrons; channel population follows the gate voltage
        let s = 1;
        if (inCh) s = Math.min(1, Math.max(0, (on - share * 0.9) * 5));
        else s = 0.55 + 0.45 * on;
        d.position.set(x, y, eZ[i]);
        d.scale.setScalar(s);
        d.updateMatrix();
        elec.setMatrixAt(i, d.matrix);
        d.scale.setScalar(s * (0.8 + 0.5 * on));
        d.updateMatrix();
        eHalo.setMatrixAt(i, d.matrix);
      }
      elec.instanceMatrix.needsUpdate = true;
      eHalo.instanceMatrix.needsUpdate = true;
      for (let i = 0; i < NH; i++) {
        const h = hB[i];
        d.position.set(h[0] + 0.03 * Math.sin(t * 0.8 + h[3]), h[1] + 0.02 * Math.cos(t * 0.6 + h[3] * 2), h[2]);
        d.scale.setScalar(1);
        d.updateMatrix();
        holes.setMatrixAt(i, d.matrix);
      }
      holes.instanceMatrix.needsUpdate = true;
      vMat.opacity = 0.08 + 0.5 * on;
      wellMat.emissiveIntensity = 0.6 + 0.3 * on;
    }
    g.rotation.y = -0.38;
    return { group: g, update: update, fit: 2.45, spin: 0.0014 };
  };
})();
