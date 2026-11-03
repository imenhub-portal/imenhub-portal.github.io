/* core_? · Microfabrication & cleanroom processing — DUV stepper exposing a 300 mm wafer */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  IMEN.courseScenes = IMEN.courseScenes || {};
  IMEN.courseScenes.fab = function (kit) {
    const T = THREE, M = kit.M;
    const g = new T.Group();
    const V = function (x, y) { return new T.Vector2(x, y); };

    /* ---------------- materials ---------------- */
    const brushed = M.metal(0x9aa6bc, 0.3);
    const dark = M.metal(0x2b3140, 0.35);
    const alu = M.aluminium();
    const gold = M.gold();
    const amber = M.glow(0xffa22e, 2.6);
    const violetGlow = M.glow(0xb06bff, 3);

    function add(geo, mat, x, y, z, parent) {
      const m = new T.Mesh(geo, mat); m.position.set(x, y, z); (parent || g).add(m); return m;
    }
    function lathe(pts, mat, seg) { return new T.Mesh(new T.LatheGeometry(pts, seg || 48), mat); }

    /* ---------------- fixed base + gantry ---------------- */
    const base = add(kit.roundedBox(4.8, 0.22, 4.2, 0.12, 0.05), dark, 0, -0.9, 0);
    add(kit.roundedBox(4.6, 0.06, 4.0, 0.1, 0.02), brushed, 0, -0.76, 0);
    add(new T.BoxGeometry(4.5, 0.03, 0.04), amber, 0, -0.93, 2.12);       // cleanroom amber accent strip
    // granite-like guide rails
    [-1.7, 1.7].forEach(function (x) { add(new T.BoxGeometry(0.1, 0.07, 3.6), M.metal(0xdfe5ee, 0.18), x, -0.71, 0); });
    [-1.7, 1.7].forEach(function (x) { add(new T.BoxGeometry(3.4, 0.07, 0.1), M.metal(0xdfe5ee, 0.18), 0, -0.71, x * 0.95); });
    // back pillars + bridge
    [-2.05, 2.05].forEach(function (x) {
      add(new T.BoxGeometry(0.36, 3.5, 0.36), brushed, x, 0.9, -1.75);
      add(new T.BoxGeometry(0.38, 0.06, 0.38), gold, x, -0.72, -1.75);
    });
    add(kit.roundedBox(4.7, 0.34, 0.5, 0.06, 0.04), brushed, 0, 2.2, -1.75);
    add(new T.BoxGeometry(0.5, 0.26, 1.9), dark, 0, 2.2, -0.85);          // arm to the lens column
    add(new T.BoxGeometry(4.3, 0.02, 0.02), amber, 0, 2.2, -1.49);

    /* ---------------- moving stage: chuck + wafer ---------------- */
    const stage = new T.Group(); g.add(stage);
    add(new T.CylinderGeometry(1.85, 1.95, 0.12, 64), dark, 0, -0.62, 0, stage);
    // chuck with ring & vacuum grooves
    const chuckTex = kit.canvasTex(512, 512, function (c, W, H) {
      c.fillStyle = "#7d8798"; c.fillRect(0, 0, W, H);
      c.strokeStyle = "#3a4152"; c.lineWidth = 3;
      for (let r = 40; r < 250; r += 28) { c.beginPath(); c.arc(W / 2, H / 2, r, 0, 6.283); c.stroke(); }
      c.lineWidth = 2;
      for (let a = 0; a < 12; a++) { c.beginPath(); c.moveTo(W / 2, H / 2); c.lineTo(W / 2 + Math.cos(a * 0.5236) * 250, H / 2 + Math.sin(a * 0.5236) * 250); c.stroke(); }
    });
    const chuck = lathe([V(0.01, -0.07), V(1.72, -0.07), V(1.78, -0.04), V(1.78, 0.0), V(1.64, 0.0), V(0.01, 0.0)], new T.MeshPhysicalMaterial({ color: 0xffffff, map: chuckTex, metalness: 0.9, roughness: 0.3, clearcoat: 0.3 }));
    chuck.position.y = -0.52; stage.add(chuck);
    add(new T.TorusGeometry(1.72, 0.035, 12, 96), gold, 0, -0.52, 0, stage).rotation.x = Math.PI / 2;
    add(new T.BoxGeometry(2.2, 0.025, 0.05), amber, 0, -0.62, 1.94, stage);
    // wafer (rounded edge, mirror)
    const waferMat = new T.MeshPhysicalMaterial({ color: 0xb4c0da, metalness: 1, roughness: 0.05, clearcoat: 1, clearcoatRoughness: 0.02, envMapIntensity: 1.5 });
    const wafer = lathe([V(0.001, -0.026), V(1.46, -0.026), V(1.5, -0.014), V(1.5, 0.014), V(1.46, 0.026), V(0.001, 0.026)], waferMat, 96);
    wafer.position.y = -0.46; stage.add(wafer);
    const WY = -0.46 + 0.026;   // wafer top surface
    // notch
    const notch = new T.Mesh(new T.CylinderGeometry(0.05, 0.05, 0.06, 14), dark);
    notch.position.set(0, -0.46, 1.5); stage.add(notch);

    /* ---------------- die grid ---------------- */
    const PX = 0.3, PZ = 0.24, dies = [];
    for (let zi = -6; zi <= 6; zi++) {
      const row = [];
      for (let xi = -6; xi <= 6; xi++) {
        const x = xi * PX, z = zi * PZ;
        const cx = Math.max(Math.abs(x) + PX * 0.46, 0), cz = Math.abs(z) + PZ * 0.46;
        if (Math.hypot(cx, cz) < 1.4) row.push([x, z]);
      }
      if (zi % 2) row.reverse();      // serpentine scan
      row.forEach(function (d) { dies.push(d); });
    }
    const ND = dies.length;
    const dieGeo = new T.BoxGeometry(PX * 0.86, 0.012, PZ * 0.86);
    const dieMat = new T.MeshPhysicalMaterial({ color: 0xffffff, metalness: 0.7, roughness: 0.22, clearcoat: 1, clearcoatRoughness: 0.05, envMapIntensity: 1.3 });
    const dm = new T.InstancedMesh(dieGeo, dieMat, ND);
    const dd = new T.Object3D();
    for (let i = 0; i < ND; i++) { dd.position.set(dies[i][0], WY + 0.006, dies[i][1]); dd.updateMatrix(); dm.setMatrixAt(i, dd.matrix); dm.setColorAt(i, new T.Color(0.2, 0.22, 0.3)); }
    stage.add(dm);
    // thin-film colour per die (iridescent), fixed random-ish
    const filmCol = [];
    for (let i = 0; i < ND; i++) {
      const hue = (0.5 + 0.6 * kit.fbm(dies[i][0] * 1.6 + 3, dies[i][1] * 1.6, 3) + (i % 5) * 0.015) % 1;
      filmCol.push(new T.Color().setHSL(hue, 0.78, 0.55));
    }
    // inner die pattern (fine lines) on top of dies for detail
    const patTex = kit.canvasTex(256, 256, function (c, W, H) {
      c.clearRect(0, 0, W, H);
      const r = kit.rng(5);
      c.strokeStyle = "rgba(255,255,255,0.5)"; c.lineWidth = 2;
      for (let i = 0; i < 26; i++) {
        const x = r() * W, y = r() * H, w = 14 + r() * 70, h = 8 + r() * 40;
        c.strokeRect(x, y, Math.min(w, W - x), Math.min(h, H - y));
      }
      c.fillStyle = "rgba(255,255,255,0.35)";
      for (let i = 0; i < 70; i++) c.fillRect(r() * W, r() * H, 4, 4);
    });
    const patMat = new T.MeshBasicMaterial({ map: patTex, transparent: true, opacity: 0.4, depthWrite: false, blending: T.AdditiveBlending, toneMapped: false });
    const patMesh = new T.InstancedMesh(new T.PlaneGeometry(PX * 0.8, PZ * 0.8), patMat, ND);
    const rotX = new T.Euler(-Math.PI / 2, 0, 0);
    for (let i = 0; i < ND; i++) { dd.position.set(dies[i][0], WY + 0.0135, dies[i][1]); dd.rotation.copy(rotX); dd.updateMatrix(); patMesh.setMatrixAt(i, dd.matrix); }
    dd.rotation.set(0, 0, 0);
    stage.add(patMesh);

    /* ---------------- projection lens column ---------------- */
    const col = new T.Group(); g.add(col);
    // nozzle
    add(new T.CylinderGeometry(0.24, 0.5, 0.42, 48), brushed, 0, 0.67, 0, col);
    add(new T.TorusGeometry(0.245, 0.03, 12, 48), gold, 0, 0.46, 0, col).rotation.x = Math.PI / 2;
    add(new T.CylinderGeometry(0.52, 0.52, 0.4, 48), brushed, 0, 1.06, 0, col);
    add(new T.TorusGeometry(0.53, 0.03, 12, 64), gold, 0, 1.28, 0, col).rotation.x = Math.PI / 2;
    // glass barrel (visible optics)
    add(new T.CylinderGeometry(0.58, 0.58, 0.66, 48, 1, true), M.glass(0xbfe0ff, 0.16), 0, 1.6, 0, col).renderOrder = 8;
    [1.35, 1.6, 1.85].forEach(function (y, i) {
      const l = add(new T.SphereGeometry(0.5, 40, 12), M.glass(0xaad0ff, 0.34), 0, y, 0, col);
      l.scale.set(1, 0.13 - i * 0.015, 1);
      l.renderOrder = 8;
      add(new T.TorusGeometry(0.5, 0.025, 8, 48), alu, 0, y, 0, col).rotation.x = Math.PI / 2;
    });
    // four slender struts around glass
    for (let i = 0; i < 4; i++) {
      const a = i * Math.PI / 2 + Math.PI / 4;
      add(new T.CylinderGeometry(0.035, 0.035, 0.7, 10), alu, Math.cos(a) * 0.6, 1.6, Math.sin(a) * 0.6, col);
    }
    add(new T.CylinderGeometry(0.62, 0.62, 0.12, 48), brushed, 0, 1.97, 0, col);
    add(new T.CylinderGeometry(0.62, 0.62, 0.34, 48), dark, 0, 2.2, 0, col);
    add(new T.TorusGeometry(0.625, 0.022, 8, 64), amber, 0, 2.12, 0, col).rotation.x = Math.PI / 2;
    // reticle stage + reticle
    const rst = add(kit.roundedBox(1.5, 0.1, 1.2, 0.08, 0.03), dark, 0, 2.5, 0, col);
    const ret = new T.Group(); ret.position.set(0, 2.57, 0); col.add(ret);
    add(kit.roundedBox(1.0, 0.04, 1.0, 0.04, 0.015), M.glass(0xc8e6ff, 0.3), 0, 0, 0, ret);
    const retPat = kit.canvasTex(512, 512, function (c, W, H) {
      c.clearRect(0, 0, W, H);
      const r = kit.rng(12);
      c.fillStyle = "#9aa3b5";
      for (let i = 0; i < 90; i++) {
        const w = 12 + r() * 90, h = 8 + r() * 50, x = r() * (W - w), y = r() * (H - h);
        c.fillRect(x, y, w, 5); c.fillRect(x, y + h, w, 5);
        if (r() < 0.5) { c.fillRect(x, y, 5, h); c.fillRect(x + w, y, 5, h); }
      }
      for (let i = 0; i < 160; i++) c.fillRect(r() * W, r() * H, 6, 6);
      c.strokeStyle = "#9aa3b5"; c.lineWidth = 8; c.strokeRect(6, 6, W - 12, H - 12);
    });
    const retMat = new T.MeshPhysicalMaterial({ color: 0xffffff, map: retPat, metalness: 0.8, roughness: 0.3, transparent: true, depthWrite: false, alphaTest: 0.05 });
    const retPlane = add(new T.PlaneGeometry(0.9, 0.9), retMat, 0, -0.025, 0, ret);
    retPlane.rotation.x = Math.PI / 2;
    // reticle clamps
    [-1, 1].forEach(function (s) { add(new T.BoxGeometry(0.06, 0.07, 0.8), gold, s * 0.53, 0, 0, ret); });
    // illuminator
    add(new T.CylinderGeometry(0.5, 0.5, 0.18, 48), brushed, 0, 2.74, 0, col);
    add(new T.CylinderGeometry(0.4, 0.4, 0.7, 48), M.metal(0x7f8ba4, 0.28), 0, 3.18, 0, col);
    for (let i = 0; i < 5; i++) add(new T.TorusGeometry(0.405, 0.016, 8, 48), alu, 0, 2.95 + i * 0.12, 0, col).rotation.x = Math.PI / 2;
    add(new T.CylinderGeometry(0.34, 0.4, 0.1, 48), dark, 0, 3.58, 0, col);
    add(new T.CylinderGeometry(0.22, 0.22, 0.03, 32), violetGlow, 0, 3.64, 0, col);

    /* ---------------- beams ---------------- */
    const beamCol = 0x9a6bff;
    const beamMat = kit.M.additive(beamCol, 0.3);
    // collimated illuminator beam & reduction cone (lathe profile, y up)
    const prof = [V(0.001, 3.6), V(0.28, 3.6), V(0.3, 3.1), V(0.3, 2.52), V(0.3, 1.9), V(0.26, 1.5), V(0.16, 1.0), V(0.1, 0.62), V(0.1, 0.02), V(0.001, 0.02)];
    const beam = lathe(prof, beamMat, 40); beam.userData.noFit = true; g.add(beam);
    const beamCore = lathe([V(0.001, 3.6), V(0.1, 3.6), V(0.1, 2.52), V(0.1, 1.9), V(0.09, 1.5), V(0.05, 1.0), V(0.03, 0.62), V(0.03, 0.02), V(0.001, 0.02)], kit.M.additive(0xe4d4ff, 0.35), 24);
    beamCore.userData.noFit = true; g.add(beamCore);
    // exposure field (pattern image projected on the die)
    const expMat = new T.MeshBasicMaterial({ map: patTex, transparent: true, opacity: 0.9, depthWrite: false, blending: T.AdditiveBlending, toneMapped: false, color: 0xcdb6ff });
    const expo = new T.Mesh(new T.PlaneGeometry(PX * 0.9, PZ * 0.9), expMat);
    expo.rotation.x = -Math.PI / 2; expo.position.set(0, WY + 0.03, 0); expo.userData.noFit = true; g.add(expo);
    const expoGlow = kit.glow(0xa070ff, 1.0, 0.8); expoGlow.position.set(0, WY + 0.05, 0); g.add(expoGlow);
    const topGlow = kit.glow(0xa070ff, 1.0, 0.5); topGlow.position.set(0, 3.64, 0); g.add(topGlow);
    // lights
    const vLight = new T.PointLight(0x9a66ff, 1.2, 4.5); vLight.position.set(0, 0.3, 0); g.add(vLight);
    const aLight = new T.PointLight(0xffa030, 1.1, 9); aLight.position.set(-3.2, 0.2, 2.4); g.add(aLight);
    

    const sh = kit.contactShadow(3.8, 0.75); sh.position.y = -1.03; g.add(sh);

    /* ---------------- animation ---------------- */
    const STEP = 0.9, MOVE = 0.35, HOLD = 3.0;
    const cycle = ND * STEP + HOLD;
    const cc = new T.Color(), cur = new T.Color(0.75, 0.55, 1.5), dull = new T.Color(0.32, 0.38, 0.55);
    let lastDone = -2;
    const ease = function (x) { return x * x * (3 - 2 * x); };
    function setStage(i, f) {
      const a = dies[Math.min(i, ND - 1)], b = dies[Math.max(0, Math.min(i - 1, ND - 1))];
      const e = ease(f);
      stage.position.set(-(b[0] + (a[0] - b[0]) * e), 0, -(b[1] + (a[1] - b[1]) * e));
    }
    function update(t) {
      const ph = (t + ND * STEP * 0.4) % cycle;
      const idx = Math.min(ND - 1, Math.floor(ph / STEP));
      const inHold = ph >= ND * STEP;
      const local = inHold ? STEP : ph - idx * STEP;
      const doneCount = inHold ? ND : idx;
      let expOn = 0;
      if (inHold) {
        stage.position.x += (0 - stage.position.x) * 0.06; stage.position.z += (0 - stage.position.z) * 0.06;
      } else {
        const f = Math.min(1, local / MOVE);
        setStage(idx, idx === 0 ? 1 : f);
        expOn = local > MOVE ? Math.sin(Math.min(1, (local - MOVE) / (STEP - MOVE)) * Math.PI) : 0;
      }
      if (doneCount !== lastDone) {
        for (let i = 0; i < ND; i++) {
          if (i < doneCount) dm.setColorAt(i, filmCol[i]);
          else if (i === doneCount && !inHold) dm.setColorAt(i, cur);
          else dm.setColorAt(i, dull);
        }
        dm.instanceColor.needsUpdate = true;
        lastDone = doneCount;
      }
      // reticle scan + beam pulse
      ret.position.x = Math.sin(t * 2.2) * 0.12 * (expOn > 0.05 ? 1 : 0.2);
      const pulse = 0.35 + 0.65 * expOn;
      beamMat.opacity = 0.1 + 0.3 * pulse;
      beamCore.material.opacity = 0.1 + 0.35 * pulse;
      expMat.opacity = 0.15 + 0.85 * pulse;
      expoGlow.material.opacity = 0.2 + 0.7 * expOn;
      expoGlow.scale.setScalar(0.6 + 0.7 * expOn);
      topGlow.material.opacity = 0.3 + 0.35 * pulse;
      vLight.intensity = 0.2 + 1.8 * expOn;
      violetGlow.emissiveIntensity = 1.5 + 2 * pulse;
      aLight.intensity = 1 + 0.15 * Math.sin(t * 0.9);
    }
    update(0);
    g.rotation.y = -0.5;
    return { group: g, update: update, fit: 2.5, spin: 0.0012 };
  };
})();
