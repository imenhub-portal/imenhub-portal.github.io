/* core_3 · Semiconductor Materials & Characterization — Si crystal with defects, XRD + AFM */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  IMEN.courseScenes = IMEN.courseScenes || {};
  IMEN.courseScenes.materials = function (kit) {
    const T = THREE, M = kit.M;
    const g = new T.Group();
    const V3 = function (x, y, z) { return new T.Vector3(x, y, z); };
    const steel = M.metal(0x9aa6bc, 0.28), dark = M.metal(0x2a3040, 0.35), gold = M.gold(), alu = M.aluminium();

    function add(geo, mat, x, y, z, parent) {
      const m = new T.Mesh(geo, mat); m.position.set(x, y, z); (parent || g).add(m); return m;
    }
    function lab(s, x, y, z, h, col) {
      const l = kit.label(s, h, { color: col || "#dfeaff", size: 96 }); l.position.set(x, y, z); g.add(l); return l;
    }

    /* ---------------- diamond-cubic lattice ---------------- */
    const N = 3, Q = 4 * N, A = 0.8, q = A / 4, off = -N * A / 2;
    const key = function (x, y, z) { return (x * 32 + y) * 32 + z; };
    const sites = [], idx = {};
    for (let x = 0; x <= Q; x++) for (let y = 0; y <= Q; y++) for (let z = 0; z <= Q; z++) {
      const ev = x % 2 === 0 && y % 2 === 0 && z % 2 === 0 && (x + y + z) % 4 === 0;
      const od = x % 2 === 1 && y % 2 === 1 && z % 2 === 1 && (x + y + z) % 4 === 3;
      if (!(ev || od)) continue;
      idx[key(x, y, z)] = sites.length;
      sites.push({ q: [x, y, z], p: V3(x * q + off, y * q + off, z * q + off), kind: 0 });
    }
    // defects
    const vacK = key(4, 6, 10), dopK = key(7, 7, 9);
    const vac = sites[idx[vacK]], dop = sites[idx[dopK]];
    const vacPos = vac.p.clone(), dopPos = dop.p.clone();
    vac.kind = -1; dop.kind = 1;
    const interPos = V3(6 * q + off, 6 * q + off, 6 * q + off);   // tetrahedral interstitial
    // lattice relaxation around defects
    const tmp = V3(0, 0, 0);
    sites.forEach(function (s) {
      s.r = s.p.clone();
      tmp.copy(vacPos).sub(s.p); let d = tmp.length();
      if (d > 0.01 && d < 1.3) s.r.addScaledVector(tmp, 0.1 * Math.exp(-d * d / 0.35) / d);
      tmp.copy(s.p).sub(dopPos); d = tmp.length();
      if (d > 0.01 && d < 1.3) s.r.addScaledVector(tmp, 0.09 * Math.exp(-d * d / 0.3) / d);
      tmp.copy(s.p).sub(interPos); d = tmp.length();
      if (d > 0.01 && d < 1.3) s.r.addScaledVector(tmp, 0.12 * Math.exp(-d * d / 0.3) / d);
    });
    // build instance lists
    const atoms = sites.filter(function (s) { return s.kind >= 0; });
    const bonds = [];
    sites.forEach(function (s, i) {
      if (s.kind < 0) return;
      [[1, 1, 1], [1, -1, -1], [-1, 1, -1], [-1, -1, 1], [-1, -1, -1], [-1, 1, 1], [1, -1, 1], [1, 1, -1]].forEach(function (n) {
        const x = s.q[0] + n[0], y = s.q[1] + n[1], z = s.q[2] + n[2];
        if (x < 0 || y < 0 || z < 0 || x > Q || y > Q || z > Q) return;
        const j = idx[key(x, y, z)];
        if (j == null || j < i || sites[j].kind < 0) return;
        bonds.push([s, sites[j]]);
      });
    });
    const latt = new T.Group(); g.add(latt);
    const depthCol = function (p, c, base) {
      const k = 0.4 + 0.6 * Math.min(1, Math.max(0, (p.z + 1.3) / 2.6));
      const h = 0.7 + 0.3 * Math.min(1, Math.max(0, (p.y + 1.3) / 2.6));
      return c.copy(base).multiplyScalar(k * h);
    };
    const aMat = new T.MeshPhysicalMaterial({ color: 0xffffff, metalness: 0.35, roughness: 0.2, clearcoat: 1, clearcoatRoughness: 0.08, envMapIntensity: 1.2 });
    const aGeo = new T.SphereGeometry(0.092, 14, 10);
    const aInst = new T.InstancedMesh(aGeo, aMat, atoms.length);
    const d = new T.Object3D(), col = new T.Color(), base = new T.Color(0x5f8cff);
    atoms.forEach(function (s, i) {
      d.position.copy(s.r); d.scale.setScalar(s.kind === 1 ? 1.45 : 1); d.updateMatrix();
      aInst.setMatrixAt(i, d.matrix);
      aInst.setColorAt(i, s.kind === 1 ? col.set(0xffb040) : depthCol(s.p, col, base));
    });
    latt.add(aInst);
    const bGeo = new T.CylinderGeometry(0.028, 0.028, 1, 8, 1, true);
    const bMat = new T.MeshPhysicalMaterial({ color: 0xffffff, metalness: 0.6, roughness: 0.3, envMapIntensity: 1 });
    const bInst = new T.InstancedMesh(bGeo, bMat, bonds.length);
    const up = V3(0, 1, 0), dir = V3(0, 0, 0);
    const bc = new T.Color(0x7f9ad0);
    bonds.forEach(function (b, i) {
      dir.copy(b[1].r).sub(b[0].r);
      const len = dir.length();
      d.position.copy(b[0].r).addScaledVector(dir, 0.5);
      d.quaternion.setFromUnitVectors(up, dir.normalize());
      d.scale.set(1, len, 1); d.updateMatrix();
      bInst.setMatrixAt(i, d.matrix);
      bInst.setColorAt(i, depthCol(d.position, col, bc));
    });
    d.quaternion.identity();
    latt.add(bInst);
    // dopant halo, vacancy ghost, interstitial, donor electron
    const dopHalo = kit.glow(0xffa030, 1.5, 0.9); dopHalo.position.copy(dopPos); latt.add(dopHalo);
    const vacGhost = new T.Mesh(new T.IcosahedronGeometry(0.2, 1), new T.MeshBasicMaterial({ color: 0xff5a6e, wireframe: true, transparent: true, opacity: 0.9, toneMapped: false }));
    vacGhost.position.copy(vacPos); latt.add(vacGhost);
    const vacGlow = kit.glow(0xff4060, 1.2, 0.7); vacGlow.position.copy(vacPos); latt.add(vacGlow);
    const inter = add(new T.SphereGeometry(0.07, 14, 10), new T.MeshPhysicalMaterial({ color: 0xff6a4a, emissive: 0xaa2210, emissiveIntensity: 0.9, roughness: 0.2, clearcoat: 1 }), interPos.x, interPos.y, interPos.z, latt);
    const donor = add(new T.SphereGeometry(0.035, 8, 6), new T.MeshBasicMaterial({ color: 0x8ff6ff, toneMapped: false }), 0, 0, 0, latt);
    const donorGlow = kit.glow(0x4fe6ff, 0.28, 0.9); latt.add(donorGlow);
    // crystal volume cage
    const cage = new T.LineSegments(new T.EdgesGeometry(new T.BoxGeometry(N * A, N * A, N * A)), new T.LineBasicMaterial({ color: 0x6fc0ff, transparent: true, opacity: 0.55 }));
    latt.add(cage);
    const CY = N * A / 2;         // half-size of crystal = 1.2
    latt.position.y = 0;

    /* ---------------- stage / goniometer ---------------- */
    add(kit.roundedBox(8.8, 0.2, 3.4, 0.14, 0.05), dark, 0.2, -2.05, 0);
    add(kit.roundedBox(8.6, 0.05, 3.2, 0.12, 0.02), steel, 0.2, -1.92, 0);
    // pedestal with degree ring
    const ringTex = kit.canvasTex(1024, 64, function (c, W, H) {
      c.fillStyle = "#c8d0de"; c.fillRect(0, 0, W, H);
      c.fillStyle = "#1b2030";
      for (let i = 0; i < 360; i++) { const x = i / 360 * W, h = i % 10 === 0 ? 30 : i % 5 === 0 ? 20 : 10; c.fillRect(x, 0, 2, h); }
      c.font = "bold 18px sans-serif"; c.textAlign = "center";
      for (let i = 0; i < 360; i += 30) c.fillText(String(i), i / 360 * W, 52);
    });
    add(new T.CylinderGeometry(0.95, 1.05, 0.22, 64), new T.MeshPhysicalMaterial({ color: 0xffffff, map: ringTex, metalness: 0.9, roughness: 0.3 }), 0, -1.8, 0);
    add(new T.CylinderGeometry(0.55, 0.8, 0.4, 48), steel, 0, -1.5, 0);
    add(new T.CylinderGeometry(0.85, 0.85, 0.08, 48), gold, 0, -1.28, 0);
    add(new T.CylinderGeometry(0.88, 0.88, 0.04, 48), M.glow(0x3a9bff, 2), 0, -1.22, 0);

    /* ---------------- X-ray source (left) ---------------- */
    const xs = new T.Group(); g.add(xs);
    const cylX = function (r1, r2, len, mat, x, y, z) {
      const m = add(new T.CylinderGeometry(r1, r2, len, 48), mat, x, y, z, xs); m.rotation.z = -Math.PI / 2; return m;
    };
    cylX(0.52, 0.52, 1.4, steel, -4.5, 0, 0);
    cylX(0.55, 0.55, 0.1, gold, -3.72, 0, 0);
    cylX(0.4, 0.5, 0.5, dark, -3.4, 0, 0);
    cylX(0.18, 0.18, 0.9, alu, -2.75, 0, 0);
    cylX(0.24, 0.24, 0.06, gold, -2.3, 0, 0);
    cylX(0.16, 0.16, 0.3, M.glass(0xaadcff, 0.45), -2.2, 0, 0).renderOrder = 8;
    // cooling ribs
    for (let i = 0; i < 6; i++) cylX(0.545, 0.545, 0.035, alu, -5.0 + i * 0.16, 0, 0);
    add(new T.CylinderGeometry(0.12, 0.12, 0.05, 24), M.glow(0x4fb4ff, 2.2), -5.22, 0, 0, xs).rotation.z = Math.PI / 2;
    // support
    add(kit.roundedBox(0.5, 1.5, 0.7, 0.06, 0.03), dark, -4.2, -1.1, 0, xs);
    add(kit.roundedBox(1.0, 0.1, 1.0, 0.08, 0.03), steel, -4.2, -1.88, 0, xs);

    /* ---------------- detector (right) ---------------- */
    const DX = 4.0, DS = 2.5;
    const det = new T.Group(); det.position.set(DX, 0.1, 0); g.add(det);
    add(new T.BoxGeometry(0.3, DS + 0.35, DS + 0.35), dark, 0.0, 0, 0, det);
    add(new T.BoxGeometry(0.05, DS + 0.45, 0.08), steel, -0.16, 0, (DS + 0.35) / 2, det);
    add(new T.BoxGeometry(0.05, DS + 0.45, 0.08), steel, -0.16, 0, -(DS + 0.35) / 2, det);
    add(new T.BoxGeometry(0.05, 0.08, DS + 0.35), steel, -0.16, (DS + 0.35) / 2, 0, det);
    add(new T.BoxGeometry(0.05, 0.08, DS + 0.35), steel, -0.16, -(DS + 0.35) / 2, 0, det);
    add(kit.roundedBox(0.2, 0.2, 0.2, 0.05, 0.02), steel, 0.5, -1.6, 0, det);
    add(new T.BoxGeometry(0.9, 1.0, 0.7), dark, 0.55, -1.5, 0, det);
    const spotTex = kit.canvasTex(512, 512, function (c, W, H) {
      c.fillStyle = "#020510"; c.fillRect(0, 0, W, H);
      const r = kit.rng(31);
      // powder (Debye-Scherrer) rings
      [58, 100, 142, 186, 224].forEach(function (R, i) {
        c.strokeStyle = "rgba(120,190,255," + (0.55 - i * 0.07) + ")"; c.lineWidth = 3 - i * 0.3;
        c.beginPath(); c.arc(W / 2, H / 2, R, 0, 6.283); c.stroke();
      });
      // Laue spots
      for (let i = -4; i <= 4; i++) for (let j = -4; j <= 4; j++) {
        if (!i && !j) continue;
        const x = W / 2 + i * 52 + (r() - 0.5) * 3, y = H / 2 + j * 52 + (r() - 0.5) * 3;
        if (Math.hypot(x - W / 2, y - H / 2) > 238) continue;
        const rad = 5 + r() * 8;
        const gr = c.createRadialGradient(x, y, 0, x, y, rad * 2.2);
        gr.addColorStop(0, "rgba(255,255,255,1)"); gr.addColorStop(0.35, "rgba(120,210,255,0.8)"); gr.addColorStop(1, "rgba(40,100,255,0)");
        c.fillStyle = gr; c.beginPath(); c.arc(x, y, rad * 2.2, 0, 6.283); c.fill();
      }
      // central transmitted-beam spot
      const gg = c.createRadialGradient(W / 2, H / 2, 0, W / 2, H / 2, 26);
      gg.addColorStop(0, "rgba(255,255,255,1)"); gg.addColorStop(1, "rgba(80,160,255,0)");
      c.fillStyle = gg; c.fillRect(W / 2 - 30, H / 2 - 30, 60, 60);
    });
    const detMat = new T.MeshStandardMaterial({ color: 0x151c33, map: spotTex, emissive: 0xffffff, emissiveMap: spotTex, emissiveIntensity: 1, roughness: 0.4, metalness: 0.3 });
    const face = add(new T.PlaneGeometry(DS, DS), detMat, -0.16, 0, 0, det);
    face.rotation.y = -Math.PI / 2;
    // beam stop
    const stop = add(new T.CylinderGeometry(0.06, 0.06, 0.4, 20), gold, -0.3, 0, 0, det); stop.rotation.z = Math.PI / 2;
    add(new T.CylinderGeometry(0.012, 0.012, 0.9, 6), alu, -0.3, 0.45, 0, det);
    // glass cover
    add(new T.BoxGeometry(0.02, DS, DS), M.glass(0xaad0ff, 0.12), -0.185, 0, 0, det).renderOrder = 9;
    // bezel LEDs
    add(new T.BoxGeometry(0.03, 0.05, 0.5), M.glow(0x40ffa0, 2), -0.15, -DS / 2 - 0.1, 0.6, det);

    /* ---------------- beams ---------------- */
    const inMat = kit.M.additive(0x7ad0ff, 0.5), inCore = kit.M.additive(0xffffff, 0.8);
    const BX0 = -2.05, BX1 = 0;
    const inBeam = add(new T.CylinderGeometry(0.075, 0.075, BX1 - BX0, 24, 1, true), inMat, (BX0 + BX1) / 2, 0, 0); inBeam.rotation.z = Math.PI / 2; inBeam.userData.noFit = true;
    const inB2 = add(new T.CylinderGeometry(0.022, 0.022, BX1 - BX0, 12, 1, true), inCore, (BX0 + BX1) / 2, 0, 0); inB2.rotation.z = Math.PI / 2; inB2.userData.noFit = true;
    // transmitted + diffracted rays
    const rayMat = kit.M.additive(0x7fc6ff, 0.4), rayCore = new T.MeshBasicMaterial({ color: 0xcfeaff, transparent: true, opacity: 0.5, blending: T.AdditiveBlending, depthWrite: false, toneMapped: false });
    const rayG = new T.CylinderGeometry(0.012, 0.012, 1, 6, 1, true);
    const spots = [[0, 0], [0.52, 0.52], [-0.52, 0.52], [0.52, -0.52], [-0.52, -0.52], [1.04, 0], [-1.04, 0], [0, 1.04], [0, -1.04], [1.04, 0.52], [-1.04, -0.52], [0.52, 1.04]];
    const rays = [];
    spots.forEach(function (sp, i) {
      const tgt = V3(DX - 0.2, 0.1 + sp[1], sp[0]);
      const len = tgt.length();
      const m = new T.Mesh(rayG, i === 0 ? rayCore : kit.M.additive(0x7fc6ff, 0.3));
      m.position.copy(tgt).multiplyScalar(0.5);
      m.quaternion.setFromUnitVectors(up, tgt.clone().normalize());
      m.scale.set(i === 0 ? 3 : 1, len, i === 0 ? 3 : 1);
      m.userData.noFit = true; g.add(m); rays.push(m);
      if (i) {
        const spot = kit.glow(0x6ac8ff, 0.2, 0.0); spot.position.set(DX - 0.2, 0.1 + sp[1], sp[0]); g.add(spot); rays[rays.length - 1].userData.spot = spot;
      }
    });
    const crystalGlow = kit.glow(0x7ad0ff, 2.4, 0.5); g.add(crystalGlow);
    // travelling photons
    const NP = 10;
    const ph = new T.InstancedMesh(new T.SphereGeometry(0.04, 8, 6), new T.MeshBasicMaterial({ color: 0xffffff, toneMapped: false }), NP);
    ph.frustumCulled = false; ph.userData.noFit = true; g.add(ph);

    /* ---------------- AFM / SEM-style probe (above) ---------------- */
    const probe = new T.Group(); g.add(probe);
    const carriage = add(kit.roundedBox(0.55, 0.3, 0.5, 0.07, 0.03), dark, 0, 2.15, 0, probe);
    add(new T.CylinderGeometry(0.08, 0.08, 0.5, 16), steel, 0, 1.8, 0, probe);
    const holder = add(kit.roundedBox(0.32, 0.12, 0.22, 0.04, 0.02), steel, 0, 1.52, 0, probe);
    const cant = add(new T.BoxGeometry(0.014, 0.01, 0.5), M.metal(0xe6ecf5, 0.15), 0, 1.45, 0.3, probe);
    const tip = add(new T.ConeGeometry(0.03, 0.22, 12), M.metal(0xf0f4ff, 0.1), 0, 1.33, 0.52, probe); tip.rotation.x = Math.PI;
    const laser = kit.glow(0xff3050, 0.22, 1); laser.position.set(0, 1.45, 0.52); probe.add(laser);
    const tipGlow = kit.glow(0x7fe0ff, 0.35, 0.8); tipGlow.position.set(0, 1.21, 0.52); probe.add(tipGlow);
    const boom = add(new T.BoxGeometry(0.22, 0.2, 1), dark, 0, 2.3, -1.2); // spans back to the post
    add(kit.roundedBox(0.5, 3.9, 0.5, 0.06, 0.03), steel, 0.0, 0.15, -2.1);
    // scan trace (top surface of crystal)
    const trail = new T.Mesh(new T.PlaneGeometry(1, 0.012), kit.M.additive(0x7fe0ff, 0.5)); trail.rotation.x = -Math.PI / 2;
    trail.userData.noFit = true; trail.position.y = CY + 0.004; g.add(trail);

    /* ---------------- labels, axes ---------------- */
    function axis(dx, dy, dz, color, txt) {
      const L = 0.7, grp = new T.Group();
      const m = new T.MeshBasicMaterial({ color: color, toneMapped: false });
      const c = new T.Mesh(new T.CylinderGeometry(0.014, 0.014, L, 8), m); c.position.y = L / 2; grp.add(c);
      const h = new T.Mesh(new T.ConeGeometry(0.05, 0.14, 12), m); h.position.y = L + 0.05; grp.add(h);
      grp.quaternion.setFromUnitVectors(up, V3(dx, dy, dz));
      grp.position.set(-2.1, -1.8, 1.35); g.add(grp);
      const l = lab(txt, -2.1 + dx * (L + 0.28), -1.8 + dy * (L + 0.28), 1.35 + dz * (L + 0.28), 0.2, "#fff"); l.material.color.set(color);
    }
    axis(1, 0, 0, 0xff6a6a, "x"); axis(0, 1, 0, 0x6aff9a, "y"); axis(0, 0, 1, 0x6aa8ff, "z");
    lab("X-ray source", -4.4, 0.95, 0.3, 0.26);
    lab("detector", 4.1, 1.8, 0.3, 0.26);
    lab("AFM tip", 1.35, 2.0, 0.55, 0.24);
    lab("vacancy", vacPos.x - 0.35, vacPos.y - 1.55, 1.5, 0.001).visible = false;
    const lv = lab("V", vacPos.x - 0.3, vacPos.y + 0.28, vacPos.z + 0.7, 0.22, "#ff8a98");
    const lp = lab("P", dopPos.x + 0.3, dopPos.y + 0.3, dopPos.z + 0.65, 0.22, "#ffc870");
    const sh = kit.contactShadow(5, 0.7); sh.position.set(0, -2.17, 0); sh.scale.set(1.3, 1, 0.8); g.add(sh);

    /* ---------------- animation ---------------- */
    const dm = new T.Object3D();
    function update(t) {
      const p = Math.pow(0.5 + 0.5 * Math.sin(t * 2.4), 1.5);            // beam pulse
      inMat.opacity = 0.2 + 0.4 * p; inCore.opacity = 0.35 + 0.55 * p;
      rays.forEach(function (m, i) {
        m.material.opacity = (i ? 0.1 : 0.3) + (i ? 0.35 : 0.4) * p;
        if (m.userData.spot) { m.userData.spot.material.opacity = 0.0; }
      });
      detMat.emissiveIntensity = 0.5 + 1.5 * p;
      crystalGlow.material.opacity = 0.15 + 0.45 * p;
      for (let i = 0; i < NP; i++) {
        const u = ((t * 0.45 + i / NP) % 1);
        dm.position.set(BX0 + u * (BX1 - BX0), 0, 0); dm.scale.setScalar(0.7 + 0.5 * p); dm.updateMatrix(); ph.setMatrixAt(i, dm.matrix);
      }
      ph.instanceMatrix.needsUpdate = true;
      // defects
      vacGhost.rotation.y = t * 0.6; vacGhost.rotation.x = t * 0.4;
      vacGhost.material.opacity = 0.55 + 0.4 * Math.sin(t * 3);
      vacGlow.material.opacity = 0.3 + 0.2 * Math.sin(t * 3);
      dopHalo.material.opacity = 0.55 + 0.35 * Math.sin(t * 2 + 1);
      const a = t * 2.6;
      donor.position.set(dopPos.x + Math.cos(a) * 0.24, dopPos.y + Math.sin(a * 0.7) * 0.12, dopPos.z + Math.sin(a) * 0.24);
      donorGlow.position.copy(donor.position);
      inter.scale.setScalar(1 + 0.12 * Math.sin(t * 4));
      // AFM raster scan
      const sx = (Math.abs(((t * 0.22) % 2) - 1) * 2 - 1) * 0.9;          // fast line, triangle wave
      const sz = Math.sin(t * 0.07 * 3.1) * 0.8;                         // slow frame axis
      probe.position.set(sx, 0, sz);
      boom.scale.z = 1.0 + 0; boom.position.z = (sz + -2.1) / 2 + 0.05; boom.scale.z = Math.abs(sz + 2.1);
      boom.position.x = sx;
      laser.material.opacity = 0.7 + 0.3 * Math.sin(t * 9);
      trail.position.x = sx; trail.position.z = sz + 0.52;
      trail.scale.x = 0.0001 + 0.5; trail.material.opacity = 0.35;
      trail.scale.x = 0.9;
    }
    g.rotation.y = 0.72;
    update(0);
    return { group: g, update: update, fit: 2.75, spin: 0.001 };
  };
})();
