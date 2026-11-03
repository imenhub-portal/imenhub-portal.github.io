/* MEMS & Sensors — comb-drive accelerometer / resonator in an etched silicon cavity */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  IMEN.courseScenes = IMEN.courseScenes || {};
  IMEN.courseScenes.mems = function (kit) {
    const T = THREE;
    const g = new T.Group();
    const root = new T.Group();
    root.rotation.x = 0.52;
    g.add(root);

    const SILC = 0x8fa3d0;
    const matBody = kit.M.silicon(0x0b1228);
    const matFloor = null;
    const matRim = kit.M.silicon(0x4a5a8c); matRim.metalness = 1; matRim.roughness = 0.14;
    const matStruct = kit.M.silicon(SILC);
    matStruct.metalness = 1;
    matStruct.roughness = 0.13;
    matStruct.color.set(0xb4c4e8);
    const matStructDark = kit.M.silicon(0x8196c8); matStructDark.metalness = 1; matStructDark.roughness = 0.16;

    /* ---------- die + cavity ---------- */
    const die = new T.Mesh(kit.roundedBox(8.6, 0.4, 5.8, 0.18, 0.03), matBody);
    root.add(die); // y: -0.2 .. 0.2
    const TOP = 0.2;

    function rectPath(w, d, r) {
      const s = new T.Shape();
      const x = -w / 2, y = -d / 2;
      s.moveTo(x + r, y); s.lineTo(x + w - r, y); s.quadraticCurveTo(x + w, y, x + w, y + r);
      s.lineTo(x + w, y + d - r); s.quadraticCurveTo(x + w, y + d, x + w - r, y + d);
      s.lineTo(x + r, y + d); s.quadraticCurveTo(x, y + d, x, y + d - r);
      s.lineTo(x, y + r); s.quadraticCurveTo(x, y, x + r, y);
      return s;
    }
    // raised rim (etched cavity walls)
    const rimShape = rectPath(8.2, 5.4, 0.14);
    const hole = rectPath(7.2, 4.2, 0.08);
    rimShape.holes.push(hole);
    const rimGeo = new T.ExtrudeGeometry(rimShape, { depth: 0.14, bevelEnabled: true, bevelThickness: 0.02, bevelSize: 0.02, bevelSegments: 2, curveSegments: 5 });
    rimGeo.rotateX(-Math.PI / 2);
    const rim = new T.Mesh(rimGeo, matRim);
    rim.position.y = TOP;
    root.add(rim);

    // cavity floor: etched pattern
    const floorTex = kit.canvasTex(1024, 600, function (c, W, H) {
      const r = kit.rng(9);
      c.fillStyle = "#070d22";
      c.fillRect(0, 0, W, H);
      const img = c.getImageData(0, 0, W, H);
      for (let i = 0; i < W * H; i++) {
        const v = (r() - 0.5) * 18;
        img.data[i * 4] += v; img.data[i * 4 + 1] += v; img.data[i * 4 + 2] += v * 1.3;
      }
      c.putImageData(img, 0, 0);
      c.strokeStyle = "rgba(120,160,255,0.10)";
      c.lineWidth = 1;
      for (let x = 0; x <= W; x += 16) { c.beginPath(); c.moveTo(x, 0); c.lineTo(x, H); c.stroke(); }
      for (let y = 0; y <= H; y += 16) { c.beginPath(); c.moveTo(0, y); c.lineTo(W, y); c.stroke(); }
      c.strokeStyle = "rgba(140,190,255,0.22)";
      for (let x = 0; x <= W; x += 128) { c.beginPath(); c.moveTo(x, 0); c.lineTo(x, H); c.stroke(); }
      for (let y = 0; y <= H; y += 128) { c.beginPath(); c.moveTo(0, y); c.lineTo(W, y); c.stroke(); }
      // alignment crosses
      c.strokeStyle = "rgba(255,200,110,0.55)";
      c.lineWidth = 2;
      [[40, 40], [W - 40, 40], [40, H - 40], [W - 40, H - 40]].forEach(function (p) {
        c.beginPath(); c.moveTo(p[0] - 14, p[1]); c.lineTo(p[0] + 14, p[1]); c.moveTo(p[0], p[1] - 14); c.lineTo(p[0], p[1] + 14); c.stroke();
      });
    });
    const floor = new T.Mesh(new T.PlaneGeometry(7.2, 4.2), new T.MeshPhysicalMaterial({ map: floorTex, metalness: 0.6, roughness: 0.28, clearcoat: 0.6, clearcoatRoughness: 0.2, envMapIntensity: 0.9 }));
    floor.rotation.x = -Math.PI / 2;
    floor.position.y = TOP + 0.003;
    root.add(floor);

    /* ---------- glass lid (L-shaped, corner cut away) ---------- */
    const lidS = new T.Shape();
    const pts = [[-4.1, -2.7], [3.6, -2.7], [4.1, -2.2], [4.1, -2.1], [-3.6, -2.1], [-3.6, 2.7], [-4.1, 2.7]];
    // shape y maps to -z after rotateX(-90deg) => feed -z
    lidS.moveTo(pts[0][0], -pts[0][1]);
    for (let i = 1; i < pts.length; i++) lidS.lineTo(pts[i][0], -pts[i][1]);
    const lidGeo = new T.ExtrudeGeometry(lidS, { depth: 0.07, bevelEnabled: true, bevelThickness: 0.015, bevelSize: 0.015, bevelSegments: 2 });
    lidGeo.rotateX(-Math.PI / 2);
    const lid = new T.Mesh(lidGeo, kit.M.glass(0xb8d8ff, 0.3));
    lid.position.y = TOP + 0.16;
    lid.renderOrder = 5;
    root.add(lid);

    /* ---------- structure layer ---------- */
    const Y0 = TOP + 0.12;   // suspended layer base
    const H = 0.14;
    const sway = new T.Group();
    root.add(sway);

    // proof mass with release holes
    const MW = 2.6, MD = 1.5;
    const ms = rectPath(MW, MD, 0.06);
    for (let ix = 0; ix < 12; ix++) {
      for (let iz = 0; iz < 6; iz++) {
        const cx = (ix - 5.5) * 0.2, cz = (iz - 2.5) * 0.2;
        const h = new T.Path();
        const s = 0.045;
        h.moveTo(cx - s, cz - s); h.lineTo(cx - s, cz + s); h.lineTo(cx + s, cz + s); h.lineTo(cx + s, cz - s); h.lineTo(cx - s, cz - s);
        ms.holes.push(h);
      }
    }
    const massGeo = new T.ExtrudeGeometry(ms, { depth: H - 0.024, bevelEnabled: true, bevelThickness: 0.012, bevelSize: 0.01, bevelSegments: 1, curveSegments: 4 });
    massGeo.rotateX(-Math.PI / 2);
    const massGroup = new T.Group();
    massGroup.position.y = Y0;
    const mass = new T.Mesh(massGeo, matStruct);
    massGroup.add(mass);
    root.add(massGroup);

    // rotor fingers (instanced, ride with the mass)
    const PITCH = 0.26, FW = 0.05;
    const rotorGeo = new T.BoxGeometry(FW, H, 0.64);
    const rotor = kit.instanced(rotorGeo, matStruct, 20, function (i, d) {
      const side = i < 10 ? 1 : -1, k = i % 10;
      d.position.set((k - 4.5) * PITCH, H / 2, side * 1.0);
    });
    massGroup.add(rotor);

    // stator bars + fingers (fixed to substrate)
    const statH = Y0 + H - TOP;
    [1, -1].forEach(function (sz) {
      const bar = new T.Mesh(kit.roundedBox(3.0, statH, 0.22, 0.03, 0.012), matStructDark);
      bar.position.set(0, TOP + statH / 2, sz * 1.62);
      root.add(bar);
      const pad = new T.Mesh(kit.roundedBox(0.34, 0.03, 0.16, 0.02, 0.008), kit.M.gold());
      pad.position.set(0, TOP + statH + 0.014, sz * 1.62);
      root.add(pad);
    });
    const stGeo = new T.BoxGeometry(FW, statH, 0.58);
    const stator = kit.instanced(stGeo, matStructDark, 22, function (i, d) {
      const sz = i < 11 ? 1 : -1, j = i % 11;
      d.position.set((j - 5) * PITCH, TOP + statH / 2, sz * 1.235);
    });
    root.add(stator);

    // capacitive gap glow strips
    const gapGeo = new T.BoxGeometry(0.06, 0.12, 0.36);
    const matGapA = kit.M.additive(0x35e6ff, 0.6);
    const matGapB = kit.M.additive(0xff4fd8, 0.6);
    const gapA = kit.instanced(gapGeo, matGapA, 20, function (i, d) {
      const side = i < 10 ? 1 : -1, k = i % 10;
      d.position.set((k - 4.5) * PITCH + 0.065, Y0 + H / 2, side * 1.125);
    });
    const gapB = kit.instanced(gapGeo, matGapB, 20, function (i, d) {
      const side = i < 10 ? 1 : -1, k = i % 10;
      d.position.set((k - 4.5) * PITCH - 0.065, Y0 + H / 2, side * 1.125);
    });
    root.add(gapA, gapB);

    /* ---------- serpentine springs (deformable) ---------- */
    function filleted(P, rmax, seg) {
      const out = [P[0].clone()];
      for (let i = 1; i < P.length - 1; i++) {
        const a = P[i - 1], b = P[i], c = P[i + 1];
        const l1 = a.distanceTo(b), l2 = b.distanceTo(c);
        const r = Math.min(rmax, l1 * 0.5, l2 * 0.5);
        const p1 = b.clone().add(a.clone().sub(b).normalize().multiplyScalar(r));
        const p2 = b.clone().add(c.clone().sub(b).normalize().multiplyScalar(r));
        for (let k = 0; k <= seg; k++) {
          const t = k / seg, u = 1 - t;
          out.push(new T.Vector2(u * u * p1.x + 2 * u * t * b.x + t * t * p2.x, u * u * p1.y + 2 * u * t * b.y + t * t * p2.y));
        }
      }
      out.push(P[P.length - 1].clone());
      return out;
    }
    const XA = 2.62, XM = 1.30, SW = 0.058;
    const springs = [];
    function makeSpring(sx, sz) {
      const zc = sz * 0.4;
      const P = [new T.Vector2(0, -0.3)];
      const us = [];
      for (let k = 0; k < 6; k++) us.push(0.2 + k * 0.2);
      us.forEach(function (u, k) {
        const a = k % 2 === 0 ? -0.3 : 0.3, b = -a;
        P.push(new T.Vector2(u, a), new T.Vector2(u, b));
      });
      P.push(new T.Vector2(XA - XM, -0.3));
      const D = filleted(P, 0.1, 6);
      const N = D.length;
      // arclength
      const s = new Float32Array(N);
      for (let i = 1; i < N; i++) s[i] = s[i - 1] + D[i].distanceTo(D[i - 1]);
      const tot = s[N - 1];
      for (let i = 0; i < N; i++) s[i] /= tot;
      const pos = new Float32Array(N * 6 * 3), nor = new Float32Array(N * 6 * 3), idx = [];
      const base = new Float32Array(N * 6), par = new Float32Array(N * 6);
      for (let i = 0; i < N; i++) {
        const a = D[Math.max(0, i - 1)], b = D[Math.min(N - 1, i + 1)];
        let tx = b.x - a.x, tz = b.y - a.y;
        const l = Math.hypot(tx, tz) || 1; tx /= l; tz /= l;
        // path -> world
        const nxl = -tz, nzl = tx; // local normal
        const wx = (sx * (XA - D[i].x)), wz = (zc - sz * D[i].y);
        const nwx = -sx * nxl, nwz = -sz * nzl;
        const hw = SW / 2;
        const lx = wx + nwx * hw, lz = wz + nwz * hw, rx = wx - nwx * hw, rz = wz - nwz * hw;
        // verts: 0 topL,1 topR,2 L top,3 L bot,4 R top,5 R bot
        const V = [[lx, H, lz], [rx, H, rz], [lx, H, lz], [lx, 0, lz], [rx, H, rz], [rx, 0, rz]];
        const Nn = [[0, 1, 0], [0, 1, 0], [nwx, 0, nwz], [nwx, 0, nwz], [-nwx, 0, -nwz], [-nwx, 0, -nwz]];
        for (let k = 0; k < 6; k++) {
          const o = (i * 6 + k) * 3;
          pos[o] = V[k][0]; pos[o + 1] = V[k][1]; pos[o + 2] = V[k][2];
          nor[o] = Nn[k][0]; nor[o + 1] = Nn[k][1]; nor[o + 2] = Nn[k][2];
          base[i * 6 + k] = V[k][0]; par[i * 6 + k] = s[i];
        }
      }
      function quad(a, b, c, d, want) {
        // ensure triangle (a,b,c) faces `want`
        const ax = pos[a * 3], ay = pos[a * 3 + 1], az = pos[a * 3 + 2];
        const e1x = pos[b * 3] - ax, e1y = pos[b * 3 + 1] - ay, e1z = pos[b * 3 + 2] - az;
        const e2x = pos[c * 3] - ax, e2y = pos[c * 3 + 1] - ay, e2z = pos[c * 3 + 2] - az;
        const nx = e1y * e2z - e1z * e2y, ny = e1z * e2x - e1x * e2z, nz = e1x * e2y - e1y * e2x;
        const dot = nx * want[0] + ny * want[1] + nz * want[2];
        if (dot >= 0) idx.push(a, b, c, a, c, d); else idx.push(a, c, b, a, d, c);
      }
      for (let i = 0; i < N - 1; i++) {
        const j = i + 1;
        quad(i * 6 + 0, j * 6 + 0, j * 6 + 1, i * 6 + 1, [0, 1, 0]);
        const o2 = (i * 6 + 2) * 3;
        quad(i * 6 + 2, j * 6 + 2, j * 6 + 3, i * 6 + 3, [nor[o2], 0, nor[o2 + 2]]);
        const o4 = (i * 6 + 4) * 3;
        quad(i * 6 + 4, j * 6 + 4, j * 6 + 5, i * 6 + 5, [nor[o4], 0, nor[o4 + 2]]);
      }
      const geo = new T.BufferGeometry();
      geo.setAttribute("position", new T.BufferAttribute(pos, 3));
      geo.setAttribute("normal", new T.BufferAttribute(nor, 3));
      geo.setIndex(idx);
      geo.computeBoundingBox();
      geo.boundingBox.expandByScalar(0.1);
      geo.computeBoundingSphere();
      const mesh = new T.Mesh(geo, matStruct);
      mesh.position.y = Y0;
      mesh.frustumCulled = false;
      root.add(mesh);
      springs.push({ geo: geo, pos: pos, base: base, par: par, n: N * 6 });

      // anchor post + bond pad
      const ax = sx * (XA + 0.3), az = zc + sz * 0.0;
      const post = new T.Mesh(kit.roundedBox(0.62, statH, 0.7, 0.04, 0.015), matStructDark);
      post.position.set(ax, TOP + statH / 2, az + sz * 0.05);
      root.add(post);
      const gp = new T.Mesh(kit.roundedBox(0.34, 0.03, 0.34, 0.03, 0.008), kit.M.gold());
      gp.position.set(ax + sx * 0.02, TOP + statH + 0.014, az + sz * 0.05);
      root.add(gp);
      return gp.position.clone();
    }
    const anchorTops = [];
    [1, -1].forEach(function (sx) { [1, -1].forEach(function (sz) { anchorTops.push({ sx: sx, sz: sz, p: makeSpring(sx, sz) }); }); });

    /* ---------- bond pads + wires ---------- */
    const goldMat = kit.M.gold(); goldMat.emissive.set(0x4a3000);
    const padGeo = kit.roundedBox(0.3, 0.03, 0.3, 0.03, 0.008);
    const pads = [];
    [-1.5, -0.5, 0.5, 1.5].forEach(function (x) { pads.push(new T.Vector3(x, TOP + 0.14 + 0.02, 2.38)); });
    [-1.2, -0.3, 0.6].forEach(function (z) { pads.push(new T.Vector3(3.86, TOP + 0.14 + 0.02, z)); });
    pads.forEach(function (p) {
      const m = new T.Mesh(padGeo, goldMat);
      m.position.copy(p);
      root.add(m);
    });
    const ballGeo = new T.SphereGeometry(0.035, 12, 8);
    function wire(a, b, lift) {
      const mid = a.clone().lerp(b, 0.5);
      mid.y += lift;
      const c1 = a.clone().lerp(mid, 0.55); c1.y = a.y + lift * 0.95;
      const c2 = b.clone().lerp(mid, 0.55); c2.y = b.y + lift * 0.95;
      const curve = new T.CubicBezierCurve3(a, c1, c2, b);
      const w = new T.Mesh(new T.TubeGeometry(curve, 28, 0.013, 6, false), goldMat);
      root.add(w);
      [a, b].forEach(function (p) { const s = new T.Mesh(ballGeo, goldMat); s.position.copy(p); s.position.y += 0.005; root.add(s); });
    }
    const sTop = TOP + statH + 0.03;
    wire(new T.Vector3(-1.5, TOP + 0.17, 2.38), new T.Vector3(-1.1, sTop, 1.62), 0.55);
    wire(new T.Vector3(-0.5, TOP + 0.17, 2.38), new T.Vector3(-0.35, sTop, 1.62), 0.5);
    wire(new T.Vector3(0.5, TOP + 0.17, 2.38), new T.Vector3(0.45, sTop, 1.62), 0.5);
    wire(new T.Vector3(1.5, TOP + 0.17, 2.38), new T.Vector3(1.1, sTop, 1.62), 0.55);
    wire(new T.Vector3(3.86, TOP + 0.17, 0.6), new T.Vector3(3.0, sTop, 0.45), 0.5);
    wire(new T.Vector3(3.86, TOP + 0.17, -0.3), new T.Vector3(3.0, sTop, -0.45), 0.5);
    wire(new T.Vector3(3.86, TOP + 0.17, -1.2), new T.Vector3(1.0, sTop, -1.62), 0.7);

    /* ---------- motion axis arrow ---------- */
    const arrowMat = kit.M.additive(0xffc86a, 0.55);
    const arr = new T.Group();
    const shaft = new T.Mesh(new T.CylinderGeometry(0.014, 0.014, 2.4, 8), arrowMat);
    shaft.rotation.z = Math.PI / 2;
    const h1 = new T.Mesh(new T.ConeGeometry(0.06, 0.18, 12), arrowMat);
    h1.rotation.z = -Math.PI / 2; h1.position.x = 1.3;
    const h2 = new T.Mesh(new T.ConeGeometry(0.06, 0.18, 12), arrowMat);
    h2.rotation.z = Math.PI / 2; h2.position.x = -1.3;
    arr.add(shaft, h1, h2);
    arr.position.set(0, TOP + 0.62, 0);
    arr.children.forEach(function (c) { c.userData.noFit = true; });
    root.add(arr);

    /* ---------- accent glow under mass ---------- */
    const cg = kit.glow(0x4aa8ff, 4.2, 0.22);
    cg.position.set(0, TOP + 0.02, 0);
    root.add(cg);

    const pl = new T.PointLight(0x58b0ff, 1.6, 9); pl.position.set(0, 1.6, 0.4); root.add(pl);
    const pl2 = new T.PointLight(0xffb070, 0.9, 9); pl2.position.set(-3, 1.2, 2.5); root.add(pl2);
    const sh = kit.contactShadow(5.4, 0.7);
    sh.position.y = -1.9;
    g.add(sh);

    /* ---------- animation ---------- */
    const AMP = 0.075;
    function update(t) {
      const d = AMP * (0.85 * Math.sin(t * 1.15) + 0.15 * Math.sin(t * 2.7 + 1));
      massGroup.position.x = d;
      for (let k = 0; k < springs.length; k++) {
        const sp = springs[k], P = sp.pos;
        for (let i = 0; i < sp.n; i++) P[i * 3] = sp.base[i] + d * sp.par[i];
        sp.geo.attributes.position.needsUpdate = true;
      }
      const f = 0.5 + 0.5 * Math.sin(t * 9) * 0.15;
      const nrm = d / AMP;
      matGapA.opacity = Math.max(0.08, 0.42 + 0.5 * nrm) * (0.9 + f * 0.25);
      matGapB.opacity = Math.max(0.08, 0.42 - 0.5 * nrm) * (0.9 + f * 0.25);
      cg.material.opacity = 0.16 + 0.07 * Math.abs(nrm);
      arrowMat.opacity = 0.35 + 0.25 * Math.abs(nrm);
      arr.position.x = d * 1.5;
      root.rotation.y = Math.sin(t * 0.22) * 0.42 - 0.12;
    }
    update(0);
    return { group: g, update: update, fit: 2.1, spin: 0 };
  };
})();
