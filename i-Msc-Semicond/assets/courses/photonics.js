/* Photonics & Optoelectronics — silicon-photonic chip with waveguides, MZI, ring resonator, grating coupler, fibre, LD, LED, detector, prism */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  IMEN.courseScenes = IMEN.courseScenes || {};
  IMEN.courseScenes.photonics = function (kit) {
    const T = THREE;
    const g = new T.Group();
    const root = new T.Group();
    root.rotation.x = 0.42;
    g.add(root);

    const WY = 0.2; // waveguide plane height
    const BLUE = 0x3a7bff, CYAN = 0x22e8ff, MAG = 0xff3ad8;

    /* ---------- chip body + top surface ---------- */
    const body = new T.Mesh(kit.roundedBox(7.2, 0.3, 4.8, 0.12, 0.03), kit.M.silicon(0x0a1026));
    root.add(body);
    const W = 1024, H = 704;
    const CX = function (x) { return (x + 3.5) / 7 * W; };
    const CZ = function (z) { return (z + 2.3) / 4.6 * H; };
    const topTex = kit.canvasTex(W, H, function (c) {
      const r = kit.rng(21);
      c.fillStyle = "#0a1a3a"; c.fillRect(0, 0, W, H);
      const grd = c.createRadialGradient(W * 0.5, H * 0.5, 40, W * 0.5, H * 0.5, W * 0.7);
      grd.addColorStop(0, "rgba(40,90,170,0.35)"); grd.addColorStop(1, "rgba(0,0,0,0.35)");
      c.fillStyle = grd; c.fillRect(0, 0, W, H);
      const img = c.getImageData(0, 0, W, H);
      for (let i = 0; i < W * H; i++) { const v = (r() - 0.5) * 14; img.data[i * 4] += v; img.data[i * 4 + 1] += v; img.data[i * 4 + 2] += v * 1.4; }
      c.putImageData(img, 0, 0);
      c.strokeStyle = "rgba(110,160,255,0.09)"; c.lineWidth = 1;
      for (let x = 0; x <= W; x += 20) { c.beginPath(); c.moveTo(x, 0); c.lineTo(x, H); c.stroke(); }
      for (let y = 0; y <= H; y += 20) { c.beginPath(); c.moveTo(0, y); c.lineTo(W, y); c.stroke(); }
      c.strokeStyle = "rgba(150,200,255,0.30)"; c.lineWidth = 2;
      c.strokeRect(14, 14, W - 28, H - 28);
      c.strokeStyle = "rgba(150,200,255,0.12)"; c.strokeRect(24, 24, W - 48, H - 48);
      // electrical traces + pads
      const padsX = [];
      for (let i = 0; i < 13; i++) padsX.push(-3.0 + i * 0.5);
      c.lineCap = "round";
      padsX.forEach(function (x, i) {
        [-1, 1].forEach(function (side) {
          const zEdge = side * 2.0;
          const px = CX(x), pz = CZ(zEdge);
          c.fillStyle = "rgba(235,190,90,0.95)";
          c.fillRect(px - 15, pz - 15, 30, 30);
          c.fillStyle = "rgba(120,80,10,0.7)";
          c.fillRect(px - 8, pz - 8, 16, 16);
          // trace toward centre
          c.strokeStyle = "rgba(210,150,90,0.55)"; c.lineWidth = 3;
          c.beginPath(); c.moveTo(px, pz - side * 15);
          const tz = side * (1.55 - (i % 3) * 0.12);
          c.lineTo(px, CZ(tz));
          if (i % 2) c.lineTo(px + (i - 6) * 8, CZ(tz * 0.8));
          c.stroke();
        });
      });
      // modulator feed traces
      c.strokeStyle = "rgba(240,190,100,0.7)"; c.lineWidth = 4;
      c.beginPath(); c.moveTo(CX(0.5), CZ(-1.5)); c.lineTo(CX(0.5), CZ(-1.9)); c.stroke();
      c.beginPath(); c.moveTo(CX(0.5), CZ(-0.7)); c.lineTo(CX(0.75), CZ(-0.55)); c.lineTo(CX(0.75), CZ(-0.2)); c.stroke();
      // alignment marks
      c.strokeStyle = "rgba(255,255,255,0.35)"; c.lineWidth = 2;
      [[60, 60], [W - 60, 60], [60, H - 60], [W - 60, H - 60]].forEach(function (p) {
        c.beginPath(); c.arc(p[0], p[1], 12, 0, 6.283); c.moveTo(p[0] - 20, p[1]); c.lineTo(p[0] + 20, p[1]); c.moveTo(p[0], p[1] - 20); c.lineTo(p[0], p[1] + 20); c.stroke();
      });
      c.fillStyle = "rgba(170,210,255,0.55)"; c.font = "700 22px Inter, monospace";
      c.fillText("SiPh-PIC  λ 1310 / 1550 nm", CX(-0.9), CZ(1.95));
    });
    const topMat = new T.MeshPhysicalMaterial({ map: topTex, roughness: 0.3, metalness: 0.45, clearcoat: 1, clearcoatRoughness: 0.08, envMapIntensity: 1.0 });
    const top = new T.Mesh(new T.PlaneGeometry(7.0, 4.6), topMat);
    top.rotation.x = -Math.PI / 2;
    top.position.y = 0.156;
    root.add(top);

    /* ---------- route utilities ---------- */
    function sm(t) { return t * t * (3 - 2 * t); }
    function line(x0, z0, x1, z1, n) { const o = []; for (let i = 0; i <= n; i++) { const t = i / n; o.push([x0 + (x1 - x0) * t, z0 + (z1 - z0) * t]); } return o; }
    function sbend(x0, z0, x1, z1, n) { const o = []; for (let i = 0; i <= n; i++) { const t = i / n; o.push([x0 + (x1 - x0) * t, z0 + (z1 - z0) * sm(t)]); } return o; }
    function bump(x0, x1, z, dz, n) { const o = []; for (let i = 0; i <= n; i++) { const t = i / n; o.push([x0 + (x1 - x0) * t, z + dz * Math.sin(Math.PI * Math.pow(Math.sin(Math.PI * t / 1), 1) * 0.5 * 2 * 0.5) ]); } return o; }
    function cat() { const o = []; for (let k = 0; k < arguments.length; k++) { const a = arguments[k]; for (let i = (k ? 1 : 0); i < a.length; i++) o.push(a[i]); } return o; }
    // segment: [[x,z]...]  (y=0 local, positioned at WY by mesh)
    function Route(parts) {
      const pts = cat.apply(null, parts);
      const n = pts.length;
      this.n = n;
      this.xyz = new Float32Array(n * 3);
      for (let i = 0; i < n; i++) { this.xyz[i * 3] = pts[i][0]; this.xyz[i * 3 + 1] = WY + 0.03; this.xyz[i * 3 + 2] = pts[i][1]; }
      let L = 0;
      for (let i = 1; i < n; i++) L += Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]);
      this.len = L;
    }
    Route.prototype.at = function (u, out) {
      const f = Math.min(Math.max(u, 0), 0.9999) * (this.n - 1);
      const i = Math.floor(f), k = f - i, a = i * 3, b = a + 3;
      out.set(this.xyz[a] + (this.xyz[b] - this.xyz[a]) * k, this.xyz[a + 1] + (this.xyz[b + 1] - this.xyz[a + 1]) * k, this.xyz[a + 2] + (this.xyz[b + 2] - this.xyz[a + 2]) * k);
    };

    /* ---------- segments ---------- */
    const S_in = line(-3.6, 0, -2.3, 0, 14);
    const S_up1 = cat(sbend(-2.3, 0, -1.1, -1.1, 22), line(-1.1, -1.1, -0.5, -1.1, 6));
    const S_ld = cat(line(-2.4, -1.95, -1.6, -1.95, 8), sbend(-1.6, -1.95, -0.85, -1.1, 18), line(-0.85, -1.1, -0.5, -1.1, 4));
    function arm(dz) {
      const o = [];
      const n = 40;
      for (let i = 0; i <= n; i++) {
        const t = i / n; const x = -0.5 + 2.0 * t;
        let w;
        if (t < 0.3) w = sm(t / 0.3); else if (t > 0.7) w = sm((1 - t) / 0.3); else w = 1;
        o.push([x, -1.1 + dz * w]);
      }
      return o;
    }
    const S_mzA = arm(-0.32), S_mzB = arm(0.32);
    const S_out = line(1.5, -1.1, 2.55, -1.1, 14);
    const S_lo = cat(sbend(-2.3, 0, -1.1, 1.0, 22), line(-1.1, 1.0, 2.55, 1.0, 30));
    // fibre cable (3D curve)
    const cablePts = [new T.Vector3(-5.75, 0.95, -0.35), new T.Vector3(-5.55, 0.45, -0.2), new T.Vector3(-5.2, 0.24, 0), new T.Vector3(-4.8, WY + 0.03, 0)];
    const cableCurve = new T.CatmullRomCurve3(cablePts);
    const cableDense = [];
    for (let i = 0; i <= 16; i++) { const p = cableCurve.getPoint(i / 16); cableDense.push([p.x, p.z, p.y]); }

    // routes with 3D start (cable) -> use custom assembly
    function Route3(cableD, parts) {
      Route.call(this, parts);
      const nc = cableD.length;
      const xyz = new Float32Array((this.n + nc) * 3);
      for (let i = 0; i < nc; i++) { xyz[i * 3] = cableD[i][0]; xyz[i * 3 + 1] = cableD[i][2]; xyz[i * 3 + 2] = cableD[i][1]; }
      xyz.set(this.xyz, nc * 3);
      this.xyz = xyz; this.n += nc;
      this.len += 4.3;
    }
    Route3.prototype = Route.prototype;
    const fibreLine = line(-4.8, 0, -3.6, 0, 6);
    const pre = [fibreLine, S_in];
    const R1A = new Route3(cableDense, pre.concat([S_up1, S_mzA, S_out]));
    const R1B = new Route3(cableDense, pre.concat([S_up1, S_mzB, S_out]));
    const R3 = new Route3(cableDense, pre.concat([S_lo]));
    const R2A = new Route(cat ? [S_ld, S_mzA, S_out] : []);
    const R2B = new Route([S_ld, S_mzB, S_out]);
    // ring route
    const RC = [0.45, 0.45];
    const ringPts = [];
    for (let i = 0; i <= 90; i++) { const a = i / 90 * Math.PI * 2 + Math.PI / 2; ringPts.push([0.45 + 0.45 * Math.cos(a), 0.3 + 0.45 * Math.sin(a)]); }
    // ring center z = 0.45 so bottom (a=pi/2 => +z) touches bus at z=1.0-0.1
    const RR = new Route([ringPts]);

    /* ---------- waveguide rendering ---------- */
    const ridgeMat = kit.M.silicon(0x5a78c8); ridgeMat.metalness = 0.9; ridgeMat.roughness = 0.2;
    function curveOf(pts2) { const v = []; for (let i = 0; i < pts2.length; i += 1) v.push(new T.Vector3(pts2[i][0], 0, pts2[i][1])); return new T.CatmullRomCurve3(v); }
    const coreMats = {};
    function coreMat(c) { return coreMats[c] || (coreMats[c] = new T.MeshBasicMaterial({ color: c, toneMapped: false })); }
    const haloMats = {};
    function haloMat(c, o) { const k = c + "_" + o; return haloMats[k] || (haloMats[k] = kit.M.additive(c, o)); }
    const haloList = [];
    function guide(pts2, col, rr) {
      const cv = curveOf(pts2);
      const n = Math.max(20, Math.min(160, pts2.length * 2));
      const ridge = new T.Mesh(new T.TubeGeometry(cv, n, rr ? rr * 2.2 : 0.075, 8, false), ridgeMat);
      ridge.scale.y = 0.32; ridge.position.y = WY; root.add(ridge);
      const core = new T.Mesh(new T.TubeGeometry(cv, n, rr || 0.026, 8, false), coreMat(col));
      core.position.y = WY + 0.02; root.add(core);
      const hm = haloMat(col, 0.16);
      const halo = new T.Mesh(new T.TubeGeometry(cv, n, rr ? rr * 3 : 0.075, 8, false), hm);
      halo.position.y = WY + 0.02; halo.userData.noFit = true; root.add(halo);
      return cv;
    }
    guide(cat(fibreLine, S_in), 0x9fd8ff);
    guide(S_up1, BLUE);
    guide(S_ld, 0xa070ff);
    guide(S_mzA, CYAN); guide(S_mzB, CYAN);
    guide(S_out, CYAN);
    guide(S_lo, MAG);
    // ring
    {
      const rg = new T.TorusGeometry(0.45, 0.026, 8, 90);
      const rm = new T.Mesh(rg, coreMat(MAG)); rm.rotation.x = -Math.PI / 2; rm.position.set(0.45, WY + 0.02, 0.3); root.add(rm);
      const rr = new T.Mesh(new T.TorusGeometry(0.45, 0.075, 8, 90), ridgeMat); rr.rotation.x = -Math.PI / 2; rr.scale.z = 0.32; rr.position.set(0.45, WY, 0.3); root.add(rr);
      const rh = new T.Mesh(new T.TorusGeometry(0.45, 0.075, 8, 90), haloMat(MAG, 0.16)); rh.rotation.x = -Math.PI / 2; rh.position.set(0.45, WY + 0.02, 0.3); rh.userData.noFit = true; root.add(rh);
      // heater arc (gold)
      const hg = new T.Mesh(new T.TorusGeometry(0.45, 0.022, 6, 60, Math.PI * 1.6), kit.M.gold()); hg.rotation.x = -Math.PI / 2; hg.rotation.z = 0.7; hg.position.set(0.45, WY + 0.075, 0.3); root.add(hg);
    }
    // MZI electrodes
    const gold = kit.M.gold(); gold.emissive.set(0x3a2600);
    [-1.1 - 0.32, -1.1 + 0.32].forEach(function (z) {
      const e = new T.Mesh(kit.roundedBox(0.95, 0.04, 0.09, 0.02, 0.01), gold);
      e.position.set(0.5, WY + 0.085, z); root.add(e);
    });

    /* ---------- grating coupler + vertical beam ---------- */
    {
      const gx0 = 2.55, gx1 = 3.05;
      const sh = new T.Shape();
      sh.moveTo(gx0 - 0.5, -0.06); sh.lineTo(gx0, -0.3); sh.lineTo(gx1, -0.3); sh.lineTo(gx1, 0.3); sh.lineTo(gx0, 0.3); sh.lineTo(gx0 - 0.5, 0.06);
      const gg = new T.ExtrudeGeometry(sh, { depth: 0.03, bevelEnabled: false });
      gg.rotateX(-Math.PI / 2);
      const gm = new T.Mesh(gg, ridgeMat); gm.position.set(0, WY, -1.1); root.add(gm);
      const bars = kit.instanced(new T.BoxGeometry(0.035, 0.03, 0.55), coreMat(0x9ff4ff), 9, function (i, d) { d.position.set(gx0 + 0.03 + i * 0.055, WY + 0.045, -1.1); });
      root.add(bars);
      const gl = kit.glow(CYAN, 1.3, 0.55); gl.position.set((gx0 + gx1) / 2, WY + 0.1, -1.1); root.add(gl);
      const bm = kit.beam(0.1, 0.38, 1.35, "130,230,255", 0.55);
      bm.position.set(2.8, WY + 0.05 + 0.675, -1.1); bm.rotation.z = -0.22; root.add(bm);
      // prism
      const tri = new T.Shape();
      const s = 0.62;
      tri.moveTo(-s / 2, -s * 0.29); tri.lineTo(s / 2, -s * 0.29); tri.lineTo(0, s * 0.577); tri.lineTo(-s / 2, -s * 0.29);
      const pg = new T.ExtrudeGeometry(tri, { depth: 0.46, bevelEnabled: true, bevelThickness: 0.015, bevelSize: 0.015, bevelSegments: 2 });
      pg.translate(0, 0, -0.23);
      const pm = kit.M.glass(0xdfeeff, 0.38);
      const prism = new T.Mesh(pg, pm);
      prism.position.set(3.08, 1.62, -1.1); prism.rotation.z = 0.12; prism.renderOrder = 6; root.add(prism);
      // spectral fan
      const cols = [0xff3040, 0xff8a28, 0xffe030, 0x40ff70, 0x30c8ff, 0x4060ff, 0xb050ff];
      const origin = new T.Vector3(3.3, 1.7, -1.1);
      cols.forEach(function (col, i) {
        const ang = 0.05 + i * 0.065;
        const len = 1.25;
        const m = new T.Mesh(new T.CylinderGeometry(0.012, 0.03, len, 6, 1, true), kit.M.additive(col, 0.8));
        m.position.set(origin.x + Math.cos(ang) * len / 2, origin.y + Math.sin(ang) * len / 2 - 0.08 - i * 0.02, origin.z);
        m.rotation.z = ang - Math.PI / 2;
        m.userData.noFit = true; root.add(m);
      });
      const og = kit.glow(0xffffff, 0.55, 0.85); og.position.copy(origin); root.add(og);
    }

    /* ---------- photodetector ---------- */
    const detTex = kit.canvasTex(256, 160, function (c, w, h) {
      c.fillStyle = "#3a1208"; c.fillRect(0, 0, w, h);
      const gr = c.createLinearGradient(0, 0, w, h); gr.addColorStop(0, "#7a2a10"); gr.addColorStop(1, "#c2542a");
      c.fillStyle = gr; c.fillRect(10, 10, w - 20, h - 20);
      c.strokeStyle = "rgba(255,210,120,0.9)"; c.lineWidth = 3;
      for (let x = 24; x < w - 20; x += 16) { c.beginPath(); c.moveTo(x, 10); c.lineTo(x, h - 10); c.stroke(); }
      c.strokeRect(6, 6, w - 12, h - 12);
    });
    const det = new T.Mesh(kit.roundedBox(0.95, 0.07, 0.6, 0.04, 0.012), new T.MeshPhysicalMaterial({ map: detTex, metalness: 0.7, roughness: 0.3, clearcoat: 0.6, emissive: 0x401408, emissiveMap: detTex }));
    det.position.set(3.0, WY + 0.03, 1.0); root.add(det);
    const detGlow = kit.glow(0xff7a30, 1.9, 0.3); detGlow.position.set(3.0, WY + 0.2, 1.0); root.add(detGlow);
    // bond wires from detector
    const wireMat = kit.M.gold();
    function wire(a, b, lift) {
      const c1 = a.clone().lerp(b, 0.3); c1.y += lift; const c2 = a.clone().lerp(b, 0.7); c2.y += lift;
      const w = new T.Mesh(new T.TubeGeometry(new T.CubicBezierCurve3(a, c1, c2, b), 20, 0.012, 6, false), wireMat);
      root.add(w);
      [a, b].forEach(function (p) { const s = new T.Mesh(new T.SphereGeometry(0.03, 10, 8), wireMat); s.position.copy(p); root.add(s); });
    }
    wire(new T.Vector3(3.3, WY + 0.08, 1.2), new T.Vector3(3.2, 0.19, 2.0), 0.25);
    wire(new T.Vector3(3.3, WY + 0.08, 0.85), new T.Vector3(3.2, 0.19, 1.6), 0.25);

    /* ---------- laser diode ---------- */
    {
      const sub = new T.Mesh(kit.roundedBox(0.8, 0.12, 0.55, 0.03, 0.015), kit.M.gold());
      sub.position.set(-2.85, WY + 0.01, -1.95); root.add(sub);
      const ld = new T.Mesh(kit.roundedBox(0.55, 0.12, 0.3, 0.02, 0.012), kit.M.silicon(0x1a1030));
      ld.position.set(-2.8, WY + 0.12, -1.95); root.add(ld);
      const fac = new T.Mesh(new T.PlaneGeometry(0.28, 0.1), new T.MeshBasicMaterial({ color: 0xb890ff, toneMapped: false }));
      fac.rotation.y = Math.PI / 2; fac.position.set(-2.52, WY + 0.12, -1.95); root.add(fac);
      const gl = kit.glow(0x9a6cff, 0.9, 0.9); gl.position.set(-2.45, WY + 0.12, -1.95); root.add(gl);
      const topc = new T.Mesh(kit.roundedBox(0.4, 0.02, 0.1, 0.01, 0.006), kit.M.metal(0xf0d080, 0.15));
      topc.position.set(-2.82, WY + 0.19, -1.95); root.add(topc);
      wire(new T.Vector3(-2.9, WY + 0.2, -1.95), new T.Vector3(-3.15, 0.19, -1.7), 0.22);
    }
    /* ---------- LED accent ---------- */
    let ledGlow;
    {
      const base = new T.Mesh(kit.roundedBox(0.7, 0.1, 0.7, 0.05, 0.015), kit.M.ceramic(0xe8e4d8));
      base.position.set(-2.85, WY + 0.0, 1.55); root.add(base);
      const ring = new T.Mesh(new T.TorusGeometry(0.27, 0.025, 8, 36), kit.M.gold());
      ring.rotation.x = Math.PI / 2; ring.position.set(-2.85, WY + 0.06, 1.55); root.add(ring);
      const die = new T.Mesh(new T.BoxGeometry(0.16, 0.04, 0.16), new T.MeshBasicMaterial({ color: 0x9cc8ff, toneMapped: false }));
      die.position.set(-2.85, WY + 0.09, 1.55); root.add(die);
      const dome = new T.Mesh(new T.SphereGeometry(0.25, 28, 14, 0, Math.PI * 2, 0, Math.PI / 2), kit.M.glass(0xbfdcff, 0.4));
      dome.position.set(-2.85, WY + 0.06, 1.55); dome.renderOrder = 5; root.add(dome);
      ledGlow = kit.glow(0x6aa8ff, 1.6, 0.6); ledGlow.position.set(-2.85, WY + 0.18, 1.55); root.add(ledGlow);
    }

    /* ---------- fibre + ferrule ---------- */
    {
      const fib = new T.Mesh(new T.CylinderGeometry(0.045, 0.045, 0.5, 12), kit.M.glass(0xddeeff, 0.5));
      fib.rotation.z = Math.PI / 2; fib.position.set(-3.88, WY + 0.03, 0); root.add(fib);
      const fcore = new T.Mesh(new T.CylinderGeometry(0.014, 0.014, 0.5, 8), coreMat(0xd0f0ff));
      fcore.rotation.z = Math.PI / 2; fcore.position.set(-3.88, WY + 0.03, 0); root.add(fcore);
      const fer = new T.Mesh(new T.CylinderGeometry(0.17, 0.17, 0.75, 32), kit.M.metal(0xd8dde6, 0.25));
      fer.rotation.z = Math.PI / 2; fer.position.set(-4.4, WY + 0.03, 0); root.add(fer);
      const ring1 = new T.Mesh(new T.CylinderGeometry(0.2, 0.2, 0.09, 32), kit.M.gold());
      ring1.rotation.z = Math.PI / 2; ring1.position.set(-4.72, WY + 0.03, 0); root.add(ring1);
      const nose = new T.Mesh(new T.CylinderGeometry(0.1, 0.17, 0.12, 32), kit.M.metal(0xe6eaf0, 0.18));
      nose.rotation.z = Math.PI / 2; nose.position.set(-4.0, WY + 0.03, 0); root.add(nose);
      const boot = new T.Mesh(new T.CylinderGeometry(0.075, 0.17, 0.5, 24), kit.M.plastic(0x15171c, 0.5));
      boot.rotation.z = Math.PI / 2; boot.position.set(-5.03, WY + 0.03, 0); root.add(boot);
      const holder = new T.Mesh(kit.roundedBox(1.05, 0.12, 0.62, 0.05, 0.02), kit.M.metal(0x9aa3b5, 0.3));
      holder.position.set(-4.45, -0.12, 0); root.add(holder);
      const jacket = new T.Mesh(new T.TubeGeometry(cableCurve, 30, 0.07, 12, false), kit.M.glass(0xffc84a, 0.55));
      jacket.renderOrder = 4; root.add(jacket);
      const cc = new T.Mesh(new T.TubeGeometry(cableCurve, 30, 0.02, 8, false), coreMat(0xcdefff));
      root.add(cc);
      const fg = kit.glow(0xbfe4ff, 0.8, 0.7); fg.position.set(-3.62, WY + 0.03, 0); root.add(fg);
    }

    /* ---------- pulses ---------- */
    const pulseTex = kit.tex.radial([[0, "rgba(255,255,255,1)"], [0.16, "rgba(255,255,255,0.95)"], [0.4, "rgba(255,255,255,0.35)"], [1, "rgba(255,255,255,0)"]], 128);
    const pmats = {};
    function pmat(c) { return pmats[c] || (pmats[c] = new T.SpriteMaterial({ map: pulseTex, color: c, transparent: true, depthWrite: false, blending: T.AdditiveBlending, toneMapped: false })); }
    const pulses = [];
    function addPulse(route, col, phase, speed, size) {
      const s = new T.Sprite(pmat(col));
      s.userData.noFit = true;
      root.add(s);
      pulses.push({ s: s, r: route, ph: phase, sp: speed, size: size || 0.34, det: route === R3 });
    }
    const cols = [BLUE, CYAN, MAG, CYAN, BLUE];
    for (let i = 0; i < 5; i++) { const c = cols[i], ph = i / 5; addPulse(R1A, c, ph, 1.5); addPulse(R1B, c, ph, 1.5); }
    for (let i = 0; i < 4; i++) addPulse(R3, i % 2 ? MAG : 0xff7ae8, i / 4 + 0.08, 1.35);
    for (let i = 0; i < 3; i++) { addPulse(R2A, 0xa57cff, i / 3, 1.3); addPulse(R2B, 0xa57cff, i / 3, 1.3); }
    for (let i = 0; i < 4; i++) addPulse(RR, MAG, i / 4, 1.5, 0.3);

    const sh = kit.contactShadow(5.8, 0.7); sh.position.y = -1.6; g.add(sh);
    const pl = new T.PointLight(0x58a8ff, 1.2, 10); pl.position.set(0, 1.5, 0.5); root.add(pl);

    const tmpV = new T.Vector3();
    function update(t) {
      let dm = 0;
      for (let i = 0; i < pulses.length; i++) {
        const p = pulses[i];
        const u = ((t * p.sp / p.r.len) + p.ph) % 1;
        p.r.at(u, tmpV);
        p.s.position.copy(tmpV);
        const fade = Math.min(1, u * 12, (1 - u) * 12);
        p.s.scale.setScalar(p.size * (0.4 + 0.6 * fade));
        if (p.det && u > 0.86) dm = Math.max(dm, Math.sin(Math.min(1, (u - 0.86) / 0.14) * Math.PI));
      }
      detGlow.material.opacity = 0.18 + 0.7 * dm;
      ledGlow.material.opacity = 0.5 + 0.2 * Math.sin(t * 2.2);
      root.rotation.y = Math.sin(t * 0.2) * 0.45 + 0.1;
    }
    update(0);
    return { group: g, update: update, fit: 2.1, spin: 0 };
  };
})();
