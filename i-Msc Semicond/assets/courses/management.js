/* Project Management — premium 3D Gantt / roadmap board for a chip development lifecycle */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  window.IMEN.courseScenes = window.IMEN.courseScenes || {};

  IMEN.courseScenes.management = function (kit) {
    const T = THREE;
    const root = new T.Group();      // sways about Y
    const board = new T.Group();     // tilted
    root.add(board);
    const TILT = 1.3;
    board.rotation.x = TILT;
    board.position.y = 0.1;

    const BW = 4.9, BD = 3.2, TOP = 0.045;
    const X0 = -1.62, XW = 3.3;
    const X = function (t) { return X0 + t * XW; };
    const ZR = function (i) { return -0.6 + i * 0.3; };
    const ZM = -0.96, ZH = -1.27;
    const BAR_Y = TOP + 0.05;

    /* ---------- board ---------- */
    const tray = new T.Mesh(kit.roundedBox(BW + 0.22, 0.16, BD + 0.22, 0.16, 0.035), kit.M.metal(0x2c3652, 0.3));
    tray.position.y = -0.1;
    board.add(tray);
    const glass = new T.Mesh(kit.roundedBox(BW, 0.09, BD, 0.12, 0.02), new T.MeshPhysicalMaterial({ color: 0x16285a, metalness: 0.2, roughness: 0.38, clearcoat: 0.8, clearcoatRoughness: 0.25, envMapIntensity: 0.45 }));
    glass.position.y = 0;
    glass.renderOrder = 1;
    board.add(glass);
    // grid + alternating row stripes + week ticks
    const gridTex = kit.canvasTex(1024, 672, function (c, W, H) {
      c.clearRect(0, 0, W, H);
      const px = function (x) { return ((x + BW / 2) / BW) * W; };
      const pz = function (z) { return ((z + BD / 2) / BD) * H; };
      for (let i = 0; i < 7; i++) {
        c.fillStyle = i % 2 ? "rgba(150,190,255,0.0)" : "rgba(150,190,255,0.12)";
        c.fillRect(px(-2.3), pz(ZR(i) - 0.15), px(2.3) - px(-2.3), pz(0.3) - pz(0));
      }
      for (let i = 0; i <= 40; i++) {
        const x = X(i / 40);
        c.strokeStyle = i % 10 === 0 ? "rgba(170,210,255,0.75)" : i % 2 === 0 ? "rgba(140,185,255,0.32)" : "rgba(140,185,255,0.14)";
        c.lineWidth = i % 10 === 0 ? 2.5 : 1.2;
        c.beginPath(); c.moveTo(px(x), pz(ZH + 0.1)); c.lineTo(px(x), pz(1.42)); c.stroke();
      }
      c.strokeStyle = "rgba(170,210,255,0.5)";
      c.lineWidth = 1.5;
      for (let i = -1; i <= 7; i++) {
        const z = ZR(i) - 0.15;
        c.beginPath(); c.moveTo(px(-2.3), pz(z)); c.lineTo(px(2.3), pz(z)); c.stroke();
      }
      c.strokeStyle = "rgba(200,225,255,0.85)";
      c.lineWidth = 5;
      c.strokeRect(6, 6, W - 12, H - 12);
    });
    const grid = new T.Mesh(new T.PlaneGeometry(BW, BD), new T.MeshBasicMaterial({ map: gridTex, transparent: true, opacity: 0.95, depthWrite: false, toneMapped: false }));
    grid.rotation.x = -Math.PI / 2;
    grid.position.y = TOP + 0.002;
    grid.renderOrder = 1;
    board.add(grid);

    function flat(str, h, x, z, color, size) {
      const l = kit.label(str, h, { color: color || "#cfe0ff", size: size || 96 });
      l.rotation.x = -Math.PI / 2;
      l.position.set(x, TOP + 0.006, z);
      l.renderOrder = 3;
      board.add(l);
      return l;
    }
    ["Q1", "Q2", "Q3", "Q4"].forEach(function (q, i) { flat(q, 0.15, X(0.125 + i * 0.25), ZH, "#9fc4ff"); });
    ["SPEC", "RTL", "VERIFY", "LAYOUT", "FAB", "TEST", "QUAL"].forEach(function (n, i) { flat(n, 0.115, -2.02, ZR(i), "#cfe0ff"); });
    flat("MILESTONES", 0.09, -1.98, ZM, "#ffd98a");

    /* ---------- tasks ---------- */
    const COLORS = [0x2f7bff, 0x19c9e6, 0xffb53a, 0x1fd18a];
    const fillMats = COLORS.map(function (c) {
      return new T.MeshPhysicalMaterial({ color: c, emissive: c, emissiveIntensity: 0.22, roughness: 0.22, metalness: 0.0, clearcoat: 1, clearcoatRoughness: 0.08, envMapIntensity: 0.8 });
    });
    const trackMats = COLORS.map(function (c) { const m = kit.M.glass(c, 0.22); return m; });
    const tasks = [
      { r: 0, s: 0.0, e: 0.1, c: 0 }, { r: 1, s: 0.13, e: 0.31, c: 0 }, { r: 2, s: 0.35, e: 0.53, c: 1 },
      { r: 3, s: 0.35, e: 0.55, c: 1 }, { r: 4, s: 0.59, e: 0.74, c: 2 }, { r: 5, s: 0.78, e: 0.88, c: 3 },
      { r: 6, s: 0.91, e: 1.0, c: 3 }, { r: 2, s: 0.57, e: 0.67, c: 2 }, { r: 1, s: 0.34, e: 0.46, c: 2 }
    ];
    const BH = 0.13, BDp = 0.19;
    tasks.forEach(function (tk) {
      const len = (tk.e - tk.s) * XW;
      tk.len = len;
      tk.x0 = X(tk.s);
      const tg = kit.roundedBox(len, BH * 0.7, BDp + 0.02, 0.07, 0.02);
      const track = new T.Mesh(tg, trackMats[tk.c]);
      track.position.set(tk.x0 + len / 2, TOP + 0.04, ZR(tk.r));
      track.renderOrder = 2;
      board.add(track);
      const fg = kit.roundedBox(len, BH, BDp, 0.07, 0.025);
      fg.translate(len / 2, 0, 0);
      const fill = new T.Mesh(fg, fillMats[tk.c]);
      fill.position.set(tk.x0, BAR_Y, ZR(tk.r));
      board.add(fill);
      tk.fill = fill;
    });

    /* ---------- dependency arrows ---------- */
    const arrowMat = kit.M.metal(0xdbe6ff, 0.25);
    const coneG = new T.ConeGeometry(0.032, 0.07, 14);
    function dep(a, b) {
      const A = tasks[a], B = tasks[b];
      const xe = A.x0 + A.len, xb = B.x0;
      const y = BAR_Y + BH / 2 + 0.004;
      const pts = [
        new T.Vector3(xe, y, ZR(A.r)), new T.Vector3(xe + 0.05, y, ZR(A.r)),
        new T.Vector3(xe + 0.05, y, ZR(B.r)), new T.Vector3(xb - 0.07, y, ZR(B.r))
      ];
      const path = new T.CurvePath();
      for (let i = 0; i < 3; i++) path.add(new T.LineCurve3(pts[i], pts[i + 1]));
      const m = new T.Mesh(new T.TubeGeometry(path, 36, 0.011, 8, false), arrowMat);
      board.add(m);
      const cone = new T.Mesh(coneG, arrowMat);
      cone.rotation.z = -Math.PI / 2;
      cone.position.set(xb - 0.04, y, ZR(B.r));
      board.add(cone);
    }
    dep(0, 1); dep(1, 2); dep(1, 3); dep(3, 4); dep(2, 4); dep(4, 5); dep(5, 6);

    /* ---------- critical path ---------- */
    const chain = [0, 1, 3, 4, 5, 6];
    const cy = BAR_Y + BH / 2 + 0.035;
    const cp = [];
    chain.forEach(function (ti, k) {
      const tk = tasks[ti];
      cp.push(new T.Vector3(tk.x0, cy, ZR(tk.r)));
      cp.push(new T.Vector3(tk.x0 + tk.len, cy, ZR(tk.r)));
      if (k < chain.length - 1) {
        const nx = tasks[chain[k + 1]];
        const mid = (tk.x0 + tk.len + nx.x0) / 2;
        cp.push(new T.Vector3(mid, cy, ZR(tk.r)));
        cp.push(new T.Vector3(mid, cy, ZR(nx.r)));
      }
    });
    const cpath = new T.CurvePath();
    for (let i = 0; i < cp.length - 1; i++) if (cp[i].distanceTo(cp[i + 1]) > 1e-4) cpath.add(new T.LineCurve3(cp[i], cp[i + 1]));
    const critMat = new T.MeshStandardMaterial({ color: 0x300808, emissive: 0xff4a3d, emissiveIntensity: 1.0, roughness: 0.5 });
    const crit = new T.Mesh(new T.TubeGeometry(cpath, 220, 0.017, 8, false), critMat);
    board.add(crit);
    const critGlow = new T.Mesh(new T.TubeGeometry(cpath, 120, 0.045, 8, false), kit.M.additive(0xff2a20, 0.1));
    critGlow.userData.noFit = true;
    board.add(critGlow);
    const pulses = [];
    for (let i = 0; i < 3; i++) {
      const p = kit.glow(0xff7a5a, 0.3, 1);
      board.add(p);
      pulses.push(p);
    }
    const tmpV = new T.Vector3();

    /* ---------- milestones + flag ---------- */
    const msT = [0.1, 0.31, 0.55, 0.74, 0.88, 1.0];
    const dia = [];
    const diaG = new T.OctahedronGeometry(0.085, 0);
    const lineP = [];
    msT.forEach(function (mt, i) {
      const last = i === msT.length - 1;
      const m = new T.Mesh(diaG, last ? kit.M.gold() : i % 2 ? kit.M.aluminium() : kit.M.gold());
      m.scale.set(1, 1.25, 0.8);
      m.position.set(X(mt), TOP + 0.13, ZM);
      if (last) m.scale.multiplyScalar(1.35);
      board.add(m);
      dia.push(m);
      lineP.push(X(mt), TOP + 0.01, ZM + 0.12, X(mt), TOP + 0.01, 1.3);
    });
    const lg = new T.BufferGeometry();
    lg.setAttribute("position", new T.Float32BufferAttribute(lineP, 3));
    board.add(new T.LineSegments(lg, new T.LineBasicMaterial({ color: 0xffd98a, transparent: true, opacity: 0.22 })));

    const flag = new T.Group();
    flag.position.set(1.97, TOP, -0.45);
    flag.rotation.x = -1.0;
    const pole = new T.Mesh(new T.CylinderGeometry(0.018, 0.022, 0.7, 12), kit.M.metal(0xe8eefc, 0.18));
    pole.position.y = 0.35;
    flag.add(pole);
    const cap = new T.Mesh(new T.SphereGeometry(0.03, 16, 12), kit.M.gold());
    cap.position.y = 0.71;
    flag.add(cap);
    const fgeo = new T.PlaneGeometry(0.4, 0.26, 12, 5);
    fgeo.translate(0.2, 0, 0);
    const fpos = fgeo.attributes.position;
    const fbase = fpos.array.slice();
    const flagMesh = new T.Mesh(fgeo, new T.MeshPhysicalMaterial({ color: 0x1fd18a, emissive: 0x0b6a48, emissiveIntensity: 0.5, roughness: 0.35, clearcoat: 0.6, side: T.DoubleSide }));
    flagMesh.position.y = 0.54;
    flag.add(flagMesh);
    board.add(flag);

    /* ---------- budget: coin stacks ---------- */
    const coinG = new T.CylinderGeometry(0.15, 0.15, 0.034, 40);
    const nA = 8, nB = 4;
    const coinMat = kit.M.gold(); coinMat.envMapIntensity = 2.2; coinMat.emissive = new T.Color(0x5a3a00); coinMat.emissiveIntensity = 0.6;
    const coins = new T.InstancedMesh(coinG, coinMat, nA + nB);
    const dm = new T.Object3D();
    for (let i = 0; i < nA + nB; i++) {
      const st = i < nA ? 0 : 1, k = st ? i - nA : i;
      dm.position.set(st ? 2.1 : 2.0, TOP + 0.017 + k * 0.038, st ? 0.62 : 1.02);
      if (st) dm.position.x = 1.9;
      dm.rotation.y = (k * 1.7 + st) % 6.28;
      dm.position.x += Math.sin(k * 2.3 + st) * 0.012;
      dm.position.z += Math.cos(k * 1.9 + st) * 0.012;
      dm.updateMatrix();
      coins.setMatrixAt(i, dm.matrix);
    }
    board.add(coins);
    const topCoin = flat("$", 0.15, 2.0, 1.02, "#5a3a00", 128);
    topCoin.position.y = TOP + 0.017 + (nA - 1) * 0.038 + 0.02;
    topCoin.renderOrder = 4;

    /* ---------- NOW plane ---------- */
    const nowG = new T.Group();
    const nowPlane = new T.Mesh(new T.PlaneGeometry(2.8, 0.6), new T.MeshBasicMaterial({ map: kit.tex.shaft("120,210,255"), transparent: true, opacity: 0.45, depthWrite: false, blending: T.AdditiveBlending, toneMapped: false, side: T.DoubleSide }));
    nowPlane.rotation.y = Math.PI / 2;
    nowPlane.position.y = 0.3 + TOP;
    nowPlane.userData.noFit = true;
    nowG.add(nowPlane);
    const nowCore = new T.Mesh(new T.BoxGeometry(0.014, 0.014, 2.8), kit.M.glow(0x9fe6ff, 3));
    nowCore.position.y = TOP + 0.008;
    nowG.add(nowCore);
    const nowLabel = kit.label("TODAY", 0.1, { color: "#9fe6ff", size: 96 });
    nowLabel.rotation.x = -Math.PI / 2;
    nowLabel.position.set(0, TOP + 0.01, ZH);
    nowG.add(nowLabel);
    nowG.position.z = 0.05;
    board.add(nowG);

    /* ---------- plinth + rear struts ---------- */
    const plinth = new T.Mesh(kit.roundedBox(5.3, 0.2, 1.1, 0.2, 0.04), kit.M.metal(0x2a3350, 0.3));
    plinth.position.set(0, -1.78, 0.3);
    root.add(plinth);
    const plateTop = new T.Mesh(kit.roundedBox(4.9, 0.05, 0.28, 0.1, 0.012), kit.M.metal(0xa7b5d3, 0.3));
    plateTop.position.set(0, -1.65, 0.42);
    root.add(plateTop);
    [-1.9, 1.9].forEach(function (x) {
      const a = new T.Vector3(x, 0.9, -0.5), b = new T.Vector3(x, -1.7, -1.05);
      const len = a.distanceTo(b);
      const s = new T.Mesh(new T.CylinderGeometry(0.04, 0.05, len, 14), kit.M.metal(0x9aa8c8, 0.28));
      s.position.copy(a).add(b).multiplyScalar(0.5);
      s.quaternion.setFromUnitVectors(new T.Vector3(0, 1, 0), a.clone().sub(b).normalize());
      root.add(s);
    });
    const sh = kit.contactShadow(3.4, 0.85);
    sh.position.set(0, -1.9, 0);
    sh.scale.set(1.15, 0.9, 1);
    root.add(sh);

    /* ---------- animation ---------- */
    const CYC = 16;
    const startZ = ZH;
    return {
      group: root,
      fit: 2.1,
      spin: 0,
      update: function (t) {
        root.rotation.y = Math.sin(t * 0.32) * 0.5 - 0.12;
        const c = t % CYC;
        const prog = c < 12.5 ? c / 12.5 : 1;
        const retract = c > 14 ? Math.max(0, 1 - (c - 14) / 1.4) : 1;
        const nowX = X(prog * 1.02 - 0.01);
        nowG.position.x = nowX;
        const showNow = c < 13 ? 1 : Math.max(0, 1 - (c - 13) / 0.6);
        nowG.visible = showNow > 0.01 && c < 14;
        nowPlane.material.opacity = 0.4 * showNow * (0.85 + 0.15 * Math.sin(t * 4));
        for (let i = 0; i < tasks.length; i++) {
          const tk = tasks[i];
          let f = (nowX - tk.x0) / tk.len;
          f = f < 0 ? 0 : f > 1 ? 1 : f;
          f = f * f * (3 - 2 * f);
          if (c >= 14) f = f * retract;
          tk.fill.scale.x = Math.max(f, 0.0001);
          tk.fill.visible = f > 0.02;
        }
        for (let i = 0; i < dia.length; i++) {
          dia[i].rotation.y = t * 0.8 + i;
          dia[i].position.y = TOP + 0.13 + Math.sin(t * 1.4 + i) * 0.012;
        }
        for (let i = 0; i < pulses.length; i++) {
          cpath.getPointAt(((t * 0.11 + i / pulses.length) % 1), tmpV);
          pulses[i].position.copy(tmpV);
          pulses[i].position.y += 0.02;
        }
        critMat.emissiveIntensity = 1.0 + 0.3 * Math.sin(t * 2.4);
        for (let i = 0; i < fpos.count; i++) {
          const x = fbase[i * 3], y = fbase[i * 3 + 1];
          fpos.setZ(i, Math.sin(x * 12 - t * 3.2 + y * 3) * 0.022 * (x / 0.4));
        }
        fpos.needsUpdate = true;
        fgeo.computeVertexNormals();
      }
    };
  };
})();
