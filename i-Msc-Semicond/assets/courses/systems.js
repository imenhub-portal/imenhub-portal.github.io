/* core_6 - Microelectronic Systems : 2.5D AI-accelerator module (SoC + 4x HBM on a silicon interposer) */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  window.IMEN.courseScenes = window.IMEN.courseScenes || {};

  window.IMEN.courseScenes.systems = function (kit) {
    const T = THREE;
    const g = new T.Group();
    const mod = new T.Group();
    mod.rotation.x = 0.4;
    g.add(mod);

    /* ---------- helpers ---------- */
    function rrShape(s, w, d, r) {
      const x = -w / 2, y = -d / 2;
      s.moveTo(x + r, y);
      s.lineTo(x + w - r, y);
      s.quadraticCurveTo(x + w, y, x + w, y + r);
      s.lineTo(x + w, y + d - r);
      s.quadraticCurveTo(x + w, y + d, x + w - r, y + d);
      s.lineTo(x + r, y + d);
      s.quadraticCurveTo(x, y + d, x, y + d - r);
      s.lineTo(x, y + r);
      s.quadraticCurveTo(x, y, x + r, y);
      return s;
    }
    function ringGeo(ow, od, iw, id, h, r, bevel) {
      const s = rrShape(new T.Shape(), ow, od, r);
      s.holes.push(rrShape(new T.Path(), iw, id, Math.max(0.005, r * 0.5)));
      const geo = new T.ExtrudeGeometry(s, { depth: Math.max(0.001, h - bevel * 2), bevelEnabled: true, bevelThickness: bevel, bevelSize: bevel * 0.8, bevelSegments: 2, curveSegments: 5 });
      geo.rotateX(-Math.PI / 2);
      geo.translate(0, bevel, 0);
      return geo;
    }
    function capUV(tex, w, d) {
      tex.wrapS = tex.wrapT = T.RepeatWrapping;
      tex.repeat.set(1 / w, 1 / d);
      tex.offset.set(0.5, 0.5);
      return tex;
    }
    function worldTex(W, w, d, draw, linear) {
      const H = Math.round((W * d) / w);
      const t = kit.canvasTex(W, H, function (c) {
        c.save();
        c.scale(W / w, H / d);
        c.translate(w / 2, d / 2);
        draw(c);
        c.restore();
      }, { linear: !!linear });
      return capUV(t, w, d);
    }
    const glowDot = kit.glow(0x59c8ff, 1, 0.8);

    /* ---------- layout constants ---------- */
    const SBW = 5.2, SBD = 3.9, SBH = 0.12;
    const IPW = 4.2, IPD = 2.9, IPH = 0.1;
    const Y_SUB_TOP = SBH / 2;
    const Y_IP = Y_SUB_TOP + 0.08 + IPH / 2;
    const Y_IP_TOP = Y_IP + IPH / 2;
    const Y_DIE = Y_IP_TOP + 0.02;
    const SOC = 1.3;
    const HBM_W = 0.6, HBM_D = 0.93;
    const HBM = [[-1.55, 0.55], [-1.55, -0.55], [1.55, 0.55], [1.55, -0.55]];

    /* lanes between dies (interposer wiring) */
    const lanes = [];
    for (let s = -1; s <= 1; s += 2) {
      [0.55, -0.55].forEach(function (zc) {
        for (let k = 0; k < 12; k++) {
          const z = zc + (k - 5.5) * 0.058;
          lanes.push({ x0: s * 0.6, z0: z, x1: s * 1.28, z1: z, hbm: true, dir: k % 2 });
        }
      });
    }
    for (let s = -1; s <= 1; s += 2) {
      for (let k = 0; k < 14; k++) {
        const x = (k - 6.5) * 0.09;
        lanes.push({ x0: x, z0: s * 0.6, x1: x, z1: s * 0.95, hbm: false, dir: k % 2 });
      }
    }

    /* ================= SUBSTRATE ================= */
    function subDraw(bump) {
      return function (c) {
        const r = kit.rng(31);
        c.fillStyle = bump ? "#4c4c4c" : "#0a3f33";
        c.fillRect(-SBW / 2, -SBD / 2, SBW, SBD);
        if (!bump) {
          for (let i = 0; i < 2400; i++) {
            c.fillStyle = "rgba(" + (r() < 0.5 ? "0,0,0," : "90,230,200,") + r() * 0.05 + ")";
            c.fillRect((r() - 0.5) * SBW, (r() - 0.5) * SBD, 0.01 + r() * 0.04, 0.01 + r() * 0.04);
          }
        }
        c.lineCap = "round"; c.lineJoin = "round";
        c.strokeStyle = bump ? "#bbb" : "rgba(214,176,100,0.55)";
        c.lineWidth = 0.012;
        for (let i = 0; i < 150; i++) {
          let x = (r() - 0.5) * SBW, y = (r() - 0.5) * SBD;
          c.beginPath(); c.moveTo(x, y);
          for (let s = 0; s < 4; s++) {
            const d = r(), len = 0.1 + r() * 0.3;
            if (d < 0.33) x += len; else if (d < 0.66) y += len; else { x += len * 0.7; y += len * 0.7; }
            c.lineTo(x, y);
          }
          c.stroke();
          c.fillStyle = bump ? "#eee" : "#e3c27a";
          c.beginPath(); c.arc(x, y, 0.022, 0, 6.3); c.fill();
        }
        if (!bump) {
          c.fillStyle = "rgba(235,245,235,0.8)";
          c.font = "bold 0.1px monospace";
          c.fillText("IMEN-X1  2.5D AI ACCEL", -2.45, 1.86);
          c.strokeStyle = "rgba(235,245,235,0.7)";
          c.lineWidth = 0.014;
          [[-1,-1],[1,-1],[1,1],[-1,1]].forEach(function (p) {
            const x = p[0] * 2.5, y = p[1] * 1.8;
            c.beginPath(); c.moveTo(x - p[0] * 0.14, y); c.lineTo(x, y); c.lineTo(x, y - p[1] * 0.14); c.stroke();
          });
        }
      };
    }
    const subCap = new T.MeshPhysicalMaterial({
      map: worldTex(1024, SBW, SBD, subDraw(false)),
      bumpMap: worldTex(1024, SBW, SBD, subDraw(true), true), bumpScale: 1.4,
      roughness: 0.4, metalness: 0.05, clearcoat: 0.7, clearcoatRoughness: 0.25, envMapIntensity: 0.6
    });
    const substrate = new T.Mesh(kit.roundedBox(SBW, SBH, SBD, 0.08, 0.015), [subCap, kit.M.plastic(0x155c49, 0.5)]);
    mod.add(substrate);

    // decoupling caps around the stiffener (instanced)
    const capPos = [];
    for (let z = -1.4; z <= 1.41; z += 0.2) capPos.push([-2.43, z, 0], [2.43, z, 0]);
    for (let x = -2.2; x <= 2.21; x += 0.2) capPos.push([x, -1.8, 1], [x, 1.8, 1]);
    const capBody = kit.instanced(new T.BoxGeometry(0.12, 0.05, 0.065), kit.M.ceramic(0x93724a), capPos.length, function (i, d) {
      d.position.set(capPos[i][0], Y_SUB_TOP + 0.025, capPos[i][1]);
      d.rotation.y = capPos[i][2] ? 0 : Math.PI / 2;
    });
    const capEnd = kit.instanced(new T.BoxGeometry(0.026, 0.054, 0.07), kit.M.solder(), capPos.length * 2, function (i, d) {
      const p = capPos[i >> 1], s = i & 1 ? 1 : -1, vert = !p[2];
      d.position.set(p[0] + (vert ? 0 : s * 0.047), Y_SUB_TOP + 0.027, p[1] + (vert ? s * 0.047 : 0));
      d.rotation.y = vert ? Math.PI / 2 : 0;
    });
    mod.add(capBody, capEnd);

    /* stiffener / heat-spreader frame */
    const frameMat = new T.MeshPhysicalMaterial({ color: 0xd5dbe6, metalness: 1, roughness: 0.28, clearcoat: 0.3, clearcoatRoughness: 0.2, envMapIntensity: 1.4 });
    const frame = new T.Mesh(ringGeo(4.55, 3.3, 4.3, 3.05, 0.17, 0.09, 0.016), frameMat);
    frame.position.y = Y_SUB_TOP;
    mod.add(frame);
    const frameGlow = new T.Mesh(ringGeo(4.32, 3.07, 4.3, 3.05, 0.01, 0.05, 0.002), kit.M.glow(0x3fa8ff, 1.6));
    frameGlow.position.y = Y_SUB_TOP + 0.158;
    mod.add(frameGlow);

    /* C4 underfill + perimeter bumps */
    const c4Fill = new T.Mesh(kit.roundedBox(IPW + 0.1, 0.075, IPD + 0.1, 0.05, 0.015), kit.M.plastic(0x3a2410, 0.3));
    c4Fill.position.y = Y_SUB_TOP + 0.04;
    mod.add(c4Fill);
    const c4Pos = [];
    for (let x = -IPW / 2 + 0.06; x <= IPW / 2 - 0.05; x += 0.14) { c4Pos.push([x, IPD / 2 + 0.045]); c4Pos.push([x, -IPD / 2 - 0.045]); }
    for (let z = -IPD / 2 + 0.16; z <= IPD / 2 - 0.15; z += 0.14) { c4Pos.push([IPW / 2 + 0.045, z]); c4Pos.push([-IPW / 2 - 0.045, z]); }
    const c4 = kit.instanced(new T.SphereGeometry(0.04, 8, 6), kit.M.metal(0x8c93a3, 0.35), c4Pos.length, function (i, d) {
      d.position.set(c4Pos[i][0], Y_SUB_TOP + 0.04, c4Pos[i][1]);
    });
    mod.add(c4);

    /* ================= INTERPOSER ================= */
    function ipDraw(emissive) {
      return function (c) {
        const r = kit.rng(77);
        if (emissive) { c.fillStyle = "#000"; c.fillRect(-IPW / 2, -IPD / 2, IPW, IPD); }
        else {
          const gr = c.createLinearGradient(-IPW / 2, -IPD / 2, IPW / 2, IPD / 2);
          gr.addColorStop(0, "#4a5f94"); gr.addColorStop(0.5, "#6f88c4"); gr.addColorStop(1, "#43588e");
          c.fillStyle = gr; c.fillRect(-IPW / 2, -IPD / 2, IPW, IPD);
          // fine wiring field
          c.lineWidth = 0.004;
          for (let i = 0; i < 700; i++) {
            const h = r() < 0.5;
            const x = (r() - 0.5) * IPW, y = (r() - 0.5) * IPD, l = 0.05 + r() * 0.3;
            c.strokeStyle = "rgba(" + (150 + r() * 80 | 0) + "," + (180 + r() * 60 | 0) + ",255," + (0.08 + r() * 0.14) + ")";
            c.beginPath(); c.moveTo(x, y); c.lineTo(h ? x + l : x, h ? y : y + l); c.stroke();
          }
          // landing pads peeking around each die
          c.fillStyle = "rgba(240,205,120,0.9)";
          const rects = [[0, 0, SOC + 0.14, SOC + 0.14], [0, 1.12, 1.74, 0.6], [0, -1.12, 1.74, 0.6]];
          HBM.forEach(function (h) { rects.push([h[0], h[1], HBM_W + 0.14, HBM_D + 0.14]); });
          rects.forEach(function (R) {
            c.strokeStyle = "rgba(220,235,255,0.35)"; c.lineWidth = 0.006;
            c.strokeRect(R[0] - R[2] / 2, R[1] - R[3] / 2, R[2], R[3]);
            for (let x = R[0] - R[2] / 2 + 0.02; x < R[0] + R[2] / 2; x += 0.045) for (let y = R[1] - R[3] / 2 + 0.02; y < R[1] + R[3] / 2; y += 0.045) {
              c.beginPath(); c.arc(x, y, 0.009, 0, 6.3); c.fill();
            }
          });
        }
        c.lineCap = "round";
        lanes.forEach(function (L) {
          c.strokeStyle = emissive ? "rgba(70,190,255,0.85)" : "rgba(120,210,255,0.55)";
          c.lineWidth = emissive ? 0.011 : 0.007;
          c.beginPath(); c.moveTo(L.x0, L.z0); c.lineTo(L.x1, L.z1); c.stroke();
        });
      };
    }
    const ipCap = new T.MeshPhysicalMaterial({
      map: worldTex(1024, IPW, IPD, ipDraw(false)),
      emissiveMap: worldTex(1024, IPW, IPD, ipDraw(true)), emissive: 0xffffff, emissiveIntensity: 0.9,
      metalness: 0.7, roughness: 0.22, clearcoat: 0.8, clearcoatRoughness: 0.1, envMapIntensity: 1.0
    });
    const interposer = new T.Mesh(kit.roundedBox(IPW, IPH, IPD, 0.04, 0.012), [ipCap, kit.M.silicon(0x24335a)]);
    interposer.position.y = Y_IP;
    mod.add(interposer);

    /* ================= SoC ================= */
    function socDraw(c, W, H) {
      const r = kit.rng(5);
      c.fillStyle = "#050b1b"; c.fillRect(0, 0, W, H);
      const u = W / 1024;
      c.save(); c.scale(u, u);
      c.lineJoin = "round";
      function blockBox(x, y, w, h, col, a) {
        c.fillStyle = col; c.globalAlpha = a == null ? 0.18 : a; c.fillRect(x, y, w, h);
        c.globalAlpha = 0.95; c.strokeStyle = col; c.lineWidth = 2; c.strokeRect(x, y, w, h); c.globalAlpha = 1;
      }
      // seal ring
      c.strokeStyle = "#c7d8ff"; c.lineWidth = 6; c.strokeRect(10, 10, 1004, 1004);
      c.strokeStyle = "#5a78c4"; c.lineWidth = 2; c.strokeRect(22, 22, 980, 980);
      // memory PHY strips (left & right) - face the HBM stacks
      [30, 928].forEach(function (x) {
        blockBox(x, 40, 66, 944, "#12d4ff", 0.12);
        for (let y = 50; y < 970; y += 13) {
          c.fillStyle = "rgba(18,212,255," + (0.35 + r() * 0.5) + ")";
          c.fillRect(x + 6, y, 54, 8);
          c.fillStyle = "#c9f4ff";
          c.fillRect(x + 8 + (r() * 40 | 0), y + 2, 4, 4);
        }
      });
      // SerDes / IO strips (top & bottom)
      [30, 928].forEach(function (y) {
        blockBox(110, y, 804, 66, "#ffb340", 0.1);
        for (let x = 120; x < 904; x += 28) {
          c.fillStyle = "rgba(255,179,64," + (0.3 + r() * 0.5) + ")";
          c.fillRect(x, y + 8, 22, 50);
          c.strokeStyle = "rgba(255,220,150,0.8)"; c.lineWidth = 1;
          c.beginPath(); c.arc(x + 11, y + 33, 7, 0, 6.3); c.stroke();
        }
      });
      // CPU cluster 4x2
      blockBox(110, 110, 804, 250, "#2f7bff", 0.08);
      for (let i = 0; i < 4; i++) for (let j = 0; j < 2; j++) {
        const x = 124 + i * 196, y = 124 + j * 116;
        blockBox(x, y, 186, 106, "#4d8dff", 0.22);
        // L1 sram
        c.fillStyle = "rgba(255,179,64,0.65)";
        for (let a = 0; a < 6; a++) for (let b = 0; b < 4; b++) c.fillRect(x + 8 + a * 9, y + 8 + b * 9, 6, 6);
        // ALU / FPU lanes
        for (let k = 0; k < 9; k++) {
          c.fillStyle = "rgba(" + (90 + r() * 40 | 0) + "," + (170 + r() * 60 | 0) + ",255," + (0.4 + r() * 0.5) + ")";
          c.fillRect(x + 76, y + 8 + k * 10, 20 + r() * 80, 6);
        }
      }
      // shared L3 cache band
      blockBox(110, 372, 804, 96, "#ffb340", 0.1);
      for (let x = 118; x < 906; x += 11) for (let y = 380; y < 460; y += 10) {
        c.fillStyle = "rgba(255," + (170 + r() * 50 | 0) + ",70," + (0.35 + r() * 0.45) + ")";
        c.fillRect(x, y, 8, 7);
      }
      // NPU MAC array
      blockBox(110, 480, 804, 410, "#ff4fd8", 0.07);
      for (let i = 0; i < 48; i++) for (let j = 0; j < 24; j++) {
        const x = 122 + i * 16.3, y = 492 + j * 16.3;
        const hot = r();
        c.fillStyle = hot > 0.8 ? "rgba(255,255,255,0.9)" : "hsla(" + (290 + (i / 48) * 60) + ",100%," + (50 + hot * 20) + "%," + (0.45 + hot * 0.45) + ")";
        c.fillRect(x, y, 12, 12);
        c.fillStyle = "rgba(5,11,27,0.6)"; c.fillRect(x + 4, y + 4, 4, 4);
      }
      // NoC mesh
      c.strokeStyle = "rgba(210,225,255,0.7)"; c.lineWidth = 2;
      c.beginPath();
      [372, 472, 892].forEach(function (y) { c.moveTo(110, y); c.lineTo(914, y); });
      [110, 312, 508, 704, 914].forEach(function (x) { c.moveTo(x, 360); c.lineTo(x, 480); });
      c.stroke();
      c.fillStyle = "rgba(255,255,255,0.88)";
      c.font = "700 34px monospace"; c.textBaseline = "middle";
      c.fillText("CPU x8", 120, 150 - 6 + 130);
      c.fillText("NPU  MAC ARRAY", 330, 935 - 20 - 80 - 10);
      c.font = "600 26px monospace"; c.fillText("L3 CACHE", 130, 420);
      c.restore();
    }
    const socTex = kit.canvasTex(1024, 1024, socDraw);
    capUV(socTex, SOC, SOC);
    const socMat = new T.MeshPhysicalMaterial({
      map: socTex, emissiveMap: socTex, emissive: 0xffffff, emissiveIntensity: 0.45,
      metalness: 0.6, roughness: 0.22, clearcoat: 0.9, clearcoatRoughness: 0.06, envMapIntensity: 0.9
    });
    const SOC_H = 0.16;
    const soc = new T.Mesh(kit.roundedBox(SOC, SOC_H, SOC, 0.02, 0.012), [socMat, kit.M.silicon(0x1d2850)]);
    soc.position.y = Y_DIE + SOC_H / 2;
    mod.add(soc);
    const socRimMat = kit.M.glow(0x44d0ff, 2.6);
    const socRim = new T.Mesh(ringGeo(SOC + 0.045, SOC + 0.045, SOC - 0.015, SOC - 0.015, 0.012, 0.025, 0.003), socRimMat);
    socRim.position.y = Y_DIE + SOC_H - 0.004;
    mod.add(socRim);
    const halo = kit.glow(0x3fa0ff, 2.4, 0.4);
    halo.position.y = Y_DIE + SOC_H + 0.05;
    mod.add(halo);

    /* I/O chiplets */
    const ioTex = kit.tex.dieLayout(512, 41, ["#08101f", "#ffb340", "#ff6bd8", "#12d4ff", "#e6dcc0"]);
    capUV(ioTex, 1.74, 0.6);
    const ioMat = new T.MeshPhysicalMaterial({ map: ioTex, emissiveMap: ioTex, emissive: 0xffffff, emissiveIntensity: 0.3, metalness: 0.6, roughness: 0.25, clearcoat: 0.8, clearcoatRoughness: 0.1, envMapIntensity: 0.9 });
    [1.12, -1.12].forEach(function (z) {
      const io = new T.Mesh(kit.roundedBox(1.74, 0.12, 0.46, 0.02, 0.01), [ioMat, kit.M.silicon(0x1d2850)]);
      // the cap UV window is 1.74 x 0.46; the texture was sized for 0.6, so squash via repeat
      io.position.set(0, Y_DIE + 0.06, z);
      mod.add(io);
    });
    ioTex.repeat.set(1 / 1.74, 1 / 0.46);

    /* ================= HBM STACKS ================= */
    const hbmTex = worldTex(512, HBM_W, HBM_D, function (c) {
      const r = kit.rng(8);
      c.fillStyle = "#0a1330"; c.fillRect(-0.3, -0.465, 0.6, 0.93);
      c.lineWidth = 0.004;
      for (let side = -1; side <= 1; side += 2) for (let k = 0; k < 4; k++) {
        const x = side > 0 ? 0.045 : -0.285, y = -0.43 + k * 0.215;
        c.fillStyle = "rgba(60,120,255,0.28)"; c.fillRect(x, y, 0.24, 0.2);
        c.strokeStyle = "rgba(110,170,255,0.9)"; c.strokeRect(x, y, 0.24, 0.2);
        for (let yy = y + 0.012; yy < y + 0.2; yy += 0.012) {
          c.strokeStyle = "rgba(150,200,255," + (0.25 + r() * 0.25) + ")";
          c.beginPath(); c.moveTo(x + 0.004, yy); c.lineTo(x + 0.236, yy); c.stroke();
        }
      }
      // TSV spine
      c.fillStyle = "rgba(255,179,64,0.2)"; c.fillRect(-0.04, -0.45, 0.08, 0.9);
      c.fillStyle = "#ffd27a";
      for (let y = -0.44; y < 0.45; y += 0.0225) for (let x = -0.026; x <= 0.027; x += 0.026) { c.beginPath(); c.arc(x, y, 0.0063, 0, 6.3); c.fill(); }
      c.fillStyle = "rgba(255,255,255,0.8)";
      c.font = "bold 0.05px monospace";
      c.save(); c.translate(-0.15, 0); c.rotate(-Math.PI / 2); c.fillText("HBM3", -0.06, 0.02); c.restore();
    });
    const hbmCap = new T.MeshPhysicalMaterial({
      map: hbmTex, emissiveMap: hbmTex, emissive: 0xffffff, emissiveIntensity: 0.35,
      metalness: 0.65, roughness: 0.2, clearcoat: 0.9, clearcoatRoughness: 0.08, envMapIntensity: 0.9
    });
    const NS = 8;
    const sliceGeo = kit.roundedBox(HBM_W, 0.032, HBM_D, 0.012, 0.006);
    const slices = new T.InstancedMesh(sliceGeo, [hbmCap, kit.M.silicon(0x1a2550)], NS * 4);
    const spacers = new T.InstancedMesh(new T.BoxGeometry(HBM_W - 0.03, 0.013, HBM_D - 0.03), new T.MeshStandardMaterial({ color: 0xd98a2a, emissive: 0xb35a00, emissiveIntensity: 0.55, metalness: 0.6, roughness: 0.3 }), 7 * 4);
    const bases = new T.InstancedMesh(kit.roundedBox(HBM_W + 0.04, 0.05, HBM_D + 0.04, 0.015, 0.008), kit.M.silicon(0x1a2447), 4);
    const dm = new T.Object3D();
    const slY0 = Y_DIE + 0.05 + 0.016 + 0.004;
    const tsvPos = [];
    HBM.forEach(function (h, s) {
      dm.rotation.set(0, 0, 0); dm.scale.set(1, 1, 1);
      dm.position.set(h[0], Y_DIE + 0.025, h[1]); dm.updateMatrix(); bases.setMatrixAt(s, dm.matrix);
      for (let i = 0; i < NS; i++) {
        dm.position.set(h[0], slY0 + i * 0.047, h[1]); dm.updateMatrix(); slices.setMatrixAt(s * NS + i, dm.matrix);
        if (i > 0) {
          dm.position.set(h[0], slY0 + i * 0.047 - 0.0235, h[1]); dm.updateMatrix(); spacers.setMatrixAt(s * 7 + i - 1, dm.matrix);
          // micro-bump / TSV dots around the perimeter of this gap
          const yy = slY0 + i * 0.047 - 0.0235;
          for (let x = -HBM_W / 2 + 0.05; x <= HBM_W / 2 - 0.04; x += 0.055) { tsvPos.push([h[0] + x, yy, h[1] + HBM_D / 2 - 0.004, 0]); tsvPos.push([h[0] + x, yy, h[1] - HBM_D / 2 + 0.004, 0]); }
          for (let z = -HBM_D / 2 + 0.05; z <= HBM_D / 2 - 0.04; z += 0.055) { tsvPos.push([h[0] + HBM_W / 2 - 0.004, yy, h[1] + z, 1]); tsvPos.push([h[0] - HBM_W / 2 + 0.004, yy, h[1] + z, 1]); }
        }
      }
    });
    const tsv = kit.instanced(new T.BoxGeometry(0.024, 0.017, 0.024), kit.M.gold(), tsvPos.length, function (i, d) {
      d.position.set(tsvPos[i][0], tsvPos[i][1], tsvPos[i][2]);
    });
    mod.add(bases, slices, spacers, tsv);

    /* ================= DATA PULSES ================= */
    const PY = Y_IP_TOP + 0.004;
    const pg = new T.PlaneGeometry(0.15, 0.02);
    pg.rotateX(-Math.PI / 2);
    const hbmLanes = lanes.filter(function (l) { return l.hbm; });
    const ioLanes = lanes.filter(function (l) { return !l.hbm; });
    const pulsesCy = new T.InstancedMesh(pg, kit.M.additive(0x7fdcff, 0.95), hbmLanes.length * 2);
    const pulsesAm = new T.InstancedMesh(pg, kit.M.additive(0xffb860, 0.95), ioLanes.length * 2);
    pulsesCy.frustumCulled = pulsesAm.frustumCulled = false;
    pulsesCy.userData.noFit = pulsesAm.userData.noFit = true;
    mod.add(pulsesCy, pulsesAm);
    function updPulses(im, arr, t, spd) {
      for (let i = 0; i < arr.length; i++) {
        const L = arr[i];
        for (let k = 0; k < 2; k++) {
          let p = (t * spd + i * 0.137 + k * 0.5) % 1;
          const q = L.dir ? 1 - p : p;
          dm.position.set(L.x0 + (L.x1 - L.x0) * q, PY, L.z0 + (L.z1 - L.z0) * q);
          dm.rotation.set(0, L.hbm ? 0 : Math.PI / 2, 0);
          const e = Math.sin(p * Math.PI);
          dm.scale.set(0.35 + e, 1, 1);
          dm.updateMatrix();
          im.setMatrixAt(i * 2 + k, dm.matrix);
        }
      }
      im.instanceMatrix.needsUpdate = true;
    }

    /* ================= LABELS ================= */
    function label(str, x, y, z, ly) {
      const t = kit.tex.text(str, { size: 72, color: "#d6e6ff", weight: "600" });
      const sp = new T.Sprite(new T.SpriteMaterial({ map: t.tex, transparent: true, depthWrite: false, depthTest: false, toneMapped: false, opacity: 0.92 }));
      sp.scale.set(0.2 * t.aspect, 0.2, 1);
      sp.position.set(x, y, z);
      sp.userData.noFit = true;
      const ln = new T.Line(new T.BufferGeometry().setFromPoints([new T.Vector3(x, y - 0.12, z), new T.Vector3(x, ly, z)]), new T.LineBasicMaterial({ color: 0x7fb0ff, transparent: true, opacity: 0.5 }));
      ln.userData.noFit = true;
      mod.add(sp, ln);
    }
    label("AI SoC", 0, 1.55, 0, Y_DIE + SOC_H);
    label("HBM3 stack", -2.1, 1.35, 0.55, Y_DIE + 0.7);
    label("HBM3 stack", 2.1, 1.35, -0.55, Y_DIE + 0.7);
    label("Si interposer", 1.1, 0.7, 1.8, Y_IP_TOP);

    const sh = kit.contactShadow(3.6, 0.6);
    sh.position.y = -0.12;
    mod.add(sh);
    sh.rotation.x = -Math.PI / 2;

    return {
      group: g,
      fit: 2.05,
      spin: 0.0022,
      update: function (t) {
        updPulses(pulsesCy, hbmLanes, t, 0.55);
        updPulses(pulsesAm, ioLanes, t, 0.7);
        socMat.emissiveIntensity = 0.42 + 0.14 * Math.sin(t * 1.6);
        socRimMat.emissiveIntensity = 2.4 + 0.9 * Math.sin(t * 2.2);
        halo.material.opacity = 0.34 + 0.1 * Math.sin(t * 1.6);
        ipCap.emissiveIntensity = 0.85 + 0.25 * Math.sin(t * 2.4);
        mod.rotation.x = 0.4 + 0.05 * Math.sin(t * 0.3);
      }
    };
  };
})();
