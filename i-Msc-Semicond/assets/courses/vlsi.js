/* elective_1 - VLSI & IoT Design : naked IC die (photo-micrograph look), gold bond wires, lead frame, wireless IoT motif */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  window.IMEN.courseScenes = window.IMEN.courseScenes || {};

  window.IMEN.courseScenes.vlsi = function (kit) {
    const T = THREE;
    const g = new T.Group();
    const rig = new T.Group();
    rig.rotation.x = 0.32;
    g.add(rig);

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
    function mergeGeos(list) {
      let vc = 0, ic = 0;
      list.forEach(function (q) { vc += q.attributes.position.count; ic += q.index.count; });
      const pos = new Float32Array(vc * 3), nor = new Float32Array(vc * 3), idx = new Uint32Array(ic);
      let vo = 0, io = 0;
      list.forEach(function (q) {
        pos.set(q.attributes.position.array, vo * 3);
        nor.set(q.attributes.normal.array, vo * 3);
        for (let i = 0; i < q.index.count; i++) idx[io + i] = q.index.array[i] + vo;
        vo += q.attributes.position.count;
        io += q.index.count;
        q.dispose();
      });
      const m = new T.BufferGeometry();
      m.setAttribute("position", new T.BufferAttribute(pos, 3));
      m.setAttribute("normal", new T.BufferAttribute(nor, 3));
      m.setIndex(new T.BufferAttribute(idx, 1));
      return m;
    }

    /* ---------- constants ---------- */
    const DIE = 2.6, DH = 0.16, HALF = DIE / 2;
    const Y_LF = 0.05;                 // top of lead frame
    const Y_DIE0 = Y_LF + 0.025;       // die bottom
    const Y_TOP = Y_DIE0 + DH;         // die top
    const NPAD = 20, PITCH = 0.11, PADV = 1.2;
    const TIPV = 1.62, OUTV = 1.97;

    /* ================= DIE FLOORPLAN TEXTURE ================= */
    const routes = [];   // channel segments (canvas px) for signal pulses
    function dieDraw(c, W, H) {
      const r = kit.rng(23);
      const bg = c.createLinearGradient(0, 0, W, H);
      bg.addColorStop(0, "#0b1230"); bg.addColorStop(0.5, "#101a42"); bg.addColorStop(1, "#080d26");
      c.fillStyle = bg; c.fillRect(0, 0, W, H);
      // power mesh
      c.lineWidth = 6;
      for (let x = 90; x < W - 80; x += 64) { c.strokeStyle = "rgba(90,150,255,0.10)"; c.beginPath(); c.moveTo(x, 80); c.lineTo(x, H - 80); c.stroke(); }
      for (let y = 90; y < H - 80; y += 64) { c.strokeStyle = "rgba(255,190,90,0.10)"; c.beginPath(); c.moveTo(80, y); c.lineTo(W - 80, y); c.stroke(); }

      const blocks = [];
      function split(x, y, w, h, d) {
        if (d >= 5 || (w < 130 && h < 130) || (d >= 2 && r() < 0.14)) { blocks.push([x, y, w, h]); return; }
        const horiz = w > h * 1.25 ? false : (h > w * 1.25 ? true : r() < 0.5);
        const f = 0.36 + r() * 0.28;
        if (horiz) { split(x, y, w, h * f, d + 1); split(x, y + h * f, w, h * (1 - f), d + 1); }
        else { split(x, y, w * f, h, d + 1); split(x + w * f, y, w * (1 - f), h, d + 1); }
      }
      split(90, 90, W - 180, H - 180, 0);
      const hues = [35, 205, 170, 305, 250, 120];
      blocks.forEach(function (b, bi) {
        const x = b[0] + 4, y = b[1] + 4, w = b[2] - 8, h = b[3] - 8;
        if (w < 20 || h < 20) return;
        routes.push({ o: 0, f: y - 4, a: b[0], b: b[0] + b[2] });
        routes.push({ o: 1, f: x - 4, a: b[1], b: b[1] + b[3] });
        const type = (bi * 7 + (r() * 6 | 0)) % 6;
        const hue = hues[type];
        c.fillStyle = "hsla(" + hue + ",85%,55%,0.16)"; c.fillRect(x, y, w, h);
        c.strokeStyle = "hsla(" + hue + ",95%,68%,0.95)"; c.lineWidth = 2; c.strokeRect(x, y, w, h);
        c.save();
        c.beginPath(); c.rect(x + 2, y + 2, w - 4, h - 4); c.clip();
        if (type === 0) {            // SRAM: dense bit-cell array
          for (let yy = y + 4; yy < y + h - 3; yy += 5) for (let xx = x + 14; xx < x + w - 3; xx += 5) {
            c.fillStyle = "hsla(" + (30 + r() * 22) + ",95%," + (45 + r() * 25) + "%," + (0.4 + r() * 0.5) + ")";
            c.fillRect(xx, yy, 3.4, 3.4);
          }
          c.fillStyle = "hsla(35,90%,40%,0.5)"; c.fillRect(x + 2, y + 2, 10, h - 4);
          c.strokeStyle = "rgba(255,230,170,0.5)"; c.lineWidth = 1;
          for (let xx = x + 14 + 40; xx < x + w; xx += 40) { c.beginPath(); c.moveTo(xx, y); c.lineTo(xx, y + h); c.stroke(); }
        } else if (type === 1) {     // standard-cell rows
          for (let yy = y + 3, k = 0; yy < y + h - 8; yy += 9, k++) {
            c.fillStyle = k % 2 ? "rgba(40,110,255,0.35)" : "rgba(20,200,255,0.25)";
            c.fillRect(x + 2, yy, w - 4, 8);
            c.strokeStyle = "rgba(190,225,255,0.55)"; c.lineWidth = 1;
            for (let xx = x + 4 + r() * 8; xx < x + w - 2; xx += 5 + r() * 14) { c.beginPath(); c.moveTo(xx, yy); c.lineTo(xx, yy + 8); c.stroke(); }
          }
        } else if (type === 2) {     // analog: capacitor array + resistor snakes
          for (let yy = y + 6; yy < y + h - 20; yy += 20) for (let xx = x + 6; xx < x + w * 0.55 - 14; xx += 20) {
            c.fillStyle = "hsla(" + (160 + r() * 30) + ",80%,55%,0.55)"; c.fillRect(xx, yy, 15, 15);
            c.strokeStyle = "rgba(220,255,240,0.8)"; c.strokeRect(xx + 3, yy + 3, 9, 9);
          }
          c.strokeStyle = "rgba(120,255,210,0.85)"; c.lineWidth = 2;
          for (let yy = y + 8; yy < y + h - 12; yy += 26) {
            c.beginPath(); c.moveTo(x + w * 0.58, yy);
            for (let k = 0; k < 8; k++) { c.lineTo(x + w * 0.58 + (k % 2 ? 6 : w * 0.38), yy + k * 3); }
            c.stroke();
          }
        } else if (type === 3) {     // RF block: square spiral inductor
          const cx = x + w / 2, cy = y + h / 2, R = Math.min(w, h) / 2 - 6;
          c.strokeStyle = "rgba(255,90,220,0.95)"; c.lineWidth = 3;
          c.beginPath();
          for (let k = 0; k < 6; k++) {
            const a = R - k * (R / 7);
            c.rect(cx - a, cy - a, a * 2, a * 2);
          }
          c.stroke();
          c.fillStyle = "#ffd0f4"; c.fillRect(cx - 5, cy - 5, 10, 10);
          c.strokeStyle = "rgba(255,190,240,0.5)"; c.lineWidth = 1; c.strokeRect(x + 5, y + 5, w - 10, h - 10);
        } else if (type === 4) {     // dense logic / DSP
          c.lineWidth = 1.2;
          for (let yy = y + 3; yy < y + h; yy += 4) {
            c.strokeStyle = "hsla(" + (240 + r() * 40) + ",90%,70%," + (0.3 + r() * 0.5) + ")";
            c.beginPath(); c.moveTo(x + 2, yy); c.lineTo(x + w - 2, yy); c.stroke();
          }
          for (let xx = x + 6; xx < x + w; xx += 12) {
            c.strokeStyle = "rgba(190,150,255," + (0.2 + r() * 0.4) + ")";
            c.beginPath(); c.moveTo(xx, y + 2); c.lineTo(xx, y + h - 2); c.stroke();
          }
        } else {                     // PLL / ring oscillator
          const cx = x + w / 2, cy = y + h / 2, R = Math.min(w, h) / 2 - 8;
          c.strokeStyle = "rgba(120,255,160,0.9)"; c.lineWidth = 2;
          [R, R * 0.7, R * 0.4].forEach(function (rr) { c.beginPath(); c.arc(cx, cy, rr, 0, 6.3); c.stroke(); });
          c.fillStyle = "#c8ffd8";
          for (let k = 0; k < 18; k++) { const a = k / 18 * 6.283; c.fillRect(cx + Math.cos(a) * R * 0.85 - 3, cy + Math.sin(a) * R * 0.85 - 3, 6, 6); }
        }
        c.restore();
      });
      // I/O cells + pad openings around the edge
      for (let side = 0; side < 4; side++) {
        c.save(); c.translate(W / 2, H / 2); c.rotate(side * Math.PI / 2); c.translate(-W / 2, -H / 2);
        for (let k = 0; k < NPAD; k++) {
          const px = W / 2 + (k - 9.5) * (PITCH / DIE * W);
          c.fillStyle = "rgba(255,190,80,0.22)"; c.fillRect(px - 20, 6, 40, 72);
          c.strokeStyle = "rgba(255,210,120,0.8)"; c.lineWidth = 1.5; c.strokeRect(px - 20, 6, 40, 72);
          c.fillStyle = "rgba(255,235,170,0.95)"; c.fillRect(px - 17, 22, 34, 34);
        }
        c.restore();
      }
      // seal ring + corner alignment marks
      c.strokeStyle = "#d9e4ff"; c.lineWidth = 5; c.strokeRect(8, 8, W - 16, H - 16);
      c.strokeStyle = "#6c86d8"; c.lineWidth = 2; c.strokeRect(20, 20, W - 40, H - 40);
      c.strokeStyle = "rgba(255,255,255,0.9)"; c.lineWidth = 3;
      [[96, 96], [W - 96, 96], [96, H - 96], [W - 96, H - 96]].forEach(function (p) {
        c.beginPath(); c.moveTo(p[0] - 14, p[1]); c.lineTo(p[0] + 14, p[1]); c.moveTo(p[0], p[1] - 14); c.lineTo(p[0], p[1] + 14); c.stroke();
      });
      c.fillStyle = "rgba(255,255,255,0.9)"; c.font = "700 22px monospace";
      c.fillText("IMEN  VLSI-IoT  rev.A", 110, H - 100);
      // thin-film interference: broad hue-shifted washes
      c.globalCompositeOperation = "soft-light";
      const film = c.createLinearGradient(0, 0, W, H);
      [["0", "#ff3da8"], ["0.2", "#ffb340"], ["0.4", "#3dff8f"], ["0.6", "#3dc8ff"], ["0.8", "#7a5cff"], ["1", "#ff3da8"]].forEach(function (s) { film.addColorStop(+s[0], s[1]); });
      c.globalAlpha = 0.38; c.fillStyle = film; c.fillRect(0, 0, W, H);
      c.globalAlpha = 1; c.globalCompositeOperation = "source-over";
    }
    const dieTex = kit.canvasTex(1024, 1024, dieDraw);
    capUV(dieTex, DIE, DIE);
    const dieMat = new T.MeshPhysicalMaterial({
      map: dieTex, emissiveMap: dieTex, emissive: 0xffffff, emissiveIntensity: 0.28,
      metalness: 0.55, roughness: 0.2, clearcoat: 1, clearcoatRoughness: 0.05, envMapIntensity: 1.0
    });
    const die = new T.Mesh(kit.roundedBox(DIE, DH, DIE, 0.02, 0.012), [dieMat, kit.M.silicon(0x7d89aa)]);
    die.position.y = Y_DIE0 + DH / 2;
    rig.add(die);

    // iridescent sheen sweeping over the die
    const sheenTex = kit.canvasTex(256, 256, function (c, W, H) {
      const gr = c.createLinearGradient(0, 0, W, 0);
      for (let i = 0; i <= 6; i++) gr.addColorStop(i / 6, "hsla(" + (i / 6 * 360) + ",100%,60%,1)");
      c.fillStyle = gr; c.fillRect(0, 0, W, H);
      const fade = c.createLinearGradient(0, 0, 0, H);
      fade.addColorStop(0, "rgba(0,0,0,1)"); fade.addColorStop(0.5, "rgba(0,0,0,0.25)"); fade.addColorStop(1, "rgba(0,0,0,1)");
      c.fillStyle = fade; c.fillRect(0, 0, W, H);
    });
    sheenTex.wrapS = T.RepeatWrapping;
    sheenTex.repeat.set(1, 1);
    const sheen = new T.Mesh(new T.PlaneGeometry(DIE - 0.02, DIE - 0.02), kit.M.additive(0xffffff, 0.2, sheenTex));
    sheen.rotation.x = -Math.PI / 2;
    sheen.position.y = Y_TOP + 0.0025;
    sheen.userData.noFit = true;
    rig.add(sheen);

    // die-attach slab
    const attach = new T.Mesh(kit.roundedBox(DIE + 0.1, 0.03, DIE + 0.1, 0.03, 0.008), new T.MeshPhysicalMaterial({ color: 0xa9b1c0, metalness: 0.9, roughness: 0.45 }));
    attach.position.y = Y_LF + 0.005;
    rig.add(attach);

    /* ================= LEAD FRAME ================= */
    const cuMat = new T.MeshPhysicalMaterial({ color: 0xe0a37c, metalness: 1, roughness: 0.28, clearcoat: 0.2, clearcoatRoughness: 0.25, envMapIntensity: 1.5 });
    const agMat = new T.MeshPhysicalMaterial({ color: 0xe4e8ef, metalness: 1, roughness: 0.18, envMapIntensity: 1.5 });
    const paddle = new T.Mesh(kit.roundedBox(2.95, 0.05, 2.95, 0.06, 0.01), agMat);
    paddle.position.y = Y_LF / 2;
    rig.add(paddle);
    const frame = new T.Mesh(ringGeo(4.3, 4.3, 3.92, 3.92, 0.05, 0.1, 0.01), cuMat);
    rig.add(frame);
    // tie bars (diagonal)
    for (let i = 0; i < 4; i++) {
      const a = Math.PI / 4 + i * Math.PI / 2;
      const bar = new T.Mesh(new T.BoxGeometry(0.1, 0.05, 0.85), cuMat);
      const rc = 2.4;
      bar.position.set(Math.sin(a) * rc * 1.0, 0.025, Math.cos(a) * rc * 1.0);
      bar.rotation.y = a;
      rig.add(bar);
    }
    // fingers + pads + wires
    const fingerData = [], padData = [];
    for (let side = 0; side < 4; side++) {
      const a = side * Math.PI / 2, ca = Math.cos(a), sa = Math.sin(a);
      const W = function (u, v) { return [u * ca + v * sa, -u * sa + v * ca]; };
      for (let k = 0; k < NPAD; k++) {
        const u = (k - 9.5) * PITCH;
        const p0 = W(u, TIPV), p1 = W(u * 1.14, OUTV);
        const pd = W(u, PADV);
        fingerData.push({ p0: p0, p1: p1 });
        padData.push({ pd: pd, p0: p0 });
      }
    }
    const NF = fingerData.length;
    const fingers = kit.instanced(new T.BoxGeometry(0.062, 0.05, 1), cuMat, NF, function (i, d) {
      const f = fingerData[i];
      const dx = f.p1[0] - f.p0[0], dz = f.p1[1] - f.p0[1], L = Math.hypot(dx, dz);
      d.position.set((f.p0[0] + f.p1[0]) / 2, 0.025, (f.p0[1] + f.p1[1]) / 2);
      d.rotation.y = Math.atan2(dx, dz);
      d.scale.set(1, 1, L);
    });
    rig.add(fingers);
    const fTips = kit.instanced(new T.BoxGeometry(0.07, 0.056, 0.16), agMat, NF, function (i, d) {
      const f = fingerData[i];
      const dx = f.p1[0] - f.p0[0], dz = f.p1[1] - f.p0[1], L = Math.hypot(dx, dz);
      d.position.set(f.p0[0] + dx / L * 0.07, 0.025, f.p0[1] + dz / L * 0.07);
      d.rotation.y = Math.atan2(dx, dz);
    });
    rig.add(fTips);
    // gold bond pads on the die
    const padMeshes = kit.instanced(new T.BoxGeometry(0.075, 0.014, 0.075), kit.M.gold(), NF, function (i, d) {
      d.position.set(padData[i].pd[0], Y_TOP + 0.007, padData[i].pd[1]);
      d.rotation.y = Math.atan2(padData[i].p0[0] - padData[i].pd[0], 0) * 0; // pads are square; keep axis-aligned
    });
    rig.add(padMeshes);
    // ball bonds
    const ballBonds = kit.instanced(new T.SphereGeometry(0.03, 10, 7), kit.M.gold(), NF, function (i, d) {
      d.position.set(padData[i].pd[0], Y_TOP + 0.016, padData[i].pd[1]);
      d.scale.set(1, 0.7, 1);
    });
    rig.add(ballBonds);
    // bond wires (merged into one mesh)
    const wireGeos = [];
    const rr = kit.rng(99);
    for (let i = 0; i < NF; i++) {
      const pd = padData[i].pd, tp = padData[i].p0;
      const A = new T.Vector3(pd[0], Y_TOP + 0.02, pd[1]);
      const B = new T.Vector3(tp[0], Y_LF + 0.032, tp[1]);
      const dir = new T.Vector3().subVectors(B, A);
      const hgt = 0.18 + rr() * 0.07;
      const C1 = new T.Vector3().copy(A).addScaledVector(dir, 0.18); C1.y += hgt;
      const C2 = new T.Vector3().copy(A).addScaledVector(dir, 0.72); C2.y += hgt * 0.55;
      const curve = new T.CubicBezierCurve3(A, C1, C2, B);
      wireGeos.push(new T.TubeGeometry(curve, 18, 0.0072, 5, false));
    }
    const wires = new T.Mesh(mergeGeos(wireGeos), new T.MeshPhysicalMaterial({ color: 0xffc94d, metalness: 1, roughness: 0.14, envMapIntensity: 1.7 }));
    rig.add(wires);

    /* ================= SIGNAL PULSES ON THE DIE ================= */
    const toW = function (px) { return (px / 1024 - 0.5) * DIE; };
    const NPUL = 36;
    const pg = new T.PlaneGeometry(0.2, 0.016);
    pg.rotateX(-Math.PI / 2);
    const pulseA = new T.InstancedMesh(pg, kit.M.additive(0x7fe6ff, 1), NPUL);
    const pulseB = new T.InstancedMesh(pg, kit.M.additive(0xffc15a, 1), NPUL);
    pulseA.frustumCulled = pulseB.frustumCulled = false;
    pulseA.userData.noFit = pulseB.userData.noFit = true;
    rig.add(pulseA, pulseB);
    const pr = kit.rng(5);
    const pulseInfo = [];
    for (let i = 0; i < NPUL * 2; i++) {
      pulseInfo.push({ r: routes[(pr() * routes.length) | 0], ph: pr(), sp: 0.12 + pr() * 0.2 });
    }
    const dm = new T.Object3D();
    const PY = Y_TOP + 0.005;
    function updPulse(im, off, t) {
      for (let i = 0; i < NPUL; i++) {
        const P = pulseInfo[off + i], R = P.r;
        const p = (t * P.sp + P.ph) % 1;
        const along = toW(R.a + (R.b - R.a) * p);
        const fix = toW(R.f);
        dm.position.set(R.o ? fix : along, PY, R.o ? along : fix);
        dm.rotation.set(0, R.o ? Math.PI / 2 : 0, 0);
        dm.scale.set(0.5 + Math.sin(p * Math.PI), 1, 1);
        dm.updateMatrix();
        im.setMatrixAt(i, dm.matrix);
      }
      im.instanceMatrix.needsUpdate = true;
    }

    /* ================= IoT / RADIO MOTIF ================= */
    const NODE_Y = 1.15;
    const core = new T.Mesh(new T.SphereGeometry(0.075, 16, 12), kit.M.glow(0x7fe6ff, 3));
    core.position.set(0, NODE_Y, 0);
    core.userData.noFit = true;
    rig.add(core);
    const coreGlow = kit.glow(0x4cc8ff, 0.9, 0.9);
    coreGlow.position.copy(core.position);
    rig.add(coreGlow);
    const antRing = new T.Mesh(new T.TorusGeometry(0.17, 0.008, 8, 40), kit.M.glow(0x9fe6ff, 2.2));
    antRing.position.copy(core.position);
    antRing.userData.noFit = true;
    rig.add(antRing);
    const shaft = kit.beam(0.55, 0.02, NODE_Y - Y_TOP, "110,200,255", 0.2);
    shaft.position.y = Y_TOP + (NODE_Y - Y_TOP) / 2;
    rig.add(shaft);

    const waves = [];
    for (let i = 0; i < 5; i++) {
      const m = new T.Mesh(new T.RingGeometry(0.965, 1.0, 96), kit.M.additive(0x6fd2ff, 0.5));
      m.rotation.x = -Math.PI / 2;
      m.position.set(0, NODE_Y, 0);
      m.userData.noFit = true;
      rig.add(m);
      waves.push(m);
    }
    // vertical radio arcs ("wi-fi" glyph) pulsing above the node
    const arcs = [];
    for (let i = 0; i < 3; i++) {
      const m = new T.Mesh(new T.RingGeometry(0.28 + i * 0.16, 0.3 + i * 0.16, 40, 1, Math.PI * 0.25, Math.PI * 0.5), kit.M.additive(0xc6f2ff, 0.6));
      m.position.set(0, NODE_Y + 0.28, 0);
      m.userData.noFit = true;
      rig.add(m);
      arcs.push(m);
    }
    // remote IoT nodes with data links
    const NN = 4;
    const nodes = [], nodeLinks = [], linkPulses = [];
    const hexGeo = new T.TorusGeometry(0.11, 0.014, 8, 6);
    const dotGeo = new T.SphereGeometry(0.035, 10, 8);
    const hexMat = kit.M.glow(0x8fe0ff, 1.8), dotMat = kit.M.glow(0xffc15a, 2.6);
    const nodeGlow = kit.glow(0xffb85a, 0.55, 0.85);
    for (let i = 0; i < NN; i++) {
      const grp = new T.Group();
      const hx = new T.Mesh(hexGeo, hexMat); hx.userData.noFit = true;
      const dt = new T.Mesh(dotGeo, dotMat); dt.userData.noFit = true;
      const gl = nodeGlow.clone();
      grp.add(hx, dt, gl);
      rig.add(grp);
      nodes.push(grp);
      const lg = new T.BufferGeometry();
      lg.setAttribute("position", new T.BufferAttribute(new Float32Array(6), 3));
      const ln = new T.Line(lg, new T.LineBasicMaterial({ color: 0x6fc8ff, transparent: true, opacity: 0.4, depthWrite: false }));
      ln.frustumCulled = false;
      ln.userData.noFit = true;
      rig.add(ln);
      nodeLinks.push(ln);
      const lp = kit.glow(0xe0f8ff, 0.2, 1);
      rig.add(lp);
      linkPulses.push(lp);
    }

    const sh = kit.contactShadow(3.4, 0.6);
    sh.position.y = -0.04;
    g.add(sh);

    return {
      group: g,
      fit: 2.05,
      spin: 0.0022,
      update: function (t) {
        updPulse(pulseA, 0, t);
        updPulse(pulseB, NPUL, t);
        sheenTex.offset.x = (t * 0.03) % 1;
        sheen.material.opacity = 0.14 + 0.07 * Math.sin(t * 0.8);
        dieMat.emissiveIntensity = 0.26 + 0.08 * Math.sin(t * 1.3);
        rig.rotation.x = 0.32 + 0.04 * Math.sin(t * 0.31);
        // expanding radio rings
        for (let i = 0; i < waves.length; i++) {
          const p = (t / 4.2 + i / waves.length) % 1;
          const s = 0.25 + p * 2.2;
          waves[i].scale.set(s, s, 1);
          waves[i].material.opacity = 0.55 * Math.pow(1 - p, 1.6) * Math.min(1, p * 8);
        }
        for (let i = 0; i < arcs.length; i++) {
          const p = (t * 0.9 - i * 0.3) % 1.6;
          arcs[i].material.opacity = Math.max(0, Math.sin(Math.max(0, p) / 1.6 * Math.PI)) * 0.9;
          arcs[i].rotation.y = 0;
        }
        coreGlow.material.opacity = 0.7 + 0.25 * Math.sin(t * 3);
        antRing.rotation.x = t * 0.9;
        antRing.rotation.y = t * 0.6;
        // IoT nodes orbit slowly and exchange packets with the chip
        for (let i = 0; i < NN; i++) {
          const a = t * 0.18 + i * Math.PI * 2 / NN;
          const rad = 2.75 + 0.15 * Math.sin(t * 0.5 + i);
          const x = Math.cos(a) * rad, z = Math.sin(a) * rad, y = 0.55 + 0.25 * Math.sin(t * 0.7 + i * 1.7);
          nodes[i].position.set(x, y, z);
          nodes[i].rotation.y = t * 0.8 + i;
          nodes[i].rotation.x = 0.5;
          const pa = nodeLinks[i].geometry.attributes.position.array;
          pa[0] = 0; pa[1] = NODE_Y; pa[2] = 0; pa[3] = x; pa[4] = y; pa[5] = z;
          nodeLinks[i].geometry.attributes.position.needsUpdate = true;
          const p = (t * 0.35 + i * 0.27) % 1;
          const q = i % 2 ? p : 1 - p;
          linkPulses[i].position.set(x * q, NODE_Y + (y - NODE_Y) * q, z * q);
          linkPulses[i].material.opacity = Math.sin(p * Math.PI);
        }
      }
    };
  };
})();
