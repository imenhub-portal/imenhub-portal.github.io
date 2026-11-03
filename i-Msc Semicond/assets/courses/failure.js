/* Failure Analysis & Reliability — decapped IC with crack, delaminated corner, hotspot, magnifier, FIB probe, cross-section trench */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  IMEN.courseScenes = IMEN.courseScenes || {};
  IMEN.courseScenes.failure = function (kit) {
    const T = THREE;
    const g = new T.Group();
    const root = new T.Group();
    root.rotation.x = 0.46;
    g.add(root);

    const RED = 0xff2a2a;
    const PCB_TOP = 0.1;
    const DIE_BASE = PCB_TOP + 0.0;
    const DIE_TOP = 0.38;
    const S = new T.Vector3(0.45, DIE_TOP, 0.35); // failure site

    /* ---------- helpers ---------- */
    function sq(w, d, r) {
      const s = new T.Shape();
      const x = -w / 2, y = -d / 2;
      s.moveTo(x + r, y); s.lineTo(x + w - r, y); s.quadraticCurveTo(x + w, y, x + w, y + r);
      s.lineTo(x + w, y + d - r); s.quadraticCurveTo(x + w, y + d, x + w - r, y + d);
      s.lineTo(x + r, y + d); s.quadraticCurveTo(x, y + d, x, y + d - r);
      s.lineTo(x, y + r); s.quadraticCurveTo(x, y, x + r, y);
      return s;
    }
    function mergeGeos(list) {
      let nv = 0, ni = 0;
      list.forEach(function (gm) { nv += gm.attributes.position.count; ni += gm.index.count; });
      const pos = new Float32Array(nv * 3), nor = new Float32Array(nv * 3), idx = new Uint32Array(ni);
      let vo = 0, io = 0;
      list.forEach(function (gm) {
        pos.set(gm.attributes.position.array, vo * 3);
        nor.set(gm.attributes.normal.array, vo * 3);
        const ix = gm.index.array;
        for (let i = 0; i < ix.length; i++) idx[io + i] = ix[i] + vo;
        vo += gm.attributes.position.count; io += ix.length;
        gm.dispose();
      });
      const o = new T.BufferGeometry();
      o.setAttribute("position", new T.BufferAttribute(pos, 3));
      o.setAttribute("normal", new T.BufferAttribute(nor, 3));
      o.setIndex(new T.BufferAttribute(idx, 1));
      return o;
    }
    function ln(pts, r, mat, y, seg) {
      const v = pts.map(function (p) { return new T.Vector3(p[0], y, p[1]); });
      const m = new T.Mesh(new T.TubeGeometry(new T.CatmullRomCurve3(v, false, "catmullrom", 0.2), seg || 40, r, 6, false), mat);
      return m;
    }

    /* ---------- PCB ---------- */
    const pcbBody = new T.Mesh(kit.roundedBox(5.6, 0.2, 4.2, 0.14, 0.03), kit.M.pcb(0x0a3a30));
    root.add(pcbBody);
    const pt = kit.tex.pcb(1024, 17, "#0b4a3a");
    const pcbTop = new T.Mesh(new T.PlaneGeometry(5.4, 4.0), new T.MeshPhysicalMaterial({ map: pt.map, bumpMap: pt.bump, bumpScale: 1.5, roughness: 0.45, metalness: 0.15, clearcoat: 0.6, clearcoatRoughness: 0.25, envMapIntensity: 0.6 }));
    pcbTop.rotation.x = -Math.PI / 2; pcbTop.position.y = PCB_TOP + 0.002; root.add(pcbTop);

    /* ---------- package (decapped) ---------- */
    const WALL = 0.42;
    const pkShape = new T.Shape();
    // (x,z) -> shape (x,-z); front-right corner cut
    [[-1.8, -1.8], [1.8, -1.8], [1.8, 1.0], [1.0, 1.8], [-1.8, 1.8]].forEach(function (p, i) {
      if (i === 0) pkShape.moveTo(p[0], -p[1]); else pkShape.lineTo(p[0], -p[1]);
    });
    pkShape.lineTo(-1.8, 1.8);
    const hole = sq(2.5, 2.5, 0.05); hole.autoClose = true;
    pkShape.holes.push(hole);
    const pkGeo = new T.ExtrudeGeometry(pkShape, { depth: WALL - 0.04, bevelEnabled: true, bevelThickness: 0.02, bevelSize: 0.02, bevelSegments: 2 });
    pkGeo.rotateX(-Math.PI / 2);
    const epoxy = kit.M.epoxy(0x101218); epoxy.envMapIntensity = 0.7; epoxy.roughness = 0.6;
    const pk = new T.Mesh(pkGeo, epoxy);
    pk.position.y = PCB_TOP + 0.02;
    root.add(pk);
    // corner piece (delaminated)
    const cShape = new T.Shape();
    cShape.moveTo(1.8, -1.0); cShape.lineTo(1.8, -1.8); cShape.lineTo(1.0, -1.8); cShape.lineTo(1.8, -1.0);
    const cGeo = new T.ExtrudeGeometry(cShape, { depth: WALL - 0.04, bevelEnabled: true, bevelThickness: 0.02, bevelSize: 0.02, bevelSegments: 2 });
    cGeo.rotateX(-Math.PI / 2);
    cGeo.translate(-1.533, 0, -1.533);
    const corner = new T.Mesh(cGeo, epoxy);
    corner.position.set(1.533 + 0.05, PCB_TOP + 0.02 + 0.09, 1.533 + 0.05);
    corner.quaternion.setFromAxisAngle(new T.Vector3(1, 0, -1).normalize(), -0.17);
    root.add(corner);
    // package marking
    const mark = kit.label("IC-7421  FA", 0.2, { color: "#8fa2c8", size: 96 });
    mark.rotation.x = -Math.PI / 2; mark.position.set(-0.2, PCB_TOP + 0.02 + WALL + 0.003, -1.52);
    root.add(mark);
    const pin1 = new T.Mesh(new T.CircleGeometry(0.06, 20), new T.MeshBasicMaterial({ color: 0x6f7c99 }));
    pin1.rotation.x = -Math.PI / 2; pin1.position.set(-1.55, PCB_TOP + 0.02 + WALL + 0.003, -1.55); root.add(pin1);
    // crack through the epoxy wall
    const dark = new T.MeshBasicMaterial({ color: 0x020203 });
    const crackTop = PCB_TOP + 0.02 + WALL + 0.004;
    root.add(ln([[1.12, 1.32], [1.25, 1.18], [1.2, 1.0], [1.38, 0.85], [1.34, 0.6], [1.5, 0.45]], 0.012, dark, crackTop, 30));
    root.add(ln([[1.25, 1.18], [1.45, 1.22], [1.55, 1.1]], 0.008, dark, crackTop, 12));

    // leads (gull-wing)
    const leadMat = kit.M.solder();
    const leadGeo = new T.BoxGeometry(0.11, 0.035, 0.5);
    const leads = [];
    for (let i = 0; i < 16; i++) {
      const s = -1.5 + i * 0.2;
      [[s, -1.8 - 0.14, 0], [s, 1.8 + 0.14, 0], [-1.8 - 0.14, s, 1], [1.8 + 0.14, s, 1]].forEach(function (a) {
        if (a[2] === 0 && a[1] > 0 && s > 0.95) return; // cut corner
        if (a[2] === 1 && a[0] > 0 && s > 0.95) return;
        leads.push(a);
      });
    }
    const leadInst = kit.instanced(leadGeo, leadMat, leads.length, function (i, d) {
      const a = leads[i];
      if (a[2] === 0) d.position.set(a[0], PCB_TOP + 0.06, a[1]);
      else { d.position.set(a[0], PCB_TOP + 0.06, a[1]); d.rotation.y = Math.PI / 2; }
    });
    root.add(leadInst);

    /* ---------- die ---------- */
    const dieMat = kit.M.silicon(0x203060);
    const dieSlab = new T.Mesh(kit.roundedBox(2.1, 0.12, 2.1, 0.03, 0.01), dieMat);
    dieSlab.position.y = PCB_TOP + 0.06 + 0.0; root.add(dieSlab);
    const layoutTex = kit.tex.dieLayout(1024, 29, ["#07102a", "#2f6bff", "#18c0ff", "#ffb340", "#c8d6ff"]);
    layoutTex.wrapS = layoutTex.wrapT = T.ClampToEdgeWrapping;
    layoutTex.repeat.set(1 / 2.1, 1 / 2.1); layoutTex.offset.set(0.5, 0.5);
    const topShape = sq(2.1, 2.1, 0.02);
    const tr = new T.Path();
    const TX0 = -0.95, TX1 = -0.1, TZ0 = 0.38, TZ1 = 0.82; // trench rect (x,z)
    tr.moveTo(TX0, -TZ1); tr.lineTo(TX0, -TZ0); tr.lineTo(TX1, -TZ0); tr.lineTo(TX1, -TZ1); tr.lineTo(TX0, -TZ1);
    topShape.holes.push(tr);
    const topGeo = new T.ExtrudeGeometry(topShape, { depth: DIE_TOP - (PCB_TOP + 0.12), bevelEnabled: false });
    topGeo.rotateX(-Math.PI / 2);
    const capMat = new T.MeshPhysicalMaterial({ map: layoutTex, metalness: 0.55, roughness: 0.26, clearcoat: 0.8, clearcoatRoughness: 0.12, envMapIntensity: 1.0 });
    const dieTop = new T.Mesh(topGeo, [capMat, dieMat]);
    dieTop.position.y = PCB_TOP + 0.12; root.add(dieTop);
    // trench floor darker patch + cross-section face
    const trF = new T.Mesh(new T.PlaneGeometry(TX1 - TX0, TZ1 - TZ0), new T.MeshStandardMaterial({ color: 0x0a0f20, metalness: 0.8, roughness: 0.35 }));
    trF.rotation.x = -Math.PI / 2; trF.position.set((TX0 + TX1) / 2, PCB_TOP + 0.123, (TZ0 + TZ1) / 2); root.add(trF);
    const xsTex = kit.canvasTex(1024, 256, function (c, w, h) {
      c.fillStyle = "#5b6890"; c.fillRect(0, 0, w, h);                       // substrate
      const layers = [[0.0, 0.34, "#4c5a86"], [0.34, 0.5, "#7c97c9"], [0.5, 0.62, "#9ab4e0"], [0.62, 0.78, "#6e88bd"], [0.78, 0.9, "#a4bde6"], [0.9, 1, "#d6e2ff"]];
      layers.forEach(function (l) { c.fillStyle = l[2]; c.fillRect(0, h * (1 - l[1]), w, h * (l[1] - l[0])); });
      c.strokeStyle = "rgba(10,20,50,0.55)"; c.lineWidth = 2;
      layers.forEach(function (l) { c.beginPath(); c.moveTo(0, h * (1 - l[1])); c.lineTo(w, h * (1 - l[1])); c.stroke(); });
      // metal lines (cross-section dots) on 3 levels + vias
      const lv = [0.56, 0.72, 0.86];
      lv.forEach(function (y, k) {
        for (let x = 30 + k * 18; x < w - 20; x += 82) {
          c.fillStyle = "#e09a62"; c.fillRect(x, h * (1 - y) - 9, 46, 16);
          c.fillStyle = "rgba(255,230,190,0.7)"; c.fillRect(x, h * (1 - y) - 9, 46, 3);
        }
      });
      for (let x = 52; x < w - 40; x += 164) {
        c.fillStyle = "#c98655"; c.fillRect(x + 10, h * (1 - 0.86), 10, h * 0.3); // via
        c.fillRect(x + 28, h * (1 - 0.72) + 4, 10, h * 0.16);
      }
      // broken via with void
      c.fillStyle = "#06080f"; c.beginPath(); c.ellipse(560, h * 0.4, 20, 22, 0, 0, 6.283); c.fill();
      c.strokeStyle = "#ff3a2a"; c.lineWidth = 4; c.beginPath(); c.ellipse(560, h * 0.4, 28, 30, 0, 0, 6.283); c.stroke();
      c.strokeStyle = "#0a0a0f"; c.lineWidth = 3; c.beginPath(); c.moveTo(540, h * 0.26); c.lineTo(572, h * 0.36); c.lineTo(552, h * 0.46); c.lineTo(584, h * 0.58); c.stroke();
    });
    const xsFace = new T.Mesh(new T.PlaneGeometry(TX1 - TX0, 0.26), new T.MeshStandardMaterial({ map: xsTex, metalness: 0.4, roughness: 0.4, emissive: 0x111826, emissiveMap: xsTex, emissiveIntensity: 0.6 }));
    xsFace.position.set((TX0 + TX1) / 2, PCB_TOP + 0.123 + 0.13, TZ0 + 0.003); root.add(xsFace);
    // trench frame highlight
    const tf = new T.Mesh(new T.PlaneGeometry(TX1 - TX0 + 0.03, 0.025), new T.MeshBasicMaterial({ color: 0x66b8ff, transparent: true, opacity: 0.6 }));
    tf.rotation.x = -Math.PI / 2; tf.position.set((TX0 + TX1) / 2, DIE_TOP + 0.004, TZ1 + 0.02); root.add(tf);

    /* ---------- pads, leadframe fingers, wires ---------- */
    const padPos = [];
    for (let i = 0; i < 15; i++) {
      const s = -0.86 + i * 0.123;
      padPos.push([s, 0.93, 0, i], [s, -0.93, 0, i], [0.93, s, 1, i], [-0.93, s, 1, i]);
    }
    const goldM = kit.M.gold(); goldM.emissive.set(0x3d2800);
    const pads = kit.instanced(new T.BoxGeometry(0.075, 0.012, 0.075), goldM, padPos.length, function (i, d) { const p = padPos[i]; d.position.set(p[0], DIE_TOP + 0.007, p[1]); });
    root.add(pads);
    const fingerGeo = new T.BoxGeometry(0.075, 0.18, 0.2);
    const fingerMat = kit.M.copper();
    const fingers = kit.instanced(fingerGeo, fingerMat, padPos.length, function (i, d) {
      const p = padPos[i];
      const fx = p[2] ? Math.sign(p[0]) * 1.17 : p[0];
      const fz = p[2] ? p[1] : Math.sign(p[1]) * 1.17;
      d.position.set(fx, PCB_TOP + 0.09, fz);
      if (p[2]) d.rotation.y = Math.PI / 2;
    });
    root.add(fingers);
    const wgeos = [];
    let brokenIdx = -1;
    padPos.forEach(function (p, i) {
      const fx = p[2] ? Math.sign(p[0]) * 1.15 : p[0];
      const fz = p[2] ? p[1] : Math.sign(p[1]) * 1.15;
      const a = new T.Vector3(p[0], DIE_TOP + 0.012, p[1]);
      const b = new T.Vector3(fx, PCB_TOP + 0.185, fz);
      if (p[2] === 0 && p[1] > 0 && p[3] === 11) { brokenIdx = i; return; }
      const c = a.clone().lerp(b, 0.5); c.y += 0.16;
      wgeos.push(new T.TubeGeometry(new T.QuadraticBezierCurve3(a, c, b), 10, 0.0065, 5, false));
    });
    const wires = new T.Mesh(mergeGeos(wgeos), kit.M.gold());
    root.add(wires);
    // broken wire: lifted stub + loose end
    const bp = padPos[brokenIdx];
    {
      const a = new T.Vector3(bp[0], DIE_TOP + 0.012, bp[1]);
      const e = new T.Vector3(bp[0] + 0.02, DIE_TOP + 0.26, bp[1] + 0.12);
      const m1 = new T.Vector3(bp[0] - 0.02, DIE_TOP + 0.22, bp[1] + 0.03);
      root.add(new T.Mesh(new T.TubeGeometry(new T.QuadraticBezierCurve3(a, m1, e), 12, 0.0065, 5, false), kit.M.gold()));
      const b = new T.Vector3(bp[0], PCB_TOP + 0.185, 1.15);
      const e2 = new T.Vector3(bp[0] + 0.02, DIE_TOP + 0.2, 1.1);
      root.add(new T.Mesh(new T.TubeGeometry(new T.QuadraticBezierCurve3(b, new T.Vector3(bp[0], DIE_TOP + 0.2, 1.14), e2), 8, 0.0065, 5, false), kit.M.gold()));
      const spk = kit.glow(0xffa040, 0.25, 0.8); spk.position.copy(e); root.add(spk);
    }

    /* ---------- die crack + defect ---------- */
    const crackY = DIE_TOP + 0.006;
    const crackPts = [[0.98, 0.58], [0.82, 0.5], [0.72, 0.36], [0.55, 0.4], [0.45, 0.28], [0.3, 0.3], [0.2, 0.12], [0.05, 0.1]];
    const crackGlowMat = kit.M.additive(RED, 0.55);
    root.add(ln(crackPts, 0.011, dark, crackY, 50));
    const cg = ln(crackPts, 0.026, crackGlowMat, crackY + 0.002, 50); cg.userData.noFit = true; root.add(cg);
    [[[0.72, 0.36], [0.78, 0.2], [0.7, 0.05]], [[0.45, 0.28], [0.52, 0.12], [0.4, 0.0]], [[0.3, 0.3], [0.22, 0.45]]].forEach(function (b) {
      root.add(ln(b, 0.007, dark, crackY, 14));
    });
    // spalled chips
    const rr = kit.rng(5);
    const chipInst = kit.instanced(new T.BoxGeometry(0.05, 0.015, 0.035), kit.M.silicon(0x6277aa), 14, function (i, d) {
      d.position.set(0.3 + rr() * 0.45, DIE_TOP + 0.01, 0.1 + rr() * 0.5);
      d.rotation.set(rr() * 0.4, rr() * 6, rr() * 0.4);
      d.scale.setScalar(0.5 + rr());
    });
    root.add(chipInst);
    // debris near delaminated corner
    const debris = kit.instanced(new T.BoxGeometry(0.07, 0.03, 0.05), epoxy, 10, function (i, d) {
      d.position.set(1.9 + rr() * 0.5, PCB_TOP + 0.02, 1.9 + rr() * 0.5);
      d.rotation.set(rr(), rr() * 6, rr());
    });
    root.add(debris);
    const gapGlow = kit.glow(0xff8a3a, 0.8, 0.5); gapGlow.position.set(1.62, PCB_TOP + 0.2, 1.62); root.add(gapGlow);

    /* ---------- thermal hotspot ---------- */
    function heatField(u, v) {
      const a = Math.exp(-((u - 0.5) * (u - 0.5) / 0.018 + (v - 0.5) * (v - 0.5) / 0.03));
      const b = 0.55 * Math.exp(-((u - 0.68) * (u - 0.68) / 0.012 + (v - 0.4) * (v - 0.4) / 0.012));
      const c = 0.35 * Math.exp(-((u - 0.33) * (u - 0.33) / 0.02 + (v - 0.62) * (v - 0.62) / 0.012));
      return a + b + c;
    }
    const ramp = [[0, 0, 0, 0, 0], [0.08, 30, 40, 200, 0.25], [0.25, 0, 190, 255, 0.55], [0.42, 60, 255, 90, 0.7], [0.58, 255, 240, 40, 0.8], [0.74, 255, 130, 20, 0.9], [0.92, 255, 40, 20, 1], [1.2, 255, 230, 200, 1]];
    function rampAt(f) {
      for (let i = 1; i < ramp.length; i++) if (f <= ramp[i][0]) {
        const a = ramp[i - 1], b = ramp[i], k = (f - a[0]) / (b[0] - a[0]);
        return [a[1] + (b[1] - a[1]) * k, a[2] + (b[2] - a[2]) * k, a[3] + (b[3] - a[3]) * k, a[4] + (b[4] - a[4]) * k];
      }
      return ramp[ramp.length - 1].slice(1);
    }
    const heatTex = kit.canvasTex(256, 256, function (c, w, h) {
      const img = c.createImageData(w, h);
      for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
        let u = x / w, v = y / h;
        const f = heatField(u, v);
        const col = rampAt(f);
        let al = col[3];
        const edge = Math.min(u, 1 - u, v, 1 - v) * 5; al *= Math.min(1, edge);
        const cont = Math.abs(((f * 9) % 1) - 0.5) < 0.06 && f > 0.08 ? 1.35 : 1;
        const o = (y * w + x) * 4;
        img.data[o] = Math.min(255, col[0] * cont); img.data[o + 1] = Math.min(255, col[1] * cont); img.data[o + 2] = Math.min(255, col[2] * cont); img.data[o + 3] = al * 255;
      }
      c.putImageData(img, 0, 0);
    });
    const heatMat = new T.MeshBasicMaterial({ map: heatTex, transparent: true, opacity: 0.85, depthWrite: false, blending: T.AdditiveBlending, toneMapped: false, side: T.DoubleSide });
    const heat = new T.Mesh(new T.PlaneGeometry(1.8, 1.8), heatMat);
    heat.rotation.x = -Math.PI / 2; heat.position.set(S.x, S.y + 0.3, S.z); heat.renderOrder = 4; heat.userData.noFit = true; root.add(heat);
    const heatTex2 = kit.canvasTex(256, 128, function (c, w, h) {
      const img = c.createImageData(w, h);
      for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
        const u = x / w, v = 1 - y / h; // v=0 bottom
        const f = 1.15 * Math.exp(-((u - 0.5) * (u - 0.5) / 0.03)) * Math.exp(-(v * v) / 0.28) + 0.2 * Math.exp(-((u - 0.78) * (u - 0.78)) / 0.02) * Math.exp(-(v * v) / 0.15);
        const col = rampAt(f);
        const al = col[3] * Math.min(1, Math.min(u, 1 - u) * 6);
        const o = (y * w + x) * 4;
        img.data[o] = col[0]; img.data[o + 1] = col[1]; img.data[o + 2] = col[2]; img.data[o + 3] = al * 255;
      }
      c.putImageData(img, 0, 0);
    });
    const heatV = new T.Mesh(new T.PlaneGeometry(1.7, 0.85), new T.MeshBasicMaterial({ map: heatTex2, transparent: true, opacity: 0.75, depthWrite: false, blending: T.AdditiveBlending, toneMapped: false, side: T.DoubleSide }));
    heatV.position.set(S.x, S.y + 0.43, S.z); heatV.renderOrder = 4; heatV.userData.noFit = true; root.add(heatV);
    const hotGlow = kit.glow(0xff5a20, 1.6, 0.5); hotGlow.position.set(S.x, S.y + 0.03, S.z); root.add(hotGlow);

    /* ---------- fault flags ---------- */
    const flags = [];
    function flag(x, y, z, h) {
      const rod = new T.Mesh(new T.CylinderGeometry(0.008, 0.008, h, 6), kit.M.metal(0xc8d0e0, 0.25));
      rod.position.set(x, y + h / 2, z); root.add(rod);
      const ball = new T.Mesh(new T.SphereGeometry(0.05, 14, 10), new T.MeshBasicMaterial({ color: RED, toneMapped: false }));
      ball.position.set(x, y + h, z); root.add(ball);
      const gl = kit.glow(RED, 0.4, 0.8); gl.position.copy(ball.position); root.add(gl);
      const ring = new T.Mesh(new T.TorusGeometry(0.1, 0.007, 6, 28), new T.MeshBasicMaterial({ color: RED, transparent: true, opacity: 0.8, toneMapped: false }));
      ring.rotation.x = -Math.PI / 2; ring.position.set(x, y + 0.012, z); root.add(ring);
      flags.push({ gl: gl, ring: ring });
    }
    flag(0.05, DIE_TOP, 0.1, 0.42);
    flag(bp[0] + 0.02, DIE_TOP, 1.35, 0.32);
    flag(1.38, PCB_TOP + WALL + 0.02, 0.5, 0.32);

    /* ---------- magnifier ---------- */
    const lens = new T.Group();
    const LR = 0.92;
    const rimMat = kit.M.metal(0xb9c1ce, 0.28);
    const rim = new T.Mesh(new T.TorusGeometry(LR + 0.05, 0.075, 18, 72), rimMat); lens.add(rim);
    const rim2 = new T.Mesh(new T.TorusGeometry(LR - 0.02, 0.022, 10, 72), kit.M.gold()); rim2.position.z = 0.02; lens.add(rim2);
    const viewTex = kit.canvasTex(512, 512, function (c, w, h) {
      const r = kit.rng(3);
      const bg = c.createRadialGradient(w / 2, h / 2, 20, w / 2, h / 2, w * 0.7);
      bg.addColorStop(0, "#1a2c58"); bg.addColorStop(1, "#050a1c");
      c.fillStyle = bg; c.fillRect(0, 0, w, h);
      c.lineCap = "round";
      // layout metal at high magnification
      for (let k = 0; k < 2; k++) {
        c.strokeStyle = k ? "rgba(255,170,80,0.65)" : "rgba(120,170,255,0.7)"; c.lineWidth = k ? 9 : 14;
        for (let i = 0; i < 14; i++) {
          const x = r() * w, y = r() * h;
          c.beginPath(); c.moveTo(x, y);
          if (k) c.lineTo(x + (r() - 0.5) * 300, y); else c.lineTo(x, y + (r() - 0.5) * 300);
          c.stroke();
        }
      }
      c.fillStyle = "rgba(20,30,60,0.6)";
      for (let i = 0; i < 30; i++) c.fillRect(r() * w, r() * h, 20 + r() * 60, 12 + r() * 40);
      // crack, large
      const path = [[40, 460], [120, 380], [170, 300], [260, 290], [300, 220], [380, 190], [430, 110], [480, 60]];
      function stroke(wd, col, blur) {
        c.save(); c.shadowColor = "#ff3020"; c.shadowBlur = blur; c.strokeStyle = col; c.lineWidth = wd;
        c.beginPath(); c.moveTo(path[0][0], path[0][1]); for (let i = 1; i < path.length; i++) c.lineTo(path[i][0], path[i][1]); c.stroke(); c.restore();
      }
      stroke(16, "rgba(255,60,40,0.55)", 28); stroke(7, "#05060a", 0);
      c.strokeStyle = "#05060a"; c.lineWidth = 4;
      [[[170, 300], [140, 230], [180, 170]], [[300, 220], [330, 290], [380, 330]], [[260, 290], [230, 360], [250, 420]]].forEach(function (b) {
        c.beginPath(); c.moveTo(b[0][0], b[0][1]); c.lineTo(b[1][0], b[1][1]); c.lineTo(b[2][0], b[2][1]); c.stroke();
      });
      // reticle
      c.strokeStyle = "rgba(255,255,255,0.55)"; c.lineWidth = 2;
      c.beginPath(); c.arc(w / 2, h / 2, 190, 0, 6.283); c.stroke();
      for (let i = 0; i < 4; i++) { const a = i * Math.PI / 2; c.beginPath(); c.moveTo(w / 2 + Math.cos(a) * 170, h / 2 + Math.sin(a) * 170); c.lineTo(w / 2 + Math.cos(a) * 215, h / 2 + Math.sin(a) * 215); c.stroke(); }
      c.fillStyle = "rgba(255,255,255,0.75)"; c.font = "700 20px monospace"; c.fillText("x2400   5.0 kV   200 nm", 150, 490);
      // scanlines
      c.fillStyle = "rgba(255,255,255,0.03)";
      for (let y = 0; y < h; y += 4) c.fillRect(0, y, w, 1);
    });
    const view = new T.Mesh(new T.CircleGeometry(LR - 0.04, 64), new T.MeshBasicMaterial({ map: viewTex, transparent: true, opacity: 0.96, toneMapped: false }));
    lens.add(view);
    const glassGeo = new T.SphereGeometry(1, 56, 28); glassGeo.scale(LR - 0.02, LR - 0.02, 0.12);
    const glass = new T.Mesh(glassGeo, kit.M.glass(0xdcecff, 0.16)); glass.position.z = 0.0; glass.renderOrder = 7; lens.add(glass);
    // handle
    const ang = -Math.PI + Math.atan2(0.85, 0.55);
    const dirx = Math.cos(ang), diry = Math.sin(ang);
    const hg = new T.Group();
    const neck = new T.Mesh(new T.CylinderGeometry(0.06, 0.075, 0.55, 18), kit.M.metal(0xc2c9d4, 0.28)); neck.position.y = 0.3; hg.add(neck);
    const collar = new T.Mesh(new T.CylinderGeometry(0.11, 0.1, 0.1, 24), kit.M.gold()); collar.position.y = 0.62; hg.add(collar);
    const grip = new T.Mesh(new T.CylinderGeometry(0.11, 0.095, 1.15, 28), kit.M.plastic(0x111318, 0.65)); grip.position.y = 1.25; hg.add(grip);
    for (let i = 0; i < 6; i++) { const rg = new T.Mesh(new T.TorusGeometry(0.103 - i * 0.002, 0.012, 6, 24), kit.M.plastic(0x1e2430, 0.5)); rg.rotation.x = Math.PI / 2; rg.position.y = 0.8 + i * 0.17; hg.add(rg); }
    const cap = new T.Mesh(new T.SphereGeometry(0.095, 18, 12), kit.M.gold()); cap.position.y = 1.83; hg.add(cap);
    hg.position.set(dirx * (LR + 0.1), diry * (LR + 0.1), 0);
    hg.rotation.z = ang - Math.PI / 2;
    lens.add(hg);
    lens.position.set(-0.3, 1.7, 0.1);
    lens.rotation.set(-1.0, -0.1, 0);
    root.add(lens);
    const lensGlow = kit.glow(0x5aa8ff, 3.2, 0.14); lensGlow.position.set(-0.3, 1.7, 0.0); root.add(lensGlow);

    /* ---------- FIB / laser probe ---------- */
    const T0 = new T.Vector3(2.1, 1.75, -0.9);
    const probe = new T.Group();
    const col = new T.Group();
    const nozz = new T.Mesh(new T.CylinderGeometry(0.16, 0.03, 0.32, 24), kit.M.metal(0xd5dae4, 0.25)); nozz.position.y = 0.16; col.add(nozz);
    const gr = new T.Mesh(new T.CylinderGeometry(0.19, 0.19, 0.06, 24), kit.M.gold()); gr.position.y = 0.35; col.add(gr);
    const hs = new T.Mesh(new T.CylinderGeometry(0.17, 0.17, 0.55, 28), kit.M.metal(0xaeb6c4, 0.3)); hs.position.y = 0.66; col.add(hs);
    const coil = new T.Mesh(new T.CylinderGeometry(0.22, 0.22, 0.3, 28), kit.M.copper()); coil.position.y = 1.0; col.add(coil);
    for (let i = 0; i < 5; i++) { const rg = new T.Mesh(new T.TorusGeometry(0.22, 0.014, 6, 28), kit.M.metal(0xe0c080, 0.3)); rg.rotation.x = Math.PI / 2; rg.position.y = 0.9 + i * 0.06; col.add(rg); }
    const ccap = new T.Mesh(new T.CylinderGeometry(0.12, 0.2, 0.18, 24), kit.M.metal(0x7b8499, 0.3)); ccap.position.y = 1.24; col.add(ccap);
    const hdir = new T.Vector3().subVectors(T0, new T.Vector3(S.x, S.y, S.z)).normalize();
    col.quaternion.setFromUnitVectors(new T.Vector3(0, 1, 0), hdir);
    col.position.copy(T0);
    probe.add(col);
    root.add(probe);
    // beam (along local +Z, scaled to length)
    const beamGeo = new T.CylinderGeometry(0.012, 0.012, 1, 8, 1, true); beamGeo.rotateX(Math.PI / 2); beamGeo.translate(0, 0, 0.5);
    const beamMat = kit.M.additive(0x8fd0ff, 0.9);
    const beam = new T.Mesh(beamGeo, beamMat); beam.userData.noFit = true; root.add(beam);
    const beamHalo = new T.Mesh(beamGeo, kit.M.additive(0x4a8cff, 0.2)); beamHalo.userData.noFit = true; root.add(beamHalo);
    const hit = kit.glow(0xbfe8ff, 0.5, 0.95); root.add(hit);
    const hit2 = kit.glow(0xff7a30, 0.9, 0.5); root.add(hit2);
    // scan line
    const scan = new T.Mesh(new T.PlaneGeometry(0.035, 1.0), kit.M.additive(0x9fe0ff, 0.5));
    scan.rotation.x = -Math.PI / 2; scan.userData.noFit = true; root.add(scan);
    // sparks
    const sparkMat = new T.SpriteMaterial({ map: kit.tex.radial([[0, "rgba(255,255,255,1)"], [0.3, "rgba(255,200,120,0.7)"], [1, "rgba(255,120,40,0)"]], 64), color: 0xffffff, transparent: true, depthWrite: false, blending: T.AdditiveBlending, toneMapped: false });
    const sparks = [];
    for (let i = 0; i < 16; i++) {
      const s = new T.Sprite(sparkMat); s.userData.noFit = true; root.add(s);
      const a = i * 2.399;
      sparks.push({ s: s, vx: Math.cos(a) * (0.25 + (i % 4) * 0.12), vz: Math.sin(a) * (0.25 + (i % 3) * 0.1), vy: 0.5 + (i % 5) * 0.18, ph: i / 16 });
    }
    const m4 = new T.Matrix4(), up = new T.Vector3(0, 1, 0), H = new T.Vector3();

    const sh = kit.contactShadow(5.2, 0.7); sh.position.y = -1.7; g.add(sh);
    const pl = new T.PointLight(0xff6a30, 1.4, 5); pl.position.set(S.x, S.y + 0.5, S.z); root.add(pl);
    const pl2 = new T.PointLight(0x6aa8ff, 1.0, 12); pl2.position.set(-1.5, 2.5, 2.5); root.add(pl2);

    function update(t) {
      const sx = Math.sin(t * 0.9);
      H.set(S.x + 0.42 * sx, S.y + 0.004, S.z - 0.12 + 0.12 * Math.sin(t * 0.45));
      const len = T0.distanceTo(H);
      m4.lookAt(H, T0, up);
      beam.quaternion.setFromRotationMatrix(m4); beam.position.copy(T0); beam.scale.set(1, 1, len);
      beamHalo.quaternion.copy(beam.quaternion); beamHalo.position.copy(T0); beamHalo.scale.set(3, 3, len);
      beamMat.opacity = 0.7 + 0.25 * Math.sin(t * 31) * Math.sin(t * 7);
      hit.position.copy(H); hit.position.y += 0.02; hit.scale.setScalar(0.4 + 0.15 * Math.sin(t * 23));
      hit2.position.copy(H); hit2.position.y += 0.02;
      scan.position.set(H.x, DIE_TOP + 0.01, S.z - 0.1);
      for (let i = 0; i < sparks.length; i++) {
        const p = sparks[i];
        const k = (t * 1.1 + p.ph) % 1;
        p.s.position.set(H.x + p.vx * k, H.y + p.vy * k - 1.4 * k * k, H.z + p.vz * k);
        p.s.scale.setScalar(0.09 * (1 - k) + 0.01);
      }
      const pu = 0.5 + 0.5 * Math.sin(t * 2.3);
      heatMat.opacity = 0.7 + 0.2 * pu;
      heat.position.y = S.y + 0.3 + 0.025 * Math.sin(t * 1.2);
      hotGlow.material.opacity = 0.35 + 0.25 * pu;
      crackGlowMat.opacity = 0.4 + 0.25 * pu;
      for (let i = 0; i < flags.length; i++) {
        const f = flags[i], q = 0.5 + 0.5 * Math.sin(t * 3 + i * 1.7);
        f.gl.material.opacity = 0.5 + 0.4 * q; f.ring.scale.setScalar(1 + 0.25 * q); f.ring.material.opacity = 0.9 - 0.5 * q;
      }
      gapGlow.material.opacity = 0.3 + 0.2 * Math.sin(t * 2 + 1);
      lens.position.y = 1.7 + 0.05 * Math.sin(t * 0.9);
      lensGlow.position.y = lens.position.y;
      root.rotation.y = Math.sin(t * 0.2) * 0.4 - 0.1;
    }
    update(0);
    return { group: g, update: update, fit: 2.1, spin: 0 };
  };
})();
