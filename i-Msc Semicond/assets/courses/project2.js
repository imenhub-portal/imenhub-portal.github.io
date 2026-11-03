/* Project II - BUILD & TEST: wafer probe-station vignette with scope hologram, PASS badge and medal */
(function () {
  "use strict";
  IMEN.courseScenes.project2 = function (kit) {
    const T = THREE;
    const g = new T.Group();
    const GOLD = 0xf3c35a, CYAN = 0x5fe3ff, GREEN = 0x3dff9a;
    const Y = new T.Vector3(0, 1, 0);
    function seg(a, b, ra, rb, mat, parent) {
      const len = a.distanceTo(b);
      const m = new T.Mesh(new T.CylinderGeometry(rb, ra, len, 12), mat);
      m.position.copy(a).add(b).multiplyScalar(0.5);
      m.quaternion.setFromUnitVectors(Y, b.clone().sub(a).normalize());
      (parent || g).add(m); return m;
    }

    /* ---------- base slab ---------- */
    const slab = new T.Mesh(kit.roundedBox(7.0, 0.2, 3.7, 0.16, 0.04), kit.M.metal(0x26335f, 0.34));
    slab.position.y = -0.1; g.add(slab);
    const slabTrim = new T.Mesh(kit.roundedBox(7.1, 0.035, 3.8, 0.18, 0.012), kit.M.gold());
    slabTrim.position.y = -0.17; g.add(slabTrim);
    const slabTop = new T.Mesh(new T.PlaneGeometry(6.8, 3.5), new T.MeshPhysicalMaterial({
      map: kit.canvasTex(1024, 540, function (c, W, H) {
        c.fillStyle = "#0e1c46"; c.fillRect(0, 0, W, H);
        for (let x = 0; x <= W; x += 24) { c.strokeStyle = x % 120 ? "rgba(110,160,255,0.14)" : "rgba(150,200,255,0.35)"; c.beginPath(); c.moveTo(x, 0); c.lineTo(x, H); c.stroke(); }
        for (let y = 0; y <= H; y += 24) { c.strokeStyle = y % 120 ? "rgba(110,160,255,0.14)" : "rgba(150,200,255,0.35)"; c.beginPath(); c.moveTo(0, y); c.lineTo(W, y); c.stroke(); }
        c.strokeStyle = "rgba(243,195,90,0.8)"; c.lineWidth = 3; c.strokeRect(10, 10, W - 20, H - 20);
      }), roughness: 0.3, metalness: 0.3, clearcoat: 0.8, envMapIntensity: 0.6
    }));
    slabTop.rotation.x = -Math.PI / 2; slabTop.position.y = 0.003; g.add(slabTop);

    /* ---------- chuck + wafer ---------- */
    const chuck = new T.Mesh(new T.CylinderGeometry(2.05, 2.15, 0.22, 96), kit.M.metal(0x9aa6c0, 0.3));
    chuck.position.y = 0.11; g.add(chuck);
    const chuckRing = new T.Mesh(new T.TorusGeometry(2.06, 0.03, 10, 96), kit.M.gold());
    chuckRing.rotation.x = Math.PI / 2; chuckRing.position.y = 0.225; g.add(chuckRing);
    const WR = 1.88;
    const waferTex = kit.canvasTex(1024, 1024, function (c, W, H) {
      const cx = W / 2, R = W / 2;
      const base = c.createRadialGradient(cx, cx, 10, cx, cx, R);
      base.addColorStop(0, "#26346a"); base.addColorStop(1, "#141d44");
      c.fillStyle = base; c.fillRect(0, 0, W, H);
      const dw = 56, dh = 56, rr = kit.rng(9);
      for (let y = -R; y < R; y += dh) for (let x = -R; x < R; x += dw) {
        const x0 = cx + x, y0 = cx + y;
        const corners = [[x0, y0], [x0 + dw, y0], [x0, y0 + dh], [x0 + dw, y0 + dh]];
        let ok = true;
        corners.forEach(function (p) { if (Math.hypot(p[0] - cx, p[1] - cx) > R - 22) ok = false; });
        if (!ok) continue;
        const hue = 205 + ((x0 + y0) / W) * 40 + rr() * 14;
        c.fillStyle = "hsl(" + hue + ",70%," + (28 + rr() * 14) + "%)";
        c.fillRect(x0 + 2, y0 + 2, dw - 4, dh - 4);
        c.strokeStyle = "rgba(190,215,255,0.5)"; c.lineWidth = 1;
        c.strokeRect(x0 + 6.5, y0 + 6.5, dw - 13, dh - 13);
        c.fillStyle = "rgba(255,220,150,0.7)";
        for (let k = 0; k < 4; k++) c.fillRect(x0 + 10 + k * 9, y0 + 10, 4, 4);
      }
      c.strokeStyle = "rgba(220,235,255,0.7)"; c.lineWidth = 5;
      c.beginPath(); c.arc(cx, cx, R - 4, 0, 6.283); c.stroke();
    });
    const wafer = new T.Mesh(new T.CircleGeometry(WR, 96), new T.MeshPhysicalMaterial({ map: waferTex, metalness: 0.7, roughness: 0.2, clearcoat: 1, clearcoatRoughness: 0.08, envMapIntensity: 1.3 }));
    wafer.rotation.x = -Math.PI / 2; wafer.position.y = 0.2635; g.add(wafer);
    const waferEdge = new T.Mesh(new T.CylinderGeometry(WR, WR, 0.04, 96, 1, true), kit.M.metal(0xb7c0d6, 0.15));
    waferEdge.position.y = 0.242; g.add(waferEdge);
    const waferBody = new T.Mesh(new T.CylinderGeometry(WR, WR, 0.04, 96), kit.M.silicon(0x1d2850));
    waferBody.position.y = 0.242; g.add(waferBody);
    const wafGroup = new T.Group(); // for subtle rotation
    // vacuum groove rings visible under wafer edge
    for (let i = 0; i < 3; i++) {
      const r = new T.Mesh(new T.TorusGeometry(2.0 - i * 0.04 + 0.0, 0.006, 6, 96), kit.M.metal(0x667088, 0.4));
      r.rotation.x = Math.PI / 2; r.position.y = 0.224; r.visible = i === 0; wafGroup.add(r);
    }
    g.add(wafGroup);

    /* ---------- test chip ---------- */
    const DIE = 1.3, DT = 0.09, dieY = 0.265;
    const die = new T.Mesh(kit.roundedBox(DIE, DT, DIE, 0.03, 0.015), kit.M.silicon(0x2b3f86));
    die.position.y = dieY + DT / 2; g.add(die);
    const dieTop = new T.Mesh(new T.PlaneGeometry(DIE - 0.34, DIE - 0.34), new T.MeshPhysicalMaterial({
      map: kit.tex.dieLayout(1024, 21, ["#0a1a46", "#3d7bff", "#2fe0ff", "#ffc04a", "#d8e4ff"]),
      metalness: 0.55, roughness: 0.22, clearcoat: 1, clearcoatRoughness: 0.1, envMapIntensity: 1.2
    }));
    dieTop.rotation.x = -Math.PI / 2; dieTop.position.y = dieY + DT + 0.002; g.add(dieTop);
    // bond pads (16)
    const padPos = [];
    const PI = DIE / 2 - 0.1;
    for (let i = 0; i < 4; i++) {
      const u = -0.38 + i * 0.253;
      padPos.push([u, -PI], [u, PI], [-PI, u], [PI, u]);
    }
    const pads = kit.instanced(new T.BoxGeometry(0.1, 0.014, 0.1), kit.M.gold(), padPos.length, function (i, d) {
      d.position.set(padPos[i][0], dieY + DT + 0.007, padPos[i][1]);
    });
    g.add(pads);
    const padRing = kit.instanced(new T.BoxGeometry(0.115, 0.006, 0.115), kit.M.metal(0xbfc6d6, 0.2), padPos.length, function (i, d) {
      d.position.set(padPos[i][0], dieY + DT + 0.002, padPos[i][1]);
    });
    g.add(padRing);

    /* ---------- probe card ---------- */
    const pcShape = new T.Shape();
    pcShape.absarc(0, 0, 2.3, 0, Math.PI * 2, false);
    const hole = new T.Path(); hole.absarc(0, 0, 1.4, 0, Math.PI * 2, true);
    pcShape.holes.push(hole);
    const pcGeo = new T.ExtrudeGeometry(pcShape, { depth: 0.08, bevelEnabled: true, bevelThickness: 0.02, bevelSize: 0.02, bevelSegments: 2, curveSegments: 72 });
    pcGeo.rotateX(-Math.PI / 2);
    const PCY = 1.9;
    const pcMat = new T.MeshPhysicalMaterial({ color: 0x0f3a8f, roughness: 0.4, clearcoat: 0.7, clearcoatRoughness: 0.2, envMapIntensity: 0.6 });
    const card = new T.Mesh(pcGeo, pcMat);
    card.position.y = PCY; g.add(card);
    const cardRing = new T.Mesh(new T.TorusGeometry(1.34, 0.025, 8, 96), kit.M.gold());
    cardRing.rotation.x = Math.PI / 2; cardRing.position.y = PCY + 0.1; g.add(cardRing);
    const cardRing2 = new T.Mesh(new T.TorusGeometry(2.3, 0.012, 6, 96), kit.M.gold());
    cardRing2.rotation.x = Math.PI / 2; cardRing2.position.y = PCY + 0.1; g.add(cardRing2);
    // pcb pads / traces on card
    const cpads = kit.instanced(new T.BoxGeometry(0.12, 0.014, 0.2), kit.M.gold(), 36, function (i, d) {
      const a = i / 36 * Math.PI * 2;
      d.position.set(Math.cos(a) * 1.85, PCY + 0.105, Math.sin(a) * 1.85);
      d.rotation.y = -a;
    });
    g.add(cpads);
    for (let i = 0; i < 36; i += 1) {
      const a = i / 36 * Math.PI * 2;
      const tr = new T.Mesh(new T.BoxGeometry(0.42, 0.004, 0.025), new T.MeshStandardMaterial({ color: 0xd9b25c, metalness: 1, roughness: 0.3 }));
      tr.position.set(Math.cos(a) * 1.57, PCY + 0.103, Math.sin(a) * 1.57);
      tr.rotation.y = -a; g.add(tr);
    }
    // mounting screws
    const screwGeo = new T.CylinderGeometry(0.07, 0.07, 0.06, 6);
    for (let i = 0; i < 6; i++) {
      const a = i / 6 * Math.PI * 2 + 0.26;
      const s = new T.Mesh(screwGeo, kit.M.metal(0xdfe4ee, 0.2));
      s.position.set(Math.cos(a) * 2.18, PCY + 0.125, Math.sin(a) * 2.18); g.add(s);
    }

    /* ---------- tungsten needles ---------- */
    const tung = new T.MeshPhysicalMaterial({ color: 0xd2d8e2, metalness: 1, roughness: 0.14, envMapIntensity: 1.5 });
    const mount = kit.M.epoxy(0x161a2a);
    const chosen = [0, 1, 2, 3, 4, 6, 8, 10, 12, 13, 14, 15];
    const tipGlows = [];
    chosen.forEach(function (idx, n) {
      const p = padPos[idx];
      const tip = new T.Vector3(p[0], dieY + DT + 0.016, p[1]);
      const out = new T.Vector3(p[0], 0, p[1]);
      const rad = out.length() || 1; out.normalize();
      const elbow = tip.clone().addScaledVector(out, 0.45); elbow.y = tip.y + 0.55;
      const att = out.clone().multiplyScalar(1.58); att.y = PCY - 0.01;
      // ceramic mounting block
      const blk = new T.Mesh(new T.BoxGeometry(0.2, 0.09, 0.16), mount);
      blk.position.set(att.x, PCY - 0.07, att.z); blk.rotation.y = -Math.atan2(att.z, att.x); g.add(blk);
      seg(new T.Vector3(att.x, PCY - 0.07, att.z), elbow, 0.026, 0.017, tung);
      seg(elbow, tip, 0.017, 0.004, tung);
      const jt = new T.Mesh(new T.SphereGeometry(0.019, 10, 8), tung); jt.position.copy(elbow); g.add(jt);
      const gl = kit.glow(CYAN, 0.32, 0.0); gl.position.copy(tip).setY(tip.y + 0.01); g.add(gl); tipGlows.push(gl);
    });

    /* ---------- scope hologram panel ---------- */
    const sp = new T.Group(); sp.position.set(3.05, 1.35, 0.15); sp.rotation.y = -0.4; g.add(sp);
    const SW = 2.1, SH = 1.45;
    const sf = new T.Mesh(kit.roundedBox(SW + 0.12, 0.05, SH + 0.12, 0.07, 0.014), kit.M.gold());
    sf.rotation.x = Math.PI / 2; sf.position.z = -0.03; sp.add(sf);
    const cvs = kit.canvas(512, 352); const cx = cvs.getContext("2d");
    const stex = new T.CanvasTexture(cvs); stex.encoding = T.sRGBEncoding;
    const screen = new T.Mesh(new T.PlaneGeometry(SW, SH), new T.MeshBasicMaterial({ map: stex, toneMapped: false }));
    sp.add(screen);
    const rs = kit.rng(77);
    const NT = 36, NP = 40, lv = [];
    for (let i = 0; i < NT; i++) lv.push([rs() < 0.5 ? 0 : 1, rs() < 0.5 ? 0 : 1, (rs() - 0.5) * 0.04, rs()]);
    function drawScope(t) {
      const W = 512, H = 352;
      cx.fillStyle = "#06122f"; cx.fillRect(0, 0, W, H);
      cx.lineWidth = 1;
      for (let x = 30; x < W; x += 36) { cx.strokeStyle = "rgba(120,170,255,0.22)"; cx.beginPath(); cx.moveTo(x, 40); cx.lineTo(x, H - 40); cx.stroke(); }
      for (let y = 40; y < H - 30; y += 34) { cx.strokeStyle = "rgba(120,170,255,0.22)"; cx.beginPath(); cx.moveTo(30, y); cx.lineTo(W - 30, y); cx.stroke(); }
      cx.fillStyle = "rgba(255,210,110,0.95)"; cx.font = "bold 17px Inter, system-ui, sans-serif";
      cx.fillText("DUT  EYE DIAGRAM", 30, 28);
      cx.fillStyle = "rgba(120,255,190,0.95)"; cx.fillText("BER 0.0e0", W - 134, 28);
      cx.lineWidth = 2; cx.strokeStyle = "rgba(95,227,255,0.45)";
      const amp = (H - 130) / 2, mid = H / 2 + 6;
      for (let i = 0; i < NT; i++) {
        const l = lv[i];
        cx.beginPath();
        for (let k = 0; k <= NP; k++) {
          const x = k / NP;
          let s = (x - 0.25) / 0.5; s = s < 0 ? 0 : s > 1 ? 1 : s; s = s * s * (3 - 2 * s);
          const v = l[0] + (l[1] - l[0]) * s;
          const y = mid + (0.5 - v) * 2 * amp * 0.85 + Math.sin(t * 2 + l[3] * 9 + x * 4) * 3 + l[2] * 100 * Math.sin(x * 3.1416);
          const px = 36 + x * (W - 72);
          if (k === 0) cx.moveTo(px, y); else cx.lineTo(px, y);
        }
        cx.stroke();
      }
      cx.strokeStyle = "rgba(255,255,255,0.35)"; cx.lineWidth = 1;
      cx.beginPath(); cx.moveTo(36 + (W - 72) * 0.5, 40); cx.lineTo(36 + (W - 72) * 0.5, H - 40); cx.stroke();
      // scan line
      const sx = ((t * 0.3) % 1) * (W - 72) + 36;
      cx.fillStyle = "rgba(160,220,255,0.12)"; cx.fillRect(sx - 3, 40, 6, H - 80);
      stex.needsUpdate = true;
    }
    drawScope(0);
    // glass cover
    const cover = new T.Mesh(new T.PlaneGeometry(SW, SH), kit.M.glass(0xaee0ff, 0.08)); cover.position.z = 0.012; sp.add(cover);
    // PASS badge
    const badge = new T.Mesh(kit.roundedBox(1.05, 0.06, 0.34, 0.14, 0.02), kit.M.gold());
    badge.rotation.x = Math.PI / 2; badge.position.set(0.2, -SH / 2 - 0.35, 0.02); sp.add(badge);
    const bface = new T.Mesh(new T.PlaneGeometry(0.97, 0.26), new T.MeshBasicMaterial({ color: 0x0a3a1f, toneMapped: false }));
    bface.position.set(0.2, -SH / 2 - 0.35, 0.056); sp.add(bface);
    const ptxt = kit.label("PASS", 0.2, { color: "#5dffa8", size: 120, weight: "800" });
    ptxt.position.set(0.28, -SH / 2 - 0.35, 0.06); sp.add(ptxt);
    const led = new T.Mesh(new T.SphereGeometry(0.06, 20, 14), kit.M.glow(GREEN, 3));
    led.position.set(-0.17, -SH / 2 - 0.35, 0.09); sp.add(led);
    const ledGlow = kit.glow(GREEN, 0.7, 0.8); ledGlow.position.copy(led.position); ledGlow.position.z += 0.05; sp.add(ledGlow);
    // stand posts
    [-0.6, 0.6].forEach(function (x) {
      const p = new T.Mesh(new T.CylinderGeometry(0.02, 0.02, 0.8, 10), kit.M.metal(0xc8d0e0, 0.2));
      p.position.set(x, -SH / 2 - 0.1 - 0.4 + 0.1, -0.05); p.userData.x = 1; sp.add(p);
      p.visible = false;
    });

    /* ---------- medal + laurel ---------- */
    const md = new T.Group(); md.position.set(-3.05, 0, 0.1); md.rotation.y = -0.25; g.add(md);
    const pedestal = new T.Mesh(new T.CylinderGeometry(0.55, 0.65, 0.16, 48), kit.M.metal(0x2f3f78, 0.3));
    pedestal.position.y = 0.08; md.add(pedestal);
    const pedRing = new T.Mesh(new T.TorusGeometry(0.56, 0.02, 8, 48), kit.M.gold()); pedRing.rotation.x = Math.PI / 2; pedRing.position.y = 0.16; md.add(pedRing);
    const post = new T.Mesh(new T.CylinderGeometry(0.05, 0.07, 0.55, 16), kit.M.gold()); post.position.y = 0.43; md.add(post);
    const medalG = new T.Group(); medalG.scale.setScalar(1.2); medalG.position.y = 1.3; md.add(medalG);
    const disc = new T.Mesh(new T.CylinderGeometry(0.55, 0.55, 0.09, 64), kit.M.gold());
    disc.rotation.x = Math.PI / 2; medalG.add(disc);
    const rim = new T.Mesh(new T.TorusGeometry(0.55, 0.045, 12, 64), kit.M.gold()); medalG.add(rim);
    const faceTex = kit.canvasTex(512, 512, function (c, W, H) {
      const gr = c.createRadialGradient(256, 256, 20, 256, 256, 250);
      gr.addColorStop(0, "#ffe9a8"); gr.addColorStop(1, "#d9a63a");
      c.fillStyle = gr; c.fillRect(0, 0, W, H);
      c.strokeStyle = "#8a5c10"; c.lineWidth = 6; c.beginPath(); c.arc(256, 256, 206, 0, 6.283); c.stroke();
      c.lineWidth = 2.5; c.beginPath(); c.arc(256, 256, 190, 0, 6.283); c.stroke();
      c.fillStyle = "#b07a1c"; c.beginPath();
      for (let i = 0; i < 10; i++) { const a = -Math.PI / 2 + i * Math.PI / 5, r = i % 2 ? 62 : 150; c.lineTo(256 + Math.cos(a) * r, 256 + Math.sin(a) * r); }
      c.closePath(); c.fill();
      c.fillStyle = "#fff0bd"; c.beginPath();
      for (let i = 0; i < 10; i++) { const a = -Math.PI / 2 + i * Math.PI / 5, r = i % 2 ? 46 : 118; c.lineTo(256 + Math.cos(a) * r, 256 + Math.sin(a) * r); }
      c.closePath(); c.fill();
    });
    [1, -1].forEach(function (s) {
      const f = new T.Mesh(new T.CircleGeometry(0.5, 64), new T.MeshPhysicalMaterial({ map: faceTex, metalness: 0.5, roughness: 0.35, clearcoat: 0.6, envMapIntensity: 1.0, emissive: 0x7a5410, emissiveMap: faceTex }));
      f.position.z = s * 0.048; if (s < 0) f.rotation.y = Math.PI; medalG.add(f);
    });
    // laurel
    const NL = 22;
    const leafGeo = new T.SphereGeometry(0.075, 12, 8); leafGeo.scale(1.7, 0.45, 0.8);
    const laurel = kit.instanced(leafGeo, kit.M.gold(), NL * 2, function (i, d) {
      const side = i < NL ? 1 : -1, k = (i % NL) / (NL - 1);
      const a = Math.PI * (1.05 + k * 0.85); // sweep around bottom
      const r = 0.74;
      d.position.set(side * Math.cos(a) * -r * 1.0, Math.sin(a) * r, side * 0 + 0.0);
      d.position.x = side * Math.abs(Math.cos(a)) * r * (Math.cos(a) < 0 ? 1 : 1);
      d.position.x = side * Math.cos(Math.PI * (k * 0.9 + 0.08) * 1.0 - Math.PI * 0.0) * 0;
      const th = -Math.PI / 2 + side * (0.3 + k * 1.9);
      d.position.set(Math.cos(th) * r, Math.sin(th) * r, 0);
      d.rotation.z = th + Math.PI / 2 + (i % 2 ? 0.6 : -0.6);
    });
    medalG.add(laurel);
    // ribbon
    const ribMat = new T.MeshPhysicalMaterial({ color: 0x2457d6, roughness: 0.5, side: T.DoubleSide, clearcoat: 0.3 });
    [-1, 1].forEach(function (s) {
      const rb = new T.Mesh(new T.PlaneGeometry(0.2, 0.7), ribMat);
      rb.position.set(s * 0.2, 0.9, 0); rb.rotation.z = s * -0.38; medalG.add(rb);
    });
    const loop = new T.Mesh(new T.TorusGeometry(0.06, 0.017, 8, 20), kit.M.gold()); loop.position.set(0, 0.62, 0); medalG.add(loop);
    const mglow = kit.glow(0xffd270, 2.0, 0.18); mglow.position.set(0, 0, -0.1); medalG.add(mglow);

    g.add(kit.contactShadow(3.9, 0.8)).position.y = -0.2;

    const nT = tipGlows.length;
    function update(t) {
      // scope (~20 fps)
      if (((t * 20) | 0) !== update._f) { update._f = (t * 20) | 0; drawScope(t); }
      sp.position.y = 1.35 + Math.sin(t * 0.8) * 0.05;
      const p = 0.7 + 0.3 * Math.sin(t * 3);
      led.material.emissiveIntensity = 2.4 + 1.4 * p; ledGlow.material.opacity = 0.5 + 0.4 * p;
      medalG.rotation.y = Math.sin(t * 0.6) * 0.9;
      medalG.position.y = 1.3 + Math.sin(t * 0.9) * 0.04;
      for (let i = 0; i < nT; i++) {
        const ph = ((t * 0.7 - i / nT) % 1 + 1) % 1;
        tipGlows[i].material.opacity = Math.max(0, 1 - ph * 3.5) * 0.9;
      }
    }
    update._f = -1;
    return { group: g, update: update, fit: 2.2, spin: 0.002 };
  };
})();
