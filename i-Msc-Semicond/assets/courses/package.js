/* core_5 - Assembly, Test & Packaging : exploded flip-chip BGA package */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  window.IMEN.courseScenes = window.IMEN.courseScenes || {};

  window.IMEN.courseScenes.package = function (kit) {
    const T = THREE;
    const g = new T.Group();
    const stack = new T.Group();
    stack.rotation.x = 0.2;
    g.add(stack);

    /* ---------- local helpers ---------- */
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
    // frame with a rectangular hole, extruded along +Y (base at y=0)
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
    // draw in world units (origin at centre) onto a canvas texture that maps 1:1 to a roundedBox cap
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
    function lay(mesh, y) { mesh.position.y = y; stack.add(mesh); return mesh; }

    /* ---------- shared glow sprite ---------- */
    const glowCy = kit.glow(0x5ad2ff, 0.12, 0.9);
    const glowGreen = kit.glow(0x6dffb0, 0.7, 0.9);

    /* ================= SUBSTRATE ================= */
    const SW = 3.2, SH = 0.1;
    function subDraw(bump) {
      return function (c) {
        const r = kit.rng(9);
        c.fillStyle = bump ? "#4a4a4a" : "#0b5a3b";
        c.fillRect(-SW / 2, -SW / 2, SW, SW);
        if (!bump) {
          for (let i = 0; i < 1800; i++) {
            c.fillStyle = "rgba(" + (r() < 0.5 ? "0,0,0," : "120,255,190,") + (r() * 0.05) + ")";
            c.fillRect((r() - 0.5) * SW, (r() - 0.5) * SW, 0.01 + r() * 0.03, 0.01 + r() * 0.03);
          }
        }
        const gold = bump ? "#cfcfcf" : "rgba(226,184,92,0.92)";
        const goldBright = bump ? "#fff" : "#f1d27a";
        c.lineCap = "round";
        c.lineJoin = "round";
        // fan-out traces from the landing zone to the ball vias
        c.strokeStyle = gold;
        c.lineWidth = 0.011;
        const Z = 0.72;
        for (let side = 0; side < 4; side++) {
          for (let k = 0; k < 24; k++) {
            const u = -0.69 + k * 0.06;
            const a = side * Math.PI / 2;
            const ca = Math.cos(a), sa = Math.sin(a);
            // local coords: u along edge, v outward
            const loc = function (uu, vv) { return [uu * ca - vv * sa, uu * sa + vv * ca]; };
            const tgtU = Math.max(-1.3, Math.min(1.3, u * 1.9 + (r() - 0.5) * 0.1));
            const tgtV = 1.0 + Math.floor(r() * 3) * 0.2;
            const p0 = loc(u, Z), p1 = loc(u, Z + 0.06 + r() * 0.1);
            const bend = Math.abs(tgtU - u) * 0.9;
            const p2 = loc(tgtU, Z + 0.06 + r() * 0.1 + bend);
            const p3 = loc(tgtU, tgtV);
            c.beginPath();
            c.moveTo(p0[0], p0[1]);
            c.lineTo(p1[0], p1[1]);
            c.lineTo(p2[0], p2[1]);
            c.lineTo(p3[0], p3[1]);
            c.stroke();
            c.fillStyle = goldBright;
            c.beginPath();
            c.arc(p3[0], p3[1], 0.026, 0, 6.3);
            c.fill();
          }
        }
        // die-attach landing grid
        c.fillStyle = bump ? "#777" : "#0e6a45";
        c.fillRect(-0.76, -0.76, 1.52, 1.52);
        c.strokeStyle = bump ? "#aaa" : "rgba(240,215,140,0.55)";
        c.lineWidth = 0.008;
        c.strokeRect(-0.76, -0.76, 1.52, 1.52);
        c.fillStyle = goldBright;
        for (let i = 0; i < 24; i++) for (let j = 0; j < 24; j++) {
          c.beginPath();
          c.arc(-0.69 + i * 0.06, -0.69 + j * 0.06, 0.016, 0, 6.3);
          c.fill();
        }
        // vias on the ball pitch
        c.fillStyle = bump ? "#bbb" : "rgba(214,176,92,0.9)";
        for (let i = 0; i < 14; i++) for (let j = 0; j < 14; j++) {
          const x = (i - 6.5) * 0.2, y = (j - 6.5) * 0.2;
          if (Math.abs(x) < 0.85 && Math.abs(y) < 0.85) continue;
          c.beginPath();
          c.arc(x, y, 0.024, 0, 6.3);
          c.fill();
        }
        // silkscreen
        if (!bump) {
          c.strokeStyle = "rgba(235,245,235,0.8)";
          c.lineWidth = 0.012;
          [[-1,-1],[1,-1],[1,1],[-1,1]].forEach(function (p) {
            const x = p[0] * 1.5, y = p[1] * 1.5;
            c.beginPath();
            c.moveTo(x - p[0] * 0.12, y); c.lineTo(x, y); c.lineTo(x, y - p[1] * 0.12);
            c.stroke();
          });
          c.fillStyle = "rgba(235,245,235,0.85)";
          c.beginPath(); c.moveTo(-1.5, -1.5); c.lineTo(-1.36, -1.5); c.lineTo(-1.5, -1.36); c.closePath(); c.fill();
        }
      };
    }
    const subCap = new T.MeshPhysicalMaterial({
      map: worldTex(1024, SW, SW, subDraw(false)),
      bumpMap: worldTex(1024, SW, SW, subDraw(true), true), bumpScale: 1.6,
      roughness: 0.4, metalness: 0.05, clearcoat: 0.7, clearcoatRoughness: 0.22, envMapIntensity: 0.6
    });
    const subSide = kit.M.plastic(0x1a7f55, 0.5);
    const substrate = new T.Mesh(kit.roundedBox(SW, SH, SW, 0.05, 0.012), [subCap, subSide]);
    lay(substrate, 0);

    // passives on the substrate (instanced 0402 caps)
    const capPos = [];
    for (let i = 0; i < 6; i++) {
      const z = -1.0 + i * 0.4;
      capPos.push([-1.46, z, 0], [1.46, z, 0], [z, -1.46, 1], [z, 1.46, 1]);
    }
    const capBody = kit.instanced(new T.BoxGeometry(0.1, 0.04, 0.052), kit.M.ceramic(0x8a6a44), capPos.length, function (i, d) {
      d.position.set(capPos[i][0], SH / 2 + 0.02, capPos[i][1]);
      d.rotation.y = capPos[i][2] ? 0 : Math.PI / 2;
    });
    const capEnd = kit.instanced(new T.BoxGeometry(0.022, 0.044, 0.056), kit.M.solder(), capPos.length * 2, function (i, d) {
      const p = capPos[i >> 1], s = i & 1 ? 1 : -1;
      const vert = !p[2];
      d.position.set(p[0] + (vert ? 0 : s * 0.039), SH / 2 + 0.022, p[1] + (vert ? s * 0.039 : 0));
      d.rotation.y = vert ? Math.PI / 2 : 0;
    });
    substrate.add(capBody, capEnd);

    // solder-ball grid (instanced) hanging below the substrate
    const BR = 0.082;
    const balls = kit.instanced(new T.SphereGeometry(BR, 12, 8), kit.M.solder(), 196, function (i, d) {
      const x = (i % 14 - 6.5) * 0.2, z = (Math.floor(i / 14) - 6.5) * 0.2;
      d.position.set(x, 0, z);
      d.scale.set(1, 0.9, 1);
    });
    const ballLayer = new T.Group();
    ballLayer.add(balls);
    stack.add(ballLayer);
    // copper ball pads (visible as a tiny ring at the neck)
    const pads = kit.instanced(new T.CylinderGeometry(0.06, 0.06, 0.006, 10), kit.M.copper(), 196, function (i, d) {
      d.position.set((i % 14 - 6.5) * 0.2, -SH / 2 - 0.003, (Math.floor(i / 14) - 6.5) * 0.2);
    });
    substrate.add(pads);

    /* ================= MICRO-BUMPS + UNDERFILL ================= */
    const bumpLayer = new T.Group();
    const NB = 24, BP = 0.058;
    const bumps = kit.instanced(new T.CylinderGeometry(0.0185, 0.0185, 0.06, 8), kit.M.solder(), NB * NB, function (i, d) {
      d.position.set((i % NB - (NB - 1) / 2) * BP, 0, (Math.floor(i / NB) - (NB - 1) / 2) * BP);
    });
    const caps = kit.instanced(new T.CylinderGeometry(0.0140, 0.0185, 0.012, 8), kit.M.copper(), NB * NB, function (i, d) {
      d.position.set((i % NB - (NB - 1) / 2) * BP, -0.035, (Math.floor(i / NB) - (NB - 1) / 2) * BP);
    });
    const fillMat = new T.MeshPhysicalMaterial({
      color: 0xe0902f, transparent: true, opacity: 0.34, roughness: 0.12, metalness: 0, clearcoat: 1, clearcoatRoughness: 0.05,
      envMapIntensity: 1.4, depthWrite: false
    });
    const fillGeo = kit.roundedBox(1.66, 0.075, 1.66, 0.05, 0.02);
    const underfill = new T.Mesh(fillGeo, fillMat);
    underfill.renderOrder = 2;
    bumpLayer.add(bumps, caps, underfill);
    stack.add(bumpLayer);

    /* ================= DIE ================= */
    const DW = 1.5, DH = 0.1;
    const dieMap = kit.tex.dieLayout(1024, 21, ["#071022", "#2f7bff", "#12d4ff", "#ffb340", "#d0dcff"]);
    capUV(dieMap, DW, DW);
    const dieCap = new T.MeshPhysicalMaterial({
      map: dieMap, emissiveMap: dieMap, emissive: 0xffffff, emissiveIntensity: 0.22,
      metalness: 0.65, roughness: 0.2, clearcoat: 0.9, clearcoatRoughness: 0.08, envMapIntensity: 0.9
    });
    const die = new T.Mesh(kit.roundedBox(DW, DH, DW, 0.025, 0.012), [dieCap, kit.M.silicon(0x1d2744)]);
    lay(die, 0);
    // glowing seal-ring around the die top
    const rim = new T.Mesh(ringGeo(DW - 0.02, DW - 0.02, DW - 0.06, DW - 0.06, 0.012, 0.02, 0.003), kit.M.glow(0x44d0ff, 2.4));
    rim.position.y = DH / 2 - 0.003;
    die.add(rim);

    /* ================= TIM ================= */
    const timMat = new T.MeshPhysicalMaterial({ color: 0xdde3ec, metalness: 1, roughness: 0.14, bumpMap: kit.tex.grain(256, 256, 4, [6, 6]), bumpScale: 0.25, envMapIntensity: 1.3 });
    const tim = new T.Mesh(kit.roundedBox(1.62, 0.034, 1.62, 0.05, 0.01), timMat);
    lay(tim, 0);

    /* ================= LID (heat spreader) ================= */
    const lid = new T.Group();
    const lidBump = kit.canvasTex(512, 512, function (c, W, H) {
      const r = kit.rng(3);
      c.fillStyle = "#808080"; c.fillRect(0, 0, W, H);
      for (let i = 0; i < 1400; i++) {
        const v = 110 + r() * 40;
        c.strokeStyle = "rgba(" + v + "," + v + "," + v + ",0.22)";
        c.lineWidth = 0.6 + r();
        const y = r() * H, x = r() * W, l = 60 + r() * 260;
        c.beginPath(); c.moveTo(x, y); c.lineTo(x + l, y + (r() - 0.5)); c.stroke();
      }
    }, { linear: true, repeat: [0.8, 0.8] });
    const lidMat = new T.MeshPhysicalMaterial({ color: 0xf0a47c, metalness: 1, roughness: 0.3, bumpMap: lidBump, bumpScale: 0.35, clearcoat: 0.25, clearcoatRoughness: 0.2, envMapIntensity: 1.5 });
    const lidTop = new T.Mesh(kit.roundedBox(2.5, 0.15, 2.5, 0.09, 0.02), lidMat);
    lidTop.position.y = 0.075;
    const flange = new T.Mesh(ringGeo(3.0, 3.0, 2.2, 2.2, 0.06, 0.1, 0.014), kit.M.metal(0xd9a283, 0.28));
    flange.position.y = 0;
    // nickel-plated inner skirt (visible from the side)
    const skirt = new T.Mesh(ringGeo(2.62, 2.62, 2.46, 2.46, 0.08, 0.08, 0.008), kit.M.metal(0xcfd4dc, 0.18));
    skirt.position.y = -0.05;
    // laser etch
    const etch = kit.label("IMEN", 0.42, { size: 128, color: "#5a2f1c", weight: "800" });
    etch.rotation.x = -Math.PI / 2;
    etch.position.set(0, 0.1535, 0.55);
    etch.material.opacity = 0.7;
    const etch2 = kit.label("FC-BGA  ·  196 BALL  ·  Cu/Ni", 0.1, { size: 64, color: "#5a2f1c", weight: "600" });
    etch2.rotation.x = -Math.PI / 2;
    etch2.position.set(0, 0.1535, 0.9);
    etch2.material.opacity = 0.6;
    lid.add(lidTop, flange, skirt, etch, etch2);
    stack.add(lid);

    /* ================= GUIDE LINES + node dots ================= */
    const guideMat = new T.LineBasicMaterial({ color: 0x4cc8ff, transparent: true, opacity: 0.5, depthWrite: false });
    const guideGeo = new T.BufferGeometry().setFromPoints([new T.Vector3(0, 0, 0), new T.Vector3(0, 1, 0)]);
    const guides = [];
    const corners = [[-1.78, -1.78], [1.78, -1.78], [1.78, 1.78], [-1.78, 1.78]];
    corners.forEach(function (c) {
      const l = new T.Line(guideGeo, guideMat);
      l.position.set(c[0], 0, c[1]);
      l.userData.noFit = true;
      stack.add(l);
      guides.push(l);
    });
    const dots = [];
    corners.forEach(function (c) {
      for (let k = 0; k < 6; k++) {
        const s = glowCy.clone();
        s.position.set(c[0], 0, c[1]);
        stack.add(s);
        dots.push(s);
      }
    });
    // travelling pulse per corner
    const pulses = corners.map(function (c, i) {
      const s = kit.glow(0xc8f2ff, 0.28, 1);
      s.position.set(c[0], 0, c[1]);
      stack.add(s);
      return s;
    });

    /* ================= TEST PROBES (pogo pins) ================= */
    const probe = new T.Group();
    const strip = new T.Mesh(kit.roundedBox(3.3, 0.12, 0.46, 0.04, 0.012), kit.M.epoxy(0x15181f));
    strip.position.y = -0.91;
    const pinX = function (i) { return (i - 6.5) * 0.2; };
    const NP = 14;
    const discs = kit.instanced(new T.CylinderGeometry(0.085, 0.085, 0.008, 18), kit.M.gold(), NP, function (i, d) {
      d.position.set(pinX(i), -0.846, 0);
    });
    const barrels = kit.instanced(new T.CylinderGeometry(0.055, 0.055, 0.3, 14), kit.M.gold(), NP, function (i, d) {
      d.position.set(pinX(i), -0.70, 0);
    });
    const helix = [];
    for (let k = 0; k <= 40; k++) {
      const a = (k / 40) * Math.PI * 2 * 4.5;
      helix.push(new T.Vector3(Math.cos(a) * 0.034, -0.55 + (k / 40) * 0.2, Math.sin(a) * 0.034));
    }
    const springGeo = new T.TubeGeometry(new T.CatmullRomCurve3(helix), 56, 0.0075, 5, false);
    const springs = kit.instanced(springGeo, kit.M.metal(0xdde3ea, 0.2), NP, function (i, d) { d.position.set(pinX(i), 0, 0); });
    const plung = kit.instanced(new T.CylinderGeometry(0.036, 0.036, 0.14, 12), kit.M.metal(0xe6e9ee, 0.15), NP, function (i, d) {
      d.position.set(pinX(i), -0.31, 0);
    });
    const tips = kit.instanced(new T.ConeGeometry(0.046, 0.06, 6), kit.M.gold(), NP, function (i, d) {
      d.position.set(pinX(i), -0.231, 0);
    });
    probe.add(strip, discs, barrels, springs, plung, tips);
    probe.position.z = 1.3;
    probe.userData.baseY = 0;
    stack.add(probe);
    const scan = glowGreen;
    scan.position.set(0, -0.2, 1.3);
    stack.add(scan);

    /* ================= LABELS ================= */
    const labelDefs = [
      ["Heat spreader (Cu/Ni lid)", function () { return lid.position.y + 0.07; }],
      ["Thermal interface (TIM)", function () { return tim.position.y; }],
      ["Silicon die", function () { return die.position.y; }],
      ["Micro-bumps + underfill", function () { return bumpLayer.position.y; }],
      ["Organic substrate", function () { return 0; }],
      ["Solder balls (BGA)", function () { return ballLayer.position.y; }],
      ["Test probes", function () { return ballLayer.position.y - 0.5; }]
    ];
    const labelLines = new T.LineBasicMaterial({ color: 0x7fb0ff, transparent: true, opacity: 0.4, depthWrite: false });
    const labels = labelDefs.map(function (d) {
      const t = kit.tex.text(d[0], { size: 72, color: "#d6e6ff", weight: "600" });
      const sp = new T.Sprite(new T.SpriteMaterial({ map: t.tex, transparent: true, depthWrite: false, depthTest: false, toneMapped: false, opacity: 0.92 }));
      sp.center.set(0, 0.5);
      sp.scale.set(0.17 * t.aspect, 0.17, 1);
      sp.userData.noFit = true;
      const ln = new T.Line(new T.BufferGeometry().setFromPoints([new T.Vector3(0, 0, 0), new T.Vector3(0.28, 0, 0)]), labelLines);
      ln.userData.noFit = true;
      stack.add(sp, ln);
      return { sp: sp, ln: ln, fn: d[1] };
    });

    /* contact shadow */
    const sh = kit.contactShadow(2.6, 0.55);
    sh.position.y = -1.75;
    g.add(sh);

    /* ================= animate ================= */
    function layout(t) {
      const e = 1 + 0.14 * Math.sin(t * 0.7);
      ballLayer.position.y = -0.42 * e;
      substrate.position.y = 0;
      bumpLayer.position.y = 0.58 * e;
      die.position.y = 1.08 * e;
      tim.position.y = 1.42 * e;
      lid.position.y = 1.78 * e;
      const top = lid.position.y + 0.2;
      for (let i = 0; i < 4; i++) {
        guides[i].position.y = ballLayer.position.y - 0.7;
        guides[i].scale.y = top - ballLayer.position.y + 0.7;
      }
      const ys = [0, bumpLayer.position.y, die.position.y, tim.position.y, lid.position.y + 0.06, ballLayer.position.y];
      for (let i = 0; i < 4; i++) for (let k = 0; k < 6; k++) dots[i * 6 + k].position.y = ys[k];
      for (let i = 0; i < 4; i++) {
        const p = ((t * 0.18 + i * 0.25) % 1);
        pulses[i].position.y = guides[i].position.y + p * guides[i].scale.y;
        pulses[i].material.opacity = Math.sin(p * Math.PI);
      }
      return e;
    }
    function placeLabels(e) {
      for (let i = 0; i < labels.length; i++) {
        const L = labels[i];
        const y = L.fn();
        L.ln.position.set(1.95, y, 0);
        L.sp.position.set(2.25, y, 0);
      }
    }
    placeLabels(layout(0));

    return {
      group: g,
      fit: 2.05,
      spin: 0.0022,
      update: function (t) {
        const e = layout(t);
        placeLabels(e);
        stack.rotation.x = 0.08 + 0.16 * Math.sin(t * 0.35);
        const press = Math.sin(t * 1.4) * 0.018;
        probe.position.y = ballLayer.position.y + 0.117 + press;
        const sp = (t * 0.45) % 2;
        const x = (sp < 1 ? sp : 2 - sp) * 2.6 - 1.3;
        scan.position.x = x;
        scan.position.y = ballLayer.position.y - 0.083 + press;
        scan.material.opacity = 0.7 + 0.3 * Math.sin(t * 6);
        rim.material.emissiveIntensity = 2 + 0.8 * Math.sin(t * 1.8);
        dieCap.emissiveIntensity = 0.2 + 0.08 * Math.sin(t * 1.3);
      }
    };
  };
})();
