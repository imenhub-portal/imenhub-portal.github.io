/* Project I - the DESIGN stage: holographic exploded IC layout stack above a drafting table */
(function () {
  "use strict";
  IMEN.courseScenes.project1 = function (kit) {
    const T = THREE;
    const g = new T.Group();
    const BLUE = 0x3d8bff, GOLD = 0xf3c35a, CYAN = 0x5fe3ff;

    /* ---------- drafting plate ---------- */
    const plate = new T.Mesh(kit.roundedBox(5.8, 0.16, 3.3, 0.14, 0.04), kit.M.metal(0x2a3a6e, 0.35));
    plate.position.y = -0.08;
    g.add(plate);
    const trim = new T.Mesh(kit.roundedBox(5.9, 0.035, 3.4, 0.16, 0.012), kit.M.gold());
    trim.position.y = -0.14;
    g.add(trim);
    const gridTex = kit.canvasTex(1024, 576, function (c, W, H) {
      c.fillStyle = "#0b1d4e"; c.fillRect(0, 0, W, H);
      const grd = c.createRadialGradient(W / 2, H / 2, 20, W / 2, H / 2, W * 0.6);
      grd.addColorStop(0, "rgba(70,130,255,0.35)"); grd.addColorStop(1, "rgba(0,0,0,0)");
      c.fillStyle = grd; c.fillRect(0, 0, W, H);
      for (let x = 0; x <= W; x += 16) {
        c.strokeStyle = (x % 128 === 0) ? "rgba(150,200,255,0.65)" : "rgba(110,160,255,0.22)";
        c.lineWidth = (x % 128 === 0) ? 1.6 : 1;
        c.beginPath(); c.moveTo(x, 0); c.lineTo(x, H); c.stroke();
      }
      for (let y = 0; y <= H; y += 16) {
        c.strokeStyle = (y % 128 === 0) ? "rgba(150,200,255,0.65)" : "rgba(110,160,255,0.22)";
        c.lineWidth = (y % 128 === 0) ? 1.6 : 1;
        c.beginPath(); c.moveTo(0, y); c.lineTo(W, y); c.stroke();
      }
      c.strokeStyle = "rgba(243,195,90,0.8)"; c.lineWidth = 3; c.strokeRect(14, 14, W - 28, H - 28);
      c.strokeStyle = "rgba(243,195,90,0.5)"; c.lineWidth = 1.5;
      c.beginPath(); c.arc(W / 2, H / 2, 150, 0, 6.283); c.stroke();
      c.beginPath(); c.arc(W / 2, H / 2, 210, 0, 6.283); c.stroke();
      c.beginPath(); c.moveTo(W / 2 - 250, H / 2); c.lineTo(W / 2 + 250, H / 2);
      c.moveTo(W / 2, H / 2 - 250); c.lineTo(W / 2, H / 2 + 250); c.stroke();
    });
    const top = new T.Mesh(new T.PlaneGeometry(5.6, 3.15), new T.MeshPhysicalMaterial({ map: gridTex, roughness: 0.35, metalness: 0.2, clearcoat: 0.8, clearcoatRoughness: 0.15, envMapIntensity: 0.6 }));
    top.rotation.x = -Math.PI / 2; top.position.y = 0.002;
    g.add(top);

    /* ---------- projector disc + beam ---------- */
    const proj = new T.Mesh(new T.CylinderGeometry(0.55, 0.65, 0.1, 48), kit.M.gold());
    proj.position.y = 0.05; g.add(proj);
    const projTop = new T.Mesh(new T.CylinderGeometry(0.46, 0.46, 0.012, 48), kit.M.glow(CYAN, 2.2));
    projTop.position.y = 0.106; g.add(projTop);
    const beam = kit.beam(0.5, 1.3, 1.3, "95,190,255", 0.35);
    beam.position.y = 0.11 + 0.55; g.add(beam);
    const pglow = kit.glow(CYAN, 2.2, 0.6); pglow.position.y = 0.2; g.add(pglow);

    /* ---------- IC layer stack ---------- */
    function maskTex(kind, seed) {
      const r = kit.rng(seed);
      return kit.canvasTex(512, 512, function (c, W, H) {
        c.clearRect(0, 0, W, H);
        function bar(x, y, w, h, fill, edge) {
          c.fillStyle = fill; c.fillRect(x, y, w, h);
          c.strokeStyle = edge; c.lineWidth = 1.5; c.strokeRect(x, y, w, h);
        }
        if (kind === 0) { // active diffusion
          for (let i = 0; i < 26; i++) {
            const w = 50 + r() * 110, h = 24 + r() * 26, x = r() * (W - w), y = 22 + Math.floor(r() * 9) * 53;
            bar(x, y, w, h, "rgba(40,230,160,0.55)", "rgba(120,255,200,0.95)");
          }
        } else if (kind === 1) { // poly gates
          for (let x = 24; x < W; x += 26 + Math.floor(r() * 3) * 6) {
            const y0 = r() * H * 0.3, y1 = H * (0.65 + r() * 0.35);
            bar(x, y0, 9, y1 - y0, "rgba(255,150,60,0.75)", "rgba(255,200,120,1)");
          }
        } else if (kind === 2) { // metal1 horizontal
          for (let y = 20; y < H; y += 30) {
            let x = r() * 60;
            while (x < W - 40) {
              const w = 50 + r() * 190;
              bar(x, y, Math.min(w, W - 20 - x), 12, "rgba(60,150,255,0.7)", "rgba(150,210,255,1)");
              x += w + 20 + r() * 70;
            }
          }
        } else if (kind === 3) { // metal2 vertical
          for (let x = 34; x < W; x += 44) {
            let y = r() * 60;
            while (y < H - 40) {
              const h = 60 + r() * 200;
              bar(x, y, 15, Math.min(h, H - 20 - y), "rgba(120,235,255,0.6)", "rgba(210,250,255,1)");
              y += h + 24 + r() * 80;
            }
          }
        } else { // vias
          for (let i = 0; i < 150; i++) {
            const x = 20 + Math.floor(r() * 22) * 22, y = 20 + Math.floor(r() * 22) * 22;
            bar(x - 7, y - 7, 14, 14, "rgba(255,215,110,0.9)", "rgba(255,245,200,1)");
          }
        }
        c.strokeStyle = "rgba(255,255,255,0.55)"; c.lineWidth = 4; c.strokeRect(2, 2, W - 4, H - 4);
      });
    }
    const layerCol = [0x33e6a6, 0xff9a3c, 0x4ea2ff, 0x76ecff, 0xffd770];
    const N = 5, LS = 2.7;
    const layers = [];
    const stack = new T.Group();
    stack.position.set(0, 0.95, 0);
    g.add(stack);
    const slabGeo = kit.roundedBox(LS, 0.045, LS, 0.08, 0.012);
    const planeGeo = new T.PlaneGeometry(LS - 0.12, LS - 0.12);
    for (let i = 0; i < N; i++) {
      const L = new T.Group();
      const slab = new T.Mesh(slabGeo, kit.M.glass(0x9fd0ff, 0.2));
      L.add(slab);
      const m = new T.Mesh(planeGeo, new T.MeshBasicMaterial({ map: maskTex(i, 31 + i * 7), transparent: true, opacity: 0.92, depthWrite: false, toneMapped: false, side: T.DoubleSide }));
      m.rotation.x = -Math.PI / 2; m.position.y = 0.026;
      L.add(m);
      const edge = new T.LineSegments(new T.EdgesGeometry(slabGeo, 30), new T.LineBasicMaterial({ color: layerCol[i], transparent: true, opacity: 0.95 }));
      L.add(edge);
      L.rotation.y = 0;
      stack.add(L);
      layers.push(L);
    }
    // corner posts (thin gold guides)
    const posts = [];
    const postGeo = new T.CylinderGeometry(0.012, 0.012, 1, 8);
    const postMat = new T.MeshBasicMaterial({ color: GOLD, transparent: true, opacity: 0.8, toneMapped: false });
    [[-1, -1], [1, -1], [1, 1], [-1, 1]].forEach(function (s) {
      const p = new T.Mesh(postGeo, postMat);
      p.position.set(s[0] * (LS / 2 + 0.03), 0, s[1] * (LS / 2 + 0.03));
      stack.add(p); posts.push(p);
    });
    // glowing via beads travelling between layers
    const BEADS = 18;
    const beadGeo = new T.SphereGeometry(0.022, 8, 8);
    const beads = new T.InstancedMesh(beadGeo, new T.MeshBasicMaterial({ color: 0xfff0b0, toneMapped: false }), BEADS);
    stack.add(beads);
    const beadPos = [];
    const rr = kit.rng(5);
    for (let i = 0; i < BEADS; i++) beadPos.push([(rr() - 0.5) * (LS - 0.4), (rr() - 0.5) * (LS - 0.4), rr()]);
    const dum = new T.Object3D();

    /* ---------- waveform panel ---------- */
    const wp = new T.Group();
    wp.position.set(-3.15, 1.3, 0.0); wp.rotation.y = 0.45;
    g.add(wp);
    const pw = 1.7, ph = 1.2;
    const frame = new T.Mesh(kit.roundedBox(pw + 0.1, 0.045, ph + 0.1, 0.06, 0.012), kit.M.gold());
    frame.rotation.x = Math.PI / 2; frame.position.z = -0.03; wp.add(frame);
    const back = new T.Mesh(new T.PlaneGeometry(pw, ph), new T.MeshBasicMaterial({
      map: kit.canvasTex(512, 362, function (c, W, H) {
        c.fillStyle = "#08183f"; c.fillRect(0, 0, W, H);
        for (let x = 0; x <= W; x += 32) { c.strokeStyle = x % 128 ? "rgba(120,170,255,0.2)" : "rgba(160,210,255,0.5)"; c.beginPath(); c.moveTo(x, 0); c.lineTo(x, H); c.stroke(); }
        for (let y = 0; y <= H; y += 32) { c.strokeStyle = y % 128 ? "rgba(120,170,255,0.2)" : "rgba(160,210,255,0.5)"; c.beginPath(); c.moveTo(0, y); c.lineTo(W, y); c.stroke(); }
        c.fillStyle = "rgba(243,195,90,0.9)"; c.font = "bold 20px Inter, system-ui, sans-serif"; c.fillText("SIM  V(t)", 16, 28);
      }), toneMapped: false
    }));
    wp.add(back);
    const WP = 90;
    const waves = [];
    [[CYAN, 0.32], [GOLD, 0], [0xff6aa8, -0.32]].forEach(function (cfg, k) {
      [0, 0.012].forEach(function (off) {
        const geo = new T.BufferGeometry();
        const arr = new Float32Array(WP * 3);
        for (let i = 0; i < WP; i++) arr[i * 3] = (i / (WP - 1) - 0.5) * (pw - 0.18);
        geo.setAttribute("position", new T.BufferAttribute(arr, 3));
        const ln = new T.Line(geo, new T.LineBasicMaterial({ color: cfg[0], transparent: true, opacity: off ? 0.5 : 1, blending: T.AdditiveBlending, depthWrite: false, toneMapped: false }));
        ln.position.set(0, cfg[1] + off, 0.012);
        ln.frustumCulled = false;
        wp.add(ln); waves.push({ arr: arr, geo: geo, k: k });
      });
    });

    /* ---------- lightbulb (idea) ---------- */
    const bulb = new T.Group();
    bulb.position.set(2.95, 1.4, 0.1); g.add(bulb);
    const prof = [];
    for (let i = 0; i <= 20; i++) { const a = (i / 20) * Math.PI * 0.82; prof.push(new T.Vector2(Math.sin(a) * 0.4, -Math.cos(a) * 0.4 + 0.4)); }
    prof.unshift(new T.Vector2(0.17, -0.18)); prof.splice(1, 0, new T.Vector2(0.17, -0.1));
    const bulbGlass = new T.Mesh(new T.LatheGeometry(prof, 48), kit.M.glass(0xffeebb, 0.28));
    bulbGlass.position.y = 0.0; bulb.add(bulbGlass);
    const fil = new T.Mesh(new T.TorusGeometry(0.1, 0.012, 8, 32, Math.PI * 1.4), kit.M.glow(0xffd070, 3));
    fil.position.y = 0.38; fil.rotation.z = Math.PI * 0.8; bulb.add(fil);
    const stem = new T.Mesh(new T.CylinderGeometry(0.012, 0.012, 0.3, 8), kit.M.metal(0xdddddd));
    stem.position.y = 0.05; bulb.add(stem);
    for (let i = 0; i < 4; i++) {
      const r = new T.Mesh(new T.TorusGeometry(0.17, 0.028, 10, 32), kit.M.brushed(0xc9a24a, 0.3));
      r.rotation.x = Math.PI / 2; r.position.y = -0.22 - i * 0.065; bulb.add(r);
    }
    const tip = new T.Mesh(new T.SphereGeometry(0.07, 16, 12), kit.M.gold()); tip.position.y = -0.5; bulb.add(tip);
    const bglow = kit.glow(0xffd270, 2.0, 0.85); bglow.position.y = 0.38; bulb.add(bglow);
    const rays = [];
    for (let i = 0; i < 8; i++) {
      const a = i / 8 * Math.PI * 2;
      const rm = new T.Mesh(new T.BoxGeometry(0.16, 0.014, 0.014), kit.M.glow(0xffd270, 2.4));
      rm.position.set(Math.cos(a) * 0.68, 0.38 + Math.sin(a) * 0.68, 0);
      rm.rotation.z = a; bulb.add(rm); rays.push(rm);
    }

    /* ---------- drafting compass ---------- */
    const comp = new T.Group();
    comp.position.set(1.85, 0, 0.8); comp.rotation.y = -0.3; g.add(comp);
    function seg(a, b, ra, rb, mat, parent) {
      const len = a.distanceTo(b);
      const m = new T.Mesh(new T.CylinderGeometry(rb, ra, len, 14), mat);
      m.position.copy(a).add(b).multiplyScalar(0.5);
      m.quaternion.setFromUnitVectors(new T.Vector3(0, 1, 0), b.clone().sub(a).normalize());
      (parent || comp).add(m); return m;
    }
    const H = new T.Vector3(0, 1.25, 0), P1 = new T.Vector3(-0.3, 0.01, 0), P2 = new T.Vector3(0.3, 0.01, 0);
    const steel = kit.M.brushed(0xd5dae4, 0.28);
    seg(H, P1.clone().add(new T.Vector3(0, 0.12, 0)), 0.03, 0.026, steel);
    seg(P1.clone().add(new T.Vector3(0, 0.12, 0)), P1, 0.014, 0.002, kit.M.metal(0xeeeeee, 0.1));
    const dir2 = P2.clone().sub(H).normalize();
    seg(H, P2.clone().add(new T.Vector3(0, 0.2, 0)), 0.03, 0.026, steel);
    seg(P2.clone().add(new T.Vector3(0, 0.2, 0)), P2.clone().add(new T.Vector3(0, 0.1, 0)), 0.03, 0.03, kit.M.gold());
    seg(P2.clone().add(new T.Vector3(0, 0.1, 0)), P2, 0.026, 0.004, kit.M.plastic(0x3a2a1a, 0.6));
    const hinge = new T.Mesh(new T.SphereGeometry(0.06, 20, 16), kit.M.gold()); hinge.position.copy(H); comp.add(hinge);
    const knob = new T.Mesh(new T.CylinderGeometry(0.025, 0.035, 0.22, 20), kit.M.gold()); knob.position.set(0, H.y + 0.15, 0); comp.add(knob);
    const knobTop = new T.Mesh(new T.SphereGeometry(0.035, 12, 10), kit.M.gold()); knobTop.position.set(0, H.y + 0.26, 0); comp.add(knobTop);
    const brace = new T.Mesh(new T.TorusGeometry(0.19, 0.01, 8, 32, Math.PI), steel);
    brace.position.set(0, 0.72, 0); brace.rotation.z = Math.PI; comp.add(brace);
    const arc = new T.Mesh(new T.TorusGeometry(0.6, 0.008, 6, 64, Math.PI * 0.75), new T.MeshBasicMaterial({ color: GOLD, toneMapped: false, transparent: true, opacity: 0.9 }));
    arc.rotation.x = -Math.PI / 2; arc.position.copy(P1).setY(0.006); comp.add(arc);

    /* ---------- ruler ---------- */
    const rulerTex = kit.canvasTex(1024, 128, function (c, W, H) {
      c.fillStyle = "rgba(60,120,230,0.55)"; c.fillRect(0, 0, W, H);
      c.fillStyle = "rgba(255,255,255,0.9)"; c.strokeStyle = "rgba(255,255,255,0.9)";
      c.font = "bold 20px Inter, system-ui, sans-serif";
      for (let i = 0; i <= 40; i++) {
        const x = 12 + i * 24.6, l = i % 10 === 0 ? 46 : i % 5 === 0 ? 34 : 22;
        c.lineWidth = 2; c.beginPath(); c.moveTo(x, H); c.lineTo(x, H - l); c.stroke();
        if (i % 10 === 0 && i < 40) c.fillText(String(i / 10), x + 4, H - 52);
      }
    });
    const ruler = new T.Mesh(new T.BoxGeometry(2.6, 0.03, 0.34), [
      kit.M.glass(0x6aa8ff, 0.5), kit.M.glass(0x6aa8ff, 0.5),
      new T.MeshPhysicalMaterial({ map: rulerTex, transparent: true, roughness: 0.2, clearcoat: 1, envMapIntensity: 1.2 }),
      kit.M.glass(0x6aa8ff, 0.5), kit.M.glass(0x6aa8ff, 0.5), kit.M.glass(0x6aa8ff, 0.5)]);
    ruler.position.set(-0.7, 0.02, 1.3); ruler.rotation.y = 0.05; g.add(ruler);

    /* ---------- magnifier accent ---------- */
    const mag = new T.Group(); mag.position.set(-1.9, 0.07, 0.85); mag.rotation.set(0, 0.6, 0); g.add(mag);
    const lens = new T.Mesh(new T.CylinderGeometry(0.34, 0.34, 0.025, 48), kit.M.glass(0xbfe0ff, 0.25));
    lens.position.set(0, 0.08, 0); mag.add(lens);
    const rim = new T.Mesh(new T.TorusGeometry(0.345, 0.03, 12, 48), kit.M.gold());
    rim.rotation.x = Math.PI / 2; rim.position.y = 0.08; mag.add(rim);
    const handle = new T.Mesh(new T.CylinderGeometry(0.035, 0.04, 0.55, 14), kit.M.plastic(0x1c2a52, 0.35));
    handle.rotation.z = Math.PI / 2; handle.position.set(0.62, 0.08, 0); mag.add(handle);
    lens.rotation.z = 0;
    mag.rotation.z = 0;

    g.add(kit.contactShadow(3.6, 0.8)).position.y = -0.17;

    const spreadBase = 0.36;
    function update(t) {
      const breath = 0.5 + 0.5 * Math.sin(t * 0.7);
      const gap = spreadBase + 0.2 * breath;
      for (let i = 0; i < N; i++) layers[i].position.y = (i - (N - 1) / 2) * gap;
      const hh = gap * (N - 1);
      posts.forEach(function (p) { p.scale.y = hh; });
      stack.position.y = 1.45 + Math.sin(t * 0.9) * 0.03;
      stack.rotation.y = Math.sin(t * 0.35) * 0.12; stack.rotation.x = 0.26;
      for (let i = 0; i < BEADS; i++) {
        const b = beadPos[i], u = (t * 0.25 + b[2]) % 1;
        dum.position.set(b[0], (u - 0.5) * hh, b[1]);
        const s = Math.sin(u * 3.1416); dum.scale.setScalar(0.4 + s);
        dum.updateMatrix(); beads.setMatrixAt(i, dum.matrix);
      }
      beads.instanceMatrix.needsUpdate = true;
      beam.scale.set(1, 1 + 0.15 * breath, 1);
      for (let w = 0; w < waves.length; w++) {
        const W = waves[w], a = W.arr, k = W.k;
        for (let i = 0; i < WP; i++) {
          const x = i / (WP - 1), ph = x * 12 - t * 1.6;
          let y;
          if (k === 0) y = Math.sin(ph) * 0.2;
          else if (k === 1) y = Math.sin(ph * 0.7) * Math.exp(-((x * 3) % 1.5)) * 0.24;
          else y = Math.max(-0.17, Math.min(0.17, Math.sin(ph * 0.6) * 0.5));
          a[i * 3 + 1] = y;
        }
        W.geo.attributes.position.needsUpdate = true;
      }
      wp.position.y = 1.25 + Math.sin(t * 0.8 + 1) * 0.04;
      bulb.position.y = 1.35 + Math.sin(t * 1.1) * 0.06;
      const pulse = 0.85 + 0.15 * Math.sin(t * 2.4);
      bglow.material.opacity = 0.7 * pulse + 0.15; bglow.scale.setScalar(1.8 + 0.4 * pulse);
      rays.forEach(function (r, i) { r.scale.x = 0.8 + 0.5 * Math.sin(t * 2 + i); });
      mag.position.y = 0.07;
    }
    return { group: g, update: update, fit: 2.15, spin: 0.002 };
  };
})();
