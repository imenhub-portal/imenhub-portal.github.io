/* Research Methodology — response-surface figure, rendered in 3D (DoE, optimum, data points) */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  window.IMEN.courseScenes = window.IMEN.courseScenes || {};

  IMEN.courseScenes.research = function (kit) {
    const T = THREE;
    const g = new T.Group();
    const R = kit.rng(21);

    /* ---------- base: tray + glass plate with measurement grid ---------- */
    const PW = 4.5, PD = 3.2, PY = 0.06;
    const tray = new T.Mesh(kit.roundedBox(PW + 0.28, 0.16, PD + 0.28, 0.12, 0.03), kit.M.metal(0x39435e, 0.32));
    tray.position.y = -0.17;
    g.add(tray);
    const plate = new T.Mesh(kit.roundedBox(PW, 0.1, PD, 0.1, 0.018), kit.M.glass(0x2a4a8a, 0.5));
    plate.position.y = 0;
    plate.renderOrder = 1;
    g.add(plate);
    // etched measurement grid (fine + coarse)
    const gridTex = kit.canvasTex(1024, 768, function (c, W, H) {
      c.clearRect(0, 0, W, H);
      const sx = W / 24, sy = H / 18;
      for (let i = 0; i <= 24; i++) {
        c.strokeStyle = i % 4 === 0 ? "rgba(140,200,255,0.85)" : "rgba(120,170,240,0.35)";
        c.lineWidth = i % 4 === 0 ? 2.4 : 1.2;
        c.beginPath(); c.moveTo(i * sx, 0); c.lineTo(i * sx, H); c.stroke();
      }
      for (let j = 0; j <= 18; j++) {
        c.strokeStyle = j % 3 === 0 ? "rgba(140,200,255,0.85)" : "rgba(120,170,240,0.35)";
        c.lineWidth = j % 3 === 0 ? 2.4 : 1.2;
        c.beginPath(); c.moveTo(0, j * sy); c.lineTo(W, j * sy); c.stroke();
      }
      c.strokeStyle = "rgba(190,225,255,0.9)";
      c.lineWidth = 6;
      c.strokeRect(3, 3, W - 6, H - 6);
    });
    const gridPlane = new T.Mesh(new T.PlaneGeometry(PW - 0.1, PD - 0.1), new T.MeshBasicMaterial({ map: gridTex, transparent: true, opacity: 0.8, depthWrite: false, toneMapped: false }));
    gridPlane.rotation.x = -Math.PI / 2;
    gridPlane.position.y = PY;
    gridPlane.renderOrder = 1;
    g.add(gridPlane);

    function flatLabel(str, h, x, z, rot, color) {
      const l = kit.label(str, h, { color: color || "#bcd8ff", size: 96 });
      l.rotation.x = -Math.PI / 2;
      l.rotation.z = rot || 0;
      l.position.set(x, PY + 0.004, z);
      g.add(l);
      return l;
    }
    flatLabel("TEMPERATURE  →", 0.15, 0, PD / 2 - 0.2, 0);
    flatLabel("TIME  →", 0.15, -PW / 2 + 0.2, 0, Math.PI / 2);

    /* ---------- response surface ---------- */
    const SX = 1.5, SZ = 1.2, N = 64;
    const geo = new T.PlaneGeometry(SX * 2, SZ * 2, N, N);
    geo.rotateX(-Math.PI / 2);
    const pos = geo.attributes.position;
    const NV = pos.count;
    const col = new Float32Array(NV * 3);
    geo.setAttribute("color", new T.BufferAttribute(col, 3));
    const U = new Float32Array(NV), V = new Float32Array(NV);
    for (let i = 0; i < NV; i++) { U[i] = pos.getX(i) / SX; V[i] = pos.getZ(i) / SZ; }

    const Y0 = 0.5, YS = 0.95;
    function gau(a, b, s) { return Math.exp(-(a * a + b * b) / (2 * s * s)); }
    function H(u, v, t) {
      const br = Math.sin(t * 0.55);
      return 0.95 * gau(u - 0.35, v + 0.15, 0.38 + 0.02 * br) +
        (0.52 + 0.14 * Math.sin(t * 0.7 + 1)) * gau(u + 0.5, v - 0.45, 0.3) -
        (0.34 + 0.06 * br) * gau(u + 0.45, v + 0.5, 0.28) +
        0.045 * Math.sin(5 * u + t * 0.8) * Math.cos(4 * v - t * 0.5) + 0.06 * u * 0.5;
    }
    function worldY(h) { return Y0 + (h + 0.4) * YS; }
    // heat gradient (deep blue -> cyan -> green -> yellow -> red)
    const stops = [[0, 0.0, 0.02, 0.3], [0.22, 0.0, 0.14, 0.75], [0.42, 0.0, 0.55, 0.5], [0.6, 0.25, 0.7, 0.04], [0.8, 0.85, 0.42, 0.0], [1, 0.8, 0.04, 0.02]];
    function heat(f, out, o) {
      f = f < 0 ? 0 : f > 1 ? 1 : f;
      let i = 0;
      while (i < stops.length - 2 && f > stops[i + 1][0]) i++;
      const a = stops[i], b = stops[i + 1], k = (f - a[0]) / (b[0] - a[0]);
      out[o] = a[1] + (b[1] - a[1]) * k;
      out[o + 1] = a[2] + (b[2] - a[2]) * k;
      out[o + 2] = a[3] + (b[3] - a[3]) * k;
    }
    function updateSurface(t) {
      for (let i = 0; i < NV; i++) {
        const h = H(U[i], V[i], t);
        pos.setY(i, worldY(h));
        heat((h + 0.25) / 1.25, col, i * 3);
      }
      pos.needsUpdate = true;
      geo.attributes.color.needsUpdate = true;
      geo.computeVertexNormals();
    }
    updateSurface(0);
    const surfTex = kit.canvasTex(512, 512, function (c, W, H2) {
      c.fillStyle = "#ffffff"; c.fillRect(0, 0, W, H2);
      c.strokeStyle = "rgba(10,20,50,0.55)";
      c.lineWidth = 3;
      for (let i = 0; i <= 16; i++) {
        const p = (i / 16) * W;
        c.beginPath(); c.moveTo(p, 0); c.lineTo(p, H2); c.stroke();
        c.beginPath(); c.moveTo(0, p); c.lineTo(W, p); c.stroke();
      }
    });
    surfTex.generateMipmaps = true;
    const surfMat = new T.MeshPhysicalMaterial({
      vertexColors: true, map: surfTex, metalness: 0.0, roughness: 0.45, clearcoat: 0.6, clearcoatRoughness: 0.12,
      side: T.DoubleSide, transparent: true, opacity: 0.95, envMapIntensity: 0.35
    });
    surfMat.onBeforeCompile = function (sh) {
      sh.vertexShader = sh.vertexShader.replace("#include <common>", "#include <common>\nvarying float vH;")
        .replace("#include <begin_vertex>", "#include <begin_vertex>\nvH = position.y;");
      sh.fragmentShader = sh.fragmentShader.replace("#include <common>", "#include <common>\nvarying float vH;")
        .replace("#include <color_fragment>", "#include <color_fragment>\n" +
          "float q = vH * 11.0; float d = abs(fract(q) - 0.5); float ln = smoothstep(0.455, 0.5, d);\n" +
          "diffuseColor.rgb = mix(diffuseColor.rgb, vec3(0.02, 0.04, 0.1), ln * 0.7);");
    };
    const surface = new T.Mesh(geo, surfMat);
    surface.renderOrder = 2;
    g.add(surface);

    // plot volume (bounding frame)
    const boxH = 1.85;
    const frame = new T.LineSegments(new T.EdgesGeometry(new T.BoxGeometry(SX * 2 + 0.1, boxH, SZ * 2 + 0.1)),
      new T.LineBasicMaterial({ color: 0x7fb0ff, transparent: true, opacity: 0.4 }));
    frame.position.y = PY + boxH / 2;
    g.add(frame);
    // vertical axis ticks on back-left corner post
    const post = new T.Mesh(new T.CylinderGeometry(0.014, 0.014, boxH, 10), kit.M.metal(0xbcd0f5, 0.2));
    post.position.set(-SX - 0.05, PY + boxH / 2, -SZ - 0.05);
    g.add(post);
    const tickG = new T.BoxGeometry(0.12, 0.012, 0.012);
    for (let i = 0; i <= 8; i++) {
      const tk = new T.Mesh(tickG, kit.M.metal(0xbcd0f5, 0.2));
      tk.position.set(-SX - 0.05 + 0.06, PY + (i / 8) * boxH, -SZ - 0.05);
      g.add(tk);
    }
    const yl = kit.label("YIELD  →", 0.15, { color: "#ffd98a", size: 96 });
    yl.rotation.z = Math.PI / 2;
    yl.position.set(-SX - 0.3, PY + 1.0, -SZ - 0.05);
    g.add(yl);

    /* ---------- experimental data points + drop lines ---------- */
    const NP = 34;
    const pu = [], pv = [], pr = [];
    for (let i = 0; i < NP; i++) {
      pu.push((R() * 2 - 1) * 0.92);
      pv.push((R() * 2 - 1) * 0.92);
      const above = R() < 0.62;
      pr.push((above ? 1 : -1) * (0.05 + R() * 0.2));
    }
    const dotMat = new T.MeshPhysicalMaterial({ color: 0xffffff, metalness: 0.0, roughness: 0.12, clearcoat: 1, clearcoatRoughness: 0.05, envMapIntensity: 1.2 });
    const dots = new T.InstancedMesh(new T.SphereGeometry(0.05, 18, 14), dotMat, NP);
    const cA = new T.Color(0xffc24a), cB = new T.Color(0x6fd8ff);
    for (let i = 0; i < NP; i++) dots.setColorAt(i, pr[i] > 0 ? cA : cB);
    dots.instanceColor.needsUpdate = true;
    g.add(dots);
    const dropPos = new Float32Array(NP * 6);
    const dropGeo = new T.BufferGeometry();
    dropGeo.setAttribute("position", new T.BufferAttribute(dropPos, 3));
    const drops = new T.LineSegments(dropGeo, new T.LineBasicMaterial({ color: 0xaecbff, transparent: true, opacity: 0.55 }));
    drops.frustumCulled = false;
    g.add(drops);
    const feet = kit.instanced(new T.CircleGeometry(0.045, 16), kit.M.additive(0x7fc4ff, 0.8), NP, function (i, d) {
      d.position.set(pu[i] * SX, PY + 0.003, pv[i] * SZ);
      d.rotation.x = -Math.PI / 2;
    });
    feet.userData.noFit = true;
    g.add(feet);
    const dm = new T.Object3D();
    function updatePoints(t) {
      for (let i = 0; i < NP; i++) {
        const y = worldY(H(pu[i], pv[i], t)) + pr[i];
        dm.position.set(pu[i] * SX, y, pv[i] * SZ);
        dm.updateMatrix();
        dots.setMatrixAt(i, dm.matrix);
        dropPos[i * 6] = pu[i] * SX; dropPos[i * 6 + 1] = y; dropPos[i * 6 + 2] = pv[i] * SZ;
        dropPos[i * 6 + 3] = pu[i] * SX; dropPos[i * 6 + 4] = PY + 0.004; dropPos[i * 6 + 5] = pv[i] * SZ;
      }
      dots.instanceMatrix.needsUpdate = true;
      dropGeo.attributes.position.needsUpdate = true;
    }
    updatePoints(0);

    /* ---------- optimum beacon ---------- */
    const OU = 0.35, OV = -0.15;
    const opt = new T.Group();
    opt.position.x = OU * SX; opt.position.z = OV * SZ;
    const ring = new T.Mesh(new T.TorusGeometry(0.17, 0.018, 14, 56), new T.MeshStandardMaterial({ color: 0x331a00, emissive: 0xffc23a, emissiveIntensity: 2.6, roughness: 0.4 }));
    ring.rotation.x = Math.PI / 2;
    opt.add(ring);
    const pulse = new T.Mesh(new T.RingGeometry(0.17, 0.2, 56), kit.M.additive(0xffd36a, 0.7));
    pulse.rotation.x = -Math.PI / 2;
    pulse.userData.noFit = true;
    opt.add(pulse);
    const beam = kit.beam(0.06, 0.012, 0.95, "255,210,110", 0.55);
    beam.position.y = 0.0;
    beam.scale.y = 1;
    beam.geometry.translate(0, 0.0, 0);
    opt.add(beam);
    const star = new T.Mesh(new T.OctahedronGeometry(0.075, 0), kit.M.gold());
    star.scale.y = 1.5;
    opt.add(star);
    const og = kit.glow(0xffc24a, 0.9, 0.9);
    opt.add(og);
    g.add(opt);

    /* ---------- 3-factor DoE lattice ---------- */
    const doe = new T.Group();
    const CS = 0.62;
    doe.position.set(PW / 2 - 0.4, 0.0, 0.85);
    const cube = new T.Group();
    cube.position.y = 0.85;
    doe.add(cube);
    cube.add(new T.LineSegments(new T.EdgesGeometry(new T.BoxGeometry(CS, CS, CS)), new T.LineBasicMaterial({ color: 0x9cc4ff, transparent: true, opacity: 0.8 })));
    const faces = new T.Mesh(new T.BoxGeometry(CS, CS, CS), kit.M.glass(0x7fb0ff, 0.1));
    cube.add(faces);
    const lat = new T.InstancedMesh(new T.SphereGeometry(0.036, 14, 10), new T.MeshPhysicalMaterial({ color: 0xffffff, roughness: 0.15, clearcoat: 1, metalness: 0.2 }), 27);
    let li = 0;
    const cc = new T.Color();
    for (let a = -1; a <= 1; a++) for (let b = -1; b <= 1; b++) for (let c = -1; c <= 1; c++) {
      const nz = Math.abs(a) + Math.abs(b) + Math.abs(c);
      dm.position.set(a * CS / 2, b * CS / 2, c * CS / 2);
      dm.scale.setScalar(nz === 3 ? 1.25 : 1);
      dm.updateMatrix();
      lat.setMatrixAt(li, dm.matrix);
      cc.setHex(nz === 3 ? 0xffc24a : nz === 0 ? 0xff5a5a : nz === 2 ? 0x56a8ff : 0x6fe0d0);
      lat.setColorAt(li, cc);
      li++;
    }
    dm.scale.setScalar(1);
    cube.add(lat);
    const cglow = kit.glow(0x6aa8ff, 1.6, 0.35);
    cube.add(cglow);
    const hot = kit.glow(0xffe08a, 0.28, 1);
    cube.add(hot);
    const pillar = new T.Mesh(new T.CylinderGeometry(0.035, 0.05, 0.5, 20), kit.M.metal(0x8896b4, 0.3));
    pillar.position.y = 0.3;
    doe.add(pillar);
    const foot = new T.Mesh(new T.CylinderGeometry(0.28, 0.3, 0.04, 36), kit.M.metal(0x9aa8c8, 0.25));
    foot.position.y = PY + 0.02;
    doe.add(foot);
    g.add(doe);
    const cl = kit.label("DoE  3³", 0.12, { color: "#cfe2ff", size: 96 });
    cl.position.set(doe.position.x, 1.42, doe.position.z);
    g.add(cl);

    /* ---------- magnifying lens accent (front-left) ---------- */
    const lens = new T.Group();
    const rim = new T.Mesh(new T.TorusGeometry(0.3, 0.035, 18, 56), kit.M.gold());
    lens.add(rim);
    const lg = new T.Mesh(new T.CircleGeometry(0.29, 48), kit.M.glass(0xcfe9ff, 0.2));
    lens.add(lg);
    const handle = new T.Mesh(new T.CylinderGeometry(0.035, 0.045, 0.62, 16), kit.M.plastic(0x1b2340, 0.35));
    handle.rotation.z = Math.PI / 4;
    handle.position.set(0.3, -0.3, 0);
    lens.add(handle);
    lens.position.set(-PW / 2 + 0.75, 0.42, PD / 2 - 0.35);
    lens.rotation.set(-1.15, 0.3, 0.1);
    g.add(lens);

    /* ---------- grounding ---------- */
    const sh = kit.contactShadow(3.6, 0.9);
    sh.scale.set(1.15, 0.85, 1);
    sh.position.y = -0.27;
    g.add(sh);

    /* ---------- animation ---------- */
    let last = -1;
    return {
      group: g,
      fit: 2.1,
      spin: 0.0022,
      update: function (t) {
        const step = Math.floor(t * 30);
        if (step !== last) { // ~30 Hz surface morph
          last = step;
          updateSurface(t);
          updatePoints(t);
          const hp = worldY(H(OU, OV, t));
          opt.position.y = hp + 0.02;
        }
        const p = (t * 0.55) % 1;
        pulse.scale.setScalar(1 + p * 2.6);
        pulse.material.opacity = 0.7 * (1 - p);
        ring.rotation.z = t * 0.5;
        star.rotation.y = t * 1.3;
        star.position.y = 0.42 + Math.sin(t * 2) * 0.03;
        og.position.y = 0.42;
        og.scale.setScalar(0.8 + 0.12 * Math.sin(t * 3));
        beam.position.y = 0.08;
        cube.rotation.y = t * 0.35;
        cube.rotation.x = 0.35 + Math.sin(t * 0.4) * 0.1;
        const k = Math.floor(t * 2.2) % 27;
        const a = Math.floor(k / 9) - 1, b = Math.floor((k % 9) / 3) - 1, c = (k % 3) - 1;
        hot.position.set(a * CS / 2, b * CS / 2, c * CS / 2);
        lens.position.y = 0.42 + Math.sin(t * 0.8) * 0.04;
      }
    };
  };
})();
