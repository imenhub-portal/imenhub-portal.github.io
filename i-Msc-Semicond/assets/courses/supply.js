/* Semiconductor Supply Chain — luminous dot-matrix globe, glowing trade arcs, containers, wafer cassette & chip tray */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  window.IMEN.courseScenes = window.IMEN.courseScenes || {};

  IMEN.courseScenes.supply = function (kit) {
    const T = THREE;
    const g = new T.Group();
    const R = 1.45, GY = 0.95, GROUND = -1.15;

    /* ---------- continents (lon,lat polygons) ---------- */
    const LAND = [
      [[-168, 66], [-156, 71], [-125, 70], [-95, 72], [-82, 69], [-64, 60], [-56, 52], [-66, 44], [-76, 35], [-81, 25], [-90, 29], [-97, 26], [-97, 19], [-88, 16], [-83, 9], [-79, 8], [-86, 12], [-95, 16], [-105, 22], [-112, 30], [-117, 32], [-124, 40], [-124, 48], [-135, 58], [-150, 60], [-165, 60]],
      [[-78, 8], [-60, 10], [-50, 0], [-35, -6], [-40, -22], [-48, -28], [-58, -38], [-66, -46], [-70, -54], [-75, -48], [-72, -30], [-70, -18], [-81, -5], [-80, 3]],
      [[-55, 60], [-45, 60], [-20, 70], [-20, 82], [-60, 82], [-70, 76]],
      [[-10, 36], [-9, 43], [-2, 48], [-5, 50], [5, 52], [8, 56], [10, 58], [5, 62], [15, 69], [28, 71], [40, 68], [40, 56], [30, 45], [28, 41], [22, 37], [14, 40], [10, 44], [3, 43], [-6, 37]],
      [[-17, 21], [-10, 30], [-5, 36], [10, 37], [20, 32], [32, 31], [35, 28], [43, 12], [51, 12], [41, -2], [40, -15], [35, -24], [27, -34], [18, -34], [12, -18], [9, -1], [8, 4], [-8, 4], [-17, 14]],
      [[28, 41], [36, 37], [35, 30], [43, 15], [55, 17], [58, 23], [66, 25], [73, 20], [77, 8], [80, 15], [88, 22], [92, 20], [98, 10], [100, 3], [104, 1], [108, 12], [107, 21], [120, 22], [122, 31], [120, 38], [128, 36], [130, 42], [142, 50], [140, 58], [163, 62], [180, 67], [180, 70], [130, 72], [110, 77], [80, 73], [60, 70], [45, 68], [40, 56], [30, 45]],
      [[36, 30], [43, 30], [48, 29], [56, 26], [58, 23], [55, 17], [43, 13], [35, 28]],
      [[114, -22], [122, -18], [130, -12], [137, -12], [142, -11], [146, -19], [153, -26], [150, -37], [141, -38], [130, -32], [115, -34]],
      [[95, 5], [98, 4], [106, -6], [103, -5]], [[109, 1], [114, 6], [119, 5], [118, -2], [111, -3]],
      [[106, -6], [114, -7], [114, -8.5], [106, -7.5]], [[131, -1], [141, -3], [150, -10], [141, -9], [134, -4]],
      [[119, -1], [125, 1], [122, -5]], [[120, 18], [123, 18], [126, 7], [122, 8]],
      [[130, 32], [135, 34], [140, 36], [142, 40], [141, 45], [145, 44], [138, 38], [133, 33]],
      [[-6, 50], [2, 52], [-3, 58], [-6, 56]], [[44, -25], [48, -24], [50, -14], [47, -13]],
      [[172, -35], [178, -38], [174, -42], [168, -46], [172, -41]], [[-12, 52], [-6, 52], [-6, 55], [-10, 54]],
      [[-85, 22], [-74, 20], [-78, 22]], [[-24, 65], [-14, 65], [-18, 63]]
    ];
    function polyPath(c, poly, sx, sy) {
      c.beginPath();
      poly.forEach(function (p, i) {
        const x = (p[0] + 180) * sx, y = (90 - p[1]) * sy;
        if (i) c.lineTo(x, y); else c.moveTo(x, y);
      });
      c.closePath();
    }
    // 1° mask for dot sampling
    const mc = kit.canvas(360, 180);
    const mg = mc.getContext("2d");
    mg.fillStyle = "#000"; mg.fillRect(0, 0, 360, 180);
    mg.fillStyle = "#fff";
    LAND.forEach(function (p) { polyPath(mg, p, 1, 1); mg.fill(); });
    const mask = mg.getImageData(0, 0, 360, 180).data;
    function isLand(lat, lon) {
      const x = Math.min(359, Math.max(0, Math.floor(lon + 180))), y = Math.min(179, Math.max(0, Math.floor(90 - lat)));
      return mask[(y * 360 + x) * 4] > 128;
    }

    const TW = 1024, TH = 512;
    const colorTex = kit.canvasTex(TW, TH, function (c, W, H) {
      const gr = c.createLinearGradient(0, 0, 0, H);
      gr.addColorStop(0, "#0e2250"); gr.addColorStop(0.5, "#0a1a42"); gr.addColorStop(1, "#0e2250");
      c.fillStyle = gr; c.fillRect(0, 0, W, H);
      c.fillStyle = "#1b3c80";
      LAND.forEach(function (p) { polyPath(c, p, W / 360, H / 180); c.fill(); });
    });
    const emisTex = kit.canvasTex(TW, TH, function (c, W, H) {
      c.fillStyle = "#000"; c.fillRect(0, 0, W, H);
      // graticule
      c.strokeStyle = "rgba(70,130,230,0.35)"; c.lineWidth = 1;
      for (let lon = -180; lon <= 180; lon += 30) { const x = (lon + 180) / 360 * W; c.beginPath(); c.moveTo(x, 0); c.lineTo(x, H); c.stroke(); }
      for (let lat = -60; lat <= 60; lat += 30) { const y = (90 - lat) / 180 * H; c.beginPath(); c.moveTo(0, y); c.lineTo(W, y); c.stroke(); }
      const step = 2.3;
      for (let lat = -84; lat <= 84; lat += step) {
        const dl = step / Math.max(0.15, Math.cos(lat * Math.PI / 180));
        for (let lon = -180; lon < 180; lon += dl) {
          const x = (lon + 180) / 360 * W, y = (90 - lat) / 180 * H;
          if (isLand(lat, lon)) {
            c.fillStyle = "rgba(120,225,255,0.95)";
            c.beginPath(); c.arc(x, y, 2.1, 0, 6.283); c.fill();
          } else {
            c.fillStyle = "rgba(60,110,200,0.35)";
            c.beginPath(); c.arc(x, y, 0.7, 0, 6.283); c.fill();
          }
        }
      }
    });

    /* ---------- globe ---------- */
    const mount = new T.Group();                // axial tilt
    mount.position.y = GY;
    mount.rotation.z = -0.36;
    g.add(mount);
    const globe = new T.Group();                // spins
    mount.add(globe);
    const sphereMat = new T.MeshPhysicalMaterial({
      color: 0xffffff, map: colorTex, emissive: 0xffffff, emissiveMap: emisTex, emissiveIntensity: 1.7,
      metalness: 0.15, roughness: 0.32, clearcoat: 1, clearcoatRoughness: 0.12, envMapIntensity: 0.7
    });
    const sphere = new T.Mesh(new T.SphereGeometry(R, 96, 64), sphereMat);
    globe.add(sphere);

    const rimMat = new T.ShaderMaterial({
      transparent: true, depthWrite: false, blending: T.AdditiveBlending, side: T.FrontSide,
      uniforms: { c: { value: new T.Color(0x4aa8ff) } },
      vertexShader: "varying vec3 vN; varying vec3 vV; void main(){ vN = normalize(normalMatrix*normal); vec4 mv = modelViewMatrix*vec4(position,1.0); vV = normalize(-mv.xyz); gl_Position = projectionMatrix*mv; }",
      fragmentShader: "uniform vec3 c; varying vec3 vN; varying vec3 vV; void main(){ float f = pow(1.0 - max(dot(vN,vV),0.0), 2.6); gl_FragColor = vec4(c*f*0.8, f*0.6); }"
    });
    const rim = new T.Mesh(new T.SphereGeometry(R * 1.012, 64, 48), rimMat);
    rim.userData.noFit = true;
    mount.add(rim);
    const atmMat = new T.ShaderMaterial({
      transparent: true, depthWrite: false, blending: T.AdditiveBlending, side: T.BackSide,
      uniforms: { c: { value: new T.Color(0x3d8bff) } },
      vertexShader: "varying vec3 vN; void main(){ vN = normalize(normalMatrix*normal); gl_Position = projectionMatrix*modelViewMatrix*vec4(position,1.0); }",
      fragmentShader: "uniform vec3 c; varying vec3 vN; void main(){ float f = pow(max(0.72 - dot(vN, vec3(0.0,0.0,1.0)), 0.0), 3.2); gl_FragColor = vec4(c*f*0.8, f*0.5); }"
    });
    const atm = new T.Mesh(new T.SphereGeometry(R * 1.2, 64, 48), atmMat);
    atm.userData.noFit = true;
    mount.add(atm);

    // meridian mount (classic desk-globe arc) + pins at the poles
    const arc = new T.Mesh(new T.TorusGeometry(R * 1.1, 0.024, 16, 96, Math.PI), kit.M.brushed(0xc2cadb, 0.3));
    arc.rotation.z = Math.PI / 2;               // spans south -> north pole on the +x/-x side
    arc.rotation.y = -0.9;
    mount.add(arc);
    [-1, 1].forEach(function (s) {
      const cap = new T.Mesh(new T.CylinderGeometry(0.06, 0.07, 0.1, 20), kit.M.gold());
      cap.position.y = s * R * 1.09;
      mount.add(cap);
    });

    /* ---------- hubs, pins, arcs, pulses ---------- */
    const hubs = [
      [24.8, 121, 0], [33.4, -112, 0], [51, 13.7, 0], [37.3, 127.5, 0],     // fabs (gold)
      [5.4, 100.3, 1], [30.6, 104, 1],                                      // assembly & test (cyan)
      [37.4, -122, 2], [22.5, 114, 2], [42.3, -83, 2], [13, 77.6, 2], [-23.5, -46.6, 2] // customers (green)
    ];
    const HC = [0xffc24a, 0x38d9ff, 0x2ee69a];
    function ll(lat, lon, r) {
      const a = lat * Math.PI / 180, b = lon * Math.PI / 180;
      return new T.Vector3(Math.cos(a) * Math.cos(b) * r, Math.sin(a) * r, -Math.cos(a) * Math.sin(b) * r);
    }
    const NH = hubs.length;
    const up = new T.Vector3(0, 1, 0), zz = new T.Vector3(0, 0, 1);
    const dm = new T.Object3D();
    const stems = new T.InstancedMesh(new T.CylinderGeometry(0.006, 0.006, 0.1, 6), kit.M.metal(0xdbe6ff, 0.2), NH);
    const heads = new T.InstancedMesh(new T.SphereGeometry(0.03, 16, 12), new T.MeshBasicMaterial({ color: 0xffffff, toneMapped: false }), NH);
    const rings = new T.InstancedMesh(new T.RingGeometry(0.035, 0.048, 40), new T.MeshBasicMaterial({ color: 0xffffff, transparent: true, depthWrite: false, blending: T.AdditiveBlending, toneMapped: false, side: T.DoubleSide }), NH);
    const hubN = [], cc = new T.Color();
    hubs.forEach(function (h, i) {
      const n = ll(h[0], h[1], 1);
      hubN.push(n);
      dm.quaternion.setFromUnitVectors(up, n);
      dm.scale.set(1, 1, 1);
      dm.position.copy(n).multiplyScalar(R + 0.05);
      dm.updateMatrix(); stems.setMatrixAt(i, dm.matrix);
      dm.position.copy(n).multiplyScalar(R + 0.105);
      dm.updateMatrix(); heads.setMatrixAt(i, dm.matrix);
      heads.setColorAt(i, cc.setHex(HC[h[2]]));
      const gl = kit.glow(HC[h[2]], 0.34, 0.9);
      gl.position.copy(n).multiplyScalar(R + 0.105);
      globe.add(gl);
    });
    globe.add(stems); globe.add(heads); globe.add(rings);
    rings.userData.noFit = true;

    const routes = [[0, 4], [3, 4], [0, 6], [1, 6], [1, 8], [2, 8], [4, 7], [4, 9], [4, 6], [5, 7], [2, 9], [0, 10], [2, 4], [5, 8]];
    const NS = 56;
    const arcs = [];
    const _a = new T.Vector3(), _b = new T.Vector3();
    routes.forEach(function (rt, k) {
      const A = hubN[rt[0]], B = hubN[rt[1]];
      const ang = A.angleTo(B);
      const pts = [];
      const sa = Math.sin(ang);
      for (let i = 0; i <= NS; i++) {
        const s = i / NS;
        _a.copy(A).multiplyScalar(Math.sin((1 - s) * ang) / sa);
        _b.copy(B).multiplyScalar(Math.sin(s * ang) / sa);
        const p = _a.add(_b).clone().normalize().multiplyScalar(R + 0.1 + (0.06 + 0.3 * ang / Math.PI) * Math.sin(Math.PI * s) * R);
        pts.push(p);
      }
      const col = HC[hubs[rt[0]][2]];
      const curve = new T.CatmullRomCurve3(pts);
      const tube = new T.Mesh(new T.TubeGeometry(curve, 64, 0.0085, 6, false), kit.M.glow(col, 1.25));
      globe.add(tube);
      const halo = new T.Mesh(new T.TubeGeometry(curve, 48, 0.02, 6, false), kit.M.additive(col, 0.2));
      halo.userData.noFit = true;
      globe.add(halo);
      const flat = new Float32Array((NS + 1) * 3);
      pts.forEach(function (p, i) { flat[i * 3] = p.x; flat[i * 3 + 1] = p.y; flat[i * 3 + 2] = p.z; });
      arcs.push({ p: flat, col: col, off: (k * 0.37) % 1, speed: 0.07 + (k % 4) * 0.012 });
    });
    // chip glints travelling along the arcs (2 per route)
    const PP = routes.length * 2;
    const ppos = new Float32Array(PP * 3), pcol = new Float32Array(PP * 3);
    const pgeo = new T.BufferGeometry();
    pgeo.setAttribute("position", new T.BufferAttribute(ppos, 3));
    pgeo.setAttribute("color", new T.BufferAttribute(pcol, 3));
    const pts = new T.Points(pgeo, new T.PointsMaterial({
      size: 0.2, map: kit.tex.radial([[0, "rgba(255,255,255,1)"], [0.2, "rgba(255,255,255,0.55)"], [1, "rgba(255,255,255,0)"]]),
      vertexColors: true, transparent: true, depthWrite: false, blending: T.AdditiveBlending, sizeAttenuation: true
    }));
    pts.frustumCulled = false;
    pts.userData.noFit = true;
    globe.add(pts);
    const chipsI = new T.InstancedMesh(new T.BoxGeometry(0.045, 0.014, 0.045), new T.MeshBasicMaterial({ color: 0xffffff, toneMapped: false }), PP);
    chipsI.frustumCulled = false;
    chipsI.userData.noFit = true;
    for (let i = 0; i < PP; i++) { const a = arcs[i >> 1]; chipsI.setColorAt(i, cc.setHex(0xffffff).lerp(new T.Color(a.col), 0.25)); }
    globe.add(chipsI);
    const _q = new T.Vector3();

    /* ---------- ground platter ---------- */
    const platTex = kit.canvasTex(1024, 1024, function (c, W, H) {
      c.fillStyle = "#000"; c.fillRect(0, 0, W, H);
      c.translate(W / 2, H / 2);
      [[0.96, 3, 0.9], [0.9, 1.5, 0.5], [0.62, 2, 0.45], [0.36, 1.5, 0.35], [0.2, 1.5, 0.3]].forEach(function (r) {
        c.strokeStyle = "rgba(90,170,255," + r[2] + ")"; c.lineWidth = r[1];
        c.beginPath(); c.arc(0, 0, r[0] * W / 2, 0, 6.283); c.stroke();
      });
      for (let i = 0; i < 180; i++) {
        const a = i / 180 * 6.283, long = i % 5 === 0;
        c.strokeStyle = "rgba(120,190,255," + (long ? 0.8 : 0.4) + ")"; c.lineWidth = long ? 2.5 : 1.2;
        c.beginPath(); c.moveTo(Math.cos(a) * W * 0.455, Math.sin(a) * W * 0.455); c.lineTo(Math.cos(a) * W * (long ? 0.43 : 0.44), Math.sin(a) * W * (long ? 0.43 : 0.44)); c.stroke();
      }
      c.setLineDash([14, 14]); c.strokeStyle = "rgba(255,194,74,0.7)"; c.lineWidth = 2.5;
      c.beginPath(); c.arc(0, 0, 0.78 * W / 2, 0, 6.283); c.stroke();
    });
    const platter = new T.Mesh(new T.CylinderGeometry(2.7, 2.78, 0.12, 120),
      new T.MeshStandardMaterial({ color: 0x0c1430, metalness: 0.7, roughness: 0.32, emissive: 0xffffff, emissiveMap: platTex, emissiveIntensity: 0.8, envMapIntensity: 0.8 }));
    platter.position.y = GROUND - 0.06;
    g.add(platter);
    const edge = new T.Mesh(new T.TorusGeometry(2.74, 0.02, 10, 140), kit.M.glow(0x5ab0ff, 1.6));
    edge.rotation.x = Math.PI / 2;
    edge.position.y = GROUND + 0.005;
    g.add(edge);

    // pedestal
    const ped = new T.Mesh(new T.CylinderGeometry(0.22, 0.6, 0.42, 48), kit.M.metal(0x4a5877, 0.28));
    ped.position.y = GROUND + 0.21;
    g.add(ped);
    const pring = new T.Mesh(new T.TorusGeometry(0.22, 0.03, 12, 48), kit.M.gold());
    pring.rotation.x = Math.PI / 2;
    pring.position.y = GROUND + 0.43;
    g.add(pring);

    /* ---------- container stack ---------- */
    const ribTex = kit.canvasTex(256, 128, function (c, W, H) {
      c.fillStyle = "#cfcfcf"; c.fillRect(0, 0, W, H);
      for (let x = 0; x < W; x += 8) {
        const gr = c.createLinearGradient(x, 0, x + 8, 0);
        gr.addColorStop(0, "#fff"); gr.addColorStop(0.5, "#999"); gr.addColorStop(1, "#fff");
        c.fillStyle = gr; c.fillRect(x, 0, 8, H);
      }
      c.fillStyle = "rgba(255,255,255,0.0)";
    });
    const ribBump = kit.canvasTex(256, 128, function (c, W, H) {
      for (let x = 0; x < W; x += 8) {
        const gr = c.createLinearGradient(x, 0, x + 8, 0);
        gr.addColorStop(0, "#fff"); gr.addColorStop(0.5, "#222"); gr.addColorStop(1, "#fff");
        c.fillStyle = gr; c.fillRect(x, 0, 8, H);
      }
    }, { linear: true });
    const cols = [0x2f6bff, 0xff8a2a, 0x1fb58a, 0xd8403a, 0x8fa0c0, 0xe6c23a];
    const cmats = cols.map(function (c) {
      return new T.MeshStandardMaterial({ color: c, map: ribTex, bumpMap: ribBump, bumpScale: 1.2, roughness: 0.5, metalness: 0.35, envMapIntensity: 0.9 });
    });
    const CL = 0.6, CW = 0.24, CH = 0.24;
    const cbox = new T.BoxGeometry(CL, CH, CW);
    // rib axis: map u runs along length (x)
    const cstack = new T.Group();
    const layout = [[0, 0, 0], [0, 0, 1], [0, 0, 2], [1, 0, 0], [1, 0, 1], [1, 0, 2], [0, 1, 0], [0, 1, 1], [1, 1, 1], [0, 2, 1], [1, 1, 0], [1, 1, 2]];
    layout.forEach(function (p, i) {
      const m = new T.Mesh(cbox, cmats[(i * 5 + 1) % cmats.length]);
      m.position.set((p[0] - 0.5) * (CL + 0.012), CH / 2 + p[1] * (CH + 0.004), (p[2] - 1) * (CW + 0.012));
      cstack.add(m);
    });
    // corner castings on the top layer for detail
    const cast = new T.InstancedMesh(new T.BoxGeometry(0.035, 0.035, 0.035), kit.M.metal(0x2a3040, 0.5), 12);
    let ci = 0;
    [[0, 1, 0], [0, 1, 1], [1, 1, 1], [1, 1, 0]].forEach(function (p) {
      [[-1, -1], [1, 1]].forEach(function (s) {
        if (ci < 12) {
          dm.quaternion.identity(); dm.scale.set(1, 1, 1);
          dm.position.set((p[0] - 0.5) * (CL + 0.012) + s[0] * (CL / 2 - 0.01), CH + 1 * (CH + 0.004) + CH - 0.01, (p[2] - 1) * (CW + 0.012) + s[1] * (CW / 2 - 0.01));
          dm.updateMatrix(); cast.setMatrixAt(ci++, dm.matrix);
        }
      });
    });
    cast.count = ci;
    cstack.add(cast);
    cstack.position.set(-1.75, GROUND, 0.95);
    cstack.rotation.y = 0.5;
    g.add(cstack);

    /* ---------- JEDEC chip tray ---------- */
    const tray = new T.Group();
    const tbase = new T.Mesh(kit.roundedBox(1.0, 0.05, 0.66, 0.03, 0.012), kit.M.plastic(0x1a2030, 0.45));
    tbase.position.y = 0.025;
    tray.add(tbase);
    const NX = 6, NZ = 4;
    const pockets = kit.instanced(new T.BoxGeometry(0.12, 0.014, 0.12), kit.M.plastic(0x2b3550, 0.5), NX * NZ, function (i, d) {
      d.position.set(((i % NX) - (NX - 1) / 2) * 0.152, 0.056, (Math.floor(i / NX) - (NZ - 1) / 2) * 0.152);
    });
    tray.add(pockets);
    const chipI = kit.instanced(kit.roundedBox(0.092, 0.026, 0.092, 0.01, 0.006), kit.M.epoxy(0x101218), NX * NZ, function (i, d) {
      d.position.set(((i % NX) - (NX - 1) / 2) * 0.152, 0.07, (Math.floor(i / NX) - (NZ - 1) / 2) * 0.152);
      d.rotation.y = 0;
    });
    tray.add(chipI);
    // gold pad glints on chips
    const padI = kit.instanced(new T.BoxGeometry(0.02, 0.003, 0.02), kit.M.gold(), NX * NZ, function (i, d) {
      d.position.set(((i % NX) - (NX - 1) / 2) * 0.152 - 0.03, 0.0845, (Math.floor(i / NX) - (NZ - 1) / 2) * 0.152 - 0.03);
    });
    tray.add(padI);
    tray.position.set(1.75, GROUND, 0.85);
    tray.rotation.y = -0.45;
    g.add(tray);

    /* ---------- wafer cassette ---------- */
    const cas = new T.Group();
    const cbase = new T.Mesh(kit.roundedBox(0.5, 0.04, 0.5, 0.05, 0.01), kit.M.metal(0x7a88a8, 0.25));
    cbase.position.y = 0.02;
    cas.add(cbase);
    const cglass = new T.Mesh(new T.BoxGeometry(0.44, 0.5, 0.44), kit.M.glass(0x9fd0ff, 0.14));
    cglass.position.y = 0.3;
    cglass.renderOrder = 5;
    cas.add(cglass);
    [[-1, -1], [1, -1], [-1, 1], [1, 1]].forEach(function (s) {
      const r = new T.Mesh(new T.BoxGeometry(0.02, 0.5, 0.02), kit.M.metal(0xd0d8ea, 0.2));
      r.position.set(s[0] * 0.22, 0.3, s[1] * 0.22);
      cas.add(r);
    });
    const topc = new T.Mesh(kit.roundedBox(0.5, 0.03, 0.5, 0.05, 0.008), kit.M.metal(0x7a88a8, 0.25));
    topc.position.y = 0.565;
    cas.add(topc);
    const wafG = new T.CylinderGeometry(0.175, 0.175, 0.008, 48);
    const wafers = kit.instanced(wafG, new T.MeshPhysicalMaterial({ color: 0x8a9cd0, metalness: 1, roughness: 0.12, clearcoat: 0.5, envMapIntensity: 1.5 }), 12, function (i, d) {
      d.position.y = 0.09 + i * 0.037;
    });
    cas.add(wafers);
    cas.position.set(0.9, GROUND, 1.6);
    cas.rotation.y = 0.35;
    g.add(cas);

    /* ---------- orbit ring + cargo plane ---------- */
    const orb = new T.Group();
    orb.position.y = GY;
    orb.rotation.set(0.62, 0, -0.3);
    g.add(orb);
    const OR = 2.35;
    const orbRing = new T.Mesh(new T.TorusGeometry(OR, 0.007, 8, 160), kit.M.additive(0x7fc0ff, 0.55));
    orbRing.rotation.x = Math.PI / 2;
    orb.add(orbRing);
    const plane = new T.Group();
    const white = kit.M.plastic(0xf4f7ff, 0.25);
    const fus = new T.Mesh(new T.CylinderGeometry(0.02, 0.02, 0.16, 12), white);
    fus.rotation.z = Math.PI / 2;
    plane.add(fus);
    const nose = new T.Mesh(new T.ConeGeometry(0.02, 0.05, 12), white);
    nose.rotation.z = -Math.PI / 2; nose.position.x = 0.105;
    plane.add(nose);
    const tailc = new T.Mesh(new T.ConeGeometry(0.02, 0.06, 12), white);
    tailc.rotation.z = Math.PI / 2; tailc.position.x = -0.11;
    plane.add(tailc);
    const wing = new T.Mesh(new T.BoxGeometry(0.06, 0.006, 0.26), kit.M.plastic(0x2f6bff, 0.3));
    wing.position.x = 0.01;
    plane.add(wing);
    const tailw = new T.Mesh(new T.BoxGeometry(0.03, 0.005, 0.09), kit.M.plastic(0x2f6bff, 0.3));
    tailw.position.x = -0.1;
    plane.add(tailw);
    const fin = new T.Mesh(new T.BoxGeometry(0.04, 0.05, 0.005), kit.M.plastic(0x2f6bff, 0.3));
    fin.position.set(-0.1, 0.03, 0);
    plane.add(fin);
    const pg = kit.glow(0xffd98a, 0.14, 1);
    pg.position.x = -0.14;
    plane.add(pg);
    plane.scale.setScalar(1.35);
    orb.add(plane);

    const sh = kit.contactShadow(3.0, 0.7);
    sh.position.y = GROUND - 0.125;
    g.add(sh);

    /* ---------- animation ---------- */
    const GLOBE0 = -Math.PI / 2 - (100 * Math.PI) / 180;
    return {
      group: g,
      fit: 2.2,
      spin: 0.0012,
      update: function (t) {
        globe.rotation.y = GLOBE0 + t * 0.1;
        // glints
        for (let k = 0; k < arcs.length; k++) {
          const a = arcs[k];
          for (let j = 0; j < 2; j++) {
            const i = k * 2 + j;
            const u = ((t * a.speed + a.off + j * 0.5) % 1) * NS;
            const i0 = Math.floor(u), f = u - i0, i1 = Math.min(NS, i0 + 1);
            const x = a.p[i0 * 3] + (a.p[i1 * 3] - a.p[i0 * 3]) * f;
            const y = a.p[i0 * 3 + 1] + (a.p[i1 * 3 + 1] - a.p[i0 * 3 + 1]) * f;
            const z = a.p[i0 * 3 + 2] + (a.p[i1 * 3 + 2] - a.p[i0 * 3 + 2]) * f;
            ppos[i * 3] = x; ppos[i * 3 + 1] = y; ppos[i * 3 + 2] = z;
            const fade = Math.sin(Math.PI * (u / NS));
            cc.setHex(a.col);
            pcol[i * 3] = cc.r * fade; pcol[i * 3 + 1] = cc.g * fade; pcol[i * 3 + 2] = cc.b * fade;
            dm.position.set(x, y, z);
            _q.set(a.p[i1 * 3] - a.p[i0 * 3], a.p[i1 * 3 + 1] - a.p[i0 * 3 + 1], a.p[i1 * 3 + 2] - a.p[i0 * 3 + 2]);
            dm.quaternion.identity();
            dm.lookAt(dm.position.x + _q.x, dm.position.y + _q.y, dm.position.z + _q.z);
            dm.scale.setScalar(0.4 + 0.9 * fade);
            dm.updateMatrix();
            chipsI.setMatrixAt(i, dm.matrix);
          }
        }
        pgeo.attributes.position.needsUpdate = true;
        pgeo.attributes.color.needsUpdate = true;
        chipsI.instanceMatrix.needsUpdate = true;
        // hub pulses
        for (let i = 0; i < NH; i++) {
          const ph = (t * 0.45 + i * 0.173) % 1;
          dm.quaternion.setFromUnitVectors(zz, hubN[i]);
          dm.position.copy(hubN[i]).multiplyScalar(R + 0.006);
          dm.scale.setScalar(0.6 + ph * 2.6);
          dm.updateMatrix();
          rings.setMatrixAt(i, dm.matrix);
          rings.setColorAt(i, cc.setHex(HC[hubs[i][2]]).multiplyScalar(1 - ph));
        }
        rings.instanceMatrix.needsUpdate = true;
        rings.instanceColor.needsUpdate = true;
        // plane
        const a = t * 0.22;
        plane.position.set(Math.cos(a) * OR, 0, Math.sin(a) * OR);
        plane.rotation.y = -a - Math.PI / 2;
        plane.rotation.x = 0.12;
      }
    };
  };
})();
