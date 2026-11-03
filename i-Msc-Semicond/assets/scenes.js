/* ==========================================================
   IMEN · M.Sc. Semiconductor Engineering
   three.js scenes — pinned to r128 (cdnjs) per house guidelines
   ========================================================== */
(function () {
  "use strict";

  const REDUCED = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const C = {
    gold: 0xd99a00,
    cyan: 0x0ea5e9,
    blue: 0x2563eb,
    steel: 0x93b4e8,
    deep: 0x0b3a8f,
    white: 0xffffff,
    orange: 0xf59e0b,
    red: 0xdc2626
  };

  const mat = (color, o) => {
    o = o || {};
    return new THREE.MeshStandardMaterial({
      color: color,
      metalness: o.metalness == null ? 0.55 : o.metalness,
      roughness: o.roughness == null ? 0.3 : o.roughness,
      emissive: o.emissive == null ? 0x000000 : o.emissive,
      emissiveIntensity: o.emissiveIntensity == null ? 1 : o.emissiveIntensity,
      transparent: !!o.transparent,
      opacity: o.opacity == null ? 1 : o.opacity,
      wireframe: !!o.wireframe
    });
  };

  const lineMat = (color, opacity) =>
    new THREE.LineBasicMaterial({
      color: color,
      transparent: true,
      opacity: opacity == null ? 1 : opacity
    });

  function lines(points, color, opacity) {
    const geo = new THREE.BufferGeometry().setFromPoints(points);
    return new THREE.LineSegments(geo, lineMat(color, opacity));
  }

  function addLights(scene) {
    scene.add(new THREE.AmbientLight(0xffffff, 0.55));
    scene.add(new THREE.HemisphereLight(0xffffff, 0xdbe8ff, 0.85));
    const key = new THREE.DirectionalLight(0xffffff, 0.9);
    key.position.set(4, 6, 4);
    scene.add(key);
    const gold = new THREE.PointLight(C.gold, 0.45, 22);
    gold.position.set(-3.2, 1.6, 2.4);
    scene.add(gold);
    const cyan = new THREE.PointLight(C.cyan, 0.4, 22);
    cyan.position.set(3.2, -1.2, 3);
    scene.add(cyan);
  }

  function sizeOf(renderer, camera, canvas) {
    const w = Math.max(canvas.clientWidth || canvas.parentElement.clientWidth || 1, 1);
    const h = Math.max(canvas.clientHeight || canvas.parentElement.clientHeight || 1, 1);
    renderer.setSize(w, h, false);
    camera.aspect = w / h;
    camera.updateProjectionMatrix();
  }

  /* =========================================================
     ORBIT — pointer / touch drag with inertia, pinch + wheel zoom
     ========================================================= */
  function attachOrbit(canvas, target, opts) {
    opts = opts || {};
    const baseX = target.rotation.x;
    const state = {
      rx: target.rotation.x,
      ry: target.rotation.y,
      vx: 0,
      vy: 0,
      dragging: false,
      zoom: 1,
      pointers: new Map(),
      pinch: 0
    };
    canvas.style.cursor = "grab";
    // pan-y lets a vertical swipe scroll the page when a finger starts on the illustration (phones)
    canvas.style.touchAction = opts.touchAction || "none";

    function down(e) {
      state.pointers.set(e.pointerId, [e.clientX, e.clientY]);
      if (state.pointers.size === 1) {
        state.dragging = true;
        state.vx = 0;
        state.vy = 0;
        canvas.style.cursor = "grabbing";
      } else if (state.pointers.size === 2) {
        const pts = Array.from(state.pointers.values());
        state.pinch = Math.hypot(pts[0][0] - pts[1][0], pts[0][1] - pts[1][1]);
      }
      if (canvas.setPointerCapture) {
        try { canvas.setPointerCapture(e.pointerId); } catch (err) {}
      }
    }
    function move(e) {
      if (!state.pointers.has(e.pointerId)) return;
      const prev = state.pointers.get(e.pointerId);
      const dx = e.clientX - prev[0];
      const dy = e.clientY - prev[1];
      state.pointers.set(e.pointerId, [e.clientX, e.clientY]);
      if (state.pointers.size === 1 && state.dragging) {
        const k = 0.0062;
        state.ry += dx * k;
        state.rx += dy * k;
        state.vy = dx * k;
        state.vx = dy * k;
      } else if (state.pointers.size === 2) {
        const pts = Array.from(state.pointers.values());
        const d = Math.hypot(pts[0][0] - pts[1][0], pts[0][1] - pts[1][1]);
        if (state.pinch > 0 && d > 0) {
          state.zoom *= state.pinch / d;
          state.zoom = Math.min(1.6, Math.max(0.55, state.zoom));
          if (opts.onZoom) opts.onZoom(state.zoom);
        }
        state.pinch = d;
      }
    }
    function up(e) {
      state.pointers.delete(e.pointerId);
      if (state.pointers.size === 0) {
        state.dragging = false;
        canvas.style.cursor = "grab";
      }
      state.pinch = 0;
    }
    function wheel(e) {
      e.preventDefault();
      state.zoom *= e.deltaY > 0 ? 1.07 : 0.935;
      state.zoom = Math.min(1.6, Math.max(0.55, state.zoom));
      if (opts.onZoom) opts.onZoom(state.zoom);
    }
    canvas.addEventListener("pointerdown", down);
    canvas.addEventListener("pointermove", move);
    canvas.addEventListener("pointerup", up);
    canvas.addEventListener("pointercancel", up);
    canvas.addEventListener("lostpointercapture", up);
    canvas.addEventListener("wheel", wheel, { passive: false });

    return {
      step: function (idleSpin) {
        if (!state.dragging) {
          state.ry += state.vy;
          state.rx += state.vx;
          state.vx *= 0.93;
          state.vy *= 0.93;
          if (idleSpin) {
            if (!(opts.idlePause && opts.idlePause())) state.ry += idleSpin;
          }
        }
        state.rx = Math.min(baseX + 1.25, Math.max(baseX - 1.25, state.rx));
        target.rotation.x = state.rx;
        target.rotation.y = state.ry;
        if (opts.onStep) opts.onStep();
      },
      dispose: function () {
        canvas.removeEventListener("pointerdown", down);
        canvas.removeEventListener("pointermove", move);
        canvas.removeEventListener("pointerup", up);
        canvas.removeEventListener("pointercancel", up);
        canvas.removeEventListener("lostpointercapture", up);
        canvas.removeEventListener("wheel", wheel);
      }
    };
  }

  /* =========================================================
     HERO — silicon wafer, die grid, dopant particles
     ========================================================= */
  function initHeroScene(canvas) {
    if (!window.THREE) return null;

    const renderer = new THREE.WebGLRenderer({
      canvas: canvas,
      antialias: true,
      alpha: true,
      powerPreference: "high-performance"
    });
    renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 1.25));
    renderer.outputEncoding = THREE.sRGBEncoding;
    renderer.toneMapping = THREE.ACESFilmicToneMapping;
    renderer.toneMappingExposure = 1.05;

    const scene = new THREE.Scene();

    const camera = new THREE.PerspectiveCamera(42, 1, 0.1, 60);
    camera.position.set(0, 2.6, 7.2);
    camera.lookAt(0, 0, 0);

    const root = new THREE.Group();
    root.rotation.x = 0.22;
    root.rotation.z = 0.06;
    scene.add(root);

    // wafer body — deep royal blue so it reads crisply on a bright surface
    const wafer = new THREE.Mesh(
      new THREE.CylinderGeometry(2.35, 2.35, 0.16, 256),
      mat(0x0d3577, { metalness: 0.55, roughness: 0.32, emissive: 0x061c44, emissiveIntensity: 0.3 })
    );
    root.add(wafer);

    // iridescent top surface — thin-film interference shader (real wafer look)
    const iriUniforms = {
      uTime: { value: 0 },
      uPulseT: { value: -10 },   // pulse progress in seconds; -10 = idle
      uCamPos: { value: new THREE.Vector3() },
      uBoost: { value: 0 }       // extra brightness during the pulse
    };
    const iriMat = new THREE.ShaderMaterial({
      uniforms: iriUniforms,
      side: THREE.DoubleSide,
      vertexShader: [
        "varying vec3 vWorld;",
        "varying vec2 vLocal;",
        "varying vec3 vNormalW;",
        "void main() {",
        "  vLocal = position.xy;",              // circle geometry is in XY before rotation
        "  vNormalW = normalize(mat3(modelMatrix) * normal);",
        "  vec4 w = modelMatrix * vec4(position, 1.0);",
        "  vWorld = w.xyz;",
        "  gl_Position = projectionMatrix * viewMatrix * w;",
        "}"
      ].join("\n"),
      fragmentShader: [
        "uniform float uTime;",
        "uniform float uPulseT;",
        "uniform float uBoost;",
        "uniform vec3 uCamPos;",
        "varying vec3 vWorld;",
        "varying vec2 vLocal;",
        "varying vec3 vNormalW;",
        "",
        "vec3 spectral(float x) {",            // smooth rainbow ramp
        "  return 0.5 + 0.5 * cos(6.28318 * (x + vec3(0.0, 0.33, 0.67)));",
        "}",
        "void main() {",
        "  float r = length(vLocal) / 2.34;",
        "  float ang = atan(vLocal.y, vLocal.x);",
        "  vec3 V = normalize(uCamPos - vWorld);",
        "  vec3 N = normalize(vNormalW);",   // wafer up
        "  float fres = pow(1.0 - abs(dot(V, N)), 1.6);",
        "",
        "  // thin-film interference: hue shifts with radius, angle & view",
        "  float film = r * 1.6 + ang / 6.28318 * 1.2 + fres * 1.1 + uTime * 0.02;",
        "  vec3 iri = spectral(film);",
        "",
        "  // deep silicon base tint",
        "  vec3 base = mix(vec3(0.05, 0.12, 0.34), vec3(0.12, 0.26, 0.60), fres);",
        "  vec3 col = mix(base, iri, 0.28 + fres * 0.45);",
        "",
        "  // expanding rainbow pulse wavefront",
        "  float pw = uPulseT * 2.6;",          // wavefront radius
        "  float band = exp(-pow((r * 2.34 - pw) * 2.6, 2.0));",
        "  vec3 pulseCol = spectral(r * 1.8 - uPulseT * 0.9);",
        "  col += pulseCol * band * 2.2;",
        "  col += spectral(film + 0.33) * exp(-pow((r * 2.34 - pw * 0.7) * 4.0, 2.0)) * 0.9;",
        "  col *= 1.0 + uBoost * 0.7;",
        "",
        "  gl_FragColor = vec4(col, 1.0);",
        "}"
      ].join("\n")
    });
    const top = new THREE.Mesh(new THREE.CircleGeometry(2.34, 256), iriMat);
    top.rotation.x = -Math.PI / 2;
    top.position.y = 0.085;
    root.add(top);

    // bevelled rim
    const rim = new THREE.Mesh(
      new THREE.TorusGeometry(2.34, 0.022, 16, 320),
      mat(0x8fc0ff, { metalness: 0.7, roughness: 0.25, emissive: 0x2f6bff, emissiveIntensity: 0.35 })
    );
    rim.rotation.x = Math.PI / 2;
    rim.position.y = 0.088;
    root.add(rim);

    // die grid — finer pitch for a more precise wafer look
    const gp = [];
    const step = 0.21;
    for (let x = -2.3; x <= 2.3; x += step) {
      const half = Math.sqrt(Math.max(0, 2.28 * 2.28 - x * x));
      gp.push(new THREE.Vector3(x, 0.095, -half), new THREE.Vector3(x, 0.095, half));
    }
    for (let z = -2.3; z <= 2.3; z += step) {
      const half = Math.sqrt(Math.max(0, 2.28 * 2.28 - z * z));
      gp.push(new THREE.Vector3(-half, 0.095, z), new THREE.Vector3(half, 0.095, z));
    }
    root.add(lines(gp, 0x9cc8ff, 0.42));

    // highlighted dies — a few glowing chips scattered across the grid
    const dieGeo = new THREE.BoxGeometry(step * 0.82, 0.022, step * 0.82);
    for (let i = 0; i < 22; i++) {
      const gx = Math.round((Math.random() * 16 - 8)) * step;
      const gz = Math.round((Math.random() * 16 - 8)) * step;
      if (gx * gx + gz * gz > 2.2 * 2.2) continue;
      const d = new THREE.Mesh(
        dieGeo,
        mat(Math.random() > 0.72 ? 0xf0c44a : 0x3f8dff, {
          metalness: 0.6,
          roughness: 0.3,
          emissive: Math.random() > 0.5 ? 0x1d4fa8 : 0x123a72,
          emissiveIntensity: 0.7
        })
      );
      d.position.set(gx, 0.098, gz);
      root.add(d);
    }

    // concentric rings
    [0.75, 1.25, 1.75, 2.27].forEach(function (r, i) {
      const ring = new THREE.Mesh(
        new THREE.TorusGeometry(r, 0.007 + i * 0.002, 10, 200),
        mat(i === 3 ? 0xf0c44a : 0x6fd0ff, {
          metalness: 0.7,
          roughness: 0.26,
          emissive: i === 3 ? C.gold : 0x2ba9e0,
          emissiveIntensity: i === 3 ? 0.8 : 0.55
        })
      );
      ring.rotation.x = Math.PI / 2;
      ring.position.y = 0.1;
      root.add(ring);
    });

    // central chip — DRAGGABLE QFP-style package the user places onto the socket
    const chipSub = new THREE.Mesh(
      new THREE.BoxGeometry(0.86, 0.06, 0.86),
      mat(0x0b2a58, { metalness: 0.7, roughness: 0.3 })
    );
    chipSub.position.y = 0.115;
    root.add(chipSub);

    // QFP package: dark epoxy body + iridescent wafer window + gull-wing leads
    const chip = new THREE.Group();

    const chipMat = mat(0x23262b, { metalness: 0.25, roughness: 0.5 });
    const chipBody = new THREE.Mesh(new THREE.BoxGeometry(0.8, 0.16, 0.8), chipMat);
    chip.add(chipBody);

    // window — IMEN emblem on a clean white disc (true colours, no tone-mapping)
    const chipWindow = new THREE.Mesh(
      new THREE.CircleGeometry(0.31, 64),
      new THREE.MeshBasicMaterial({ color: 0xffffff, toneMapped: false })
    );
    chipWindow.rotation.x = -Math.PI / 2;
    chipWindow.position.y = 0.085;
    chip.add(chipWindow);
    const logoTex = new THREE.TextureLoader().load("assets/imen-emblem.png", function () {
      logoMesh.visible = true;
    });
    logoTex.encoding = THREE.sRGBEncoding;
    logoTex.anisotropy = 4;
    const logoMesh = new THREE.Mesh(
      new THREE.PlaneGeometry(0.5, 0.5),
      new THREE.MeshBasicMaterial({ map: logoTex, transparent: true, toneMapped: false })
    );
    logoMesh.rotation.x = -Math.PI / 2;
    logoMesh.position.y = 0.088;
    logoMesh.visible = false; // shown once the texture has loaded
    chip.add(logoMesh);

    // window bezel
    const bezel = new THREE.Mesh(
      new THREE.TorusGeometry(0.325, 0.018, 10, 64),
      mat(0x3a3e45, { metalness: 0.4, roughness: 0.45 })
    );
    bezel.rotation.x = Math.PI / 2;
    bezel.position.y = 0.085;
    chip.add(bezel);

    // corner dimples (mould marks)
    const dimpleGeo = new THREE.CylinderGeometry(0.045, 0.045, 0.02, 20);
    const dimpleMat = mat(0x16181c, { metalness: 0.2, roughness: 0.6 });
    [[-0.33, -0.33], [0.33, -0.33], [-0.33, 0.33], [0.33, 0.33]].forEach(function (c) {
      const dimple = new THREE.Mesh(dimpleGeo, dimpleMat);
      dimple.position.set(c[0], 0.085, c[1]);
      chip.add(dimple);
    });

    // gull-wing leads — shoulder + foot per pin, 12 pins per side
    const leadMat = mat(0xd8dee6, { metalness: 0.95, roughness: 0.22 });
    const shoulderGeo = new THREE.BoxGeometry(0.038, 0.014, 0.1);
    const footGeo = new THREE.BoxGeometry(0.038, 0.012, 0.09);
    for (let side = 0; side < 4; side++) {
      for (let i = 0; i < 12; i++) {
        const t = -0.32 + (i / 11) * 0.64;
        const lead = new THREE.Group();
        const shoulder = new THREE.Mesh(shoulderGeo, leadMat);
        shoulder.position.z = 0.05;
        shoulder.rotation.x = 0.5;
        const foot = new THREE.Mesh(footGeo, leadMat);
        foot.position.set(0, -0.055, 0.125);
        lead.add(shoulder, foot);
        lead.position.set(t, 0.0, 0.4);
        lead.rotation.y = (side * Math.PI) / 2;
        chip.add(lead);
      }
    }

    chip.position.y = 0.21;
    root.add(chip);

    const chipEdge = new THREE.LineSegments(
      new THREE.EdgesGeometry(new THREE.BoxGeometry(0.8, 0.16, 0.8)),
      lineMat(0x565d68, 0.9)
    );
    chipEdge.position.copy(chip.position);
    root.add(chipEdge);

    // die home (socket, centre) + start (off-centre, hovering)
    const SOCKET = new THREE.Vector3(0, 0.24, 0);
    const DIE_HOME = SOCKET.clone();
    const DIE_START = new THREE.Vector3(1.18, 0.68, 0.88);
    chip.position.copy(DIE_START);
    chipEdge.position.copy(DIE_START);
    chipSub.position.set(0, 0.115, 0);

    // socket highlight — dashed target ring at the centre
    const socketRing = new THREE.Mesh(
      new THREE.TorusGeometry(0.62, 0.016, 10, 112),
      mat(C.gold, { emissive: C.gold, emissiveIntensity: 1.1, transparent: true, opacity: 0.85 })
    );
    socketRing.rotation.x = Math.PI / 2;
    socketRing.position.y = 0.105;
    root.add(socketRing);
    const socketPad = new THREE.Mesh(
      new THREE.BoxGeometry(0.72, 0.012, 0.72),
      new THREE.MeshBasicMaterial({ color: C.gold, transparent: true, opacity: 0.14 })
    );
    socketPad.position.y = 0.098;
    root.add(socketPad);

    // socket light shine — soft additive glow on the wafer, a rising light shaft and a warm point light
    function radialTex() {
      const c = document.createElement("canvas");
      c.width = c.height = 128;
      const g = c.getContext("2d");
      const grd = g.createRadialGradient(64, 64, 0, 64, 64, 64);
      grd.addColorStop(0, "rgba(255,244,200,1)");
      grd.addColorStop(0.3, "rgba(255,206,92,0.6)");
      grd.addColorStop(0.65, "rgba(255,170,40,0.16)");
      grd.addColorStop(1, "rgba(255,170,40,0)");
      g.fillStyle = grd;
      g.fillRect(0, 0, 128, 128);
      return new THREE.CanvasTexture(c);
    }
    function shaftTex() {
      const c = document.createElement("canvas");
      c.width = 8;
      c.height = 128;
      const g = c.getContext("2d");
      const grd = g.createLinearGradient(0, 128, 0, 0); // bottom -> top
      grd.addColorStop(0, "rgba(255,236,170,0.95)");
      grd.addColorStop(0.5, "rgba(255,205,90,0.35)");
      grd.addColorStop(1, "rgba(255,190,60,0)");
      g.fillStyle = grd;
      g.fillRect(0, 0, 8, 128);
      return new THREE.CanvasTexture(c);
    }
    const shineGlow = new THREE.Mesh(
      new THREE.PlaneGeometry(3.0, 3.0),
      new THREE.MeshBasicMaterial({
        map: radialTex(), transparent: true, opacity: 0.7, depthWrite: false,
        blending: THREE.AdditiveBlending, toneMapped: false
      })
    );
    shineGlow.rotation.x = -Math.PI / 2;
    shineGlow.position.set(SOCKET.x, 0.108, SOCKET.z);
    root.add(shineGlow);
    const shineShaft = new THREE.Mesh(
      new THREE.CylinderGeometry(0.5, 0.78, 2.2, 40, 1, true),
      new THREE.MeshBasicMaterial({
        map: shaftTex(), transparent: true, opacity: 0.4, depthWrite: false, side: THREE.DoubleSide,
        blending: THREE.AdditiveBlending, toneMapped: false
      })
    );
    shineShaft.position.set(SOCKET.x, 0.1 + 1.1, SOCKET.z);
    root.add(shineShaft);
    const shineLight = new THREE.PointLight(0xffc860, 1.4, 6);
    shineLight.position.set(SOCKET.x, 0.9, SOCKET.z);
    root.add(shineLight);

    // die drop shadow (moves with the die)
    const shadow = new THREE.Mesh(
      new THREE.CircleGeometry(0.52, 48),
      new THREE.MeshBasicMaterial({ color: 0x04122e, transparent: true, opacity: 0.3 })
    );
    shadow.rotation.x = -Math.PI / 2;
    shadow.position.set(DIE_START.x, 0.096, DIE_START.z);
    root.add(shadow);

    // pulse wavefront — expanding ring fired on placement
    const pulse = new THREE.Mesh(
      new THREE.TorusGeometry(1, 0.03, 10, 120),
      new THREE.MeshBasicMaterial({ color: 0xffe27a, transparent: true, opacity: 0 })
    );
    pulse.rotation.x = Math.PI / 2;
    pulse.position.y = 0.105;
    root.add(pulse);

    // sequential pad flash
    const padMats = [];
    const padGeo = new THREE.BoxGeometry(0.1, 0.02, 0.028);
    for (let i = 0; i < 32; i++) {
      const a = (i / 32) * Math.PI * 2;
      const pm = mat(0x9cc8ff, { metalness: 0.8, roughness: 0.3, emissive: 0x2f6bff, emissiveIntensity: 0.5 });
      padMats.push(pm);
      const pad = new THREE.Mesh(padGeo, pm);
      pad.position.set(Math.cos(a) * 0.56, 0.115, Math.sin(a) * 0.56);
      pad.rotation.y = -a;
      root.add(pad);
    }

    // floating dopant particles above the wafer
    const sparks = [];
    for (let i = 0; i < 26; i++) {
      const a = Math.random() * Math.PI * 2;
      const r = 0.5 + Math.random() * 1.8;
      const s = new THREE.Mesh(
        new THREE.SphereGeometry(0.02 + Math.random() * 0.022, 10, 10),
        mat(Math.random() > 0.5 ? C.cyan : 0xf0c44a, { emissive: Math.random() > 0.5 ? C.cyan : C.gold, emissiveIntensity: 1.4, metalness: 0.1 })
      );
      s.userData = { a: a, r: r, h: 0.2 + Math.random() * 0.9, sp: 0.15 + Math.random() * 0.25 };
      sparks.push(s);
      root.add(s);
    }

    addLights(scene);

    let visible = true;
    const io =
      "IntersectionObserver" in window
        ? new IntersectionObserver(
            function (entries) {
              visible = entries[0].isIntersecting;
            },
            { threshold: 0.02 }
          )
        : null;
    if (io) io.observe(canvas);

    function layout() {
      const w = Math.max(canvas.clientWidth || canvas.parentElement.clientWidth || 1, 1);
      const h = Math.max(canvas.clientHeight || canvas.parentElement.clientHeight || 1, 1);
      renderer.setSize(w, h, false);
      const aspect = w / h;
      camera.aspect = aspect;

      if (aspect < 0.8) {
        root.scale.setScalar(0.74);
        camera.position.z = 9.4;
      } else if (aspect < 1.15) {
        root.scale.setScalar(0.94);
        camera.position.z = 8.4;
      } else {
        root.scale.setScalar(1.12);
        camera.position.z = 7.6;
      }
      camera.updateProjectionMatrix();
    }

    const ro = new ResizeObserver(layout);
    ro.observe(canvas.parentElement || canvas);
    layout();

    // interactive orbit — drag to spin/tilt, pinch or wheel to zoom
    const baseZ = { z: camera.position.z };
    /* ---------- die placement mini-interaction ---------- */
    const die = { placed: false, dragging: false, hovering: true, t: 0 };
    // static wafer view — orbiting is disabled so it never fights the die drag
    const orbit = { step: function () {}, dispose: function () {} };
    void baseZ;
    const raycaster = new THREE.Raycaster();
    const pointerNdc = new THREE.Vector2();
    const hitPoint = new THREE.Vector3();
    // drag plane — the wafer's own surface (root-local y = 0)
    function makeDragPlane(height = 0.55) {
      const o = new THREE.Vector3(0, height, 0);
      const x = new THREE.Vector3(1, height, 0);
      const z = new THREE.Vector3(0, height, 1);
      root.localToWorld(o);
      root.localToWorld(x);
      root.localToWorld(z);
      return new THREE.Plane().setFromCoplanarPoints(o, x, z);
    }
    let dragPlane = makeDragPlane();
    let onUnlock = null;
    let onPlace = null;
    let onFrame = null;
    let pendingPtr = null; // latest pointer position, applied once per frame
    let dragTarget = null;
    const grabOffset = new THREE.Vector2();
    let activePointer = null;
    const originalTouchAction = canvas.style.touchAction;
    canvas.style.touchAction = "none";

    function toRootLocal(v) {
      const inv = new THREE.Matrix4().copy(root.matrixWorld).invert();
      return v.clone().applyMatrix4(inv);
    }
    function pointerToWafer(e) {
      const rect = canvas.getBoundingClientRect();
      pointerNdc.set(
        ((e.clientX - rect.left) / rect.width) * 2 - 1,
        -((e.clientY - rect.top) / rect.height) * 2 + 1
      );
      raycaster.setFromCamera(pointerNdc, camera);
      if (raycaster.ray.intersectPlane(dragPlane, hitPoint)) {
        return toRootLocal(hitPoint);
      }
      return null;
    }
    function overDie(e) {
      const rect = canvas.getBoundingClientRect();
      pointerNdc.set(
        ((e.clientX - rect.left) / rect.width) * 2 - 1,
        -((e.clientY - rect.top) / rect.height) * 2 + 1
      );
      raycaster.setFromCamera(pointerNdc, camera);
      // generous hit target — 0.9 radius sphere around the die centre
      const hit = new THREE.Vector3();
      const world = chip.position.clone();
      root.localToWorld(world);
      return raycaster.ray.distanceToPoint(world) < 1.05;
    }
    function dieDown(e) {
      if (activePointer !== null || (e.button != null && e.button !== 0)) return;
      if (die.placed) return;
      if (!overDie(e)) return;
      activePointer = e.pointerId;
      root.updateMatrixWorld(true);
      dragPlane = makeDragPlane(chip.position.y);
      const grab = pointerToWafer(e);
      grabOffset.set(grab ? chip.position.x - grab.x : 0, grab ? chip.position.z - grab.z : 0);
      die.dragging = true;
      die.hovering = false;
      pendingPtr = { clientX: e.clientX, clientY: e.clientY };
      applyDrag();
      easeChip();
      if (e.cancelable) e.preventDefault();
      e.stopImmediatePropagation();
      canvas.style.cursor = "grabbing";
      // keep receiving moves even if the pointer leaves the canvas
      try { canvas.setPointerCapture(e.pointerId); } catch (err) {}
    }
    function dieMove(e) {
      if (!die.dragging || e.pointerId !== activePointer) return;
      pendingPtr = { clientX: e.clientX, clientY: e.clientY };
      applyDrag();
      easeChip();
      if (REDUCED) renderer.render(scene, camera);
      if (e.cancelable) e.preventDefault();
      e.stopImmediatePropagation();
    }
    function applyDrag() {
      if (!die.dragging || !pendingPtr) return;
      const p = pointerToWafer(pendingPtr);
      pendingPtr = null;
      if (!p) return;
      p.x += grabOffset.x;
      p.z += grabOffset.y;
      const r = Math.hypot(p.x, p.z);
      const maxR = 2.05;
      if (r > maxR) {
        p.x *= maxR / r;
        p.z *= maxR / r;
      }
      dragTarget = { x: p.x, z: p.z };
    }
    function easeChip() {
      if (!die.dragging || !dragTarget) return;
      const k = 1; // track the pointer directly, without trailing interpolation
      const x = chip.position.x + (dragTarget.x - chip.position.x) * k;
      const z = chip.position.z + (dragTarget.z - chip.position.z) * k;
      chip.position.set(x, chip.position.y, z);
      chipEdge.position.copy(chip.position);
      shadow.position.set(x, 0.096, z);
      shadow.material.opacity = 0.34;
    }
    function dieUp(e) {
      if (!die.dragging || e.pointerId !== activePointer) return;
      pendingPtr = { clientX: e.clientX, clientY: e.clientY };
      applyDrag();
      easeChip();
      die.dragging = false;
      activePointer = null;
      canvas.style.cursor = "grab";
      try { canvas.releasePointerCapture(e.pointerId); } catch (err) {}
      applyDrag();
      if (dragTarget) {
        chip.position.set(dragTarget.x, 0.55, dragTarget.z);
        chipEdge.position.copy(chip.position);
        shadow.position.set(dragTarget.x, 0.096, dragTarget.z);
      }
      dragTarget = null;
      const d = Math.hypot(chip.position.x - SOCKET.x, chip.position.z - SOCKET.z);
      if (d < 0.7) {
        place();
      }
      if (REDUCED) renderer.render(scene, camera);
      e.stopImmediatePropagation();
    }
    function place() {
      if (die.placed) return;
      die.placed = true;
      chip.position.copy(DIE_HOME);
      chipEdge.position.copy(DIE_HOME);
      shadow.material.opacity = 0.22;
      shadow.position.set(DIE_HOME.x, 0.096, DIE_HOME.z);
      die.pulseT = 0;
      // onPlace fires the instant the die is seated, together with the wafer pulse
      if (onPlace) {
        try { onPlace(); } catch (err) {}
      }
      // onUnlock fires when the pulse animation completes (see frame loop)
    }
    // accessibility fallback — clicking the socket also places the die
    function tapPlace(e) {
      if (die.placed || die.dragging) return;
      const rect = canvas.getBoundingClientRect();
      pointerNdc.set(
        ((e.clientX - rect.left) / rect.width) * 2 - 1,
        -((e.clientY - rect.top) / rect.height) * 2 + 1
      );
      raycaster.setFromCamera(pointerNdc, camera);
      if (raycaster.intersectObject(socketPad, false).length) place();
    }
    // stop the page from scrolling / cancelling the pointer while the chip is grabbed on touch screens
    function dieTouchStart(e) {
      if (die.placed || !e.touches.length) return;
      if (overDie(e.touches[0])) e.preventDefault();
    }
    function dieTouchMove(e) {
      if (die.dragging) e.preventDefault();
    }
    canvas.addEventListener("touchstart", dieTouchStart, { passive: false });
    canvas.addEventListener("touchmove", dieTouchMove, { passive: false });
    canvas.addEventListener("pointerdown", dieDown, true);
    canvas.addEventListener("pointermove", dieMove, true);
    canvas.addEventListener("pointerup", dieUp, true);
    canvas.addEventListener("pointercancel", dieUp, true);
    canvas.addEventListener("click", tapPlace);
    function disposeDie() {
      canvas.style.touchAction = originalTouchAction;
      canvas.removeEventListener("pointerdown", dieDown, true);
      canvas.removeEventListener("pointermove", dieMove, true);
      canvas.removeEventListener("pointerup", dieUp, true);
      canvas.removeEventListener("pointercancel", dieUp, true);
      canvas.removeEventListener("click", tapPlace);
      canvas.removeEventListener("touchstart", dieTouchStart);
      canvas.removeEventListener("touchmove", dieTouchMove);
    }

    // Motion is intentionally minimal: one very slow rotation, no float, no parallax.
    if (REDUCED) {
      renderer.render(scene, camera);
      return {
        onUnlock: function (fn) { onUnlock = fn; },
        onPlace: function (fn) { onPlace = fn; },
        onFrame: function (fn) { onFrame = fn; },
        place: function () {
          place();
          renderer.render(scene, camera);
          if (onUnlock) {
            try { onUnlock(); } catch (err) {}
          }
        },
        isPlaced: function () { return die.placed; },
        dispose: function () {
          disposeDie();
          orbit.dispose();
          if (io) io.disconnect();
          ro.disconnect();
          renderer.dispose();
        }
      };
    }

    const clock = new THREE.Clock();
    let raf = 0;
    let lastNow = performance.now();

    function frame() {
      raf = requestAnimationFrame(frame);
      const now = performance.now();
      const dt = Math.min(0.05, (now - lastNow) / 1000); // frame-rate independent animation
      lastNow = now;
      if (!visible) return;
      applyDrag();
      easeChip();
      const t = clock.getElapsedTime();
      orbit.step(0.0009);

      // iridescent wafer uniforms
      iriUniforms.uTime.value = t;
      iriUniforms.uCamPos.value.copy(camera.position);

      // die idle motion / snap animation / pulse wavefront
      if (!die.placed) {
        if (!die.dragging) {
          chip.position.y = 0.55 + Math.sin(t * 1.6) * 0.07;
          chip.rotation.y = Math.sin(t * 0.9) * 0.14;
          chipEdge.position.copy(chip.position);
          chipEdge.rotation.copy(chip.rotation);
          shadow.position.set(chip.position.x, 0.096, chip.position.z);
          shadow.material.opacity = 0.24 + Math.sin(t * 1.6) * 0.06;
          socketRing.material.opacity = 0.55 + Math.sin(t * 2.2) * 0.3;
          socketRing.scale.setScalar(1 + Math.sin(t * 2.2) * 0.05);
        }
        // shine breathes, and flares up as the chip gets closer to the socket
        const near = Math.max(0, 1 - Math.hypot(chip.position.x - SOCKET.x, chip.position.z - SOCKET.z) / 1.8);
        const br = 0.5 + 0.5 * Math.sin(t * 2.2);
        const flare = 1 + near * 0.9;
        shineGlow.material.opacity = Math.min(1, (0.5 + br * 0.3) * flare);
        shineGlow.scale.setScalar(1 + br * 0.08 + near * 0.18);
        shineShaft.material.opacity = Math.min(0.9, (0.26 + br * 0.16) * flare);
        shineLight.intensity = (1.0 + br * 0.7) * flare;
      } else if (die.pulseT != null && die.pulseT < 1.6) {
        // snap-in
        die.pulseT += 1.68 * dt;
        const k = Math.min(1, die.pulseT);
        const ease = 1 - Math.pow(1 - k, 3);
        chip.position.y = 0.55 + (DIE_HOME.y - 0.55) * Math.min(1, ease * 2.2);
        chip.rotation.y *= 0.9;
        chipEdge.position.copy(chip.position);
        chipEdge.rotation.copy(chip.rotation);
        // rainbow pulse across the whole wafer
        iriUniforms.uPulseT.value = k * 1.4;
        iriUniforms.uBoost.value = Math.sin(k * Math.PI) * 0.8;
        // expanding pulse ring
        const pr = 0.4 + ease * 2.05;
        pulse.scale.set(pr, pr, 1);
        pulse.material.opacity = (1 - ease) * 0.95;
        pulse.material.color.setHSL((0.13 + ease * 0.85) % 1, 0.9, 0.6);
        // pads flash outward in sequence
        padMats.forEach(function (pm, i) {
          const hit = ease * 34 - i;
          pm.emissiveIntensity = hit > 0 && hit < 2.2 ? 2.2 : 0.5;
        });
        chipMat.emissive = new THREE.Color(0xd99a00);
        chipMat.emissiveIntensity = Math.sin(Math.min(1, die.pulseT) * Math.PI) * 0.55;
        socketRing.material.opacity = Math.max(0, 0.85 - ease);
        shineGlow.material.opacity = Math.max(0, 0.9 * (1 - ease));
        shineShaft.material.opacity = Math.max(0, 0.5 * (1 - ease));
        shineLight.intensity = Math.max(0, 1.6 * (1 - ease));
        if (die.pulseT >= 1.6 && !die.unlocked) {
          die.unlocked = true;
          if (onUnlock) {
            try { onUnlock(); } catch (err) {}
          }
        }
      } else {
        pulse.material.opacity = 0;
        shineGlow.material.opacity = 0;
        shineShaft.material.opacity = 0;
        shineLight.intensity = 0;
        iriUniforms.uPulseT.value = -10;
        iriUniforms.uBoost.value = Math.max(0, iriUniforms.uBoost.value - 0.03);
        chipMat.emissive = new THREE.Color(0xd99a00);
        chipMat.emissiveIntensity = 0.18 + Math.sin(t * 2) * 0.06;
      }

      sparks.forEach(function (s) {
        const u = s.userData;
        s.position.set(
          Math.cos(u.a + t * u.sp) * u.r,
          u.h + Math.sin(t * 0.8 + u.a * 3) * 0.12,
          Math.sin(u.a + t * u.sp) * u.r
        );
      });
      renderer.render(scene, camera);
      if (onFrame) onFrame();
    }

    // warm-up: compile every program and draw one frame in the "pulse" state now,
    // so the first drop of the chip does not pay for it in the middle of the interaction
    try {
      renderer.compile(scene, camera);
      iriUniforms.uPulseT.value = 0.4;
      iriUniforms.uBoost.value = 0.4;
      pulse.material.opacity = 0.05;
      renderer.render(scene, camera);
      iriUniforms.uPulseT.value = -10;
      iriUniforms.uBoost.value = 0;
      pulse.material.opacity = 0;
    } catch (err) {}
    frame();

    return {
      onUnlock: function (fn) { onUnlock = fn; },
      onPlace: function (fn) { onPlace = fn; },
      onFrame: function (fn) { onFrame = fn; },
      place: place,
      isPlaced: function () { return die.placed; },
      debugScreenPos: function () {
        const v = chip.position.clone();
        root.localToWorld(v);
        v.project(camera);
        const rect = canvas.getBoundingClientRect();
        return {
          x: rect.left + ((v.x + 1) / 2) * rect.width,
          y: rect.top + ((1 - v.y) / 2) * rect.height
        };
      },
      // position of the die relative to the canvas element (for the floating label)
      dieLabelPos: function () {
        const v = chip.position.clone();
        v.y += 0.32; // float above the die
        root.localToWorld(v);
        v.project(camera);
        return {
          x: ((v.x + 1) / 2) * canvas.clientWidth,
          y: ((1 - v.y) / 2) * canvas.clientHeight
        };
      },
      dispose: function () {
        cancelAnimationFrame(raf);
        disposeDie();
        orbit.dispose();
        if (io) io.disconnect();
        ro.disconnect();
        renderer.dispose();
      }
    };
  }

  /* =========================================================
     COURSE ILLUSTRATIONS
     ========================================================= */
  function buildScene(key) {
    // hyper-real illustrations registered by assets/courses/*.js (see course-kit.js)
    const reg = window.IMEN && window.IMEN.courseScenes && window.IMEN.courseScenes[key];
    if (reg && window.IMEN.kit) {
      try {
        const out = reg(window.IMEN.kit);
        const holder = window.IMEN.kit.fit(out.group, out.fit || 2.05);
        return { group: holder, update: out.update || function () {}, real: true, spin: out.spin };
      } catch (err) {
        console.warn("Course scene '" + key + "' failed, using fallback:", err);
      }
    }
    const g = new THREE.Group();
    let update = function () {};

    if (key === "device") {
      const p = new THREE.Mesh(new THREE.BoxGeometry(1.25, 1.25, 0.55), mat(0x1d4fa8, { emissive: 0x0a2b66, emissiveIntensity: 0.6 }));
      p.position.x = -0.68;
      const n = new THREE.Mesh(new THREE.BoxGeometry(1.25, 1.25, 0.55), mat(C.orange, { emissive: 0x7a3d00, emissiveIntensity: 0.55 }));
      n.position.x = 0.68;
      const j = new THREE.Mesh(new THREE.BoxGeometry(0.07, 1.28, 0.58), mat(C.cyan, { emissive: C.cyan, emissiveIntensity: 1.4, metalness: 0.2 }));
      g.add(p, n, j);
      // metallisation contacts on both faces
      [-1, 1].forEach(function (sx) {
        const contact = new THREE.Mesh(new THREE.BoxGeometry(0.1, 1.32, 0.62), mat(0xf0c44a, { metalness: 0.95, roughness: 0.2, emissive: 0x8a5c00, emissiveIntensity: 0.4 }));
        contact.position.x = sx * 1.34;
        g.add(contact);
        const lead = new THREE.Mesh(new THREE.CylinderGeometry(0.035, 0.035, 0.7, 14), mat(0xf0c44a, { metalness: 1, roughness: 0.25 }));
        lead.rotation.z = Math.PI / 2;
        lead.position.x = sx * 1.72;
        g.add(lead);
      });
      // doping lattice dots inside each region
      for (let i = 0; i < 30; i++) {
        const side = i % 2 ? -1 : 1;
        const d = new THREE.Mesh(new THREE.SphereGeometry(0.026, 8, 8), mat(side < 0 ? 0x8fc0ff : 0xffd27a, { emissive: side < 0 ? 0x2f6bff : C.gold, emissiveIntensity: 0.9, metalness: 0.2 }));
        d.position.set(side * (0.25 + Math.random() * 1.0), (Math.random() - 0.5) * 1.05, (Math.random() - 0.5) * 0.45);
        g.add(d);
      }

      const carriers = [];
      for (let i = 0; i < 14; i++) {
        const s = new THREE.Mesh(new THREE.SphereGeometry(0.05, 12, 12), mat(i % 2 ? C.cyan : C.gold, { emissive: i % 2 ? C.cyan : C.gold, emissiveIntensity: 1.2, metalness: 0.1 }));
        s.userData.phase = Math.random();
        s.userData.y = (Math.random() - 0.5) * 0.95;
        s.userData.z = (Math.random() - 0.5) * 0.4;
        carriers.push(s);
        g.add(s);
      }
      update = function (t) {
        carriers.forEach(function (s) {
          const k = (t * 0.28 + s.userData.phase) % 1;
          s.position.set(-1.2 + k * 2.4, s.userData.y, s.userData.z);
          s.material.opacity = 0.35 + Math.sin(k * Math.PI) * 0.65;
          s.material.transparent = true;
        });
      };
    } else if (key === "fab") {
      const wafer = new THREE.Mesh(new THREE.CylinderGeometry(1.55, 1.55, 0.12, 120), mat(0x123a72, { metalness: 0.9, roughness: 0.22, emissive: 0x071d42, emissiveIntensity: 0.7 }));
      g.add(wafer);
      // die grid on the wafer
      const gp = [];
      const step = 0.22;
      for (let x = -1.5; x <= 1.5; x += step) {
        const half = Math.sqrt(Math.max(0, 1.5 * 1.5 - x * x));
        gp.push(new THREE.Vector3(x, 0.065, -half), new THREE.Vector3(x, 0.065, half));
        gp.push(new THREE.Vector3(-half, 0.065, x), new THREE.Vector3(half, 0.065, x));
      }
      g.add(lines(gp, 0x9cc8ff, 0.35));
      [0.55, 0.95, 1.35, 1.52].forEach(function (r) {
        const ring = new THREE.Mesh(new THREE.TorusGeometry(r, 0.008, 8, 140), mat(C.cyan, { emissive: C.cyan, emissiveIntensity: 0.8 }));
        ring.rotation.x = Math.PI / 2;
        ring.position.y = 0.07;
        g.add(ring);
      });
      const beam = new THREE.Mesh(
        new THREE.ConeGeometry(0.5, 1.7, 40, 1, true),
        new THREE.MeshBasicMaterial({ color: C.cyan, transparent: true, opacity: 0.16, side: THREE.DoubleSide, blending: THREE.AdditiveBlending, depthWrite: false })
      );
      beam.position.y = 1.15;
      g.add(beam);
      const spot = new THREE.Mesh(new THREE.CircleGeometry(0.16, 32), new THREE.MeshBasicMaterial({ color: 0xe6ffff, transparent: true, opacity: 0.9 }));
      spot.rotation.x = -Math.PI / 2;
      spot.position.y = 0.09;
      g.add(spot);
      // exposure head above the beam
      const head = new THREE.Mesh(new THREE.CylinderGeometry(0.34, 0.42, 0.34, 32), mat(0x0b2a58, { metalness: 0.85, roughness: 0.25, emissive: 0x123a72, emissiveIntensity: 0.4 }));
      head.position.y = 2.05;
      g.add(head);
      const lens = new THREE.Mesh(new THREE.CylinderGeometry(0.2, 0.2, 0.06, 32), mat(0x8fc0ff, { metalness: 0.4, roughness: 0.05, emissive: C.cyan, emissiveIntensity: 0.8 }));
      lens.position.y = 1.87;
      g.add(lens);
      // floating particles in the beam
      const dust = [];
      for (let i = 0; i < 12; i++) {
        const d = new THREE.Mesh(new THREE.SphereGeometry(0.018, 8, 8), new THREE.MeshBasicMaterial({ color: 0xcffcff, transparent: true, opacity: 0.85 }));
        d.userData = { a: Math.random() * Math.PI * 2, r: Math.random() * 0.34, sp: 0.5 + Math.random() };
        dust.push(d);
        g.add(d);
      }
      update = function (t) {
        wafer.rotation.y = t * 0.35;
        beam.material.opacity = 0.12 + Math.sin(t * 2.4) * 0.05;
        spot.material.opacity = 0.6 + Math.sin(t * 2.4) * 0.25;
        g.position.y = Math.sin(t * 0.8) * 0.05;
        dust.forEach(function (d) {
          const u = d.userData;
          const yy = 0.35 + ((t * 0.25 * u.sp + u.a) % 1) * 1.6;
          d.position.set(Math.cos(u.a) * u.r, yy, Math.sin(u.a) * u.r);
          d.material.opacity = 0.85 * (1 - (yy - 0.35) / 1.6);
        });
      };
    } else if (key === "materials") {
      const pts = [];
      const nodes = [];
      for (let x = -1; x <= 1; x++) {
        for (let y = -1; y <= 1; y++) {
          for (let z = -1; z <= 1; z++) {
            const s = new THREE.Mesh(new THREE.SphereGeometry(0.085, 16, 16), mat(z === 0 ? C.gold : C.cyan, { emissive: z === 0 ? C.gold : C.cyan, emissiveIntensity: 0.7, metalness: 0.2 }));
            s.position.set(x * 0.62, y * 0.62, z * 0.62);
            nodes.push(s);
            g.add(s);
            pts.push(s.position.clone());
          }
        }
      }
      const links = [];
      for (let i = 0; i < pts.length; i++) {
        for (let j = i + 1; j < pts.length; j++) {
          if (pts[i].distanceTo(pts[j]) < 0.7) links.push(pts[i], pts[j]);
        }
      }
      g.add(lines(links, C.blue, 0.35));
      // unit-cell wireframe
      const cell = new THREE.LineSegments(new THREE.EdgesGeometry(new THREE.BoxGeometry(1.24, 1.24, 1.24)), lineMat(0xcfe8ff, 0.5));
      g.add(cell);
      const scan = new THREE.Mesh(new THREE.TorusGeometry(1.05, 0.015, 8, 110), mat(C.gold, { emissive: C.gold, emissiveIntensity: 1.2 }));
      scan.rotation.x = Math.PI / 2;
      g.add(scan);
      update = function (t) {
        scan.position.y = Math.sin(t * 0.9) * 0.75;
        scan.material.emissiveIntensity = 0.8 + Math.sin(t * 3) * 0.4;
      };
    } else if (key === "package") {
      const base = new THREE.Mesh(new THREE.BoxGeometry(2.1, 0.16, 1.5), mat(0x0b2a58, { metalness: 0.7 }));
      base.position.y = -0.75;
      g.add(base);
      const layers = [
        { w: 1.7, h: 0.18, d: 1.2, y: -0.55, c: 0x1d4fa8 },
        { w: 1.35, h: 0.18, d: 0.95, y: -0.34, c: 0x2f6bff },
        { w: 1.0, h: 0.18, d: 0.7, y: -0.13, c: 0x4be1ff },
        { w: 0.62, h: 0.22, d: 0.62, y: 0.1, c: 0xf5c542 }
      ];
      const meshes = layers.map(function (l) {
        const m = new THREE.Mesh(new THREE.BoxGeometry(l.w, l.h, l.d), mat(l.c, { metalness: 0.75, roughness: 0.25, emissive: l.c, emissiveIntensity: l.c === C.gold ? 0.5 : 0.18 }));
        m.position.y = l.y;
        g.add(m);
        const edge = new THREE.LineSegments(new THREE.EdgesGeometry(m.geometry), lineMat(0xcfe8ff, 0.35));
        edge.position.copy(m.position);
        g.add(edge);
        return { m: m, edge: edge, y: l.y };
      });
      for (let i = 0; i < 12; i++) {
        const pin = new THREE.Mesh(new THREE.BoxGeometry(0.07, 0.07, 0.2), mat(C.gold, { metalness: 1, roughness: 0.25, emissive: C.gold, emissiveIntensity: 0.4 }));
        const side = i < 6 ? -1 : 1;
        pin.position.set(-1.05 + (i % 6) * 0.42, -0.86, side * 0.78);
        g.add(pin);
      }
      // wire bonds from top die down to the substrate
      const bondPts = [];
      for (let i = 0; i < 8; i++) {
        const a = (i / 8) * Math.PI * 2;
        const x0 = Math.cos(a) * 0.28;
        const z0 = Math.sin(a) * 0.28;
        const x1 = Math.cos(a) * 0.62;
        const z1 = Math.sin(a) * 0.62;
        const mid = new THREE.Vector3((x0 + x1) / 2, 0.42, (z0 + z1) / 2);
        const curve = new THREE.QuadraticBezierCurve3(new THREE.Vector3(x0, 0.22, z0), mid, new THREE.Vector3(x1, -0.53, z1));
        curve.getPoints(14).forEach(function (p, idx, arr) {
          if (idx < arr.length - 1) bondPts.push(p, arr[idx + 1]);
        });
      }
      g.add(lines(bondPts, 0xf0c44a, 0.9));
      update = function (t) {
        meshes.forEach(function (o, i) {
          o.m.position.y = o.y + Math.sin(t * 1.1 + i * 0.6) * 0.015;
          o.edge.position.y = o.m.position.y;
        });
      };
    } else if (key === "systems") {
      const board = new THREE.Mesh(new THREE.BoxGeometry(2.3, 0.12, 1.7), mat(0x0a2450, { metalness: 0.6 }));
      board.position.y = -0.5;
      g.add(board);
      const soc = new THREE.Mesh(new THREE.BoxGeometry(0.72, 0.22, 0.72), mat(C.gold, { metalness: 0.9, roughness: 0.2, emissive: C.gold, emissiveIntensity: 0.35 }));
      soc.position.y = -0.33;
      g.add(soc);
      const chips = [
        [-0.75, -0.18, 0.42],
        [0.75, 0.05, 0.38],
        [-0.72, 0.32, 0.36],
        [0.72, -0.55, 0.34]
      ];
      const chipMeshes = chips.map(function (p) {
        const c = new THREE.Mesh(new THREE.BoxGeometry(p[2], 0.16, p[2]), mat(C.cyan, { emissive: C.cyan, emissiveIntensity: 0.3, metalness: 0.7 }));
        c.position.set(p[0], -0.36, p[1]);
        g.add(c);
        return c;
      });
      const pts = [];
      chips.forEach(function (p) {
        pts.push(new THREE.Vector3(0, -0.36, 0), new THREE.Vector3(p[0], -0.36, p[1]));
      });
      g.add(lines(pts, C.cyan, 0.45));
      // vias
      for (let i = 0; i < 14; i++) {
        const v = new THREE.Mesh(new THREE.CylinderGeometry(0.03, 0.03, 0.06, 10), mat(0xf0c44a, { metalness: 1, roughness: 0.3 }));
        v.position.set(-1.05 + Math.random() * 2.1, -0.42, -0.75 + Math.random() * 1.5);
        g.add(v);
      }
      update = function (t) {
        chipMeshes.forEach(function (c, i) {
          c.material.emissiveIntensity = 0.25 + (Math.sin(t * 2 + i) + 1) * 0.22;
        });
        soc.rotation.y = t * 0.4;
      };
    } else if (key === "research") {
      const core = new THREE.Mesh(new THREE.SphereGeometry(0.26, 24, 24), mat(C.gold, { emissive: C.gold, emissiveIntensity: 0.9, metalness: 0.4 }));
      g.add(core);
      const pts = [];
      const nodes = [];
      for (let i = 0; i < 7; i++) {
        const a = (i / 7) * Math.PI * 2;
        const n = new THREE.Mesh(new THREE.SphereGeometry(0.1, 16, 16), mat(C.cyan, { emissive: C.cyan, emissiveIntensity: 0.9, metalness: 0.3 }));
        n.position.set(Math.cos(a) * 1.35, Math.sin(a * 2) * 0.32, Math.sin(a) * 1.35);
        n.userData.base = n.position.clone();
        nodes.push(n);
        g.add(n);
        pts.push(new THREE.Vector3(0, 0, 0), n.position.clone());
      }
      g.add(lines(pts, C.blue, 0.5));
      const orbitRings = [];
      [0.9, 1.35].forEach(function (r, i) {
        const ring = new THREE.Mesh(new THREE.TorusGeometry(r, 0.008, 8, 120), mat(i ? C.cyan : C.gold, { emissive: i ? C.cyan : C.gold, emissiveIntensity: 0.7, transparent: true, opacity: 0.6 }));
        ring.rotation.x = Math.PI / 2 + i * 0.35;
        orbitRings.push(ring);
        g.add(ring);
      });
      update = function (t) {
        nodes.forEach(function (n, i) {
          n.position.y = n.userData.base.y + Math.sin(t * 1.4 + i) * 0.08;
        });
        orbitRings.forEach(function (r, i) {
          r.rotation.z = t * (0.2 + i * 0.12);
        });
        core.scale.setScalar(1 + Math.sin(t * 2) * 0.06);
      };
    } else if (key === "management") {
      const bars = [];
      const widths = [1.5, 1.1, 1.8, 0.9, 1.35, 1.65];
      const colors = [C.blue, C.cyan, C.gold, C.blue, C.cyan, C.gold];
      widths.forEach(function (w, i) {
        const b = new THREE.Mesh(new THREE.BoxGeometry(w, 0.16, 0.16), mat(colors[i], { emissive: colors[i], emissiveIntensity: 0.35, metalness: 0.6 }));
        b.position.set(-0.9 + (2 - w) / 2, 0.65 - i * 0.28, 0);
        b.userData.baseX = b.position.x;
        b.userData.w = w;
        bars.push(b);
        g.add(b);
      });
      update = function (t) {
        bars.forEach(function (b, i) {
          const k = 0.85 + (Math.sin(t * 1.2 + i * 0.7) + 1) * 0.12;
          b.scale.x = k;
          b.position.x = b.userData.baseX - (b.userData.w * (k - 1)) / 2;
        });
      };
    } else if (key === "project") {
      const platforms = [
        { w: 1.9, y: -0.7, c: 0x1d4fa8 },
        { w: 1.3, y: -0.2, c: 0x2f6bff },
        { w: 0.75, y: 0.3, c: C.cyan }
      ];
      const meshes = platforms.map(function (p) {
        const m = new THREE.Mesh(new THREE.BoxGeometry(p.w, 0.14, p.w), mat(p.c, { metalness: 0.7, roughness: 0.25, emissive: p.c, emissiveIntensity: 0.2 }));
        m.position.y = p.y;
        g.add(m);
        return m;
      });
      const beacon = new THREE.Mesh(new THREE.SphereGeometry(0.16, 20, 20), mat(C.gold, { emissive: C.gold, emissiveIntensity: 1.3 }));
      beacon.position.y = 0.75;
      g.add(beacon);
      const beam = new THREE.Mesh(new THREE.CylinderGeometry(0.02, 0.02, 0.4, 12), mat(C.gold, { emissive: C.gold, emissiveIntensity: 0.7 }));
      beam.position.y = 0.53;
      g.add(beam);
      update = function (t) {
        beacon.scale.setScalar(1 + Math.sin(t * 2.2) * 0.12);
        meshes.forEach(function (m, i) {
          m.position.y += Math.sin(t * 0.9 + i) * 0.0006;
        });
      };
    } else if (key === "vlsi") {
      const board = new THREE.Mesh(new THREE.BoxGeometry(2.4, 0.1, 1.8), mat(0x061f45, { metalness: 0.5 }));
      board.position.y = -0.5;
      g.add(board);
      const pts = [];
      const traces = [
        [[-1.0, -0.4], [0, -0.4]],
        [[0, -0.4], [0, 0.4]],
        [[0, 0.4], [1.0, 0.4]],
        [[-1.0, 0.2], [-0.4, 0.2], [-0.4, -0.1], [0.4, -0.1], [0.4, 0.6], [1.0, 0.6]]
      ];
      traces.forEach(function (tr) {
        for (let i = 0; i < tr.length - 1; i++) {
          pts.push(new THREE.Vector3(tr[i][0], -0.43, tr[i][1]), new THREE.Vector3(tr[i + 1][0], -0.43, tr[i + 1][1]));
        }
      });
      g.add(lines(pts, C.cyan, 0.6));
      const chips = [[-1.0, -0.4], [0, 0.4], [1.0, 0.4], [1.0, 0.6]].map(function (p) {
        const c = new THREE.Mesh(new THREE.BoxGeometry(0.32, 0.14, 0.32), mat(C.gold, { emissive: C.gold, emissiveIntensity: 0.35, metalness: 0.8 }));
        c.position.set(p[0], -0.38, p[1]);
        g.add(c);
        return c;
      });
      const pulses = [];
      for (let i = 0; i < 5; i++) {
        const p = new THREE.Mesh(new THREE.SphereGeometry(0.045, 12, 12), new THREE.MeshBasicMaterial({ color: 0xe8ffff }));
        p.userData.off = i / 5;
        pulses.push(p);
        g.add(p);
      }
      update = function (t) {
        chips.forEach(function (c, i) {
          c.material.emissiveIntensity = 0.3 + (Math.sin(t * 2.6 + i * 1.4) + 1) * 0.25;
        });
        pulses.forEach(function (p) {
          const k = (t * 0.35 + p.userData.off) % 1;
          p.position.set(-1.1 + k * 2.2, -0.36, Math.sin(k * Math.PI * 2) * 0.35);
        });
      };
    } else if (key === "mems") {
      const base = new THREE.Mesh(new THREE.BoxGeometry(1.9, 0.18, 0.9), mat(0x0b2a58, { metalness: 0.7 }));
      base.position.y = -0.55;
      g.add(base);
      const anchor = new THREE.Mesh(new THREE.BoxGeometry(0.18, 0.4, 0.4), mat(0x2f6bff, { emissive: 0x123a72, emissiveIntensity: 0.5 }));
      anchor.position.set(-0.85, -0.3, 0);
      g.add(anchor);
      const arm = new THREE.Group();
      arm.position.set(-0.85, -0.1, 0);
      const beam = new THREE.Mesh(new THREE.BoxGeometry(1.35, 0.08, 0.14), mat(C.cyan, { emissive: C.cyan, emissiveIntensity: 0.35, metalness: 0.5 }));
      beam.position.x = 0.68;
      arm.add(beam);
      const mass = new THREE.Mesh(new THREE.BoxGeometry(0.34, 0.28, 0.34), mat(C.gold, { emissive: C.gold, emissiveIntensity: 0.3, metalness: 0.8 }));
      mass.position.x = 1.32;
      arm.add(mass);
      g.add(arm);
      // comb fingers on the proof mass + fixed electrodes
      const combGeo = new THREE.BoxGeometry(0.16, 0.05, 0.035);
      for (let i = 0; i < 8; i++) {
        const f = new THREE.Mesh(combGeo, mat(C.gold, { emissive: C.gold, emissiveIntensity: 0.35, metalness: 0.8 }));
        f.position.set(1.42, -0.05, -0.28 + i * 0.08);
        arm.add(f);
        const fx = new THREE.Mesh(combGeo, mat(C.cyan, { emissive: C.cyan, emissiveIntensity: 0.3, metalness: 0.6 }));
        fx.position.set(1.56, -0.05, -0.24 + i * 0.08);
        arm.add(fx);
      }
      update = function (t) {
        arm.rotation.z = Math.sin(t * 2.1) * 0.085;
        mass.material.emissiveIntensity = 0.25 + (Math.sin(t * 2.1 + 1) + 1) * 0.2;
      };
    } else if (key === "photonics") {
      const guide = new THREE.Mesh(
        new THREE.BoxGeometry(2.6, 0.42, 0.5),
        new THREE.MeshPhysicalMaterial({ color: 0x2f6bff, metalness: 0.1, roughness: 0.1, transparent: true, opacity: 0.22, transmission: 0.6 })
      );
      g.add(guide);
      const core = new THREE.Mesh(new THREE.BoxGeometry(2.6, 0.1, 0.14), mat(C.cyan, { emissive: C.cyan, emissiveIntensity: 1.1, metalness: 0.2 }));
      g.add(core);
      const pulses = [];
      for (let i = 0; i < 4; i++) {
        const p = new THREE.Mesh(new THREE.SphereGeometry(0.06, 14, 14), new THREE.MeshBasicMaterial({ color: 0xf3ffff }));
        p.userData.off = i / 4;
        pulses.push(p);
        g.add(p);
      }
      [1.3, -1.3].forEach(function (x) {
        const ring = new THREE.Mesh(new THREE.TorusGeometry(0.3, 0.02, 10, 48), mat(C.gold, { emissive: C.gold, emissiveIntensity: 0.9 }));
        ring.position.x = x;
        ring.rotation.y = Math.PI / 2;
        g.add(ring);
      });
      update = function (t) {
        pulses.forEach(function (p) {
          const k = (t * 0.5 + p.userData.off) % 1;
          p.position.set(-1.25 + k * 2.5, 0, 0);
          p.material.opacity = Math.sin(k * Math.PI);
          p.material.transparent = true;
        });
        core.material.emissiveIntensity = 0.9 + Math.sin(t * 3) * 0.2;
      };
    } else if (key === "failure") {
      const chip = new THREE.Mesh(new THREE.BoxGeometry(1.7, 0.22, 1.7), mat(0x0b2a58, { metalness: 0.8, roughness: 0.25 }));
      chip.position.y = -0.35;
      g.add(chip);
      const crackPts = [
        new THREE.Vector3(-0.7, -0.22, 0.2),
        new THREE.Vector3(-0.3, -0.22, 0.0),
        new THREE.Vector3(-0.05, -0.22, 0.32),
        new THREE.Vector3(0.3, -0.22, 0.1),
        new THREE.Vector3(0.65, -0.22, 0.35)
      ];
      const crack = lines(crackPts, C.red, 1);
      g.add(crack);
      const scan = new THREE.Mesh(
        new THREE.ConeGeometry(0.32, 0.9, 28, 1, true),
        new THREE.MeshBasicMaterial({ color: C.red, transparent: true, opacity: 0.18, side: THREE.DoubleSide, blending: THREE.AdditiveBlending, depthWrite: false })
      );
      scan.position.y = 0.45;
      g.add(scan);
      const spot = new THREE.Mesh(new THREE.CircleGeometry(0.1, 24), new THREE.MeshBasicMaterial({ color: 0xffdada }));
      spot.rotation.x = -Math.PI / 2;
      spot.position.y = -0.22;
      g.add(spot);
      update = function (t) {
        const x = Math.sin(t * 0.9) * 0.75;
        scan.position.x = x;
        spot.position.x = x;
        scan.material.opacity = 0.14 + Math.sin(t * 3) * 0.05;
        crack.material.opacity = 0.65 + Math.sin(t * 2.4) * 0.35;
      };
    } else if (key === "supply") {
      const globe = new THREE.Mesh(
        new THREE.SphereGeometry(1.25, 26, 18),
        new THREE.MeshBasicMaterial({ color: C.cyan, wireframe: true, transparent: true, opacity: 0.2 })
      );
      g.add(globe);
      const inner = new THREE.Mesh(new THREE.SphereGeometry(1.1, 32, 24), mat(0x061f45, { metalness: 0.2, transparent: true, opacity: 0.55 }));
      g.add(inner);
      const nodes = [];
      const pts = [];
      for (let i = 0; i < 7; i++) {
        const phi = Math.acos(2 * Math.random() - 1);
        const theta = Math.random() * Math.PI * 2;
        const v = new THREE.Vector3(
          Math.sin(phi) * Math.cos(theta),
          Math.cos(phi),
          Math.sin(phi) * Math.sin(theta)
        ).multiplyScalar(1.28);
        const n = new THREE.Mesh(new THREE.SphereGeometry(0.075, 14, 14), mat(i % 2 ? C.gold : C.cyan, { emissive: i % 2 ? C.gold : C.cyan, emissiveIntensity: 1, metalness: 0.2 }));
        n.position.copy(v);
        n.userData.base = v.clone();
        nodes.push(n);
        g.add(n);
        if (i > 0) pts.push(nodes[i - 1].position.clone().multiplyScalar(0.92), v.clone().multiplyScalar(0.92));
      }
      g.add(lines(pts, C.gold, 0.4));
      update = function (t) {
        nodes.forEach(function (n, i) {
          n.scale.setScalar(1 + Math.sin(t * 2 + i) * 0.18);
        });
      };
    } else {
      const ico = new THREE.Mesh(new THREE.IcosahedronGeometry(1.1, 1), new THREE.MeshBasicMaterial({ color: C.cyan, wireframe: true, transparent: true, opacity: 0.55 }));
      g.add(ico);
      update = function (t) {
        ico.rotation.y = t * 0.4;
        ico.rotation.x = t * 0.2;
      };
    }

    return { group: g, update: update };
  }

  /* =========================================================
     SINGLE REUSABLE VIEWER FOR THE COURSE SHEET
     ========================================================= */
  function createCourseViewer(canvas) {
    if (!window.THREE) return null;

    const renderer = new THREE.WebGLRenderer({ canvas: canvas, antialias: true, alpha: true });
    renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 1.5));
    renderer.outputEncoding = THREE.sRGBEncoding;
    renderer.toneMapping = THREE.ACESFilmicToneMapping;
    renderer.toneMappingExposure = window.IMEN && window.IMEN.kit ? 0.9 : 1.05;

    const scene = new THREE.Scene();
    const camera = new THREE.PerspectiveCamera(42, 1, 0.1, 40);
    const BASE_Z = window.IMEN && window.IMEN.kit ? 5.5 : 6.2; // closer camera: the hyper-real scenes fill the frame
    camera.position.set(0, BASE_Z * 0.26, BASE_Z);
    camera.lookAt(0, 0, 0);
    const kit = window.IMEN && window.IMEN.kit;
    if (kit) {
      // studio image-based lighting for realistic metal / glass reflections
      try { scene.environment = kit.makeEnv(renderer); } catch (err) { console.warn("env map unavailable", err); }
      scene.add(new THREE.HemisphereLight(0xbcd0ff, 0x0b1020, 0.35));
      const key = new THREE.DirectionalLight(0xffffff, 0.8);
      key.position.set(3.5, 6, 4);
      scene.add(key);
      const rim = new THREE.DirectionalLight(0x7aa8ff, 0.8);
      rim.position.set(-5, 2.5, -4);
      scene.add(rim);
    } else {
      addLights(scene);
    }

    let current = null;
    let raf = 0;
    const clock = new THREE.Clock();

    // interactive orbit on a pivot so per-scene animations keep working
    const pivot = new THREE.Group();
    scene.add(pivot);
    const orbit = attachOrbit(canvas, pivot, {
      touchAction: "pan-y",
      onZoom: function (z) {
        camera.position.z = BASE_Z * z;
      }
    });

    const ro = new ResizeObserver(function () {
      const w = Math.max(canvas.clientWidth, 1);
      const h = Math.max(canvas.clientHeight, 1);
      renderer.setSize(w, h, false);
      camera.aspect = w / h;
      camera.updateProjectionMatrix();
    });
    ro.observe(canvas);

    // phones can drop a WebGL context under memory pressure: recover instead of showing a dead canvas
    let currentKey = null;
    let disposed = false;
    function onContextLost(e) {
      e.preventDefault();
    }
    function onContextRestored() {
      if (disposed) return;
      try {
        if (kit) scene.environment = kit.makeEnv(renderer);
        if (currentKey) api.setCourse(currentKey);
      } catch (err) {
        console.warn("Could not restore the course scene:", err);
      }
    }
    canvas.addEventListener("webglcontextlost", onContextLost);
    canvas.addEventListener("webglcontextrestored", onContextRestored);

    function disposeCurrent() {
      if (!current) return;
      if (kit && current.real) {
        kit.disposeGroup(current.group);
        pivot.remove(current.group);
        scene.remove(current.group);
        current = null;
        return;
      }
      current.group.traverse(function (o) {
        if (o.geometry) o.geometry.dispose();
        if (o.material) {
          if (Array.isArray(o.material)) o.material.forEach(function (m) { m.dispose(); });
          else o.material.dispose();
        }
      });
      scene.remove(current.group);
      pivot.remove(current.group);
      current = null;
    }

    function frame() {
      raf = requestAnimationFrame(frame);
      const el = clock.getElapsedTime();
      if (current && !REDUCED) current.update(current.real ? el : el * 0.45);
      orbit.step(current && current.spin != null ? current.spin : 0.0025);
      renderer.render(scene, camera);
    }

    const api = {
      setCourse: function (key) {
        currentKey = key;
        disposeCurrent();
        current = buildScene(key);
        pivot.add(current.group);
        if (REDUCED) current.update(0);
        renderer.render(scene, camera);
      },
      stats: function () {
        const i = renderer.info;
        return { calls: i.render.calls, triangles: i.render.triangles, geometries: i.memory.geometries, textures: i.memory.textures };
      },
      start: function () {
        const w = Math.max(canvas.clientWidth, 1);
        const h = Math.max(canvas.clientHeight, 1);
        renderer.setSize(w, h, false);
        camera.aspect = w / h;
        camera.updateProjectionMatrix();
        if (!raf) frame();
      },
      dispose: function () {
        cancelAnimationFrame(raf);
        raf = 0;
        orbit.dispose();
        ro.disconnect();
        disposeCurrent();
        disposed = true;
        canvas.removeEventListener("webglcontextlost", onContextLost);
        canvas.removeEventListener("webglcontextrestored", onContextRestored);
        renderer.dispose();
        if (renderer.forceContextLoss) renderer.forceContextLoss();
      }
    };
    return api;
  }

  window.IMEN = window.IMEN || {};
  window.IMEN.initHeroScene = initHeroScene;
  window.IMEN.createCourseViewer = createCourseViewer;
})();
