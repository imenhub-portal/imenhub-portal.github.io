/* ==========================================================
   IMEN · M.Sc. Semiconductor Engineering
   Course illustration kit — shared helpers for the 14 hyper-real scenes.
   three.js r128 (cdnjs, pinned). Loaded BEFORE scenes.js and assets/courses/*.js

   A scene file registers itself like this:

     IMEN.courseScenes.mems = function (kit) {
       const g = new THREE.Group();
       ...build meshes with kit.M.* materials / kit.tex.* textures...
       return {
         group: g,
         update: function (t) { ... },   // t = elapsed REAL seconds
         // optional: fit: 2.05 (bounding-sphere radius the scene is scaled to)
         // optional: spin: 0.0025 (idle auto-rotation per frame, 0 = none)
       };
     };
   ========================================================== */
(function () {
  "use strict";
  window.IMEN = window.IMEN || {};
  const T = THREE;
  const kit = {};
  window.IMEN.kit = kit;
  window.IMEN.courseScenes = window.IMEN.courseScenes || {};

  /* ---------------- random + noise ---------------- */
  kit.rng = function (seed) {
    let s = (seed || 1) >>> 0 || 1;
    return function () {
      s = (s * 1664525 + 1013904223) >>> 0;
      return s / 4294967296;
    };
  };
  function hash2(x, y) {
    let h = (x * 374761393 + y * 668265263) | 0;
    h = (h ^ (h >>> 13)) * 1274126177;
    return ((h ^ (h >>> 16)) >>> 0) / 4294967296;
  }
  kit.noise2 = function (x, y) {
    const xi = Math.floor(x), yi = Math.floor(y);
    const xf = x - xi, yf = y - yi;
    const u = xf * xf * (3 - 2 * xf), v = yf * yf * (3 - 2 * yf);
    const a = hash2(xi, yi), b = hash2(xi + 1, yi), c = hash2(xi, yi + 1), d = hash2(xi + 1, yi + 1);
    return a + (b - a) * u + (c - a) * v + (a - b - c + d) * u * v;
  };
  kit.fbm = function (x, y, oct) {
    let a = 0.5, f = 1, s = 0;
    for (let i = 0; i < (oct || 4); i++) {
      s += a * kit.noise2(x * f, y * f);
      f *= 2;
      a *= 0.5;
    }
    return s;
  };

  /* ---------------- canvas textures ---------------- */
  kit.canvas = function (w, h) {
    const c = document.createElement("canvas");
    c.width = w;
    c.height = h;
    return c;
  };
  // draw(ctx, w, h). opts: {linear:true} for bump/roughness maps, {repeat:[x,y]}
  kit.canvasTex = function (w, h, draw, opts) {
    opts = opts || {};
    const c = kit.canvas(w, h);
    draw(c.getContext("2d"), w, h);
    const t = new T.CanvasTexture(c);
    t.encoding = opts.linear ? T.LinearEncoding : T.sRGBEncoding;
    t.anisotropy = 8;
    if (opts.repeat) {
      t.wrapS = t.wrapT = T.RepeatWrapping;
      t.repeat.set(opts.repeat[0], opts.repeat[1]);
    }
    return t;
  };

  kit.tex = {};
  // fine brushed-metal streaks (use as roughnessMap / bumpMap with linear:true)
  kit.tex.brushed = function (w, h, opts) {
    opts = opts || {};
    const r = kit.rng(opts.seed || 7);
    return kit.canvasTex(w || 512, h || 512, function (g, W, H) {
      g.fillStyle = "#808080";
      g.fillRect(0, 0, W, H);
      for (let i = 0; i < 2600; i++) {
        const y = r() * H, len = 40 + r() * W * 0.7, x = r() * W;
        const v = 90 + r() * 110;
        g.strokeStyle = "rgba(" + v + "," + v + "," + v + ",0.35)";
        g.lineWidth = 0.6 + r() * 1.2;
        g.beginPath();
        g.moveTo(x, y);
        g.lineTo(x + len, y + (r() - 0.5) * 2);
        g.stroke();
      }
    }, { linear: true, repeat: opts.repeat || [1, 1] });
  };
  // fine grain (use as bump map)
  kit.tex.grain = function (w, h, seed, repeat) {
    const r = kit.rng(seed || 3);
    return kit.canvasTex(w || 256, h || 256, function (g, W, H) {
      const img = g.createImageData(W, H);
      for (let i = 0; i < W * H; i++) {
        const v = 110 + r() * 60;
        img.data[i * 4] = img.data[i * 4 + 1] = img.data[i * 4 + 2] = v;
        img.data[i * 4 + 3] = 255;
      }
      g.putImageData(img, 0, 0);
    }, { linear: true, repeat: repeat || [4, 4] });
  };
  // IC die floorplan: standard-cell rows, SRAM blocks, bus routing. palette = [bg, a, b, c, metal]
  kit.tex.dieLayout = function (size, seed, palette) {
    const r = kit.rng(seed || 11);
    const pal = palette || ["#0a1430", "#2f6bff", "#18c0ff", "#ffb340", "#c8d6ff"];
    return kit.canvasTex(size || 1024, size || 1024, function (g, W, H) {
      g.fillStyle = pal[0];
      g.fillRect(0, 0, W, H);
      // block placement
      const blocks = [];
      for (let i = 0; i < 14; i++) {
        const bw = W * (0.1 + r() * 0.22), bh = H * (0.08 + r() * 0.2);
        const bx = r() * (W - bw), by = r() * (H - bh);
        blocks.push([bx, by, bw, bh]);
      }
      blocks.forEach(function (b, i) {
        const col = pal[1 + (i % 3)];
        g.fillStyle = col;
        g.globalAlpha = 0.22;
        g.fillRect(b[0], b[1], b[2], b[3]);
        g.globalAlpha = 0.9;
        g.strokeStyle = col;
        g.lineWidth = 2;
        g.strokeRect(b[0], b[1], b[2], b[3]);
        // fine cell rows inside
        g.globalAlpha = 0.5;
        g.lineWidth = 1;
        const step = 5 + (i % 3) * 2;
        for (let y = b[1] + step; y < b[1] + b[3]; y += step) {
          g.beginPath();
          g.moveTo(b[0] + 2, y);
          g.lineTo(b[0] + b[2] - 2, y);
          g.stroke();
        }
        for (let x = b[0] + step * 2; x < b[0] + b[2]; x += step * 6) {
          g.globalAlpha = 0.35;
          g.beginPath();
          g.moveTo(x, b[1] + 2);
          g.lineTo(x, b[1] + b[3] - 2);
          g.stroke();
        }
      });
      g.globalAlpha = 1;
      // orthogonal routing (two metal layers)
      for (let layer = 0; layer < 2; layer++) {
        g.strokeStyle = layer ? pal[4] : pal[3];
        g.globalAlpha = layer ? 0.55 : 0.7;
        g.lineWidth = layer ? 1.2 : 1.8;
        for (let i = 0; i < 90; i++) {
          let x = r() * W, y = r() * H;
          g.beginPath();
          g.moveTo(x, y);
          for (let s = 0; s < 3; s++) {
            if ((s + layer) % 2) x += (r() - 0.5) * W * 0.4;
            else y += (r() - 0.5) * H * 0.4;
            g.lineTo(x, y);
          }
          g.stroke();
        }
      }
      g.globalAlpha = 1;
    });
  };
  // PCB traces & pads on solder-mask green (colour map). Returns {map, bump}
  kit.tex.pcb = function (size, seed, base) {
    const S = size || 1024;
    const draw = function (bump) {
      return function (g, W, H) {
        const r = kit.rng(seed || 5);
        g.fillStyle = bump ? "#555" : base || "#0b5d3b";
        g.fillRect(0, 0, W, H);
        g.lineCap = "round";
        g.lineJoin = "round";
        for (let i = 0; i < 160; i++) {
          let x = r() * W, y = r() * H;
          g.strokeStyle = bump ? "#bbb" : "rgba(214,176,92,0.85)";
          g.lineWidth = bump ? 5 : 4;
          g.beginPath();
          g.moveTo(x, y);
          for (let s = 0; s < 4; s++) {
            const d = r();
            const len = 30 + r() * 140;
            if (d < 0.33) x += len;
            else if (d < 0.66) y += len;
            else { x += len * 0.7; y += len * 0.7; }
            g.lineTo(x, y);
          }
          g.stroke();
          g.fillStyle = bump ? "#eee" : "#e8cf86";
          g.beginPath();
          g.arc(x, y, bump ? 7 : 6.5, 0, 6.2832);
          g.fill();
        }
        // vias
        for (let i = 0; i < 260; i++) {
          g.fillStyle = bump ? "#ddd" : "#d6b05c";
          g.beginPath();
          g.arc(r() * W, r() * H, 3, 0, 6.2832);
          g.fill();
        }
      };
    };
    return {
      map: kit.canvasTex(S, S, draw(false)),
      bump: kit.canvasTex(S, S, draw(true), { linear: true })
    };
  };
  // soft radial gradient used for glows / contact shadows. stops: [[pos, 'rgba()'], ...]
  kit.tex.radial = function (stops, size) {
    return kit.canvasTex(size || 128, size || 128, function (g, W, H) {
      const grd = g.createRadialGradient(W / 2, H / 2, 0, W / 2, H / 2, W / 2);
      stops.forEach(function (s) { grd.addColorStop(s[0], s[1]); });
      g.fillStyle = grd;
      g.fillRect(0, 0, W, H);
    });
  };
  // vertical gradient (bottom -> top) for light shafts
  kit.tex.shaft = function (rgb) {
    const c = rgb || "255,230,160";
    return kit.canvasTex(8, 128, function (g, W, H) {
      const grd = g.createLinearGradient(0, H, 0, 0);
      grd.addColorStop(0, "rgba(" + c + ",0.95)");
      grd.addColorStop(0.5, "rgba(" + c + ",0.32)");
      grd.addColorStop(1, "rgba(" + c + ",0)");
      g.fillStyle = grd;
      g.fillRect(0, 0, W, H);
    });
  };
  // crisp text on a transparent canvas texture (labels, symbols). returns {tex, aspect}
  kit.tex.text = function (str, opts) {
    opts = opts || {};
    const px = opts.size || 128;
    const c = kit.canvas(8, 8);
    let g = c.getContext("2d");
    const font = (opts.weight || "700") + " " + px + "px " + (opts.font || "Inter, system-ui, sans-serif");
    g.font = font;
    const w = Math.ceil(g.measureText(str).width) + px * 0.4;
    const h = Math.ceil(px * 1.35);
    c.width = w;
    c.height = h;
    g = c.getContext("2d");
    g.font = font;
    g.fillStyle = opts.color || "#ffffff";
    g.textAlign = "center";
    g.textBaseline = "middle";
    g.fillText(str, w / 2, h / 2 + px * 0.04);
    const t = new T.CanvasTexture(c);
    t.encoding = T.sRGBEncoding;
    t.anisotropy = 8;
    return { tex: t, aspect: w / h };
  };

  /* ---------------- studio environment (reflections for metals / glass) ---------------- */
  kit.makeEnv = function (renderer) {
    const sc = new T.Scene();
    const dome = new T.Mesh(
      new T.SphereGeometry(50, 32, 16),
      new T.MeshBasicMaterial({
        side: T.BackSide,
        map: kit.canvasTex(16, 256, function (g, w, h) {
          const grd = g.createLinearGradient(0, 0, 0, h);
          grd.addColorStop(0, "#7a92d6");
          grd.addColorStop(0.42, "#233463");
          grd.addColorStop(0.55, "#101a3a");
          grd.addColorStop(1, "#04060e");
          g.fillStyle = grd;
          g.fillRect(0, 0, w, h);
        })
      })
    );
    sc.add(dome);
    function panel(w, h, x, y, z, inten, color) {
      const m = new T.Mesh(
        new T.PlaneGeometry(w, h),
        new T.MeshBasicMaterial({ color: new T.Color(color).multiplyScalar(inten), side: T.DoubleSide })
      );
      m.position.set(x, y, z);
      m.lookAt(0, 0, 0);
      sc.add(m);
    }
    panel(34, 20, 0, 42, 8, 7, "#ffffff");      // big overhead softbox
    panel(14, 34, -40, 8, 14, 5.5, "#d6e4ff");  // cool key strip
    panel(10, 26, 38, 6, -4, 4.5, "#ffe6c4");   // warm kicker
    panel(46, 5, 0, -4, -46, 3.5, "#8fb8ff");   // back rim strip
    panel(12, 12, 20, 22, 36, 3, "#ffffff");    // front fill
    panel(24, 3, -10, -22, 30, 1.6, "#6fa0ff"); // low bounce
    const pm = new T.PMREMGenerator(renderer);
    const rt = pm.fromScene(sc, 0.02);
    pm.dispose();
    dome.geometry.dispose();
    dome.material.map.dispose();
    dome.material.dispose();
    return rt.texture;
  };

  /* ---------------- materials ---------------- */
  function phys(o) {
    const m = new T.MeshPhysicalMaterial(o);
    // dielectrics get less image-based light so dark materials (PCB, epoxy) keep their colour
    m.envMapIntensity = o.envMapIntensity == null ? (o.metalness > 0.5 ? 1.15 : 0.55) : o.envMapIntensity;
    return m;
  }
  kit.M = {
    metal: function (color, rough) { return phys({ color: color == null ? 0xc8ced8 : color, metalness: 1, roughness: rough == null ? 0.22 : rough }); },
    brushed: function (color, rough) {
      const tx = kit.tex.brushed(512, 512, { repeat: [2, 2] });
      return phys({ color: color == null ? 0xb9c0cc : color, metalness: 1, roughness: rough == null ? 0.38 : rough, roughnessMap: tx, bumpMap: tx, bumpScale: 0.4 });
    },
    gold: function () { return phys({ color: 0xf3c35a, metalness: 1, roughness: 0.2 }); },
    copper: function () { return phys({ color: 0xd98a5f, metalness: 1, roughness: 0.26 }); },
    aluminium: function () { return phys({ color: 0xdfe5ee, metalness: 1, roughness: 0.3 }); },
    solder: function () { return phys({ color: 0xcfd3da, metalness: 1, roughness: 0.16 }); },
    silicon: function (color) { return phys({ color: color == null ? 0x2b3556 : color, metalness: 0.85, roughness: 0.16, clearcoat: 0.7, clearcoatRoughness: 0.12 }); },
    epoxy: function (color) { return phys({ color: color == null ? 0x14161b : color, metalness: 0.1, roughness: 0.52, clearcoat: 0.2, clearcoatRoughness: 0.4 }); },
    plastic: function (color, rough) { return phys({ color: color == null ? 0xffffff : color, metalness: 0, roughness: rough == null ? 0.38 : rough, clearcoat: 0.5, clearcoatRoughness: 0.2 }); },
    pcb: function (color) { return phys({ color: color == null ? 0x0f6a42 : color, metalness: 0, roughness: 0.42, clearcoat: 0.6, clearcoatRoughness: 0.25 }); },
    ceramic: function (color) { return phys({ color: color == null ? 0xe9e4d6 : color, metalness: 0, roughness: 0.55, clearcoat: 0.25 }); },
    glass: function (tint, opacity) {
      return phys({
        color: tint == null ? 0xcfe9ff : tint, metalness: 0, roughness: 0.04, transparent: true,
        opacity: opacity == null ? 0.22 : opacity, clearcoat: 1, clearcoatRoughness: 0.03,
        envMapIntensity: 1.8, side: T.DoubleSide, depthWrite: false
      });
    },
    // self-lit colour (LEDs, light paths, glowing traces)
    glow: function (color, intensity) {
      return new T.MeshStandardMaterial({ color: 0x000000, emissive: color, emissiveIntensity: intensity == null ? 2 : intensity, roughness: 1, metalness: 0 });
    },
    additive: function (color, opacity, map) {
      return new T.MeshBasicMaterial({ color: color, map: map || null, transparent: true, opacity: opacity == null ? 0.5 : opacity, depthWrite: false, blending: T.AdditiveBlending, toneMapped: false, side: T.DoubleSide });
    }
  };

  /* ---------------- geometry helpers ---------------- */
  // box with softened edges (realistic chip packages, PCBs, housings)
  kit.roundedBox = function (w, h, d, r, bevel) {
    r = Math.min(r == null ? 0.04 : r, Math.min(w, d) / 2 - 0.001);
    bevel = bevel == null ? Math.min(0.02, h / 4) : bevel;
    const s = new T.Shape();
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
    const g = new T.ExtrudeGeometry(s, { depth: Math.max(0.001, h - bevel * 2), bevelEnabled: true, bevelThickness: bevel, bevelSize: bevel * 0.8, bevelSegments: 3, curveSegments: 6 });
    g.rotateX(-Math.PI / 2);          // extrude along +Y
    g.translate(0, bevel, 0);
    g.translate(0, -h / 2 + 0, 0);
    g.computeBoundingBox();
    return g;
  };
  // tube along a polyline / curve
  kit.tube = function (pts, radius, material, segs) {
    const curve = pts.isCurve ? pts : new T.CatmullRomCurve3(pts);
    const m = new T.Mesh(new T.TubeGeometry(curve, segs || 64, radius, 10, false), material);
    return m;
  };
  // additive glow sprite
  kit.glow = function (color, size, opacity) {
    const s = new T.Sprite(new T.SpriteMaterial({
      map: kit.tex.radial([[0, "rgba(255,255,255,1)"], [0.25, "rgba(255,255,255,0.45)"], [1, "rgba(255,255,255,0)"]]),
      color: color == null ? 0x6aa8ff : color, transparent: true, opacity: opacity == null ? 0.9 : opacity,
      depthWrite: false, blending: T.AdditiveBlending, toneMapped: false
    }));
    s.scale.setScalar(size || 1);
    s.userData.noFit = true;
    return s;
  };
  // light beam / cone (additive). axis = +Y, base at y=0
  kit.beam = function (rBottom, rTop, length, rgb, opacity) {
    const m = new T.Mesh(
      new T.CylinderGeometry(rTop, rBottom, length, 40, 1, true),
      new T.MeshBasicMaterial({ map: kit.tex.shaft(rgb), transparent: true, opacity: opacity == null ? 0.5 : opacity, depthWrite: false, blending: T.AdditiveBlending, toneMapped: false, side: T.DoubleSide })
    );
    m.position.y = length / 2;
    m.userData.noFit = true;
    return m;
  };
  // fake soft contact shadow lying on y = 0
  kit.contactShadow = function (radius, opacity) {
    const m = new T.Mesh(
      new T.PlaneGeometry(radius * 2, radius * 2),
      new T.MeshBasicMaterial({ map: kit.tex.radial([[0, "rgba(0,0,0,0.75)"], [0.55, "rgba(0,0,0,0.3)"], [1, "rgba(0,0,0,0)"]]), transparent: true, opacity: opacity == null ? 0.6 : opacity, depthWrite: false })
    );
    m.rotation.x = -Math.PI / 2;
    m.userData.noFit = true;
    return m;
  };
  // flat text label plane (faces +Z)
  kit.label = function (str, height, opts) {
    const t = kit.tex.text(str, opts);
    const m = new T.Mesh(
      new T.PlaneGeometry(height * t.aspect, height),
      new T.MeshBasicMaterial({ map: t.tex, transparent: true, depthWrite: false, toneMapped: false })
    );
    return m;
  };
  // InstancedMesh convenience: setter(i, dummyObject3D) positions each instance
  kit.instanced = function (geo, material, count, setter) {
    const im = new T.InstancedMesh(geo, material, count);
    const d = new T.Object3D();
    for (let i = 0; i < count; i++) {
      d.position.set(0, 0, 0);
      d.rotation.set(0, 0, 0);
      d.scale.set(1, 1, 1);
      setter(i, d);
      d.updateMatrix();
      im.setMatrixAt(i, d.matrix);
    }
    im.instanceMatrix.needsUpdate = true;
    return im;
  };

  /* ---------------- framing + cleanup ---------------- */
  // scale & centre a scene group so its bounding sphere has radius `radius`
  kit.fit = function (group, radius) {
    const box = new T.Box3();
    const tmp = new T.Box3();
    group.updateMatrixWorld(true);
    group.traverse(function (o) {
      if (o.userData && o.userData.noFit) return;
      if (!o.geometry) return;
      if (!o.geometry.boundingBox) o.geometry.computeBoundingBox();
      tmp.copy(o.geometry.boundingBox).applyMatrix4(o.matrixWorld);
      box.union(tmp);
    });
    if (box.isEmpty()) return group;
    const sphere = new T.Sphere();
    box.getBoundingSphere(sphere);
    const k = (radius || 2.05) / Math.max(sphere.radius, 0.001);
    group.position.sub(sphere.center);
    const holder = new T.Group();
    holder.add(group);
    holder.scale.setScalar(k);
    return holder;
  };
  kit.disposeGroup = function (root) {
    root.traverse(function (o) {
      if (o.geometry) o.geometry.dispose();
      const ms = o.material ? (Array.isArray(o.material) ? o.material : [o.material]) : [];
      ms.forEach(function (m) {
        ["map", "bumpMap", "roughnessMap", "normalMap", "emissiveMap", "alphaMap", "metalnessMap"].forEach(function (k) {
          if (m[k] && m[k].dispose) m[k].dispose();
        });
        m.dispose();
      });
    });
  };
})();
