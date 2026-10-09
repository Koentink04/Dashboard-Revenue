(function () {
  function waitTHREE(cb) {
    if (window.THREE) return cb();
    var t = setInterval(function () { if (window.THREE) { clearInterval(t); cb(); } }, 40);
    setTimeout(function () { clearInterval(t); }, 15000);
  }

  function labelSprite(THREE, text, color, size) {
    var dpr = 2, w = 256, h = 64;
    var c = document.createElement('canvas');
    c.width = w * dpr; c.height = h * dpr;
    var g = c.getContext('2d');
    g.scale(dpr, dpr);
    g.clearRect(0, 0, w, h);
    g.font = '600 34px "IBM Plex Mono", ui-monospace, monospace';
    g.fillStyle = color;
    g.textAlign = 'center';
    g.textBaseline = 'middle';
    g.fillText(text, w / 2, h / 2);
    var tex = new THREE.CanvasTexture(c);
    tex.minFilter = THREE.LinearFilter;
    var sp = new THREE.Sprite(new THREE.SpriteMaterial({ map: tex, transparent: true, depthWrite: false }));
    sp.scale.set(size * 4, size, 1);
    return sp;
  }

  function hexToNum(h) { return parseInt(String(h).replace('#', ''), 16); }

  class Rev3D extends HTMLElement {
    connectedCallback() {
      if (this._init) return;
      this._init = true;
      this.style.display = 'block';
      this.style.position = 'relative';
      this.style.width = '100%';
      this.style.height = '100%';
      var self = this;
      waitTHREE(function () { self._whenSized(function () { self._build(); }); });
    }
    _whenSized(cb) {
      var self = this, tries = 0;
      (function poll() {
        if (self.clientWidth > 0 && self.clientHeight > 0) return cb();
        if (++tries > 300) return;
        requestAnimationFrame(poll);
      })();
    }
    disconnectedCallback() {
      if (this._raf) cancelAnimationFrame(this._raf);
      if (this._renderer) this._renderer.dispose();
    }
    static get observedAttributes() { return ['data', 'accent', 'mode', 'motion']; }
    attributeChangedCallback() {
      var self = this;
      if (!this._init || !this._renderer) return;
      clearTimeout(this._reb);
      this._reb = setTimeout(function () { self._teardown(); self._build(); }, 60);
    }

    _teardown() {
      if (this._raf) cancelAnimationFrame(this._raf);
      this._raf = null;
      if (this._renderer) {
        try { this._renderer.forceContextLoss(); } catch (e) {}
        try { this._renderer.dispose(); } catch (e) {}
        this._renderer = null;
      }
      this.innerHTML = '';
    }

    _build() {
      var THREE = window.THREE;
      var mode = this.getAttribute('mode') || 'wave';
      var accent = this.getAttribute('accent') || '#3D7BD9';
      var motion = this.getAttribute('motion') !== 'off'
        && !window.matchMedia('(prefers-reduced-motion: reduce)').matches;
      var data = [];
      try { data = JSON.parse(this.getAttribute('data') || '[]'); } catch (e) { data = []; }

      var w = this.clientWidth, h = this.clientHeight;
      var attrs = { alpha: true, antialias: false, depth: true, stencil: false, preserveDrawingBuffer: true, premultipliedAlpha: true, powerPreference: 'default' };
      var canvas = document.createElement('canvas');
      var ctx = canvas.getContext('webgl2', attrs) || canvas.getContext('webgl', attrs);
      if (!ctx) { this.setAttribute('data-webgl', 'unavailable'); return; }
      var renderer = new THREE.WebGLRenderer(Object.assign({ canvas: canvas, context: ctx }, attrs));
      renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
      renderer.setSize(w, h);
      canvas.style.display = 'block';
      this.appendChild(canvas);
      this._renderer = renderer;

      var scene = new THREE.Scene();
      var camera = new THREE.PerspectiveCamera(38, w / h, 0.1, 400);
      var group = new THREE.Group();
      scene.add(group);

      var accentNum = hexToNum(accent);
      var tick = function () {};

      if (mode === 'wave') {
        camera.position.set(0, 9, 34);
        camera.lookAt(0, 0, 0);
        var seg = 54;
        var geo = new THREE.PlaneGeometry(96, 46, seg, Math.round(seg * 0.5));
        var mat = new THREE.MeshBasicMaterial({ color: accentNum, wireframe: true, transparent: true, opacity: 0.32 });
        var mesh = new THREE.Mesh(geo, mat);
        mesh.rotation.x = -Math.PI / 2.35;
        mesh.position.y = -4;
        group.add(mesh);

        var glowGeo = new THREE.PlaneGeometry(96, 46, seg, Math.round(seg * 0.5));
        var glow = new THREE.Mesh(glowGeo, new THREE.MeshBasicMaterial({ color: 0x8FB6F0, wireframe: true, transparent: true, opacity: 0.07 }));
        glow.rotation.x = -Math.PI / 2.35;
        glow.position.set(0, -3.4, 0);
        group.add(glow);

        var pGeo = new THREE.BufferGeometry();
        var N = 260, pos = new Float32Array(N * 3);
        for (var i = 0; i < N; i++) {
          pos[i * 3] = (Math.random() - 0.5) * 90;
          pos[i * 3 + 1] = Math.random() * 20 - 2;
          pos[i * 3 + 2] = (Math.random() - 0.5) * 40;
        }
        pGeo.setAttribute('position', new THREE.BufferAttribute(pos, 3));
        var pts = new THREE.Points(pGeo, new THREE.PointsMaterial({ color: 0xBFD4F2, size: 0.16, transparent: true, opacity: 0.5 }));
        group.add(pts);

        var base = geo.attributes.position.array.slice();
        var gbase = glowGeo.attributes.position.array.slice();
        tick = function (t) {
          var a = geo.attributes.position.array, b = glowGeo.attributes.position.array;
          for (var i = 0; i < a.length; i += 3) {
            var x = base[i], y = base[i + 1];
            var z = Math.sin(x * 0.16 + t * 0.55) * 1.5 + Math.sin(y * 0.21 - t * 0.4) * 1.15 + Math.sin((x + y) * 0.09 + t * 0.28) * 1.1;
            a[i + 2] = z;
            b[i + 2] = z * 1.25;
          }
          geo.attributes.position.needsUpdate = true;
          glowGeo.attributes.position.needsUpdate = true;
          pts.rotation.y = t * 0.02;
          group.rotation.z = Math.sin(t * 0.1) * 0.012;
        };
      } else {
        camera.position.set(0, 17, 39);
        camera.lookAt(0, 1.5, 0);
        scene.add(new THREE.AmbientLight(0xffffff, 0.72));
        var d1 = new THREE.DirectionalLight(0xffffff, 0.85); d1.position.set(12, 26, 18); scene.add(d1);
        var d2 = new THREE.DirectionalLight(0x9EC0F0, 0.35); d2.position.set(-16, 10, -12); scene.add(d2);

        var grid = new THREE.GridHelper(72, 24, 0x2A4676, 0x1B2E4E);
        grid.position.y = -0.02;
        group.add(grid);

        var n = Math.max(data.length, 1);
        var step = 4.4, spanX = (n - 1) * step;
        var maxV = 1;
        data.forEach(function (d) { maxV = Math.max(maxV, d.target || 0, d.real || 0); });
        var H = 15;
        var bars = [];

        data.forEach(function (d, idx) {
          var x = -spanX / 2 + idx * step;
          var tH = (d.target || 0) / maxV * H;
          var rH = (d.real || 0) / maxV * H;
          var cH = (d.cash || 0) / maxV * H;

          var ghost = new THREE.Mesh(
            new THREE.BoxGeometry(2.5, tH, 2.5),
            new THREE.MeshBasicMaterial({ color: 0x7C95C0, transparent: true, opacity: 0.13 })
          );
          ghost.position.set(x, tH / 2, 0);
          group.add(ghost);
          var edge = new THREE.LineSegments(
            new THREE.EdgesGeometry(new THREE.BoxGeometry(2.5, tH, 2.5)),
            new THREE.LineBasicMaterial({ color: 0x9CB4DA, transparent: true, opacity: 0.3 })
          );
          edge.position.copy(ghost.position);
          group.add(edge);

          var future = (d.real || 0) === 0;
          if (!future) {
            var realMat = new THREE.MeshLambertMaterial({ color: accentNum });
            var real = new THREE.Mesh(new THREE.BoxGeometry(1.75, 1, 1.75), realMat);
            real.position.set(x, 0, 0);
            real.scale.y = 0.001;
            group.add(real);
            bars.push({ m: real, h: rH, delay: idx * 0.055, y0: 0 });

            var cash = new THREE.Mesh(new THREE.BoxGeometry(1.9, 1, 1.9), new THREE.MeshLambertMaterial({ color: 0x2E9E5B }));
            cash.position.set(x, 0, 0);
            cash.scale.y = 0.001;
            group.add(cash);
            bars.push({ m: cash, h: cH, delay: idx * 0.055 + 0.12, y0: 0 });

            var cap = new THREE.Mesh(
              new THREE.BoxGeometry(2.05, 0.16, 2.05),
              new THREE.MeshBasicMaterial({ color: (d.real >= d.target) ? 0x3FBF77 : 0xE0566E })
            );
            cap.position.set(x, rH + 0.1, 0);
            group.add(cap);
            cap.visible = false;
            bars.push({ cap: cap, delay: idx * 0.055 + 0.5 });
          }

          var lab = labelSprite(THREE, d.label, future ? '#5E739A' : '#C6D6EE', 1.35);
          lab.position.set(x, -0.6, 3.4);
          group.add(lab);
        });

        var t0 = null;
        tick = function (t) {
          if (t0 === null) t0 = t;
          var e = t - t0;
          bars.forEach(function (b) {
            if (b.cap) { if (e > b.delay) b.cap.visible = true; return; }
            var p = Math.min(1, Math.max(0, (e - b.delay) / 0.9));
            p = 1 - Math.pow(1 - p, 3);
            b.m.scale.y = Math.max(0.001, b.h * p);
            b.m.position.y = (b.h * p) / 2;
          });
          group.rotation.y = Math.sin(t * 0.13) * 0.16;
        };
      }

      var start = performance.now();
      var self = this;
      var lastW = w, lastH = h;
      function fit() {
        var W = self.clientWidth, H2 = self.clientHeight;
        if (!W || !H2) return;
        if (W === lastW && H2 === lastH) return;
        lastW = W; lastH = H2;
        renderer.setSize(W, H2);
        camera.aspect = W / H2;
        camera.updateProjectionMatrix();
      }
      function loop(now) {
        var t = (now - start) / 1000;
        fit();
        tick(motion ? t : 1.6);
        renderer.render(scene, camera);
        self._raf = requestAnimationFrame(loop);
      }
      this._raf = requestAnimationFrame(loop);
    }
  }

  if (!customElements.get('rev-3d')) customElements.define('rev-3d', Rev3D);
})();
