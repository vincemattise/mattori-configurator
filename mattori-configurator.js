// Minimal OrbitControls implementation (rotate + zoom + pan)
  // Pan: right/middle-click, Shift+drag (Apple mouse), or two-finger touch
  // Pinch-to-zoom on touch devices
  THREE.OrbitControls = class {
    constructor(camera, domElement) {
      this.camera = camera;
      this.domElement = domElement;
      this.target = new THREE.Vector3();
      this.enableDamping = false;
      this.dampingFactor = 0.1;
      this.rotateSpeed = 1.0;
      this.zoomSpeed = 1.0;
      this.panSpeed = 1.0;
      this.minDistance = 0;
      this.maxDistance = Infinity;
      this.maxPolarAngle = Math.PI;
      this.minPolarAngle = 0;

      // Internal state (spherical coords relative to target)
      this._spherical = new THREE.Spherical();
      this._sphericalDelta = new THREE.Spherical();
      this._panOffset = new THREE.Vector3();
      this._zoomScale = 1;
      this._rotateStart = new THREE.Vector2();
      this._panStart = new THREE.Vector2();
      this._pointerType = null; // 'rotate' | 'pan'

      // Multi-touch state
      this._pointers = new Map();       // pointerId → {x, y}
      this._prevTouchCenter = null;     // {x, y}
      this._prevPinchDist = null;       // number

      // Initialize spherical from current camera position
      const offset = new THREE.Vector3().subVectors(camera.position, this.target);
      this._spherical.setFromVector3(offset);

      this._onPointerDown = this._onPointerDown.bind(this);
      this._onPointerMove = this._onPointerMove.bind(this);
      this._onPointerUp = this._onPointerUp.bind(this);
      this._onWheel = this._onWheel.bind(this);
      this._onContextMenu = e => e.preventDefault();

      domElement.style.touchAction = 'none'; // prevent browser gestures
      domElement.addEventListener('pointerdown', this._onPointerDown);
      domElement.addEventListener('wheel', this._onWheel, { passive: false });
      domElement.addEventListener('contextmenu', this._onContextMenu);
    }

    _applyPan(dx, dy) {
      const el = this.domElement;
      const offset = new THREE.Vector3().subVectors(this.camera.position, this.target);
      let targetDist = offset.length();
      targetDist *= Math.tan((this.camera.fov / 2) * Math.PI / 180);
      const panX = 2 * dx * targetDist / el.clientHeight * this.panSpeed;
      const panY = 2 * dy * targetDist / el.clientHeight * this.panSpeed;
      const camLeft = new THREE.Vector3().setFromMatrixColumn(this.camera.matrix, 0);
      const camUp = new THREE.Vector3().setFromMatrixColumn(this.camera.matrix, 1);
      this._panOffset.addScaledVector(camLeft, -panX);
      this._panOffset.addScaledVector(camUp, panY);
    }

    _getTouchCenter() {
      const pts = Array.from(this._pointers.values());
      return {
        x: (pts[0].x + pts[1].x) / 2,
        y: (pts[0].y + pts[1].y) / 2
      };
    }

    _getPinchDist() {
      const pts = Array.from(this._pointers.values());
      const dx = pts[0].x - pts[1].x;
      const dy = pts[0].y - pts[1].y;
      return Math.sqrt(dx * dx + dy * dy);
    }

    _onPointerDown(e) {
      this.domElement.setPointerCapture(e.pointerId);
      this._pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });

      if (this._pointers.size === 2) {
        // Second finger down → switch to pan+pinch
        this._pointerType = 'touch-pan';
        this._prevTouchCenter = this._getTouchCenter();
        this._prevPinchDist = this._getPinchDist();
        return;
      }

      if (this._pointers.size > 2) return; // ignore 3+ fingers

      // Single pointer: Shift+left = pan (Apple mouse), right/middle = pan, left = rotate
      if (e.button === 0 && !e.shiftKey) {
        this._pointerType = 'rotate';
        this._rotateStart.set(e.clientX, e.clientY);
      } else {
        this._pointerType = 'pan';
        this._panStart.set(e.clientX, e.clientY);
      }
      this.domElement.addEventListener('pointermove', this._onPointerMove);
      this.domElement.addEventListener('pointerup', this._onPointerUp);
    }

    _onPointerMove(e) {
      // Update tracked pointer position
      if (this._pointers.has(e.pointerId)) {
        this._pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
      }

      // Two-finger touch: pan + pinch zoom
      if (this._pointerType === 'touch-pan' && this._pointers.size === 2) {
        const center = this._getTouchCenter();
        const dist = this._getPinchDist();

        // Pan from center movement
        if (this._prevTouchCenter) {
          const dx = center.x - this._prevTouchCenter.x;
          const dy = center.y - this._prevTouchCenter.y;
          this._applyPan(dx, dy);
        }

        // Pinch zoom from distance change
        if (this._prevPinchDist && this._prevPinchDist > 0) {
          const scale = this._prevPinchDist / dist;
          this._zoomScale *= scale;
        }

        this._prevTouchCenter = center;
        this._prevPinchDist = dist;
        return;
      }

      if (this._pointerType === 'rotate') {
        const dx = e.clientX - this._rotateStart.x;
        const dy = e.clientY - this._rotateStart.y;
        const el = this.domElement;
        this._sphericalDelta.theta -= 2 * Math.PI * dx / el.clientWidth * this.rotateSpeed;
        this._sphericalDelta.phi -= 2 * Math.PI * dy / el.clientHeight * this.rotateSpeed;
        this._rotateStart.set(e.clientX, e.clientY);
      } else if (this._pointerType === 'pan') {
        const dx = e.clientX - this._panStart.x;
        const dy = e.clientY - this._panStart.y;
        this._applyPan(dx, dy);
        this._panStart.set(e.clientX, e.clientY);
      }
    }

    _onPointerUp(e) {
      this.domElement.releasePointerCapture(e.pointerId);
      this._pointers.delete(e.pointerId);

      // If we were in two-finger mode and one finger lifts, reset
      if (this._pointerType === 'touch-pan') {
        this._prevTouchCenter = null;
        this._prevPinchDist = null;
        // If one finger remains, switch to rotate
        if (this._pointers.size === 1) {
          this._pointerType = 'rotate';
          const remaining = Array.from(this._pointers.values())[0];
          this._rotateStart.set(remaining.x, remaining.y);
        } else if (this._pointers.size === 0) {
          this._pointerType = null;
        }
        return;
      }

      if (this._pointers.size === 0) {
        this.domElement.removeEventListener('pointermove', this._onPointerMove);
        this.domElement.removeEventListener('pointerup', this._onPointerUp);
        this._pointerType = null;
      }
    }

    _onWheel(e) {
      e.preventDefault();
      if (e.deltaY > 0) {
        this._zoomScale *= Math.pow(0.95, this.zoomSpeed);
      } else {
        this._zoomScale /= Math.pow(0.95, this.zoomSpeed);
      }
    }

    update() {
      const offset = new THREE.Vector3().subVectors(this.camera.position, this.target);
      this._spherical.setFromVector3(offset);

      this._spherical.theta += this._sphericalDelta.theta;
      this._spherical.phi += this._sphericalDelta.phi;

      this._spherical.phi = Math.max(this.minPolarAngle, Math.min(this.maxPolarAngle, this._spherical.phi));
      this._spherical.makeSafe();

      if (this.camera.isOrthographicCamera) {
        const zoomFactor = this._zoomScale;
        this.camera.left *= zoomFactor;
        this.camera.right *= zoomFactor;
        this.camera.top *= zoomFactor;
        this.camera.bottom *= zoomFactor;
        this.camera.updateProjectionMatrix();
      } else {
        this._spherical.radius *= this._zoomScale;
        this._spherical.radius = Math.max(this.minDistance, Math.min(this.maxDistance, this._spherical.radius));
      }

      this.target.add(this._panOffset);

      offset.setFromSpherical(this._spherical);
      this.camera.position.copy(this.target).add(offset);
      this.camera.lookAt(this.target);

      if (this.enableDamping) {
        this._sphericalDelta.theta *= (1 - this.dampingFactor);
        this._sphericalDelta.phi *= (1 - this.dampingFactor);
      } else {
        this._sphericalDelta.set(0, 0, 0);
      }
      this._panOffset.set(0, 0, 0);
      this._zoomScale = 1;
    }

    dispose() {
      this.domElement.removeEventListener('pointerdown', this._onPointerDown);
      this.domElement.removeEventListener('wheel', this._onWheel);
      this.domElement.removeEventListener('contextmenu', this._onContextMenu);
      this.domElement.removeEventListener('pointermove', this._onPointerMove);
      this.domElement.removeEventListener('pointerup', this._onPointerUp);
    }
  };

// ============================================================
    // STATE
    // ============================================================
    let floors = [];
    let canvases = [];
    let maxWorldW = 1;
    let maxWorldH = 1;
    let originalFmlData = null;
    let originalFileName = '';

    // Force-load Nexa Bold font immediately so it's ready before any overlay appears
    var nexaFontReady = document.fonts.load('700 1em "Nexa Bold"');

    const CANVAS_W = 260;
    const CANVAS_H = 400;

    // Wizard state
    let currentWizardStep = 1;
    const TOTAL_WIZARD_STEPS = 5;
    let currentFloorReviewIndex = 0;
    var viewedFloors = new Set();
    let noFloorsMode = false;

    // Active viewers for cleanup
    let activeViewers = [];       // { renderer, controls, animId }
    let previewViewers = [];      // static thumbnail viewers in unified preview
    let floorReviewViewer = null; // step 4 active viewer
    let layoutViewers = [];       // step 5 viewers

    // ============================================================
    // DOM REFERENCES (initialized lazily via ensureDomRefs)
    // ============================================================
    var dropzone, fileInput, errorMsg, loadingOverlay, floorsGrid,
        btnExport, btnDownloadFml, toast, fileLabel, productHeroImage,
        unifiedFramePreview, frameStreet, frameCity, floorsLoading,
        unifiedFloorsOverlay, unifiedLabelsOverlay, unifiedAddressOverlay,
        wizard, wizardDots, wizardStepIndicator, btnWizardPrev, btnWizardNext,
        addressStreet, addressCity, labelsFields,
        floorReviewViewerEl, floorReviewLabel, floorIncludeCb, floorLayoutViewer,
        fundaUrlInput, btnFunda;

    var _domRefsReady = false;
    function ensureDomRefs() {
      if (_domRefsReady) return;
      _domRefsReady = true;
      dropzone = document.getElementById('dropzone');
      fileInput = document.getElementById('fileInput');
      errorMsg = document.getElementById('errorMsg');
      loadingOverlay = document.getElementById('loadingOverlay');
      floorsGrid = document.getElementById('floorsGrid');
      btnExport = document.getElementById('btnExport');
      btnDownloadFml = document.getElementById('btnDownloadFml');
      toast = document.getElementById('toast');
      fileLabel = document.getElementById('fileLabel');
      productHeroImage = document.getElementById('productHeroImage');
      unifiedFramePreview = document.getElementById('unifiedFramePreview');
      frameStreet = document.getElementById('frameStreet');
      frameCity = document.getElementById('frameCity');
      floorsLoading = document.getElementById('floorsLoading');
      unifiedFloorsOverlay = document.getElementById('unifiedFloorsOverlay');
      unifiedLabelsOverlay = document.getElementById('unifiedLabelsOverlay');
      unifiedAddressOverlay = document.getElementById('unifiedAddressOverlay');
      wizard = document.getElementById('wizard');
      wizardDots = document.getElementById('wizardDots');
      wizardStepIndicator = document.getElementById('wizardStepIndicator');
      btnWizardPrev = document.getElementById('btnWizardPrev');
      btnWizardNext = document.getElementById('btnWizardNext');
      addressStreet = document.getElementById('addressStreet');
      addressCity = document.getElementById('addressCity');
      labelsFields = document.getElementById('labelsFields');
      floorReviewViewerEl = document.getElementById('floorReviewViewer');
      floorReviewLabel = document.getElementById('floorReviewLabel');
      floorIncludeCb = document.getElementById('floorIncludeCb');
      floorLayoutViewer = document.getElementById('floorLayoutViewer');
      fundaUrlInput = document.getElementById('fundaUrl');
      btnFunda = document.getElementById('btnFunda');

      // Show admin panel only when ?admin=true is in the URL
      var adminPanel = document.getElementById('adminPanel');
      if (adminPanel) {
        var urlParams = new URLSearchParams(window.location.search);
        if (urlParams.get('admin') === 'true') {
          adminPanel.style.display = '';
        } else {
          adminPanel.style.display = 'none';
        }
      }
    }

    // ============================================================
    // BREAK OUT OF SHOPIFY CONTAINER (scrollbar-safe)
    // ============================================================
    (function breakoutConfigurator() {
      var el = document.querySelector('.mattori-configurator');
      if (!el) return;
      // Force overflow visible on all ancestors up to body
      var ancestor = el.parentElement;
      while (ancestor && ancestor !== document.body && ancestor !== document.documentElement) {
        ancestor.style.overflow = 'visible';
        ancestor = ancestor.parentElement;
      }
      // CSS sets width:100vw; JS corrects margin-left then reveals
      // Element starts visibility:hidden in HTML — no flash possible
      function applyBreakout() {
        var rect = el.getBoundingClientRect();
        el.style.marginLeft = (-rect.left) + 'px';
      }
      applyBreakout();
      el.style.visibility = '';
      el.classList.add('revealed');
      // Re-correct on resize
      window.addEventListener('resize', function() {
        el.style.marginLeft = '0px';
        requestAnimationFrame(function() {
          var rect = el.getBoundingClientRect();
          el.style.marginLeft = (-rect.left) + 'px';
        });
      });
    })();

    // ============================================================
    // FORCE RIGHT COLUMN LAYOUT (bulletproof against Shopify CSS)
    // ============================================================
    (function enforceRightColumnLayout() {
      ensureDomRefs();
      const inner = document.querySelector('.mattori-configurator .page-col-right-inner');
      if (!inner) return;
      inner.style.cssText += ';display:flex!important;flex-direction:column!important;gap:1.5rem!important;';
      // Physically reorder DOM children
      const orderMap = [
        '.product-title',
        '.product-price-block',
        '.product-description',
        '#btnStartConfigurator',
        '#wizard',
        '.product-badges',
        '.product-specs'
      ];
      const fragment = document.createDocumentFragment();
      orderMap.forEach(sel => {
        const el = inner.querySelector(sel);
        if (el) {
          el.style.cssText += ';margin:0!important;order:unset!important;position:static!important;z-index:auto!important;inset:auto!important;';
          fragment.appendChild(el);
        }
      });
      while (inner.firstChild) {
        inner.firstChild.style && (inner.firstChild.style.cssText += ';margin:0!important;');
        fragment.appendChild(inner.firstChild);
      }
      inner.appendChild(fragment);
    })();

    // ============================================================
    // UTILITY FUNCTIONS
    // ============================================================
    function clearError() { errorMsg.textContent = ''; }
    function setError(msg) { errorMsg.textContent = msg; }
    function showLoading() { loadingOverlay.classList.add('active'); }
    function hideLoading() { loadingOverlay.classList.remove('active'); }

    var toastTimer = null;
    function showToast(msg) {
      toast.textContent = msg;
      toast.classList.add('show');
      if (toastTimer) clearTimeout(toastTimer);
      toastTimer = setTimeout(() => toast.classList.remove('show'), 2200);
    }

    function median(arr) {
      if (!arr.length) return 0;
      const sorted = [...arr].sort((a, b) => a - b);
      const mid = Math.floor(sorted.length / 2);
      return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
    }

    // ============================================================
    // ADDRESS PARSING
    // ============================================================
    function parseFundaAddress(url) {
      try {
        const u = new URL(url);
        const segments = u.pathname.split('/').filter(Boolean);
        if (segments.length < 4) return null;

        const city = segments[2];
        const slug = segments[3];

        const prefixes = [
          '2-onder-1-kap', 'appartement', 'benedenwoning', 'bovenwoning', 'bungalow',
          'geschakeld', 'grachtenpand', 'herenhuis', 'hoekwoning', 'huis', 'landhuis',
          'maisonnette', 'penthouse', 'stacaravan', 'tussenwoning', 'villa',
          'vrijstaand', 'woonboot', 'woonhuis'
        ];

        let streetSlug = slug;
        for (const prefix of prefixes) {
          if (streetSlug.startsWith(prefix + '-')) {
            streetSlug = streetSlug.slice(prefix.length + 1);
            break;
          }
        }

        // Rejoin house number + suffix with hyphen (e.g. "275-1" stays "275-1")
        const parts = streetSlug.split('-');
        const merged = [];
        for (let i = 0; i < parts.length; i++) {
          if (i > 0 && /^\d+$/.test(parts[i - 1]) && /^\d+[a-z]?$/.test(parts[i]) && parts[i].length <= 3) {
            merged[merged.length - 1] += '-' + parts[i];
          } else {
            merged.push(parts[i]);
          }
        }
        const titleCased = merged.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
        const cityTitle = city.charAt(0).toUpperCase() + city.slice(1);

        return {
          street: titleCased,
          city: cityTitle + ', The Netherlands'
        };
      } catch (e) {
        return null;
      }
    }

    function parseAddressFromFML(data) {
      for (const floor of data.floors ?? []) {
        for (const design of floor.designs ?? []) {
          for (const label of design.labels ?? []) {
            if (!label.text || !label.bold) continue;
            const lines = label.text.split('\n');
            if (lines.length < 2) continue;
            const firstLine = lines[0].trim();
            const dashIdx = firstLine.lastIndexOf(' - ');
            if (dashIdx > 0) {
              return {
                street: firstLine.slice(0, dashIdx).trim(),
                city: firstLine.slice(dashIdx + 3).trim() + ', The Netherlands'
              };
            }
          }
        }
      }
      return null;
    }

    function getFloorLabel(dutchName) {
      const name = (dutchName || '').toUpperCase().trim();
      const map = {
        'BEGANE GROND': '1st floor',
        'EERSTE VERDIEPING': '2nd floor',
        'TWEEDE VERDIEPING': '3rd floor',
        'DERDE VERDIEPING': '4th floor',
        'VIERDE VERDIEPING': '5th floor',
        'KELDER': 'Basement',
        'BERGING': 'Storage',
        'SCHUUR': 'Storage',
        'ZOLDER': 'Attic',
        'GARAGE': 'Garage'
      };
      if (map[name]) return map[name];
      for (const [key, val] of Object.entries(map)) {
        if (name.includes(key)) return val;
      }
      return dutchName || 'Floor';
    }

    let currentAddress = { street: '', city: '' };
    let lastFundaUrl = '';

    // ============================================================
    // BOUNDING BOX
    // ============================================================
    function computeWallBBox(design) {
      const pts = [];
      for (const wall of design.walls ?? []) {
        pts.push(wall.a, wall.b);
        if (wall.c && wall.c.x != null && wall.c.y != null) {
          for (let t = 0.25; t <= 0.75; t += 0.25) {
            pts.push({
              x: (1-t)*(1-t)*wall.a.x + 2*(1-t)*t*wall.c.x + t*t*wall.b.x,
              y: (1-t)*(1-t)*wall.a.y + 2*(1-t)*t*wall.c.y + t*t*wall.b.y
            });
          }
        }
      }
      if (!pts.length) return null;
      return {
        minX: Math.min(...pts.map(p => p.x)),
        minY: Math.min(...pts.map(p => p.y)),
        maxX: Math.max(...pts.map(p => p.x)),
        maxY: Math.max(...pts.map(p => p.y))
      };
    }

    function isSurfaceOutsideWalls(surface, wallBBox) {
      if (!wallBBox) return false;
      const poly = surface.poly ?? [];
      if (poly.length < 3) return false;
      let cx = 0, cy = 0;
      for (const pt of poly) { cx += (pt.x ?? 0); cy += (pt.y ?? 0); }
      cx /= poly.length; cy /= poly.length;
      const MARGIN = 25;
      return cx < wallBBox.minX - MARGIN || cx > wallBBox.maxX + MARGIN ||
             cy < wallBBox.minY - MARGIN || cy > wallBBox.maxY + MARGIN;
    }

    function computeBoundingBox(design) {
      const points = [];
      const wallBBox = computeWallBBox(design);
      for (const wall of design.walls ?? []) {
        points.push({ x: wall.a.x, y: wall.a.y }, { x: wall.b.x, y: wall.b.y });
        if (wall.c && wall.c.x != null && wall.c.y != null) {
          for (let t = 0.25; t <= 0.75; t += 0.25) {
            const px = (1-t)*(1-t)*wall.a.x + 2*(1-t)*t*wall.c.x + t*t*wall.b.x;
            const py = (1-t)*(1-t)*wall.a.y + 2*(1-t)*t*wall.c.y + t*t*wall.b.y;
            points.push({ x: px, y: py });
          }
        }
      }
      for (const area of design.areas ?? []) {
        for (const pt of area.poly ?? []) points.push(pt);
      }
      for (const surface of design.surfaces ?? []) {
        if (isSurfaceOutsideWalls(surface, wallBBox)) continue;
        const tessellated = tessellateSurfacePoly(surface.poly ?? []);
        for (const pt of tessellated) points.push(pt);
      }
      for (const bal of design.balustrades ?? []) {
        points.push({ x: bal.a.x, y: bal.a.y }, { x: bal.b.x, y: bal.b.y });
        if (bal.c && bal.c.x != null && bal.c.y != null) {
          for (let t = 0.25; t <= 0.75; t += 0.25) {
            const px = (1-t)*(1-t)*bal.a.x + 2*(1-t)*t*bal.c.x + t*t*bal.b.x;
            const py = (1-t)*(1-t)*bal.a.y + 2*(1-t)*t*bal.c.y + t*t*bal.b.y;
            points.push({ x: px, y: py });
          }
        }
      }
      if (!points.length) return { minX: 0, minY: 0, maxX: 1, maxY: 1 };
      return {
        minX: Math.min(...points.map(p => p.x)),
        minY: Math.min(...points.map(p => p.y)),
        maxX: Math.max(...points.map(p => p.x)),
        maxY: Math.max(...points.map(p => p.y))
      };
    }

    // ============================================================
    // BALUSTRADE AUTO-DETECTION
    // ============================================================
    function detectBalustrades(design) {
      const balustrades = [];
      for (const item of design.items ?? []) {
        const w = item.width ?? 0;
        const h = item.height ?? 0;
        const zHeight = item.z_height ?? 0;
        const isElongated = Math.max(w, h) / Math.max(1, Math.min(w, h)) > 2.5 && Math.min(w, h) < 20;
        const isRailingHeight = zHeight > 50 && zHeight < 130;
        if (isElongated && isRailingHeight) {
          const cx = item.x ?? 0;
          const cy = item.y ?? 0;
          const angle = (item.rotation ?? 0) * Math.PI / 180;
          const halfLen = Math.max(w, h) / 2;
          const dx = Math.cos(angle) * halfLen;
          const dy = Math.sin(angle) * halfLen;
          balustrades.push({
            a: { x: cx - dx, y: cy - dy },
            b: { x: cx + dx, y: cy + dy },
            thickness: Math.min(w, h),
            height: zHeight
          });
        }
      }
      return balustrades;
    }

    // ============================================================
    // WALL POLYGON UNION (Funda-style rendering)
    // ============================================================
    // Each wall → 2D rectangle. Boolean-union all rectangles into
    // one merged outline using polygon-clipping library. Extrude to 3D.

    function wallToRect(w, extA, extB) {
      const ax = w.a.x, ay = w.a.y, bx = w.b.x, by = w.b.y;
      const dx = bx - ax, dy = by - ay;
      const len = Math.hypot(dx, dy);
      if (len < 0.1) return null;
      const ux = dx / len, uy = dy / len;
      const nx = -uy, ny = ux;
      const ht = (w.thickness ?? 20) / 2;
      const eax = ax - ux * extA, eay = ay - uy * extA;
      const ebx = bx + ux * extB, eby = by + uy * extB;
      return [
        [eax + nx * ht, eay + ny * ht],
        [ebx + nx * ht, eby + ny * ht],
        [ebx - nx * ht, eby - ny * ht],
        [eax - nx * ht, eay - ny * ht],
        [eax + nx * ht, eay + ny * ht]
      ];
    }

    function computeWallUnion(walls, allWalls) {
      const TOLERANCE = 3;

      function isDiagWall(w) {
        const dx = w.b.x - w.a.x, dy = w.b.y - w.a.y;
        const len = Math.hypot(dx, dy);
        if (len < 0.1) return false;
        return Math.min(Math.abs(dx / len), Math.abs(dy / len)) > 0.15;
      }

      // Step 1: Build junction map from ALL walls
      const junctions = new Map();
      for (const w of allWalls) {
        if (Math.hypot(w.b.x - w.a.x, w.b.y - w.a.y) < 0.1) continue;
        for (const ep of ['a', 'b']) {
          const px = w[ep].x, py = w[ep].y;
          let found = false;
          for (const [key, members] of junctions) {
            const [kx, ky] = key.split(',').map(Number);
            if (Math.hypot(px - kx, py - ky) < TOLERANCE) {
              members.push({ wall: w, endpoint: ep });
              found = true;
              break;
            }
          }
          if (!found) {
            junctions.set(`${px},${py}`, [{ wall: w, endpoint: ep }]);
          }
        }
      }

      // Step 2: For each wall, compute extension per endpoint.
      // Straight walls at L-junctions with other STRAIGHT walls → extend.
      // Skip extension if the other wall is diagonal.
      const rects = [];
      for (let i = 0; i < walls.length; i++) {
        const w = walls[i];
        const wdx = w.b.x - w.a.x, wdy = w.b.y - w.a.y;
        const wlen = Math.hypot(wdx, wdy);
        if (wlen < 0.1) continue;
        const wIsDiag = isDiagWall(w);

        let extA = 0, extB = 0;

        if (!wIsDiag) {
          for (const other of allWalls) {
            if (other === w) continue;
            if (isDiagWall(other)) continue; // skip diagonal others

            const otherHt = (other.thickness ?? 20) / 2;
            const aShares =
              Math.hypot(w.a.x - other.a.x, w.a.y - other.a.y) < TOLERANCE ||
              Math.hypot(w.a.x - other.b.x, w.a.y - other.b.y) < TOLERANCE;
            if (aShares) extA = Math.max(extA, otherHt);

            const bShares =
              Math.hypot(w.b.x - other.a.x, w.b.y - other.a.y) < TOLERANCE ||
              Math.hypot(w.b.x - other.b.x, w.b.y - other.b.y) < TOLERANCE;
            if (bShares) extB = Math.max(extB, otherHt);
          }
        }

        const r = wallToRect(w, extA, extB);
        if (r) rects.push([r]);
      }

      // Step 3: Add fill polygons ONLY at junctions involving a diagonal wall.
      // These fill the gap between straight and diagonal wall rectangles.
      for (const [key, members] of junctions) {
        if (members.length < 2) continue;
        const hasDiagonal = members.some(m => isDiagWall(m.wall));
        if (!hasDiagonal) continue; // straight-only junctions handled by extension

        const edgePoints = [];
        for (const m of members) {
          const w = m.wall;
          const dx = w.b.x - w.a.x, dy = w.b.y - w.a.y;
          const len = Math.hypot(dx, dy);
          if (len < 0.1) continue;
          const nx = -dy / len, ny = dx / len;
          const ht = (w.thickness ?? 20) / 2;
          const px = w[m.endpoint].x, py = w[m.endpoint].y;
          edgePoints.push({ x: px + nx * ht, y: py + ny * ht });
          edgePoints.push({ x: px - nx * ht, y: py - ny * ht });
        }
        if (edgePoints.length < 3) continue;

        const jx = Number(key.split(',')[0]), jy = Number(key.split(',')[1]);
        edgePoints.sort((a, b) =>
          Math.atan2(a.y - jy, a.x - jx) - Math.atan2(b.y - jy, b.x - jx)
        );

        const fillRing = edgePoints.map(p => [p.x, p.y]);
        fillRing.push([edgePoints[0].x, edgePoints[0].y]);
        rects.push([fillRing]);
      }

      if (rects.length === 0) return [];
      try {
        return polygonClipping.union(...rects);
      } catch (e) {
        console.warn('Wall union failed, falling back to individual rects', e);
        return rects;
      }
    }

    // ============================================================
    // BALUSTRADE ENDPOINT EXTENSION
    // ============================================================
    function extendBalustrades(balustrades) {
      const TOL = 15;
      return balustrades.map((bal, idx) => {
        const ax = bal.a.x, ay = bal.a.y;
        const bx = bal.b.x, by = bal.b.y;
        const dx = bx - ax, dy = by - ay;
        const len = Math.hypot(dx, dy);
        if (len < 0.1) return { ...bal };
        const ux = dx / len, uy = dy / len;
        // Only extend at endpoints connected to another balustrade, by half the OTHER's thickness
        let extA = 0, extB = 0;
        for (let i = 0; i < balustrades.length; i++) {
          if (i === idx) continue;
          const o = balustrades[i];
          const oHt = (o.thickness ?? 10) / 2;
          if (Math.hypot(ax - o.a.x, ay - o.a.y) < TOL || Math.hypot(ax - o.b.x, ay - o.b.y) < TOL) extA = Math.max(extA, oHt);
          if (Math.hypot(bx - o.a.x, by - o.a.y) < TOL || Math.hypot(bx - o.b.x, by - o.b.y) < TOL) extB = Math.max(extB, oHt);
        }
        return {
          a: { x: ax - ux * extA, y: ay - uy * extA },
          b: { x: bx + ux * extB, y: by + uy * extB },
          thickness: bal.thickness ?? 10,
          height: bal.height ?? 100
        };
      });
    }

    // ============================================================
    // BALUSTRADE CHAIN BUILDING
    // ============================================================
    function buildBalustradeChains(balustrades) {
      if (balustrades.length === 0) return [];
      const TOLERANCE = 15;
      const used = new Array(balustrades.length).fill(false);
      const chains = [];

      function findNeighbor(px, py, excludeIdx) {
        for (let i = 0; i < balustrades.length; i++) {
          if (i === excludeIdx || used[i]) continue;
          const b = balustrades[i];
          if (Math.hypot(b.a.x - px, b.a.y - py) < TOLERANCE) return { idx: i, end: 'a' };
          if (Math.hypot(b.b.x - px, b.b.y - py) < TOLERANCE) return { idx: i, end: 'b' };
        }
        return null;
      }

      for (let start = 0; start < balustrades.length; start++) {
        if (used[start]) continue;
        used[start] = true;
        const chain = [{ bal: balustrades[start], flipped: false }];

        let tip = balustrades[start].b;
        let lastIdx = start;
        while (true) {
          const nb = findNeighbor(tip.x, tip.y, lastIdx);
          if (!nb) break;
          used[nb.idx] = true;
          const bal = balustrades[nb.idx];
          const flipped = nb.end === 'b';
          chain.push({ bal, flipped });
          tip = flipped ? bal.a : bal.b;
          lastIdx = nb.idx;
        }

        tip = balustrades[start].a;
        lastIdx = start;
        while (true) {
          const nb = findNeighbor(tip.x, tip.y, lastIdx);
          if (!nb) break;
          used[nb.idx] = true;
          const bal = balustrades[nb.idx];
          const flipped = nb.end === 'a';
          chain.unshift({ bal, flipped });
          tip = flipped ? bal.b : bal.a;
          lastIdx = nb.idx;
        }

        chains.push(chain);
      }

      return chains;
    }

    function getChainEdges(chain) {
      const leftEdge = [];
      const rightEdge = [];

      for (let i = 0; i < chain.length; i++) {
        const entry = chain[i];
        const b = entry.bal;
        const a_pt = entry.flipped ? b.b : b.a;
        const b_pt = entry.flipped ? b.a : b.b;
        const dx = b_pt.x - a_pt.x, dy = b_pt.y - a_pt.y;
        const len = Math.hypot(dx, dy);
        if (len < 0.1) continue;
        const nx = -dy / len, ny = dx / len;
        const ht = (b.thickness ?? 10) / 2;

        if (i === 0) {
          leftEdge.push({ x: a_pt.x + nx * ht, y: a_pt.y + ny * ht });
          rightEdge.push({ x: a_pt.x - nx * ht, y: a_pt.y - ny * ht });
        }
        leftEdge.push({ x: b_pt.x + nx * ht, y: b_pt.y + ny * ht });
        rightEdge.push({ x: b_pt.x - nx * ht, y: b_pt.y - ny * ht });
      }

      return { leftEdge, rightEdge };
    }

    function mergeBalustradeStrips(balustrades) {
      const chains = buildBalustradeChains(balustrades);
      const strips = [];

      for (const chain of chains) {
        if (chain.length < 2) {
          const b = chain[0].bal;
          const dx = b.b.x - b.a.x, dy = b.b.y - b.a.y;
          const len = Math.hypot(dx, dy);
          if (len < 0.1) continue;
          const nx = -dy / len, ny = dx / len;
          const ht = (b.thickness ?? 10) / 2;
          strips.push([
            { x: b.a.x + nx * ht, y: b.a.y + ny * ht },
            { x: b.b.x + nx * ht, y: b.b.y + ny * ht },
            { x: b.b.x - nx * ht, y: b.b.y - ny * ht },
            { x: b.a.x - nx * ht, y: b.a.y - ny * ht }
          ]);
          continue;
        }

        const { leftEdge, rightEdge } = getChainEdges(chain);
        const poly = [...leftEdge, ...rightEdge.reverse()];
        if (poly.length >= 3) strips.push(poly);
      }

      return strips;
    }

    function buildBalustradeFillPolygons(balustrades) {
      const chains = buildBalustradeChains(balustrades);
      const fills = [];

      for (const chain of chains) {
        if (chain.length < 3) continue;

        const { leftEdge, rightEdge } = getChainEdges(chain);

        function bboxArea(pts) {
          let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
          for (const p of pts) {
            minX = Math.min(minX, p.x); maxX = Math.max(maxX, p.x);
            minY = Math.min(minY, p.y); maxY = Math.max(maxY, p.y);
          }
          return (maxX - minX) * (maxY - minY);
        }

        const outerEdge = bboxArea(leftEdge) >= bboxArea(rightEdge) ? leftEdge : rightEdge;
        const innerEdge = outerEdge === leftEdge ? rightEdge : leftEdge;

        if (outerEdge.length >= 3) {
          fills.push([...outerEdge]);
        }

        if (innerEdge.length >= 3) {
          fills.push([...innerEdge]);
        }
      }

      return fills;
    }

    // ============================================================
    // ARC WALL TESSELLATION
    // ============================================================
    function tessellateArcWall(wall, numSeg) {
      const ax = wall.a.x, ay = wall.a.y;
      const bx = wall.b.x, by = wall.b.y;
      const cx = wall.c.x, cy = wall.c.y;
      const hA = wall.az?.h ?? wall.bz?.h ?? 265;
      const hB = wall.bz?.h ?? wall.az?.h ?? 265;
      const segments = [];
      for (let i = 0; i < numSeg; i++) {
        const t0 = i / numSeg;
        const t1 = (i + 1) / numSeg;
        const x0 = (1 - t0) * (1 - t0) * ax + 2 * (1 - t0) * t0 * cx + t0 * t0 * bx;
        const y0 = (1 - t0) * (1 - t0) * ay + 2 * (1 - t0) * t0 * cy + t0 * t0 * by;
        const x1 = (1 - t1) * (1 - t1) * ax + 2 * (1 - t1) * t1 * cx + t1 * t1 * bx;
        const y1 = (1 - t1) * (1 - t1) * ay + 2 * (1 - t1) * t1 * cy + t1 * t1 * by;
        segments.push({
          a: { x: x0, y: y0 },
          b: { x: x1, y: y1 },
          thickness: wall.thickness ?? 20,
          openings: [],
          _arcSeg: true,
          _heightA: hA + (hB - hA) * t0,
          _heightB: hA + (hB - hA) * t1
        });
      }
      return segments;
    }

    function flattenWalls(walls) {
      const ARC_SEGMENTS = 16;
      const result = [];
      for (const wall of walls) {
        if (wall.c && wall.c.x != null && wall.c.y != null) {
          result.push(...tessellateArcWall(wall, ARC_SEGMENTS));
        } else {
          result.push({
            ...wall,
            _heightA: wall.az?.h ?? wall.bz?.h ?? 265,
            _heightB: wall.bz?.h ?? wall.az?.h ?? 265
          });
        }
      }
      return result;
    }

    function tessellateArcBalustrade(bal, numSeg) {
      const ax = bal.a.x, ay = bal.a.y;
      const bx = bal.b.x, by = bal.b.y;
      const cx = bal.c.x, cy = bal.c.y;
      const segments = [];
      for (let i = 0; i < numSeg; i++) {
        const t0 = i / numSeg;
        const t1 = (i + 1) / numSeg;
        const x0 = (1 - t0) * (1 - t0) * ax + 2 * (1 - t0) * t0 * cx + t0 * t0 * bx;
        const y0 = (1 - t0) * (1 - t0) * ay + 2 * (1 - t0) * t0 * cy + t0 * t0 * by;
        const x1 = (1 - t1) * (1 - t1) * ax + 2 * (1 - t1) * t1 * cx + t1 * t1 * bx;
        const y1 = (1 - t1) * (1 - t1) * ay + 2 * (1 - t1) * t1 * cy + t1 * t1 * by;
        segments.push({
          a: { x: x0, y: y0 },
          b: { x: x1, y: y1 },
          thickness: bal.thickness ?? 10,
          height: bal.height ?? 100
        });
      }
      return segments;
    }

    function flattenBalustrades(balustrades) {
      const ARC_SEGMENTS = 16;
      const result = [];
      for (const bal of balustrades) {
        if (bal.c && bal.c.x != null && bal.c.y != null) {
          result.push(...tessellateArcBalustrade(bal, ARC_SEGMENTS));
        } else {
          result.push(bal);
        }
      }
      return result;
    }

    // ============================================================
    // SURFACE POLYGON CURVE TESSELLATION
    // ============================================================
    function tessellateSurfacePoly(poly) {
      if (!poly || poly.length < 3) return poly;
      const ARC_SEGMENTS = 24;
      const result = [];
      for (let i = 0; i < poly.length; i++) {
        const curr = poly[i];
        const prev = poly[(i - 1 + poly.length) % poly.length];
        if (curr.cx != null && curr.cy != null) {
          const ax = prev.x, ay = prev.y;
          const bx = curr.x, by = curr.y;
          const cx = curr.cx, cy = curr.cy;
          for (let s = 1; s <= ARC_SEGMENTS; s++) {
            const t = s / ARC_SEGMENTS;
            const px = (1-t)*(1-t)*ax + 2*(1-t)*t*cx + t*t*bx;
            const py = (1-t)*(1-t)*ay + 2*(1-t)*t*cy + t*t*by;
            result.push({ x: px, y: py, z: curr.z ?? 0 });
          }
        } else {
          result.push({ x: curr.x, y: curr.y, z: curr.z ?? 0 });
        }
      }
      return result;
    }

    // ============================================================
    // STAIR VOID DETECTION
    // ============================================================
    function detectStairVoids(allFloorDesigns) {
      const POSITION_TOL = 5;
      const voidsByFloor = allFloorDesigns.map(() => []);

      for (let fi = 0; fi < allFloorDesigns.length; fi++) {
        const design = allFloorDesigns[fi];

        for (const surface of design.surfaces ?? []) {
          if ((surface.role ?? -1) !== 14 && !surface.isCutout) continue;
          const poly = tessellateSurfacePoly(surface.poly ?? []);
          if (poly.length >= 3) {
            voidsByFloor[fi].push(poly.map(p => ({ x: p.x, y: p.y })));
          }
        }

        if (fi > 0) {
          const prevDesign = allFloorDesigns[fi - 1];
          const prevItems = prevDesign.items ?? [];
          const currItems = design.items ?? [];

          for (const curr of currItems) {
            if (!curr.refid) continue;
            for (const prev of prevItems) {
              if (prev.refid !== curr.refid) continue;
              if (Math.abs((prev.x ?? 0) - (curr.x ?? 0)) > POSITION_TOL) continue;
              if (Math.abs((prev.y ?? 0) - (curr.y ?? 0)) > POSITION_TOL) continue;

              const VOID_MARGIN = 10;
              const w = Math.max(0, (curr.width ?? 0) - VOID_MARGIN * 2);
              const h = Math.max(0, (curr.height ?? 0) - VOID_MARGIN * 2);
              if (w < 30 || h < 30) continue;
              const cx = curr.x ?? 0;
              const cy = curr.y ?? 0;
              const rot = (curr.rotation ?? 0) * Math.PI / 180;
              const cosR = Math.cos(rot);
              const sinR = Math.sin(rot);
              const hw = w / 2, hh = h / 2;

              const corners = [
                { x: cx + cosR * (-hw) - sinR * (-hh), y: cy + sinR * (-hw) + cosR * (-hh) },
                { x: cx + cosR * ( hw) - sinR * (-hh), y: cy + sinR * ( hw) + cosR * (-hh) },
                { x: cx + cosR * ( hw) - sinR * ( hh), y: cy + sinR * ( hw) + cosR * ( hh) },
                { x: cx + cosR * (-hw) - sinR * ( hh), y: cy + sinR * (-hw) + cosR * ( hh) }
              ];
              voidsByFloor[fi].push(corners);
              break;
            }
          }
        }
      }
      return voidsByFloor;
    }

    function pointInPolygon(px, py, poly) {
      let inside = false;
      for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
        const xi = poly[i].x, yi = poly[i].y;
        const xj = poly[j].x, yj = poly[j].y;
        if ((yi > py) !== (yj > py) && px < (xj - xi) * (py - yi) / (yj - yi) + xi) {
          inside = !inside;
        }
      }
      return inside;
    }

    function polygonOverlapsVoid(poly, voids) {
      if (!voids || voids.length === 0) return false;
      let cx = 0, cy = 0;
      for (const p of poly) { cx += p.x; cy += p.y; }
      cx /= poly.length; cy /= poly.length;
      for (const v of voids) {
        if (pointInPolygon(cx, cy, v)) return true;
      }
      return false;
    }

    // ============================================================
    // FLOOR PROCESSING
    // ============================================================
    function processFloors(data) {
      floors = [];
      canvases = [];
      maxWorldW = 1;
      maxWorldH = 1;

      const validFloors = (data.floors ?? []).filter(f => f?.designs?.[0]);

      const allDesigns = validFloors.map(f => f.designs[0]);
      const voidsByFloor = detectStairVoids(allDesigns);

      const widths = [], heights = [];

      for (const floor of validFloors) {
        const bbox = computeBoundingBox(floor.designs[0]);
        widths.push(Math.max(1, bbox.maxX - bbox.minX));
        heights.push(Math.max(1, bbox.maxY - bbox.minY));
      }

      const medW = median(widths);
      const medH = median(heights);

      for (let i = 0; i < validFloors.length; i++) {
        const floor = validFloors[i];
        const design = floor.designs[0];

        if (!design.balustrades) {
          design.balustrades = detectBalustrades(design);
        }
        design.balustrades = flattenBalustrades(design.balustrades);

        const bbox = computeBoundingBox(design);
        const worldW = Math.max(1, bbox.maxX - bbox.minX);
        const worldH = Math.max(1, bbox.maxY - bbox.minY);
        const name = (floor.name || "").toLowerCase();

        if (!name.includes("situatie") && !name.includes("site")
            && worldW <= medW * 2.2 && worldH <= medH * 2.2) {
          maxWorldW = Math.max(maxWorldW, worldW);
          maxWorldH = Math.max(maxWorldH, worldH);
        }

        floors.push({
          design, bbox, worldW, worldH,
          name: floor.name || `Verdieping ${i + 1}`,
          voids: voidsByFloor[i] || []
        });
      }

      // Auto-detect excluded floors
      excludedFloors = new Set();
      for (let i = 0; i < floors.length; i++) {
        if (isLikelySituatie(floors[i])) {
          excludedFloors.add(i);
        }
      }

      // Show admin export buttons + frame toggle
      document.getElementById('uploadActions').classList.add('active');
      for (const n of [1, 2]) {
        const bt = document.getElementById('btnTest' + n);
        if (bt) bt.style.display = 'none';
      }
      var adminFrameToggle = document.getElementById('adminFrameToggle');
      if (adminFrameToggle) adminFrameToggle.style.display = '';
      // Layout controls moved to step 4 per-floor cards

      // Admin: floor dimensions overview
      var dimsEl = document.getElementById('adminFloorDims');
      if (dimsEl) {
        var html = '<div class="dims-title">Afmetingen (L × B × H in m)</div>';
        for (var fi = 0; fi < floors.length; fi++) {
          var f = floors[fi];
          var l = (f.worldW / 100).toFixed(1);
          var b = (f.worldH / 100).toFixed(1);
          var excl = excludedFloors.has(fi) ? ' excluded' : '';
          html += '<div class="dims-row' + excl + '"><span>' + f.name + '</span><span>' + l + ' × ' + b + ' × 2.8</span></div>';
        }
        dimsEl.innerHTML = html;
        dimsEl.style.display = '';
      }

      // Don't switch to unified preview yet — stays on hero image until step 2

      // Update address in preview
      updateFrameAddress();

      // Render thumbnails, then show the button
      if (unifiedFloorsOverlay) unifiedFloorsOverlay.style.display = 'none';
      if (floorsLoading) floorsLoading.classList.remove('hidden');
      requestAnimationFrame(() => {
        setTimeout(() => {
          renderPreviewThumbnails();
          updateFloorLabels();
          if (floorsLoading) floorsLoading.classList.add('hidden');
          updateWizardUI();
        }, 100);
      });
    }

    // ============================================================
    // 3D VIEWER — shared scene builder
    // ============================================================
    function parseOBJToGroups(objString) {
      const allVertices = [];
      const groups = {};
      let currentGroup = 'default';
      groups[currentGroup] = [];

      for (const line of objString.split('\n')) {
        const parts = line.trim().split(/\s+/);
        if (parts[0] === 'v' && parts.length >= 4) {
          allVertices.push(parseFloat(parts[1]), parseFloat(parts[2]), parseFloat(parts[3]));
        } else if (parts[0] === 'g' && parts.length >= 2) {
          currentGroup = parts[1];
          if (!groups[currentGroup]) groups[currentGroup] = [];
        } else if (parts[0] === 'f' && parts.length >= 4) {
          const idxs = parts.slice(1).map(p => parseInt(p.split('/')[0]) - 1);
          for (let i = 1; i < idxs.length - 1; i++) {
            groups[currentGroup].push(idxs[0], idxs[i], idxs[i + 1]);
          }
        }
      }

      function makeGeometry(faceIndices) {
        if (faceIndices.length === 0) return null;
        const geometry = new THREE.BufferGeometry();
        geometry.setAttribute('position', new THREE.BufferAttribute(new Float32Array(allVertices), 3));
        geometry.setIndex(faceIndices);
        geometry.computeVertexNormals();
        return geometry;
      }

      const wallFaces = (groups['walls'] || [])
        .concat(groups['walls_balustrades'] || [])
        .concat(groups['default'] || []);
      const allFaces = Object.values(groups).reduce((acc, g) => acc.concat(g), []);

      return {
        walls: makeGeometry(wallFaces),
        floor: makeGeometry(groups['floor'] || []),
        all: makeGeometry(allFaces)
      };
    }

    // Build a Three.js scene for a floor (reused by preview thumbnails, floor review, layout)
    // Editing floor color — subtle muted sage so walls stand out
    var EDIT_FLOOR_COLOR = 0xB8C4B0;

    function buildFloorScene(floorIndex, floorColor) {
      const floor = floors[floorIndex];
      if (!floor) return null;

      const objString = generateFloorOBJ(floor);
      const groups = parseOBJToGroups(objString);

      const allGeo = groups.all;
      allGeo.computeBoundingBox();
      const box = allGeo.boundingBox;
      const center = new THREE.Vector3();
      box.getCenter(center);
      const size = new THREE.Vector3();
      box.getSize(size);

      const offsetX = -center.x;
      const offsetY = -center.y;
      const offsetZ = -center.z;
      if (groups.walls) groups.walls.translate(offsetX, offsetY, offsetZ);
      if (groups.floor) groups.floor.translate(offsetX, offsetY, offsetZ);
      center.set(0, 0, 0);

      const scene = new THREE.Scene();
      scene.background = null;

      const wallMaterial = new THREE.MeshPhongMaterial({
        color: 0xAA9A82,
        flatShading: true,
        side: THREE.DoubleSide,
        shininess: 5,
        specular: 0x222222
      });
      var fColor = (floorColor !== undefined) ? floorColor : 0xAA9A82;
      const floorMaterial = new THREE.MeshPhongMaterial({
        color: fColor,
        flatShading: true,
        side: THREE.DoubleSide,
        shininess: 5,
        specular: 0x222222
      });

      if (groups.walls) scene.add(new THREE.Mesh(groups.walls, wallMaterial));
      if (groups.floor) scene.add(new THREE.Mesh(groups.floor, floorMaterial));

      scene.add(new THREE.AmbientLight(0xFFF8F0, 1.0));
      const dirLight = new THREE.DirectionalLight(0xFFF5E8, 0.7);
      dirLight.position.set(0, 8, 5);
      scene.add(dirLight);
      const fillLight = new THREE.DirectionalLight(0xF0EBE0, 0.4);
      fillLight.position.set(-4, 6, -1);
      scene.add(fillLight);

      // Global size based on largest floor — for uniform scaling across all viewers
      const SCALE = 0.01;
      const globalSize = new THREE.Vector3(maxWorldW * SCALE, size.y, maxWorldH * SCALE);

      return { scene, size, center, globalSize };
    }

    // ============================================================
    // STATIC THUMBNAIL RENDER (for unified preview)
    // ============================================================
    function renderStaticThumbnail(floorIndex, container) {
      renderStaticThumbnailSized(floorIndex, container, null, null, {});
    }

    // opts: { ortho: bool, floorColor: hex }
    function renderStaticThumbnailSized(floorIndex, container, forceW, forceH, opts) {
      opts = opts || {};
      var floorColor = (opts.floorColor !== undefined) ? opts.floorColor : undefined;
      const result = buildFloorScene(floorIndex, floorColor);
      if (!result) return;
      const { scene, size, center, globalSize } = result;

      // Apply rotation if set for this floor
      var rotation = getFloorRotate(floorIndex);
      if (rotation) {
        scene.rotateY(rotation * Math.PI / 180);
      }

      const width = Math.round(forceW || container.getBoundingClientRect().width) || 200;
      const height = Math.round(forceH || container.getBoundingClientRect().height) || 260;

      var camera;
      if (opts.ortho) {
        // Orthographic camera — uniform scale, flat top-down (for editing)
        var floorData = floors[floorIndex];
        var padding = 1.1;
        var pxPerUnit = width / (floorData.worldW * 0.01 * padding);
        var halfFrustumW = (width / 2) / pxPerUnit;
        var halfFrustumH = (height / 2) / pxPerUnit;
        camera = new THREE.OrthographicCamera(
          -halfFrustumW, halfFrustumW, halfFrustumH, -halfFrustumH, 0.01, 1000
        );
        camera.up.set(0, 0, -1); // Z- = up on screen (north in floorplan)
        camera.position.set(0, 50, 0);
        camera.lookAt(0, 0, 0);

        // Fine-tune alignment: push OBJ to edge of frustum so model
        // visually touches the canvas edge (compensates for OBJ bounding
        // box being slightly larger than worldW/worldH due to wall thickness)
        var align = getFloorAlign(floorIndex);
        if (align === 'top' || align === 'bottom') {
          var shiftZ = halfFrustumH - size.z / 2;
          // Camera up=(0,0,-1) means screen-up = world -Z
          // 'top' on screen → shift scene -Z, 'bottom' → shift scene +Z
          scene.position.z += (align === 'bottom') ? shiftZ : -shiftZ;
        }

        camera.updateProjectionMatrix();
      } else {
        // Perspective camera with slight tilt (for final result with depth)
        var ref = globalSize;
        var padding = 1.15;
        var halfW = (ref.x * padding) / 2;
        var halfZ = (ref.z * padding) / 2;
        var halfExtent = Math.max(halfW, halfZ);
        var FOV = 12;
        var aspect = width / height;
        camera = new THREE.PerspectiveCamera(FOV, aspect, 0.01, halfExtent * 100);
        var camDist = halfExtent / Math.tan((FOV / 2) * Math.PI / 180);
        camera.position.set(0, camDist, camDist * 0.14);
        camera.lookAt(center);
      }

      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
      renderer.setClearColor(0x000000, 0);
      renderer.setSize(width, height);
      renderer.setPixelRatio(dpr);

      // Single frame render — no animation loop
      renderer.render(scene, camera);
      container.appendChild(renderer.domElement);

      previewViewers.push({ renderer });
    }

    // ============================================================
    // LAYOUT ENGINE — computes scale + position for preview
    // ============================================================
    var currentLayout = null; // stores computed layout for admin controls
    var layoutAlign = 'center'; // 'center' or 'bottom' or 'top'
    var layoutGapFactor = 0.08; // gap as fraction of largest floor dimension (0..0.3)
    var floorSettings = {}; // per-floor: { align: 'center'|'top'|'bottom', rotate: 0..315 }

    function computeFloorLayout(zoneW, zoneH, includedIndices) {
      // Gather floor dimensions in cm
      var items = includedIndices.map(function(i) {
        var f = floors[i];
        return { index: i, w: f.worldW, h: f.worldH, name: f.name || '' };
      });

      if (items.length === 0) return { scale: 1, positions: [] };

      // Total bounding box if we stack all floors
      var totalW = 0, totalH = 0;
      for (var it of items) {
        totalW = Math.max(totalW, it.w);
        totalH = Math.max(totalH, it.h);
      }

      // Available zone aspect ratio
      var zoneAspect = zoneW / zoneH;

      // Try different layout strategies and pick the best
      var layouts = [];

      if (items.length === 1) {
        // Single floor — centered
        layouts.push(layoutSingle(items));
      } else if (items.length === 2) {
        // Two floors — try side-by-side, stacked, and diagonal
        layouts.push(layoutSideBySide(items));
        layouts.push(layoutStacked(items));
        layouts.push(layoutDiagonal(items));
      } else if (items.length === 3) {
        layouts.push(layoutTriangle(items));
        layouts.push(layoutSideBySide(items));
        layouts.push(layoutStacked(items));
      } else {
        // 4+ floors — grid-based
        layouts.push(layoutGrid(items));
        layouts.push(layoutSideBySide(items));
      }

      // Score each layout: how well does it fill the zone?
      var best = null, bestScore = -1;
      for (var layout of layouts) {
        var bbox = layoutBoundingBox(layout.positions);
        if (bbox.w === 0 || bbox.h === 0) continue;

        // Scale to fit zone
        var scaleX = zoneW / bbox.w;
        var scaleY = zoneH / bbox.h;
        var scale = Math.min(scaleX, scaleY);

        // Score: how much of the zone is filled (0..1)
        var fillRatio = (bbox.w * scale * bbox.h * scale) / (zoneW * zoneH);
        // Prefer layouts closer to zone aspect ratio
        var layoutAspect = bbox.w / bbox.h;
        var aspectMatch = 1 - Math.abs(layoutAspect - zoneAspect) / Math.max(layoutAspect, zoneAspect);
        var score = fillRatio * 0.6 + aspectMatch * 0.4;

        if (score > bestScore) {
          bestScore = score;
          best = { positions: layout.positions, bbox: bbox, scale: scale, type: layout.type };
        }
      }

      if (!best) return { scale: 1, positions: [] };

      // Center the layout in the zone
      var finalScale = best.scale * 0.88; // 12% padding
      var scaledW = best.bbox.w * finalScale;
      var scaledH = best.bbox.h * finalScale;
      var offsetX = (zoneW - scaledW) / 2 - best.bbox.x * finalScale;
      var offsetY = (zoneH - scaledH) / 2 - best.bbox.y * finalScale;

      var result = [];
      for (var pos of best.positions) {
        result.push({
          index: pos.index,
          x: pos.x * finalScale + offsetX,
          y: pos.y * finalScale + offsetY,
          w: pos.w * finalScale,
          h: pos.h * finalScale
        });
      }

      // Per-floor alignment: shift canvas Y within the layout bounding box
      // Find the total scaled height of the layout zone used by all positions
      var minResultY = Infinity, maxResultY = -Infinity;
      for (var ri = 0; ri < result.length; ri++) {
        minResultY = Math.min(minResultY, result[ri].y);
        maxResultY = Math.max(maxResultY, result[ri].y + result[ri].h);
      }
      var layoutH = maxResultY - minResultY;
      for (var ri = 0; ri < result.length; ri++) {
        var floorAlign = getFloorAlign(result[ri].index);
        if (floorAlign === 'bottom') {
          // Push to bottom of layout bounding box
          result[ri].y = minResultY + layoutH - result[ri].h;
        } else if (floorAlign === 'top') {
          // Push to top of layout bounding box
          result[ri].y = minResultY;
        }
        // 'center' keeps the position from the layout strategy
      }

      return { scale: finalScale, positions: result, type: best.type };
    }

    function layoutSingle(items) {
      var it = items[0];
      return {
        type: 'centered',
        positions: [{ index: it.index, x: 0, y: 0, w: it.w, h: it.h }]
      };
    }

    function getFloorAlign(floorIndex) {
      var fs = floorSettings[floorIndex];
      return (fs && fs.align) ? fs.align : layoutAlign;
    }

    function getFloorRotate(floorIndex) {
      var fs = floorSettings[floorIndex];
      return (fs && fs.rotate) ? fs.rotate : 0;
    }

    function alignY(align, maxH, itemH) {
      if (align === 'bottom') return maxH - itemH;
      if (align === 'top') return 0;
      return (maxH - itemH) / 2;
    }

    function layoutSideBySide(items) {
      // Place floors horizontally with gap — default center-aligned
      var gap = Math.max(items[0].w, items[0].h) * layoutGapFactor;
      var positions = [];
      var curX = 0;
      var maxH = Math.max.apply(null, items.map(function(i) { return i.h; }));
      for (var it of items) {
        // Default vertical centering; per-floor align applied in computeFloorLayout
        positions.push({ index: it.index, x: curX, y: (maxH - it.h) / 2, w: it.w, h: it.h });
        curX += it.w + gap;
      }
      return { type: 'side-by-side', positions: positions };
    }

    function layoutStacked(items) {
      // Place floors vertically with gap
      var gap = Math.max(items[0].w, items[0].h) * layoutGapFactor;
      var positions = [];
      var curY = 0;
      for (var it of items) {
        var maxW = Math.max.apply(null, items.map(function(i) { return i.w; }));
        positions.push({ index: it.index, x: (maxW - it.w) / 2, y: curY, w: it.w, h: it.h });
        curY += it.h + gap;
      }
      return { type: 'stacked', positions: positions };
    }

    function layoutDiagonal(items) {
      // Place first floor top-left, second bottom-right with slight overlap zone
      if (items.length < 2) return layoutSingle(items);
      var a = items[0], b = items[1];
      var gap = Math.max(a.w, a.h) * layoutGapFactor * 0.6;
      var positions = [
        { index: a.index, x: 0, y: 0, w: a.w, h: a.h },
        { index: b.index, x: a.w * 0.35 + gap, y: a.h * 0.35 + gap, w: b.w, h: b.h }
      ];
      // Add remaining floors (if any) stacked to the right
      var curX = Math.max(a.w, positions[1].x + b.w) + gap;
      for (var i = 2; i < items.length; i++) {
        positions.push({ index: items[i].index, x: curX, y: 0, w: items[i].w, h: items[i].h });
        curX += items[i].w + gap;
      }
      return { type: 'diagonal', positions: positions };
    }

    function layoutTriangle(items) {
      // Two on top, one centered below (or vice versa based on sizes)
      if (items.length < 3) return layoutSideBySide(items);
      var sorted = items.slice().sort(function(a, b) { return (b.w * b.h) - (a.w * a.h); });
      var gap = Math.max(sorted[0].w, sorted[0].h) * layoutGapFactor;

      // Biggest floor on top-left, second top-right, smallest centered below
      var a = sorted[0], b = sorted[1], c = sorted[2];
      var topW = a.w + gap + b.w;
      var positions = [
        { index: a.index, x: 0, y: 0, w: a.w, h: a.h },
        { index: b.index, x: a.w + gap, y: 0, w: b.w, h: b.h },
        { index: c.index, x: (topW - c.w) / 2, y: Math.max(a.h, b.h) + gap, w: c.w, h: c.h }
      ];
      // Add remaining items
      var curY = Math.max(a.h, b.h) + gap + c.h + gap;
      for (var i = 3; i < items.length; i++) {
        positions.push({ index: items[i].index, x: (topW - items[i].w) / 2, y: curY, w: items[i].w, h: items[i].h });
        curY += items[i].h + gap;
      }
      return { type: 'triangle', positions: positions };
    }

    function layoutGrid(items) {
      // Auto grid layout — default center-aligned; per-floor align applied in computeFloorLayout
      var cols = Math.ceil(Math.sqrt(items.length));
      var rows = Math.ceil(items.length / cols);
      var maxCellW = Math.max.apply(null, items.map(function(i) { return i.w; }));
      var maxCellH = Math.max.apply(null, items.map(function(i) { return i.h; }));
      var gap = maxCellW * layoutGapFactor;
      var positions = [];
      for (var i = 0; i < items.length; i++) {
        var col = i % cols;
        var row = Math.floor(i / cols);
        var cellX = col * (maxCellW + gap);
        var cellY = row * (maxCellH + gap);
        positions.push({
          index: items[i].index,
          x: cellX + (maxCellW - items[i].w) / 2,
          y: cellY + (maxCellH - items[i].h) / 2,
          w: items[i].w,
          h: items[i].h
        });
      }
      return { type: 'grid', positions: positions };
    }

    function layoutBoundingBox(positions) {
      if (!positions.length) return { x: 0, y: 0, w: 0, h: 0 };
      var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      for (var p of positions) {
        minX = Math.min(minX, p.x);
        minY = Math.min(minY, p.y);
        maxX = Math.max(maxX, p.x + p.w);
        maxY = Math.max(maxY, p.y + p.h);
      }
      return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
    }

    // ============================================================
    // RENDER PREVIEW USING LAYOUT ENGINE
    // ============================================================
    function renderPreviewThumbnails() {
      // Cleanup old preview viewers
      for (const v of previewViewers) {
        if (v.renderer) v.renderer.dispose();
      }
      previewViewers = [];

      floorsGrid.innerHTML = '';

      // Collect included floor indices
      var includedIndices = [];
      for (let i = 0; i < floors.length; i++) {
        if (!excludedFloors.has(i)) includedIndices.push(i);
      }
      if (floorOrder && floorOrder.length === includedIndices.length) {
        includedIndices = floorOrder.slice();
      }

      // Get overlay zone dimensions
      var overlayRect = floorsGrid.parentElement.getBoundingClientRect();
      var zoneW = overlayRect.width;
      var zoneH = overlayRect.height;

      if (zoneW < 10 || zoneH < 10 || includedIndices.length === 0) return;

      // Make floorsGrid a positioning container
      floorsGrid.style.position = 'relative';
      floorsGrid.style.width = '100%';
      floorsGrid.style.height = '100%';
      floorsGrid.style.display = 'block';

      // Compute layout
      currentLayout = computeFloorLayout(zoneW, zoneH, includedIndices);

      // Use ortho + colored floors during editing (steps 2-4), perspective for final (step 5)
      var useOrtho = currentWizardStep < 5;

      // Render each floor at its computed position
      for (var pi = 0; pi < currentLayout.positions.length; pi++) {
        var pos = currentLayout.positions[pi];
        var wrap = document.createElement('div');
        wrap.className = 'floor-canvas-wrap';
        wrap.style.position = 'absolute';
        wrap.style.left = pos.x + 'px';
        wrap.style.top = pos.y + 'px';
        wrap.style.width = pos.w + 'px';
        wrap.style.height = pos.h + 'px';

        floorsGrid.appendChild(wrap);

        var renderOpts = { ortho: useOrtho };
        if (useOrtho) {
          renderOpts.floorColor = EDIT_FLOOR_COLOR;
        }
        renderStaticThumbnailSized(pos.index, wrap, pos.w, pos.h, renderOpts);
      }
    }

    // ============================================================
    // INTERACTIVE 3D VIEWER (for step 4 floor review)
    // ============================================================
    function renderInteractiveViewer(floorIndex, container) {
      const result = buildFloorScene(floorIndex);
      if (!result) return null;
      const { scene, size, center, globalSize } = result;

      const rect = container.getBoundingClientRect();
      const width = Math.round(rect.width) || 400;
      const height = Math.round(rect.height) || 500;

      // Use globalSize for uniform scaling across all floors
      const ref = globalSize;
      const padding = 1.25;
      const halfW = (ref.x * padding) / 2;
      const halfZ = (ref.z * padding) / 2;
      const halfExtent = Math.max(halfW, halfZ);

      const FOV = 12;
      const aspect = width / height;
      const camera = new THREE.PerspectiveCamera(FOV, aspect, 0.01, halfExtent * 100);
      const camDist = halfExtent / Math.tan((FOV / 2) * Math.PI / 180) * 1.3;
      camera.position.set(0, camDist, camDist * 0.14);
      camera.lookAt(center);

      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
      renderer.setClearColor(0x000000, 0);
      renderer.setSize(width, height);
      renderer.setPixelRatio(dpr);
      container.appendChild(renderer.domElement);

      const controls = new THREE.OrbitControls(camera, renderer.domElement);
      controls.target.copy(center);
      controls.enableDamping = true;
      controls.dampingFactor = 0.12;
      controls.rotateSpeed = 0.8;
      controls.zoomSpeed = 1.0;
      controls.panSpeed = 0.8;
      controls.minDistance = halfExtent * 0.5;
      controls.maxDistance = halfExtent * 20;
      controls.maxPolarAngle = Math.PI * 0.85;
      controls.update();

      renderer.domElement.addEventListener('mousedown', e => e.stopPropagation());
      renderer.domElement.addEventListener('click', e => e.stopPropagation());
      renderer.domElement.addEventListener('touchstart', e => e.stopPropagation());

      let animId;
      function animate() {
        animId = requestAnimationFrame(animate);
        controls.update();
        renderer.render(scene, camera);
      }
      animate();

      return { renderer, controls, animId };
    }

    // ============================================================
    // ORTHOGRAPHIC VIEWER (for step 5 layout)
    // ============================================================
    function renderOrthographicViewer(floorIndex, container) {
      const result = buildFloorScene(floorIndex);
      if (!result) return null;
      const { scene, size, center, globalSize } = result;

      // Apply rotation if set for this floor
      var rotation = getFloorRotate(floorIndex);
      if (rotation) {
        scene.rotateY(rotation * Math.PI / 180);
      }

      const rect = container.getBoundingClientRect();
      const width = Math.round(rect.width) || 200;
      const height = Math.round(rect.height) || 260;

      // Use own size for card thumbnails — each card fills its own space
      const ref = size;
      const padding = 1.25;
      const halfW = (ref.x * padding) / 2;
      const halfZ = (ref.z * padding) / 2;
      const halfExtent = Math.max(halfW, halfZ);

      const FOV = 12;
      const aspect = width / height;
      const camera = new THREE.PerspectiveCamera(FOV, aspect, 0.01, halfExtent * 100);
      const camDist = halfExtent / Math.tan((FOV / 2) * Math.PI / 180);
      camera.position.set(0, camDist, camDist * 0.14);
      camera.lookAt(center);

      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
      renderer.setClearColor(0x000000, 0);
      renderer.setSize(width, height);
      renderer.setPixelRatio(dpr);

      renderer.render(scene, camera);
      container.appendChild(renderer.domElement);

      return { renderer };
    }

    // ============================================================
    // FRAME ADDRESS (unified preview)
    // ============================================================
    function updateFrameAddress() {
      frameStreet.textContent = addressStreet.value || '';
      frameCity.textContent = addressCity.value || '';
    }

    // ============================================================
    // LABELS (unified preview + step 6)
    // ============================================================
    let floorLabels = [];
    var labelMode = 'single'; // 'single' = 1 label total, 'per-floor' = label per floor
    var singleLabelText = 'floor plan';

    function translateFloorName(name, singleFloor) {
      const lower = name.toLowerCase().trim();
      if (/^kelder/.test(lower)) return 'basement';
      if (/^zolder/.test(lower)) return 'attic';
      if (/^berging/.test(lower)) return 'storage';
      if (/^garage/.test(lower)) return 'garage';
      if (/^dak/.test(lower)) return 'roof';
      if (/^tuin/.test(lower)) return 'garden';
      if (singleFloor) return 'floor plan';
      if (/^begane\s*grond/.test(lower)) return '1st floor';
      const ordMap = [
        [/^eerste\b/, '2nd'], [/^tweede\b/, '3rd'], [/^derde\b/, '4th'],
        [/^vierde\b/, '5th'], [/^vijfde\b/, '6th'], [/^zesde\b/, '7th'],
        [/^zevende\b/, '8th'], [/^achtste\b/, '9th'], [/^negende\b/, '10th'],
        [/^tiende\b/, '11th'],
      ];
      for (const [regex, ord] of ordMap) {
        if (regex.test(lower)) return ord + ' floor';
      }
      const numMatch = lower.match(/^(\d+)e?\s+verdieping/);
      if (numMatch) {
        const n = parseInt(numMatch[1]) + 1;
        const suf = n === 1 ? 'st' : n === 2 ? 'nd' : n === 3 ? 'rd' : 'th';
        return n + suf + ' floor';
      }
      return name.charAt(0).toLowerCase() + name.slice(1).toLowerCase();
    }

    function getIncludedFloorLabels() {
      const labels = [];
      var includedIndices = [];
      for (let i = 0; i < floors.length; i++) {
        if (excludedFloors.has(i)) continue;
        includedIndices.push(i);
      }
      // Respect custom floor order from layout step
      if (floorOrder && floorOrder.length === includedIndices.length) {
        includedIndices = floorOrder.slice();
      }
      const regularCount = includedIndices.filter(i => {
        const n = (floors[i].name || '').toLowerCase().trim();
        return !(/^(kelder|zolder|berging|garage|dak|tuin)/.test(n));
      }).length;
      const isSingle = regularCount <= 1;
      for (const i of includedIndices) {
        const rawName = floors[i].name || 'Verdieping';
        labels.push({
          index: i,
          label: translateFloorName(rawName, isSingle)
        });
      }
      return labels;
    }

    function updateFloorLabels() {
      floorLabels = getIncludedFloorLabels();

      // Update unified preview labels overlay
      unifiedLabelsOverlay.innerHTML = '';
      if (labelMode === 'single') {
        var el = document.createElement('div');
        el.className = 'label-item';
        el.textContent = singleLabelText;
        unifiedLabelsOverlay.appendChild(el);
      } else {
        for (const item of floorLabels) {
          var el = document.createElement('div');
          el.className = 'label-item';
          el.textContent = item.label;
          unifiedLabelsOverlay.appendChild(el);
        }
      }
    }

    function updateLabelsOverlayOnly() {
      unifiedLabelsOverlay.innerHTML = '';
      if (labelMode === 'single') {
        var el = document.createElement('div');
        el.className = 'label-item';
        el.textContent = singleLabelText;
        unifiedLabelsOverlay.appendChild(el);
      } else {
        for (var i = 0; i < floorLabels.length; i++) {
          var el = document.createElement('div');
          el.className = 'label-item';
          el.textContent = floorLabels[i].label;
          unifiedLabelsOverlay.appendChild(el);
        }
      }
    }

    // ============================================================
    // FLOOR EXCLUSION
    // ============================================================
    let excludedFloors = new Set();

    function isLikelySituatie(floor) {
      const name = (floor.name || '').toLowerCase();
      const excludeKeywords = ['situatie', 'site', 'tuin', 'garden', 'buitenruimte', 'omgeving', 'terrein', 'perceel'];
      return excludeKeywords.some(kw => name.includes(kw));
    }

    function isExtraFloor(floor) {
      const name = (floor.name || '').toLowerCase();
      const extraKeywords = ['berging', 'garage', 'schuur', 'zolder', 'kelder', 'storage', 'attic', 'basement'];
      return extraKeywords.some(kw => name.includes(kw));
    }

    function toggleFloorExclusion(index) {
      if (excludedFloors.has(index)) {
        excludedFloors.delete(index);
      } else {
        excludedFloors.add(index);
      }
      // Refresh unified preview thumbnails + labels
      renderPreviewThumbnails();
      updateFloorLabels();
    }

    // ============================================================
    // WIZARD
    // ============================================================
    function initWizard() {
      // Create dots
      wizardDots.innerHTML = '';
      for (let i = 1; i <= TOTAL_WIZARD_STEPS; i++) {
        const dot = document.createElement('div');
        dot.className = 'wizard-dot';
        if (i === 1) dot.classList.add('active');
        wizardDots.appendChild(dot);
      }
      updateWizardUI();
    }

    var WIZARD_STEP_TITLES = {
      1: 'Voer je Funda-link in',
      2: 'Controleer het adres',
      3: 'Bekijk de plattegronden',
      4: 'Ontwerp je indeling',
      5: 'Pas de labels aan'
    };

    function showWizardStep(n) {
      if (n < 1 || n > TOTAL_WIZARD_STEPS) return;
      currentWizardStep = n;

      // Update dynamic wizard title
      var wizTitle = document.querySelector('.mattori-configurator .wizard-title');
      if (wizTitle && WIZARD_STEP_TITLES[n]) {
        wizTitle.textContent = WIZARD_STEP_TITLES[n];
      }

      // Hide all steps
      for (let i = 1; i <= TOTAL_WIZARD_STEPS; i++) {
        const step = document.getElementById(`wizardStep${i}`);
        if (step) step.style.display = 'none';
      }

      // Remove any previous loading indicator
      var oldLoader = document.querySelector('.mattori-configurator .wizard-step-loading');
      if (oldLoader) oldLoader.remove();

      // Steps that need 3D rendering: show brief loading spinner
      var needsRender = (n === 3 || n === 4);
      if (needsRender) {
        var loader = document.createElement('div');
        loader.className = 'wizard-step-loading';
        loader.innerHTML = '<div class="step-spinner"></div>';
        var wizEl = document.getElementById('wizard');
        var navEl = document.getElementById('wizardNav');
        if (wizEl && navEl) wizEl.insertBefore(loader, navEl);
      }

      // Small delay for rendering steps so spinner is visible
      var delay = needsRender ? 80 : 0;
      setTimeout(function() {
        // Remove loader
        var ld = document.querySelector('.mattori-configurator .wizard-step-loading');
        if (ld) ld.remove();

        // Show current step
        const current = document.getElementById(`wizardStep${n}`);
        if (current) {
          current.style.display = '';
          current.style.animation = 'none';
          void current.offsetWidth;
          current.style.animation = '';
        }

        // Show/hide floors in unified preview (visible from step 4 onward)
        if (unifiedFloorsOverlay) {
          unifiedFloorsOverlay.style.display = n >= 4 ? '' : 'none';
        }

        // Show/hide address overlay (visible from step 2 onward)
        if (unifiedAddressOverlay) {
          if (n >= 2) {
            // Wait for Nexa Bold to load before showing address to prevent FOUT
            unifiedAddressOverlay.style.display = '';
            unifiedAddressOverlay.style.opacity = '0';
            nexaFontReady.then(() => {
              unifiedAddressOverlay.style.transition = 'opacity 0.15s ease';
              unifiedAddressOverlay.style.opacity = '1';
            });
          } else {
            unifiedAddressOverlay.style.display = 'none';
            unifiedAddressOverlay.style.opacity = '';
            unifiedAddressOverlay.style.transition = '';
          }
        }

        // Show/hide labels overlay (visible from step 5 onward)
        if (unifiedLabelsOverlay) {
          if (n >= 5) {
            unifiedLabelsOverlay.style.display = '';
            unifiedLabelsOverlay.style.opacity = '0';
            nexaFontReady.then(() => {
              unifiedLabelsOverlay.style.transition = 'opacity 0.15s ease';
              unifiedLabelsOverlay.style.opacity = '1';
            });
          } else {
            unifiedLabelsOverlay.style.display = 'none';
            unifiedLabelsOverlay.style.opacity = '';
            unifiedLabelsOverlay.style.transition = '';
          }
        }

        // Toggle dashed border on the active editing zone
        if (unifiedAddressOverlay) unifiedAddressOverlay.classList.toggle('zone-editing', n === 2);
        if (unifiedFloorsOverlay) unifiedFloorsOverlay.classList.toggle('zone-editing', n === 4);
        if (unifiedLabelsOverlay) unifiedLabelsOverlay.classList.toggle('zone-editing', n === 5);

        // Switch between hero image (step 1) and unified frame preview (step 2+)
        if (n === 1) {
          if (productHeroImage) productHeroImage.style.display = '';
          if (unifiedFramePreview) unifiedFramePreview.style.display = 'none';
          if (floorReviewViewerEl) floorReviewViewerEl.style.display = 'none';
        } else if (n === 3) {
          // Step 3: floor review viewer in left column
          if (productHeroImage) productHeroImage.style.display = 'none';
          if (unifiedFramePreview) unifiedFramePreview.style.display = 'none';
          if (floorReviewViewerEl) floorReviewViewerEl.style.display = '';
        } else {
          // Steps 2, 4, 5: unified frame preview
          if (productHeroImage) productHeroImage.style.display = 'none';
          if (floorReviewViewerEl) floorReviewViewerEl.style.display = 'none';
          if (floors.length > 0 && unifiedFramePreview) unifiedFramePreview.style.display = '';
        }

        // Show/hide order button + disclaimer (only on last step)
        var orderBtn = document.getElementById('btnOrder');
        var orderNotice = document.getElementById('orderNotice');
        if (orderBtn) orderBtn.style.display = n === TOTAL_WIZARD_STEPS ? '' : 'none';
        if (orderNotice) orderNotice.style.display = n === TOTAL_WIZARD_STEPS ? '' : 'none';

        // Special step initialization
        if (n === 3) {
          // Only reset viewed floors on first visit (not when returning)
          if (viewedFloors.size === 0) {
            currentFloorReviewIndex = 0;
          }
          buildThumbstrip();
          renderFloorReview();
        } else if (n === 4) {
          renderLayoutView();
          // Re-render unified preview thumbnails so they show during layout step
          renderPreviewThumbnails();
          updateFloorLabels();
        } else if (n === 5) {
          // Re-render with perspective camera (final look with depth)
          renderPreviewThumbnails();
          updateFloorLabels();
          renderLabelsFields();
        }
      }, delay);

      updateWizardUI();
    }

    function updateWizardUI() {
      // Update indicator text
      wizardStepIndicator.textContent = `Stap ${currentWizardStep} van ${TOTAL_WIZARD_STEPS}`;

      // Update dots
      const dots = wizardDots.querySelectorAll('.wizard-dot');
      dots.forEach((dot, i) => {
        const stepNum = i + 1;
        dot.classList.remove('active', 'completed');
        if (stepNum === currentWizardStep) {
          dot.classList.add('active');
        } else if (stepNum < currentWizardStep) {
          dot.classList.add('completed');
        }
      });

      // Update prev/next/order buttons
      btnWizardPrev.style.display = currentWizardStep > 1 ? '' : 'none';

      // Always reset button text (guards against leftover "Laden..." state)
      btnWizardNext.textContent = 'Volgende \u2192';
      btnWizardNext.disabled = false;

      if (currentWizardStep === TOTAL_WIZARD_STEPS) {
        // Last step: hide next, show order
        btnWizardNext.style.display = 'none';
      } else if (currentWizardStep === 1) {
        // Step 1: show next if data is loaded, or noFloorsMode button
        if (noFloorsMode) {
          btnWizardNext.textContent = 'Bestellen op goed vertrouwen \u2192';
          btnWizardNext.style.display = '';
        } else {
          btnWizardNext.textContent = 'Start met ontwerpen \u2192';
          btnWizardNext.style.display = floors.length > 0 ? '' : 'none';
        }
      } else if (currentWizardStep === 3) {
        // Step 3: hide wizard next — "Klopt, volgende" button handles navigation
        btnWizardNext.style.display = 'none';
      } else {
        btnWizardNext.style.display = '';
      }
    }

    function nextWizardStep() {
      // noFloorsMode: skip steps 2-5, go straight to cart
      if (noFloorsMode && currentWizardStep === 1) {
        submitOrder();
        return;
      }
      if (currentWizardStep < TOTAL_WIZARD_STEPS) {
        showWizardStep(currentWizardStep + 1);
      }
    }

    function prevWizardStep() {
      if (currentWizardStep > 1) {
        showWizardStep(currentWizardStep - 1);
      }
    }

    // ============================================================
    // STEP 4: Floor Review
    // ============================================================
    let thumbstripRenderers = [];

    function buildThumbstrip() {
      const strip = document.getElementById('floorReviewThumbstrip');
      if (!strip) return;

      // Dispose old renderers
      thumbstripRenderers.forEach(r => { if (r.renderer) r.renderer.dispose(); });
      thumbstripRenderers = [];
      strip.innerHTML = '';

      // Build thumbnails
      for (let i = 0; i < floors.length; i++) {
        const thumb = document.createElement('div');
        thumb.className = 'floor-thumb';
        if (i === currentFloorReviewIndex) thumb.classList.add('active');
        if (excludedFloors.has(i)) thumb.classList.add('excluded');
        thumb.addEventListener('click', () => {
          currentFloorReviewIndex = i;
          renderFloorReview();
        });
        strip.appendChild(thumb);

        // Render mini static thumbnail
        const viewer = renderStaticThumbnail(i, thumb);
        if (viewer) thumbstripRenderers.push(viewer);
      }

    }

    function updateThumbstripState() {
      const strip = document.getElementById('floorReviewThumbstrip');
      if (!strip) return;
      const thumbs = strip.querySelectorAll('.floor-thumb');
      thumbs.forEach((t, i) => {
        t.classList.toggle('active', i === currentFloorReviewIndex);
        t.classList.toggle('excluded', excludedFloors.has(i));
      });
    }

    function renderFloorReview() {
      // Cleanup previous viewer
      if (floorReviewViewer) {
        if (floorReviewViewer.animId) cancelAnimationFrame(floorReviewViewer.animId);
        if (floorReviewViewer.controls) floorReviewViewer.controls.dispose();
        if (floorReviewViewer.renderer) floorReviewViewer.renderer.dispose();
        floorReviewViewer = null;
      }
      floorReviewViewerEl.innerHTML = '';

      if (floors.length === 0) return;

      // Clamp index
      if (currentFloorReviewIndex >= floors.length) currentFloorReviewIndex = floors.length - 1;
      if (currentFloorReviewIndex < 0) currentFloorReviewIndex = 0;

      // Show loading spinner overlay
      var loadingOverlayEl = document.createElement('div');
      loadingOverlayEl.className = 'floor-review-loading-overlay';
      loadingOverlayEl.innerHTML = '<div class="review-spinner"></div>';
      floorReviewViewerEl.appendChild(loadingOverlayEl);

      // Delay render slightly so spinner is visible
      setTimeout(function() {
        // Remove spinner
        var spinner = floorReviewViewerEl.querySelector('.floor-review-loading-overlay');
        if (spinner) spinner.remove();

        // Render interactive viewer
        floorReviewViewer = renderInteractiveViewer(currentFloorReviewIndex, floorReviewViewerEl);

        // Add zoom/rotate hint label
        var hint = document.createElement('div');
        hint.className = 'floor-review-hint';
        hint.textContent = 'Klik en sleep om te draaien \u00B7 scroll om te zoomen';
        floorReviewViewerEl.appendChild(hint);

        // Show/hide excluded overlay
        updateFloorReviewExcludedOverlay();
      }, 60);

      // Track viewed floors and update wizard UI
      viewedFloors.add(currentFloorReviewIndex);
      updateWizardUI();

      // Update thumbstrip state (highlight active, don't rebuild)
      updateThumbstripState();

      // Reset issue panel
      var issueDetails = document.getElementById('floorIssueDetails');
      if (issueDetails) issueDetails.style.display = 'none';
      var issueBtn = document.getElementById('btnFloorIssue');
      if (issueBtn) issueBtn.classList.remove('active');

      // Update button texts and exclude highlight
      var confirmBtn = document.getElementById('btnFloorConfirm');
      if (confirmBtn) {
        confirmBtn.textContent = 'Wel meenemen ✓';
      }
      var excludeBtn = document.getElementById('btnFloorExclude');
      if (excludeBtn) {
        excludeBtn.classList.toggle('active', excludedFloors.has(currentFloorReviewIndex));
      }
    }

    function updateFloorReviewExcludedOverlay() {
      if (!floorReviewViewerEl) return;
      var existing = floorReviewViewerEl.querySelector('.floor-review-excluded');
      if (existing) existing.remove();
      if (excludedFloors.has(currentFloorReviewIndex)) {
        var overlay = document.createElement('div');
        overlay.className = 'floor-review-excluded';
        overlay.innerHTML = '<span>Uitgezet</span>';
        floorReviewViewerEl.appendChild(overlay);
      }
    }

    function navigateFloorReview(direction) {
      currentFloorReviewIndex += direction;
      if (currentFloorReviewIndex < 0) currentFloorReviewIndex = floors.length - 1;
      if (currentFloorReviewIndex >= floors.length) currentFloorReviewIndex = 0;
      renderFloorReview();
    }

    // "Wel meenemen" — include current floor, navigate to next or advance to step 4
    function confirmFloor() {
      ensureDomRefs();
      excludedFloors.delete(currentFloorReviewIndex); // Re-include if previously excluded
      viewedFloors.add(currentFloorReviewIndex);
      // Close any open issue panel
      var details = document.getElementById('floorIssueDetails');
      if (details) details.style.display = 'none';
      var issueBtn = document.getElementById('btnFloorIssue');
      if (issueBtn) issueBtn.classList.remove('active');
      // Update preview thumbnails
      renderPreviewThumbnails();
      updateFloorLabels();
      // Navigate to next floor, or auto-advance to step 4
      if (currentFloorReviewIndex < floors.length - 1) {
        currentFloorReviewIndex++;
        renderFloorReview();
      } else {
        // Last floor — go to step 4 automatically
        showWizardStep(4);
      }
    }

    // "Klopt niet" toggle — show/hide issue textarea
    function toggleFloorIssue() {
      var details = document.getElementById('floorIssueDetails');
      var issueBtn = document.getElementById('btnFloorIssue');
      if (!details) return;
      var isOpen = details.style.display !== 'none';
      details.style.display = isOpen ? 'none' : '';
      if (issueBtn) issueBtn.classList.toggle('active', !isOpen);
    }

    // "Niet meenemen" — exclude floor and advance
    function excludeFloor() {
      ensureDomRefs();
      excludedFloors.add(currentFloorReviewIndex);
      viewedFloors.add(currentFloorReviewIndex);
      // Close any open issue panel
      var details = document.getElementById('floorIssueDetails');
      if (details) details.style.display = 'none';
      var issueBtn = document.getElementById('btnFloorIssue');
      if (issueBtn) issueBtn.classList.remove('active');
      // Update preview thumbnails to show excluded state
      renderPreviewThumbnails();
      updateFloorLabels();
      // Navigate to next floor or advance to step 4
      if (currentFloorReviewIndex < floors.length - 1) {
        currentFloorReviewIndex++;
        renderFloorReview();
      } else {
        showWizardStep(4);
      }
    }

    // ============================================================
    // ADMIN: Toggle frame image (One.png ↔ Two.png)
    // ============================================================
    var FRAME_IMG_ONE = 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/One_6ca91ccb-41c3-4f74-b387-4f0b5f64292a.png?v=1771340361';
    var FRAME_IMG_TWO = 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Two_1a0d36a8-e861-4cec-8417-f251458bb134.png?v=1771340362';

    function toggleAdminFrame() {
      ensureDomRefs();
      var cb = document.getElementById('adminUseAltFrame');
      var img = document.getElementById('unifiedFrameImage');
      if (!cb || !img) return;
      img.src = cb.checked ? FRAME_IMG_TWO : FRAME_IMG_ONE;
    }

    function toggleLayoutAlign() {
      ensureDomRefs();
      var cb = document.getElementById('adminAlignBottom');
      layoutAlign = (cb && cb.checked) ? 'bottom' : 'center';
      // Re-render preview with new alignment
      renderPreviewThumbnails();
      updateFloorLabels();
    }

    // Clear control panel, show spinner, do work, rebuild
    function updatePreviewWithLoading(fn) {
      // Clear layout viewer and show spinner (same pattern as step transitions)
      if (floorLayoutViewer) {
        // Dispose old renderers first
        for (var v of layoutViewers) { if (v.renderer) v.renderer.dispose(); }
        layoutViewers = [];
        floorLayoutViewer.innerHTML = '';
        var loader = document.createElement('div');
        loader.className = 'wizard-step-loading';
        loader.innerHTML = '<div class="step-spinner"></div>';
        floorLayoutViewer.appendChild(loader);
      }
      setTimeout(function() {
        fn();
        renderLayoutView();
      }, 60);
    }

    function setLayoutGap(value) {
      layoutGapFactor = parseFloat(value);
      updatePreviewWithLoading(function() {
        renderPreviewThumbnails();
        updateFloorLabels();
      });
    }

    function setGlobalAlign(align) {
      layoutAlign = align;
      // Clear any per-floor overrides
      for (var key in floorSettings) {
        if (floorSettings[key]) delete floorSettings[key].align;
      }
      updatePreviewWithLoading(function() {
        renderPreviewThumbnails();
        updateFloorLabels();
      });
    }

    function rotateFloor(floorIndex) {
      if (!floorSettings[floorIndex]) floorSettings[floorIndex] = {};
      var current = getFloorRotate(floorIndex);
      floorSettings[floorIndex].rotate = current === 0 ? 180 : 0;
      updatePreviewWithLoading(function() {
        renderPreviewThumbnails();
        updateFloorLabels();
      });
    }

    // ============================================================
    // STEP 4: Layout View (orthographic top-down)
    // ============================================================
    // Track custom floor order (set by drag & drop in layout step)
    var floorOrder = null; // null = default order

    function moveFloorInOrder(fromIdx, toIdx) {
      if (!floorOrder || toIdx < 0 || toIdx >= floorOrder.length) return;
      var item = floorOrder.splice(fromIdx, 1)[0];
      floorOrder.splice(toIdx, 0, item);
      updatePreviewWithLoading(function() {
        renderPreviewThumbnails();
        updateFloorLabels();
        // Highlight moved card briefly after renderLayoutView rebuilds cards
        setTimeout(function() {
          var cards = floorLayoutViewer.querySelectorAll('.floor-layout-card');
          if (cards[toIdx]) {
            cards[toIdx].classList.add('just-moved');
            setTimeout(function() { cards[toIdx].classList.remove('just-moved'); }, 600);
          }
        }, 20);
      });
    }

    function renderLayoutView() {
      // Cleanup old layout viewers
      for (const v of layoutViewers) {
        if (v.renderer) v.renderer.dispose();
      }
      layoutViewers = [];
      floorLayoutViewer.innerHTML = '';

      // Global alignment control
      var alignRow = document.createElement('div');
      alignRow.className = 'floor-layout-gap-row';
      var alignLabel = document.createElement('span');
      alignLabel.className = 'floor-gap-label';
      alignLabel.textContent = 'Uitlijning';

      var alignGroup = document.createElement('div');
      alignGroup.className = 'floor-align-group';

      var alignValues = ['bottom', 'center', 'top'];
      var alignIcons = {
        top: '<svg width="16" height="16" viewBox="0 0 16 16"><rect x="1" y="1" width="6" height="14" rx="1.2" fill="currentColor" opacity="0.9"/><rect x="9" y="1" width="6" height="9" rx="1.2" fill="currentColor" opacity="0.45"/></svg>',
        center: '<svg width="16" height="16" viewBox="0 0 16 16"><rect x="1" y="1" width="6" height="14" rx="1.2" fill="currentColor" opacity="0.9"/><rect x="9" y="3.5" width="6" height="9" rx="1.2" fill="currentColor" opacity="0.45"/></svg>',
        bottom: '<svg width="16" height="16" viewBox="0 0 16 16"><rect x="1" y="1" width="6" height="14" rx="1.2" fill="currentColor" opacity="0.9"/><rect x="9" y="6" width="6" height="9" rx="1.2" fill="currentColor" opacity="0.45"/></svg>'
      };
      var alignTitles = { top: 'Bovenkant uitlijnen', center: 'Midden uitlijnen', bottom: 'Onderkant uitlijnen' };

      for (var ai = 0; ai < alignValues.length; ai++) {
        var alignBtn = document.createElement('button');
        alignBtn.type = 'button';
        alignBtn.className = 'floor-align-btn' + (alignValues[ai] === layoutAlign ? ' active' : '');
        alignBtn.innerHTML = alignIcons[alignValues[ai]];
        alignBtn.title = alignTitles[alignValues[ai]];

        (function(val) {
          alignBtn.addEventListener('click', function() {
            setGlobalAlign(val);
            var siblings = this.parentElement.querySelectorAll('.floor-align-btn');
            for (var s = 0; s < siblings.length; s++) siblings[s].classList.remove('active');
            this.classList.add('active');
          });
        })(alignValues[ai]);

        alignGroup.appendChild(alignBtn);
      }

      alignRow.appendChild(alignLabel);
      alignRow.appendChild(alignGroup);
      floorLayoutViewer.appendChild(alignRow);

      // Gap control with +/- buttons
      var gapRow = document.createElement('div');
      gapRow.className = 'floor-layout-gap-row';
      var gapLabel = document.createElement('span');
      gapLabel.className = 'floor-gap-label';
      gapLabel.textContent = 'Tussenruimte';

      var gapMinus = document.createElement('button');
      gapMinus.type = 'button';
      gapMinus.className = 'floor-gap-btn';
      gapMinus.textContent = '−';
      gapMinus.addEventListener('click', function() {
        setLayoutGap(Math.round(Math.max(0, layoutGapFactor * 100 - 2)) / 100);
      });

      var gapValue = document.createElement('span');
      gapValue.className = 'floor-gap-value';
      gapValue.textContent = Math.round(layoutGapFactor * 100) + '%';

      var gapPlus = document.createElement('button');
      gapPlus.type = 'button';
      gapPlus.className = 'floor-gap-btn';
      gapPlus.textContent = '+';
      gapPlus.addEventListener('click', function() {
        setLayoutGap(Math.round(Math.min(25, layoutGapFactor * 100 + 2)) / 100);
      });

      gapRow.appendChild(gapLabel);
      gapRow.appendChild(gapMinus);
      gapRow.appendChild(gapValue);
      gapRow.appendChild(gapPlus);
      floorLayoutViewer.appendChild(gapRow);

      // Build ordered list: included first (in custom order), then excluded
      var includedIndices = [];
      var excludedIndices = [];
      for (var i = 0; i < floors.length; i++) {
        if (excludedFloors.has(i)) {
          excludedIndices.push(i);
        } else {
          includedIndices.push(i);
        }
      }

      // Use custom order if set for included floors
      if (floorOrder && floorOrder.length === includedIndices.length) {
        includedIndices = floorOrder.slice();
      } else {
        floorOrder = includedIndices.slice();
      }

      // Render all floors: included first, then excluded
      var allIndices = includedIndices.concat(excludedIndices);
      var includedCount = 0;

      for (var idx = 0; idx < allIndices.length; idx++) {
        var floorIdx = allIndices[idx];
        var isExcluded = excludedFloors.has(floorIdx);
        var card = document.createElement('div');
        card.className = 'floor-layout-card' + (isExcluded ? ' excluded' : '');

        // Number (only for included)
        var numEl = document.createElement('div');
        numEl.className = 'floor-layout-number';
        if (!isExcluded) {
          includedCount++;
          numEl.textContent = includedCount;
        } else {
          numEl.textContent = '–';
        }

        // Canvas
        var canvasWrap = document.createElement('div');
        canvasWrap.className = 'floor-layout-canvas-wrap';

        // Floor name label
        var nameEl = document.createElement('div');
        nameEl.className = 'floor-layout-name';
        nameEl.textContent = floors[floorIdx].name || 'Verdieping ' + (idx + 1);

        // Include/exclude checkbox
        var includeWrap = document.createElement('div');
        includeWrap.className = 'floor-layout-include';
        var includeCb = document.createElement('input');
        includeCb.type = 'checkbox';
        includeCb.checked = !isExcluded;
        var includeLabel = document.createElement('label');
        includeLabel.textContent = 'Meenemen';

        // Closure for checkbox handler
        (function(fi) {
          includeCb.addEventListener('change', function() {
            // Show loading in layout viewer
            floorLayoutViewer.innerHTML = '';
            var loader = document.createElement('div');
            loader.className = 'wizard-step-loading';
            loader.innerHTML = '<div class="step-spinner"></div>';
            floorLayoutViewer.appendChild(loader);
            setTimeout(function() {
              toggleFloorExclusion(fi);
              renderLayoutView();
              buildThumbstrip();
            }, 80);
          });
        })(floorIdx);

        includeWrap.appendChild(includeCb);
        includeWrap.appendChild(includeLabel);

        // Per-floor controls (rotation only) — only for included floors
        var floorControls = document.createElement('div');
        floorControls.className = 'floor-layout-controls';

        if (!isExcluded) {
          // 180° flip button
          var rotBtn = document.createElement('button');
          rotBtn.type = 'button';
          rotBtn.className = 'floor-rotate-btn';
          rotBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 16 16"><path d="M3 4L8 1L13 4" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"/><path d="M13 12L8 15L3 12" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"/><line x1="8" y1="1" x2="8" y2="15" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2 1.5"/></svg>';
          rotBtn.title = '180° draaien';

          (function(fi) {
            rotBtn.addEventListener('click', function() {
              rotateFloor(fi);
            });
          })(floorIdx);

          floorControls.appendChild(rotBtn);
        }

        // Up/down arrows (only for included)
        var arrows = document.createElement('div');
        arrows.className = 'floor-layout-arrows';

        if (!isExcluded) {
          var orderIdx = includedIndices.indexOf(floorIdx);
          var btnUp = document.createElement('button');
          btnUp.type = 'button';
          btnUp.innerHTML = '&#9650;';
          btnUp.disabled = orderIdx === 0;

          var btnDown = document.createElement('button');
          btnDown.type = 'button';
          btnDown.innerHTML = '&#9660;';
          btnDown.disabled = orderIdx === includedIndices.length - 1;

          (function(currentIdx) {
            btnUp.addEventListener('click', function() { moveFloorInOrder(currentIdx, currentIdx - 1); });
            btnDown.addEventListener('click', function() { moveFloorInOrder(currentIdx, currentIdx + 1); });
          })(orderIdx);

          arrows.appendChild(btnUp);
          arrows.appendChild(btnDown);
        }

        card.appendChild(numEl);
        card.appendChild(canvasWrap);
        card.appendChild(nameEl);
        card.appendChild(includeWrap);
        card.appendChild(floorControls);
        card.appendChild(arrows);
        floorLayoutViewer.appendChild(card);

        var viewer = renderOrthographicViewer(floorIdx, canvasWrap);
        if (viewer) layoutViewers.push(viewer);
      }
    }

    // ============================================================
    // STEP 5: Labels editing
    // ============================================================
    function setLabelMode(mode) {
      labelMode = mode;
      renderLabelsFieldsContent();
      updateFloorLabels();
    }

    function renderLabelsFields() {
      floorLabels = getIncludedFloorLabels();

      var step5 = document.getElementById('wizardStep5');
      if (!step5) return;

      // Find or create container after step-description
      var container = step5.querySelector('.labels-step-container');
      if (!container) {
        container = document.createElement('div');
        container.className = 'labels-step-container';
        step5.appendChild(container);
      }
      container.innerHTML = '';

      // Toggle switch
      var toggle = document.createElement('div');
      toggle.className = 'label-mode-toggle';

      var labelSingle = document.createElement('span');
      labelSingle.className = 'label-mode-label' + (labelMode === 'single' ? ' active' : '');
      labelSingle.textContent = '1 label';

      var switchWrap = document.createElement('label');
      switchWrap.className = 'label-mode-switch';
      var switchInput = document.createElement('input');
      switchInput.type = 'checkbox';
      switchInput.checked = labelMode === 'per-floor';
      var slider = document.createElement('span');
      slider.className = 'switch-slider';
      switchWrap.appendChild(switchInput);
      switchWrap.appendChild(slider);

      var labelMulti = document.createElement('span');
      labelMulti.className = 'label-mode-label' + (labelMode === 'per-floor' ? ' active' : '');
      labelMulti.textContent = 'Per plattegrond';

      switchInput.addEventListener('change', function() {
        setLabelMode(this.checked ? 'per-floor' : 'single');
      });
      labelSingle.addEventListener('click', function() {
        switchInput.checked = false;
        setLabelMode('single');
      });
      labelMulti.addEventListener('click', function() {
        switchInput.checked = true;
        setLabelMode('per-floor');
      });

      toggle.appendChild(labelSingle);
      toggle.appendChild(switchWrap);
      toggle.appendChild(labelMulti);
      container.appendChild(toggle);

      // Fields container
      var fieldsDiv = document.createElement('div');
      fieldsDiv.className = 'labels-fields';
      fieldsDiv.id = 'labelsFieldsInner';
      container.appendChild(fieldsDiv);

      renderLabelsFieldsContent();
      updateFloorLabels();
    }

    function renderLabelsFieldsContent() {
      var fieldsDiv = document.getElementById('labelsFieldsInner');
      if (!fieldsDiv) return;
      fieldsDiv.innerHTML = '';

      if (labelMode === 'single') {
        // Single label mode — one text field
        var row = document.createElement('div');
        row.className = 'label-field-row';

        var span = document.createElement('span');
        span.textContent = 'Label';

        var inp = document.createElement('input');
        inp.type = 'text';
        inp.value = singleLabelText;
        inp.placeholder = 'floor plan';
        inp.addEventListener('input', function() {
          singleLabelText = inp.value;
          updateLabelsOverlayOnly();
        });

        row.appendChild(span);
        row.appendChild(inp);
        fieldsDiv.appendChild(row);
      } else {
        // Per-floor mode — one field per floor
        for (var li = 0; li < floorLabels.length; li++) {
          (function(idx) {
            var item = floorLabels[idx];
            var row = document.createElement('div');
            row.className = 'label-field-row';

            var span = document.createElement('span');
            span.textContent = floors[item.index].name;

            var inp = document.createElement('input');
            inp.type = 'text';
            inp.value = item.label;
            inp.placeholder = item.label;
            inp.addEventListener('input', function() {
              floorLabels[idx].label = inp.value;
              updateLabelsOverlayOnly();
            });

            row.appendChild(span);
            row.appendChild(inp);
            fieldsDiv.appendChild(row);
          })(li);
        }
      }
    }

    // ============================================================
    // FILE UPLOAD HANDLING
    // ============================================================
    async function handleFileUpload(file) {
      clearError();
      const name = file.name.toLowerCase();
      if (!name.endsWith('.fml') && !name.endsWith('.json')) {
        setError('Upload een geldig FML bestand.');
        return;
      }
      showLoading();
      try {
        const text = await file.text();
        const data = JSON.parse(text);
        if (!data?.floors?.length) throw new Error('Geen verdiepingen gevonden in het bestand.');
        if (!data.floors.some(f => f?.designs?.length)) throw new Error('Geen geldige plattegronden gevonden.');
        originalFmlData = data;
        originalFileName = file.name;
        fileLabel.textContent = `📄 ${file.name}`;
        const addr = parseAddressFromFML(data);
        if (addr) {
          addressStreet.value = addr.street;
          addressCity.value = addr.city;
          currentAddress = addr;
        } else {
          addressStreet.value = '';
          addressCity.value = '';
        }
        processFloors(data);
        showToast(`✓ ${file.name} geladen`);
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Er ging iets mis bij het laden.');
      } finally {
        hideLoading();
      }
    }

    // ============================================================
    // OBJ EXPORT
    // ============================================================
    function generateFloorOBJ(floor) {
      let vertices = [];
      let faces = [];
      let vertexIndex = 1;

      const WALL_HEIGHT = 2.8;
      const SCALE = 0.01;

      function addFace(a, b, c, d) { faces.push(`f ${a} ${b} ${c} ${d}`); }
      function addTriFace(a, b, c) { faces.push(`f ${a} ${b} ${c}`); }

      // clipA/clipB: direction {dx,dy} of intersecting wall at endpoint A/B (or null)
      // When set, the end-face is angled to align with the intersecting wall
      function createWallBox(x1, y1, x2, y2, bottomZ, topZ, halfThickness, normalX, normalY) {
        if (topZ <= bottomZ) return;
        const corners = [
          { x: x1 + normalX * halfThickness, y: y1 + normalY * halfThickness }, // A + normal
          { x: x2 + normalX * halfThickness, y: y2 + normalY * halfThickness }, // B + normal
          { x: x2 - normalX * halfThickness, y: y2 - normalY * halfThickness }, // B - normal
          { x: x1 - normalX * halfThickness, y: y1 - normalY * halfThickness }  // A - normal
        ];
        for (const c of corners) vertices.push(`v ${c.x.toFixed(4)} ${bottomZ.toFixed(4)} ${c.y.toFixed(4)}`);
        for (const c of corners) vertices.push(`v ${c.x.toFixed(4)} ${topZ.toFixed(4)} ${c.y.toFixed(4)}`);
        const base = vertexIndex;
        addFace(base + 3, base + 2, base + 1, base + 0);
        addFace(base + 4, base + 5, base + 6, base + 7);
        addFace(base + 0, base + 1, base + 5, base + 4);
        addFace(base + 2, base + 3, base + 7, base + 6);
        addFace(base + 3, base + 0, base + 4, base + 7);
        addFace(base + 1, base + 2, base + 6, base + 5);
        vertexIndex += 8;
      }

      function createOpeningFrame(x1, y1, x2, y2, bottomZ, topZ, frameSize, frameThickness, fnx, fny) {
        const fdx = x2 - x1, fdy = y2 - y1;
        const flen = Math.hypot(fdx, fdy);
        if (flen < 0.001) return;
        const fux = fdx / flen, fuy = fdy / flen;
        createWallBox(x1, y1, x1 + fux * frameSize, y1 + fuy * frameSize, bottomZ, topZ, frameThickness, fnx, fny);
        createWallBox(x2 - fux * frameSize, y2 - fuy * frameSize, x2, y2, bottomZ, topZ, frameThickness, fnx, fny);
        createWallBox(x1 + fux * frameSize, y1 + fuy * frameSize, x2 - fux * frameSize, y2 - fuy * frameSize, topZ - frameSize, topZ, frameThickness, fnx, fny);
        if (bottomZ > 0.01) {
          createWallBox(x1 + fux * frameSize, y1 + fuy * frameSize, x2 - fux * frameSize, y2 - fuy * frameSize, bottomZ, bottomZ + frameSize, frameThickness, fnx, fny);
        }
      }

      const design = floor.design;
      const wallBBox = computeWallBBox(design);
      const walls = flattenWalls(design.walls ?? []);
      const bbox = floor.bbox;
      const centerX = (bbox.minX + bbox.maxX) / 2;
      const centerY = (bbox.minY + bbox.maxY) / 2;

      // ── Polygon-union wall rendering ──────────────────────────
      // Walls WITHOUT openings → boolean union → extrude as polygon
      // Walls WITH openings → render as individual segment boxes

      const solidWalls = walls.filter(w => !(w.openings && w.openings.length > 0));
      const openingWalls = walls.filter(w => w.openings && w.openings.length > 0);

      // Union all solid walls into merged 2D polygons
      const wallUnion = computeWallUnion(solidWalls, walls);

      // Ear-clipping triangulation for concave 2D polygons
      function earClipTriangulate(pts) {
        // pts = array of {x, y} — must be in CCW order
        // Returns array of [i, j, k] index triples
        const n = pts.length;
        if (n < 3) return [];
        if (n === 3) return [[0, 1, 2]];

        // Ensure CCW winding
        let area = 0;
        for (let i = 0; i < n; i++) {
          const j = (i + 1) % n;
          area += pts[i].x * pts[j].y - pts[j].x * pts[i].y;
        }
        const indices = [];
        for (let i = 0; i < n; i++) indices.push(i);
        if (area < 0) indices.reverse(); // was CW, flip to CCW

        const tris = [];
        let remaining = indices.slice();

        function cross2d(o, a, b) {
          return (a.x - o.x) * (b.y - o.y) - (a.y - o.y) * (b.x - o.x);
        }

        function pointInTriangle(px, py, ax, ay, bx, by, cx, cy) {
          const d1 = (px - bx) * (ay - by) - (ax - bx) * (py - by);
          const d2 = (px - cx) * (by - cy) - (bx - cx) * (py - cy);
          const d3 = (px - ax) * (cy - ay) - (cx - ax) * (py - ay);
          const hasNeg = (d1 < 0) || (d2 < 0) || (d3 < 0);
          const hasPos = (d1 > 0) || (d2 > 0) || (d3 > 0);
          return !(hasNeg && hasPos);
        }

        let safety = remaining.length * 3;
        while (remaining.length > 3 && safety-- > 0) {
          let earFound = false;
          for (let i = 0; i < remaining.length; i++) {
            const prev = remaining[(i - 1 + remaining.length) % remaining.length];
            const curr = remaining[i];
            const next = remaining[(i + 1) % remaining.length];
            const p = pts[prev], c = pts[curr], nx2 = pts[next];

            // Is this a convex vertex?
            if (cross2d(p, c, nx2) <= 0) continue;

            // Does any other vertex fall inside this triangle?
            let inside = false;
            for (let j = 0; j < remaining.length; j++) {
              const vi = remaining[j];
              if (vi === prev || vi === curr || vi === next) continue;
              if (pointInTriangle(pts[vi].x, pts[vi].y, p.x, p.y, c.x, c.y, nx2.x, nx2.y)) {
                inside = true;
                break;
              }
            }
            if (inside) continue;

            tris.push([prev, curr, next]);
            remaining.splice(i, 1);
            earFound = true;
            break;
          }
          if (!earFound) break; // degenerate polygon
        }
        if (remaining.length === 3) {
          tris.push([remaining[0], remaining[1], remaining[2]]);
        }
        return tris;
      }

      // Helper: extrude a 2D polygon ring to 3D wall geometry
      function extrudeWallPoly(ring, bottomZ, topZ) {
        if (ring.length < 3) return;
        // ring = array of [x,y] in world coords; may or may not repeat first point
        const pts = ring.slice();
        // Remove closing duplicate if present
        const first = pts[0], last = pts[pts.length - 1];
        if (Math.hypot(first[0] - last[0], first[1] - last[1]) < 0.01) pts.pop();
        if (pts.length < 3) return;

        const n = pts.length;
        // Convert to OBJ coords
        const objPts = pts.map(p => ({
          x: (p[0] - centerX) * SCALE,
          y: (p[1] - centerY) * SCALE
        }));

        // Bottom + top vertices
        const baseBot = vertexIndex;
        for (const p of objPts) {
          vertices.push(`v ${p.x.toFixed(4)} ${bottomZ.toFixed(4)} ${p.y.toFixed(4)}`);
          vertexIndex++;
        }
        const baseTop = vertexIndex;
        for (const p of objPts) {
          vertices.push(`v ${p.x.toFixed(4)} ${topZ.toFixed(4)} ${p.y.toFixed(4)}`);
          vertexIndex++;
        }

        // Triangulate the polygon (handles concave shapes correctly)
        const tris = earClipTriangulate(objPts);
        for (const [a, b, c] of tris) {
          addTriFace(baseBot + a, baseBot + c, baseBot + b); // bottom face (flip winding)
          addTriFace(baseTop + a, baseTop + b, baseTop + c); // top face
        }

        // Side faces
        for (let i = 0; i < n; i++) {
          const j = (i + 1) % n;
          addFace(baseBot + i, baseBot + j, baseTop + j, baseTop + i);
        }
      }

      faces.push('g walls');

      // Render unioned solid walls
      for (const polygon of wallUnion) {
        // polygon = array of rings. First ring = outer, rest = holes
        for (let ri = 0; ri < polygon.length; ri++) {
          const ring = polygon[ri];
          if (ri === 0) {
            // Outer ring → extrude as wall
            extrudeWallPoly(ring, 0, WALL_HEIGHT);
          }
          // Holes are interior cutouts — skip for wall geometry
          // (they represent empty space inside wall outlines, rare but possible)
        }
      }

      // Render walls WITH openings as individual boxes (with L-junction extension)
      for (const wall of openingWalls) {
        // Compute L-junction extension for opening walls too (same logic as solid walls)
        const _wdx = wall.b.x - wall.a.x, _wdy = wall.b.y - wall.a.y;
        const _wlen = Math.hypot(_wdx, _wdy);
        const _wIsDiag = _wlen > 0.1 && Math.min(Math.abs(_wdx / _wlen), Math.abs(_wdy / _wlen)) > 0.15;
        let owExtA = 0, owExtB = 0;
        if (!_wIsDiag && _wlen > 0.1) {
          const _ux = _wdx / _wlen, _uy = _wdy / _wlen;
          for (const other of walls) {
            if (other === wall) continue;
            const odx = other.b.x - other.a.x, ody = other.b.y - other.a.y;
            const olen = Math.hypot(odx, ody);
            if (olen < 0.1) continue;
            if (Math.min(Math.abs(odx / olen), Math.abs(ody / olen)) > 0.15) continue; // skip diag
            const otherHt = (other.thickness ?? 20) / 2;
            if (Math.hypot(wall.a.x - other.a.x, wall.a.y - other.a.y) < 3 ||
                Math.hypot(wall.a.x - other.b.x, wall.a.y - other.b.y) < 3) {
              owExtA = Math.max(owExtA, otherHt);
            }
            if (Math.hypot(wall.b.x - other.a.x, wall.b.y - other.a.y) < 3 ||
                Math.hypot(wall.b.x - other.b.x, wall.b.y - other.b.y) < 3) {
              owExtB = Math.max(owExtB, otherHt);
            }
          }
        }
        const ax = (wall.a.x - centerX) * SCALE - (_wlen > 0.1 ? (_wdx / _wlen) * owExtA * SCALE : 0);
        const ay = (wall.a.y - centerY) * SCALE - (_wlen > 0.1 ? (_wdy / _wlen) * owExtA * SCALE : 0);
        const bx = (wall.b.x - centerX) * SCALE + (_wlen > 0.1 ? (_wdx / _wlen) * owExtB * SCALE : 0);
        const by = (wall.b.y - centerY) * SCALE + (_wlen > 0.1 ? (_wdy / _wlen) * owExtB * SCALE : 0);
        const wdx = bx - ax, wdy = by - ay;
        const wlen = Math.hypot(wdx, wdy);
        if (wlen < 0.001) continue;
        const wnx = -wdy / wlen, wny = wdx / wlen;
        const halfThick = (wall.thickness ?? 20) / 2 * SCALE;
        const wallWorldLen = Math.hypot(wall.b.x - wall.a.x, wall.b.y - wall.a.y);

        const sortedOpenings = (wall.openings ?? []).map(op => {
          const t = op.t ?? 0.5;
          const halfW = (op.width ?? 90) / 2 / wallWorldLen;
          const height = (op.height ?? (op.type === "door" ? 210 : 120)) * SCALE;
          const elevation = (op.elevation ?? (op.type === "door" ? 0 : 90)) * SCALE;
          return {
            startT: Math.max(0, t - halfW),
            endT: Math.min(1, t + halfW),
            bottomZ: elevation,
            topZ: elevation + height,
            type: op.type ?? "door"
          };
        }).sort((a, b) => a.startT - b.startT);

        let currentT = 0;
        for (const op of sortedOpenings) {
          if (op.startT > currentT) {
            const sAx = ax + wdx * currentT, sAy = ay + wdy * currentT;
            const sBx = ax + wdx * op.startT, sBy = ay + wdy * op.startT;
            createWallBox(sAx, sAy, sBx, sBy, 0, WALL_HEIGHT, halfThick, wnx, wny);
          }
          const oAx = ax + wdx * op.startT, oAy = ay + wdy * op.startT;
          const oBx = ax + wdx * op.endT, oBy = ay + wdy * op.endT;
          if (op.bottomZ > 0.01) {
            createWallBox(oAx, oAy, oBx, oBy, 0, op.bottomZ, halfThick, wnx, wny);
          }
          if (op.topZ < WALL_HEIGHT - 0.01) {
            createWallBox(oAx, oAy, oBx, oBy, op.topZ, WALL_HEIGHT, halfThick, wnx, wny);
          }
          const frameSize = 0.05;
          const frameThick = halfThick * 0.8;
          createOpeningFrame(oAx, oAy, oBx, oBy, op.bottomZ, op.topZ, frameSize, frameThick, wnx, wny);
          currentT = Math.max(currentT, op.endT);
        }
        if (currentT < 1) {
          const sAx = ax + wdx * currentT, sAy = ay + wdy * currentT;
          createWallBox(sAx, sAy, bx, by, 0, WALL_HEIGHT, halfThick, wnx, wny);
        }
      }

      faces.push('g floor');

      const FLOOR_THICKNESS = 0.30;

      function extrudePolygon(poly) {
        if (poly.length < 3) return;
        const n = poly.length;

        let cx = 0, cy = 0;
        for (const pt of poly) { cx += pt.x; cy += pt.y; }
        cx /= n; cy /= n;

        const baseBot = vertexIndex;
        for (const pt of poly) {
          vertices.push(`v ${pt.x.toFixed(4)} ${(-FLOOR_THICKNESS).toFixed(4)} ${pt.y.toFixed(4)}`);
          vertexIndex++;
        }
        vertices.push(`v ${cx.toFixed(4)} ${(-FLOOR_THICKNESS).toFixed(4)} ${cy.toFixed(4)}`);
        const botCenter = vertexIndex;
        vertexIndex++;

        const baseTop = vertexIndex;
        for (const pt of poly) {
          vertices.push(`v ${pt.x.toFixed(4)} ${(0).toFixed(4)} ${pt.y.toFixed(4)}`);
          vertexIndex++;
        }
        vertices.push(`v ${cx.toFixed(4)} ${(0).toFixed(4)} ${cy.toFixed(4)}`);
        const topCenter = vertexIndex;
        vertexIndex++;

        for (let i = 0; i < n; i++) {
          const j = (i + 1) % n;
          addTriFace(botCenter, baseBot + j, baseBot + i);
        }
        for (let i = 0; i < n; i++) {
          const j = (i + 1) % n;
          addTriFace(topCenter, baseTop + i, baseTop + j);
        }
        for (let i = 0; i < n; i++) {
          const j = (i + 1) % n;
          addFace(baseBot + i, baseBot + j, baseTop + j, baseTop + i);
        }
      }

      const floorVoids = floor.voids ?? [];

      {
        const floorSources = [];

        for (const area of design.areas ?? []) {
          const tessellated = tessellateSurfacePoly(area.poly ?? []);
          if (tessellated.length >= 3) floorSources.push(tessellated);
        }

        for (const surface of design.surfaces ?? []) {
          if (surface.isCutout) continue; // cutouts are voids, not floor sources
          if (isSurfaceOutsideWalls(surface, wallBBox)) continue;
          const sName = (surface.name ?? "").trim();
          const cName = (surface.customName ?? "").trim();
          if (!sName && !cName) continue; // skip only if BOTH name and customName are empty
          if (sName && cName && cName.toLowerCase() !== sName.toLowerCase()) continue;
          const tessellated = tessellateSurfacePoly(surface.poly ?? []);
          if (tessellated.length >= 3) floorSources.push(tessellated);
        }

        // Use the same unioned wall polygons for floor sources
        for (const polygon of wallUnion) {
          for (const ring of polygon) {
            const pts = ring.slice();
            // Remove closing duplicate if present
            if (pts.length > 1) {
              const f = pts[0], l = pts[pts.length - 1];
              if (Math.hypot(f[0] - l[0], f[1] - l[1]) < 0.01) pts.pop();
            }
            if (pts.length >= 3) {
              floorSources.push(pts.map(p => ({ x: p[0], y: p[1] })));
            }
          }
        }
        // Also add individual wall rects as floor sources — with small expansion
        // to ensure overlap with adjacent area polygons (prevents gap at boundaries)
        for (const w of walls) {
          if ((w.thickness ?? 20) < 0.1) continue; // skip zero-thickness (handled separately below)
          const r = wallToRect(w, 1, 1);
          if (r) {
            floorSources.push(r.slice(0, 4).map(p => ({ x: p[0], y: p[1] })));
          }
        }
        // Add zero-thickness walls as thin floor strips (they separate areas but have no 3D geometry)
        for (const w of walls) {
          if ((w.thickness ?? 20) > 0.1) continue; // only zero-thickness walls
          const zLen = Math.hypot(w.b.x - w.a.x, w.b.y - w.a.y);
          if (zLen < 0.1) continue;
          const zw = { a: w.a, b: w.b, thickness: 6 }; // give a small floor-only thickness
          const zr = wallToRect(zw, 0, 0);
          if (zr) {
            const pts = zr.slice(0, 4).map(p => ({ x: p[0], y: p[1] }));
            floorSources.push(pts);
          }
        }
        // Also add opening walls (not in the union) — with L-junction extension
        for (const w of openingWalls) {
          const _dx = w.b.x - w.a.x, _dy = w.b.y - w.a.y;
          const _len = Math.hypot(_dx, _dy);
          const _isDiag = _len > 0.1 && Math.min(Math.abs(_dx / _len), Math.abs(_dy / _len)) > 0.15;
          let _extA = 0, _extB = 0;
          if (!_isDiag && _len > 0.1) {
            for (const other of walls) {
              if (other === w) continue;
              const odx = other.b.x - other.a.x, ody = other.b.y - other.a.y;
              const olen = Math.hypot(odx, ody);
              if (olen < 0.1) continue;
              if (Math.min(Math.abs(odx / olen), Math.abs(ody / olen)) > 0.15) continue;
              const oHt = (other.thickness ?? 20) / 2;
              if (Math.hypot(w.a.x - other.a.x, w.a.y - other.a.y) < 3 ||
                  Math.hypot(w.a.x - other.b.x, w.a.y - other.b.y) < 3) _extA = Math.max(_extA, oHt);
              if (Math.hypot(w.b.x - other.a.x, w.b.y - other.a.y) < 3 ||
                  Math.hypot(w.b.x - other.b.x, w.b.y - other.b.y) < 3) _extB = Math.max(_extB, oHt);
            }
          }
          const r = wallToRect(w, _extA, _extB);
          if (r) {
            const pts = r.slice(0, 4).map(p => ({ x: p[0], y: p[1] }));
            floorSources.push(pts);
          }
        }

        const balStripsOBJ = mergeBalustradeStrips(design.balustrades ?? []);
        for (const strip of balStripsOBJ) {
          if (strip.length >= 3) floorSources.push(strip);
        }
        const balFillsOBJ = buildBalustradeFillPolygons(design.balustrades ?? []);
        for (const fill of balFillsOBJ) {
          if (fill.length >= 3) floorSources.push(fill);
        }

        // --- Polygon-based floor: union all sources, subtract voids, extrude ---

        // Expand each floor source slightly so adjacent polygons overlap,
        // guaranteeing the union merges them into one continuous shape.
        const FLOOR_EXPAND = 2;
        for (let si = 0; si < floorSources.length; si++) {
          const poly = floorSources[si];
          if (poly.length < 3) continue;
          const n = poly.length;
          const expanded = [];
          // Compute centroid once
          let cx = 0, cy = 0;
          for (const p of poly) { cx += p.x; cy += p.y; }
          cx /= n; cy /= n;
          for (let i = 0; i < n; i++) {
            const prev = poly[(i - 1 + n) % n];
            const curr = poly[i];
            const next = poly[(i + 1) % n];
            const e1dx = curr.x - prev.x, e1dy = curr.y - prev.y;
            const e1len = Math.hypot(e1dx, e1dy) || 1;
            const n1x = -e1dy / e1len, n1y = e1dx / e1len;
            const e2dx = next.x - curr.x, e2dy = next.y - curr.y;
            const e2len = Math.hypot(e2dx, e2dy) || 1;
            const n2x = -e2dy / e2len, n2y = e2dx / e2len;
            let nx = n1x + n2x, ny = n1y + n2y;
            const nlen = Math.hypot(nx, ny);
            if (nlen < 0.001) { expanded.push({ x: curr.x, y: curr.y }); continue; }
            nx /= nlen; ny /= nlen;
            const toCx = cx - curr.x, toCy = cy - curr.y;
            if (nx * toCx + ny * toCy > 0) { nx = -nx; ny = -ny; }
            expanded.push({ x: curr.x + nx * FLOOR_EXPAND, y: curr.y + ny * FLOOR_EXPAND });
          }
          floorSources[si] = expanded;
        }

        // Convert floorSources {x,y}[] to polygonClipping format [[[x,y],...]]
        const floorPolys = [];
        for (const src of floorSources) {
          if (src.length < 3) continue;
          const ring = src.map(p => [p.x, p.y]);
          const f = ring[0], l = ring[ring.length - 1];
          if (Math.hypot(f[0] - l[0], f[1] - l[1]) > 0.01) ring.push([f[0], f[1]]);
          floorPolys.push([ring]);
        }

        // Union all floor sources into combined polygons
        let floorResult = [];
        if (floorPolys.length > 0) {
          try {
            floorResult = polygonClipping.union(...floorPolys);
          } catch (e) {
            console.warn('Floor union failed, using individual polygons', e);
            floorResult = floorPolys;
          }
        }

        // Subtract voids (stair openings)
        for (const v of floorVoids) {
          if (v.length < 3) continue;
          const vRing = v.map(p => [p.x, p.y]);
          const vf = vRing[0], vl = vRing[vRing.length - 1];
          if (Math.hypot(vf[0] - vl[0], vf[1] - vl[1]) > 0.01) vRing.push([vf[0], vf[1]]);
          try {
            floorResult = polygonClipping.difference(floorResult, [[vRing]]);
          } catch (e) {
            console.warn('Floor void subtraction failed', e);
          }
        }

        // Extrude each result polygon
        const botY = (-FLOOR_THICKNESS).toFixed(4);
        const topY = (0).toFixed(4);

        for (const polygon of floorResult) {
          // polygon = [outerRing, ...holeRings]
          for (let ri = 0; ri < polygon.length; ri++) {
            const ring = polygon[ri];
            // Remove closing duplicate for triangulation
            const pts = ring.slice();
            if (pts.length > 1) {
              const ff = pts[0], ll = pts[pts.length - 1];
              if (Math.hypot(ff[0] - ll[0], ff[1] - ll[1]) < 0.01) pts.pop();
            }
            if (pts.length < 3) continue;

            const objPts = pts.map(p => ({
              x: (p[0] - centerX) * SCALE,
              y: (p[1] - centerY) * SCALE
            }));

            if (ri === 0) {
              // Outer ring → top and bottom face triangulation
              const tris = earClipTriangulate(objPts);
              const baseBot = vertexIndex;
              for (const pt of objPts) {
                vertices.push(`v ${pt.x.toFixed(4)} ${botY} ${pt.y.toFixed(4)}`);
              }
              vertexIndex += objPts.length;
              const baseTop = vertexIndex;
              for (const pt of objPts) {
                vertices.push(`v ${pt.x.toFixed(4)} ${topY} ${pt.y.toFixed(4)}`);
              }
              vertexIndex += objPts.length;
              for (const [a, b, c] of tris) {
                addTriFace(baseBot + a, baseBot + c, baseBot + b); // bottom (flipped winding)
                addTriFace(baseTop + a, baseTop + b, baseTop + c); // top
              }
            }

            // Side faces for every ring (outer + holes)
            const nPts = objPts.length;
            const sideBaseBot = vertexIndex;
            for (const pt of objPts) {
              vertices.push(`v ${pt.x.toFixed(4)} ${botY} ${pt.y.toFixed(4)}`);
            }
            vertexIndex += nPts;
            const sideBaseTop = vertexIndex;
            for (const pt of objPts) {
              vertices.push(`v ${pt.x.toFixed(4)} ${topY} ${pt.y.toFixed(4)}`);
            }
            vertexIndex += nPts;
            for (let i = 0; i < nPts; i++) {
              const j = (i + 1) % nPts;
              addFace(sideBaseBot + i, sideBaseBot + j, sideBaseTop + j, sideBaseTop + i);
            }
          }
        }
      }

      faces.push('g walls_balustrades');

      const extBalsOBJ = extendBalustrades(design.balustrades ?? []);
      for (const bal of extBalsOBJ) {
        const bax = (bal.a.x - centerX) * SCALE;
        const bay = (bal.a.y - centerY) * SCALE;
        const bbx = (bal.b.x - centerX) * SCALE;
        const bby = (bal.b.y - centerY) * SCALE;
        const bthickness = (bal.thickness ?? 10) * SCALE;
        const bheight = (bal.height ?? 100) * SCALE;
        const bdx = bbx - bax, bdy = bby - bay;
        const blen = Math.hypot(bdx, bdy);
        if (blen < 0.001) continue;
        const bnx = -bdy / blen;
        const bny = bdx / blen;
        const bhalfThick = bthickness / 2;

        createWallBox(bax, bay, bbx, bby, 0, bheight, bhalfThick, bnx, bny);
      }

      return [
        `# ${floor.name}`,
        "# Generated by FML Plattegrond Viewer",
        "# Scale: 1 unit = 1 meter",
        "",
        ...vertices,
        "",
        ...faces
      ].join("\n");
    }

    function sanitizeFilename(name) {
      return name.replace(/[^a-zA-Z0-9\-_ ]/g, '').replace(/\s+/g, '_').toLowerCase() || 'verdieping';
    }

    async function exportOBJ() {
      if (floors.length === 0) {
        setError('Geen plattegronden beschikbaar om te exporteren.');
        return;
      }

      const objFiles = floors.map(floor => ({
        name: sanitizeFilename(floor.name),
        content: generateFloorOBJ(floor)
      }));

      if (objFiles.length === 1) {
        const blob = new Blob([objFiles[0].content], { type: "text/plain" });
        const downloadUrl = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.download = `${objFiles[0].name}.obj`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(downloadUrl);
        showToast(`✓ ${objFiles[0].name}.obj geëxporteerd`);
        return;
      }

      if (typeof JSZip === 'undefined') {
        setError('JSZip library niet geladen. Controleer je internetverbinding.');
        return;
      }

      const zip = new JSZip();
      const usedNames = new Map();

      for (const file of objFiles) {
        let fileName = file.name;
        const count = usedNames.get(fileName) || 0;
        if (count > 0) fileName = `${fileName}_${count}`;
        usedNames.set(file.name, count + 1);

        zip.file(`${fileName}.obj`, file.content);
      }

      try {
        const zipBlob = await zip.generateAsync({ type: "blob" });
        const downloadUrl = URL.createObjectURL(zipBlob);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.download = "plattegronden.zip";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(downloadUrl);
        showToast(`✓ ${objFiles.length} OBJ bestanden geëxporteerd`);
      } catch (err) {
        setError('Er ging iets mis bij het maken van het ZIP bestand.');
      }
    }

    function downloadSingleFloorOBJ(floorIndex) {
      const floor = floors[floorIndex];
      if (!floor) return;
      const objContent = generateFloorOBJ(floor);
      const fileName = sanitizeFilename(floor.name);
      const blob = new Blob([objContent], { type: "text/plain" });
      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `${fileName}.obj`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);
      showToast(`✓ ${fileName}.obj geëxporteerd`);
    }

    // ============================================================
    // EVENT LISTENERS
    // ============================================================

    // Drag & Drop
    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('drag-over');
    });
    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('drag-over');
    });
    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('drag-over');
      const file = e.dataTransfer?.files?.[0];
      if (file) handleFileUpload(file);
    });
    dropzone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
      const file = fileInput.files?.[0];
      if (file) handleFileUpload(file);
      fileInput.value = '';
    });

    // Export
    btnExport.addEventListener('click', exportOBJ);

    // FML Download
    btnDownloadFml.addEventListener('click', () => {
      if (!originalFmlData) {
        setError('Geen FML data beschikbaar om te downloaden.');
        return;
      }
      const blob = new Blob([JSON.stringify(originalFmlData, null, 2)], { type: "application/json" });
      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = 'plattegrond.fml';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);
      showToast('✓ FML gedownload');
    });

    // Address fields — live update frame preview
    addressStreet.addEventListener('input', () => updateFrameAddress());
    addressCity.addEventListener('input', () => updateFrameAddress());

    // Wizard navigation
    btnWizardPrev.addEventListener('click', () => prevWizardStep());
    btnWizardNext.addEventListener('click', () => nextWizardStep());

    // Funda URL loading (refs in ensureDomRefs)

    function getFundaUrl() {
      return fundaUrlInput.value.trim();
    }

    fundaUrlInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') loadFromFunda();
    });
    btnFunda.addEventListener('click', () => loadFromFunda());

    // Function declaration so it hoists (Shopify addEventListener issue)
    const TEST_LINKS = {
      1: 'https://www.funda.nl/detail/koop/haarlem/appartement-prinsen-bolwerk-72/43226270/',
      2: 'https://www.funda.nl/detail/koop/amsterdam/appartement-hoofdweg-275-1/89691599/'
    };
    function pasteTestLink(n) {
      var input = document.getElementById('fundaUrl');
      if (input) input.value = TEST_LINKS[n] || TEST_LINKS[1];
    }
    for (const n of [1, 2]) {
      const btn = document.getElementById('btnTest' + n);
      if (btn) btn.addEventListener('click', () => pasteTestLink(n));
    }

    // Funda status checker
    function setFundaStatus(state, html) {
      var box = document.getElementById('fundaStatus');
      var icon = document.getElementById('fundaStatusIcon');
      var text = document.getElementById('fundaStatusText');
      if (!box) return;
      box.className = 'funda-status visible ' + state;
      if (state === 'loading') {
        icon.innerHTML = '<div class="mini-spinner"></div>';
      } else if (state === 'success') {
        icon.textContent = '';
      } else if (state === 'error') {
        icon.textContent = '✕';
      }
      text.innerHTML = html;
    }

    function hideFundaStatus() {
      var box = document.getElementById('fundaStatus');
      if (box) box.className = 'funda-status';
      // Also hide contact email if present
      var emailBtn = document.getElementById('contactEmailBtn');
      if (emailBtn) emailBtn.style.display = 'none';
    }

    function showContactEmail(fundaUrl) {
      var existing = document.getElementById('contactEmailBtn');
      if (existing) {
        // Update mailto and show
        var subject = encodeURIComponent('Frame³ bestelling — hulp nodig');
        var body = encodeURIComponent('Hoi,\n\nIk wil graag een Frame³ bestellen maar het lukt niet via de website.\n\nFunda link: ' + (fundaUrl || '(niet ingevuld)') + '\n\nKunnen jullie me helpen?\n\nAlvast bedankt!');
        existing.href = 'mailto:vince@mattori.nl?subject=' + subject + '&body=' + body;
        existing.style.display = '';
        return;
      }
      // Create email button after funda-status
      var statusBox = document.getElementById('fundaStatus');
      if (!statusBox) return;
      var subject = encodeURIComponent('Frame³ bestelling — hulp nodig');
      var body = encodeURIComponent('Hoi,\n\nIk wil graag een Frame³ bestellen maar het lukt niet via de website.\n\nFunda link: ' + (fundaUrl || '(niet ingevuld)') + '\n\nKunnen jullie me helpen?\n\nAlvast bedankt!');
      var btn = document.createElement('a');
      btn.id = 'contactEmailBtn';
      btn.className = 'btn-contact-email';
      btn.href = 'mailto:vince@mattori.nl?subject=' + subject + '&body=' + body;
      btn.textContent = 'Neem contact op →';
      statusBox.parentNode.insertBefore(btn, statusBox.nextSibling);
    }

    async function loadFromFunda() {
      ensureDomRefs();
      noFloorsMode = false; // Reset on new attempt
      const url = getFundaUrl();
      clearError();
      if (!url) { setError('Voer een Funda URL in.'); return; }
      if (!url.includes('funda.nl')) { setError('Dit is geen Funda URL.'); return; }

      setFundaStatus('loading', 'Plattegrond ophalen...');
      showLoading();
      btnFunda.disabled = true;

      try {
        const resp = await fetch('https://web-production-89353.up.railway.app/funda-fml', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ url })
        });
        const data = await resp.json();

        // Check if Funda link was valid but no interactive floor plans found
        var noPlattegrond = data.error && (data.error.toLowerCase().includes('geen plattegrond') || data.error.toLowerCase().includes('geen fml') || data.error.toLowerCase().includes('no floorplan'));
        var noValidFloors = !data?.floors?.length || !(data.floors ?? []).some(f => f?.designs?.[0]);

        if (noPlattegrond || (data.floors && noValidFloors)) {
          // noFloorsMode: Funda link valid but no interactive floor plans
          noFloorsMode = true;
          lastFundaUrl = url;
          setFundaStatus('success', '<strong>✓ Funda link herkend</strong><strong class="status-warning">✗ Geen interactieve plattegronden beschikbaar</strong><span>Geen zorgen — we bouwen je Frame³ handmatig op basis van de Funda-foto\'s.</span>');
          updateWizardUI();
          return;
        }

        if (data.error) {
          setFundaStatus('error', '<strong>Funda link niet geldig</strong><span>Controleer de link en probeer het opnieuw, of neem contact op.</span>');
          showContactEmail(url);
          btnWizardNext.style.display = 'none';
          return;
        }

        originalFmlData = data;
        originalFileName = 'funda-plattegrond.fml';
        lastFundaUrl = url;
        fileLabel.textContent = `🔗 ${url.split('/').filter(Boolean).pop() || 'funda'}`;

        const addr = parseFundaAddress(url) || parseAddressFromFML(data);
        if (addr) {
          addressStreet.value = addr.street;
          addressCity.value = addr.city;
          currentAddress = addr;
        } else {
          addressStreet.value = '';
          addressCity.value = '';
        }

        var addrStr = addr ? addr.street + ', ' + addr.city : 'Adres niet gevonden';
        setFundaStatus('success', '<strong>✓ Funda link correct</strong><strong>✓ ' + data.floors.length + ' interactieve plattegrond' + (data.floors.length === 1 ? '' : 'en') + ' gevonden</strong><span class="funda-address-line">📍 ' + addrStr + '</span>');

        processFloors(data);
      } catch (err) {
        if (err.message && (err.message.includes('Load failed') || err.message.includes('Failed to fetch'))) {
          setFundaStatus('error', '<strong>Verbinding mislukt</strong><span>Controleer de link en probeer het opnieuw, of neem contact op.</span>');
        } else {
          setFundaStatus('error', `<strong>Fout</strong><span>Controleer de link en probeer het opnieuw, of neem contact op.</span>`);
        }
        showContactEmail(url);
        btnWizardNext.style.display = 'none';
      } finally {
        hideLoading();
        btnFunda.disabled = false;
      }
    }

    // Order button — adds product to Shopify cart via Cart API
    function submitOrder() {
      ensureDomRefs();
      // Find variant ID from Shopify's hidden input
      var variantInput = document.querySelector('form[action*="/cart/add"] input[name="id"]');
      var variantId = variantInput ? variantInput.value : null;
      if (!variantId) {
        showToast('Product niet gevonden.');
        return;
      }
      // Disable the correct button (btnOrder for normal flow, btnWizardNext for noFloorsMode)
      var orderBtn = noFloorsMode ? btnWizardNext : document.getElementById('btnOrder');
      var originalText = orderBtn ? orderBtn.textContent : '';
      if (orderBtn) { orderBtn.disabled = true; orderBtn.textContent = 'Toevoegen…'; }
      var fundaLink = fundaUrlInput ? fundaUrlInput.value.trim() : '';
      var itemProperties = {};
      if (fundaLink) itemProperties['Funda link'] = fundaLink;
      if (noFloorsMode) itemProperties['Opmerking'] = 'Geen interactieve plattegronden — handmatig opbouwen';
      fetch('/cart/add.js', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items: [{ id: parseInt(variantId), quantity: 1, properties: itemProperties }] })
      })
      .then(function(res) {
        if (!res.ok) throw new Error('Status ' + res.status);
        return res.json();
      })
      .then(function() {
        // Redirect to cart page
        window.location.href = '/cart';
      })
      .catch(function(err) {
        showToast('Kon niet toevoegen aan winkelwagen.');
        if (orderBtn) { orderBtn.disabled = false; orderBtn.textContent = originalText; }
      });
    }
    // submitOrder is triggered via onclick="submitOrder()" in HTML

    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
        e.preventDefault();
        fileInput.click();
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 'e') {
        e.preventDefault();
        if (floors.length > 0) exportOBJ();
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 'd') {
        e.preventDefault();
        if (originalFmlData) btnDownloadFml.click();
      }
    });

    // Start configurator button — use event delegation for Shopify robustness
    var configuratorStarted = false;
    function startConfigurator() {
      ensureDomRefs();
      if (configuratorStarted) return;
      configuratorStarted = true;
      const btn = document.getElementById('btnStartConfigurator');
      if (btn) {
        btn.style.transition = 'opacity 0.3s ease';
        btn.style.opacity = '0';
        setTimeout(() => { btn.style.display = 'none'; }, 300);
      }

      // Animate out all collapsible info blocks with staggered delay
      const collapsibles = document.querySelectorAll('.mattori-configurator .collapsible-info');
      collapsibles.forEach((el, i) => {
        setTimeout(() => {
          el.classList.add('collapsing');
        }, i * 80);
      });

      // Hide Shopify buy button and sticky bar
      var buyBlock = document.querySelector('.mattori-configurator .buy-buttons-block');
      if (buyBlock) { buyBlock.style.transition = 'opacity 0.3s ease'; buyBlock.style.opacity = '0'; setTimeout(() => { buyBlock.style.display = 'none'; }, 300); }
      var stickyBar = document.querySelector('sticky-add-to-cart');
      if (stickyBar) stickyBar.style.display = 'none';

      // After all animations complete, show wizard
      const totalDelay = collapsibles.length * 80 + 400;
      setTimeout(() => {
        collapsibles.forEach(el => { el.style.display = 'none'; });
        wizard.style.display = '';
        initWizard();
      }, totalDelay);
    }

    // Attach via event delegation on container (survives DOM reordering)
    const configuratorRoot = document.querySelector('.mattori-configurator');
    if (configuratorRoot) {
      configuratorRoot.addEventListener('click', (e) => {
        if (e.target.closest('#btnStartConfigurator')) {
          e.preventDefault();
          startConfigurator();
        }
      });
    }

    // Also try direct attachment as fallback
    const btnStartConfigurator = document.getElementById('btnStartConfigurator');
    if (btnStartConfigurator) {
      btnStartConfigurator.addEventListener('click', (e) => {
        e.preventDefault();
        startConfigurator();
      });
    }
