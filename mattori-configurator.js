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
    var floorReviewStatus = {};   // { floorIndex: 'confirmed' | 'issue' | 'major' }
    var floorIssues = {};         // { floorIndex: 'text describing issue' }
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
        fundaUrlInput, btnFunda, frameHouseIcon;

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
      frameHouseIcon = document.getElementById('frameHouseIcon');

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
          if (i > 0 && /^\d+$/.test(parts[i - 1]) && /^(\d+[a-z]?|[a-z]{1,2})$/i.test(parts[i]) && parts[i].length <= 3) {
            merged[merged.length - 1] += '-' + parts[i];
          } else {
            merged.push(parts[i]);
          }
        }
        const titleCased = merged.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ')
          .replace(/-([a-z])/g, (_, c) => '-' + c.toUpperCase());
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

    // House icon options
    var houseIconOptions = [
      { id: 'huisje1', label: 'Huisje 1', url: 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Huisje_1.png?v=1771603403' },
      { id: 'huisje2', label: 'Huisje 2', url: 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Huisje_2.png?v=1771603403' },
      { id: 'huisje3', label: 'Huisje 3', url: 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Huisje_3.png?v=1771603403' },
      { id: 'huisje4', label: 'Huisje 4', url: 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Huisje_4.png?v=1771603403' },
      { id: 'huisje5', label: 'Huisje 5', url: 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Huisje_5.png?v=1771603403' },
      { id: 'huisje6', label: 'Huisje 6', url: 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Huisje_6.png?v=1771603403' },
      { id: 'huisje7', label: 'Huisje 7', url: 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Huisje_7.png?v=1771603403' },
      { id: 'huisje8', label: 'Huisje 8', url: 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Huisje_8.png?v=1771603402' }
    ];
    var selectedHouseIcon = 'huisje1'; // default

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
      const voidsByFloor = allFloorDesigns.map(() => []);

      for (let fi = 0; fi < allFloorDesigns.length; fi++) {
        const design = allFloorDesigns[fi];

        // Only use explicit isCutout surfaces and role=14 (stair void markers)
        for (const surface of design.surfaces ?? []) {
          if ((surface.role ?? -1) !== 14 && !surface.isCutout) continue;
          const poly = tessellateSurfacePoly(surface.poly ?? []);
          if (poly.length >= 3) {
            voidsByFloor[fi].push(poly.map(p => ({ x: p.x, y: p.y })));
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
        // Use actual bounding box to prevent clipping. Swap for 90°/270° rotation.
        var isSwapped = (rotation === 90 || rotation === 270);
        var effectiveW = isSwapped ? size.z : size.x;
        var effectiveH = isSwapped ? size.x : size.z;
        var pxPerUnitW = width / effectiveW;
        var pxPerUnitH = height / effectiveH;
        var pxPerUnit = Math.min(pxPerUnitW, pxPerUnitH);
        var halfFrustumW = (width / 2) / pxPerUnit;
        var halfFrustumH = (height / 2) / pxPerUnit;
        camera = new THREE.OrthographicCamera(
          -halfFrustumW, halfFrustumW, halfFrustumH, -halfFrustumH, 0.01, 1000
        );
        camera.up.set(0, 0, -1); // Z- = up on screen (north in floorplan)
        camera.position.set(0, 50, 0);
        camera.lookAt(0, 0, 0);

        // Fine-tune alignment: push model to canvas edge so it touches
        // the alignment line at that edge
        var modelHalfW = effectiveW / 2;
        var modelHalfH = effectiveH / 2;

        // Y-axis: push to top or bottom edge
        var alignY = getFloorAlignY(floorIndex);
        if (alignY === 'top' || alignY === 'bottom') {
          var shiftZ = halfFrustumH - modelHalfH;
          scene.position.z += (alignY === 'bottom') ? shiftZ : -shiftZ;
        }

        // X-axis: push to left or right edge
        var alignX = getFloorAlignX(floorIndex);
        if (alignX === 'left' || alignX === 'right') {
          var shiftX = halfFrustumW - modelHalfW;
          scene.position.x += (alignX === 'right') ? shiftX : -shiftX;
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
      const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true, preserveDrawingBuffer: true });
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
    var layoutAlignX = 'center'; // 'left', 'center', or 'right'
    var layoutAlignY = 'bottom'; // 'top', 'center', or 'bottom'
    var layoutGapFactor = 0.08; // gap as fraction of largest floor dimension (0..0.3)
    var layoutScaleFactor = 1.0; // user scale: 0.82 (klein), 1.0 (normaal), 1.1 (groot)
    var showGridOverlay = true; // user toggle for grid + alignment lines
    var floorSettings = {}; // per-floor: { rotate: 0|90|180|270 }

    // ============================================================
    // GRID SYSTEM — snap-to-grid for physical production
    // ============================================================
    var ZONE_PHYSICAL_W_MM = 170; // usable interior width in mm
    var ZONE_PHYSICAL_H_MM = 130; // usable interior height in mm
    var GRID_CELL_MM = 5;         // grid cell size in mm

    var gridEditMode = false;     // true when user activated drag-to-reposition
    var customPositions = null;   // null = auto-layout, or [{index, gridX, gridY}] (gridX/gridY=null means auto)
    var layoutCalculated = false; // true after user clicks "Bereken indeling" in step 4
    var _dragCleanups = [];       // cleanup functions for drag event listeners

    function getGridDimensions() {
      var overlay = floorsGrid ? floorsGrid.parentElement : null;
      if (!overlay) return { cols: 34, rows: 26, cellPx: 10, pxPerMm: 2, zoneW: 340, zoneH: 260 };
      var overlayRect = overlay.getBoundingClientRect();
      var zoneW = overlayRect.width;
      var zoneH = overlayRect.height;
      var pxPerMm = zoneW / ZONE_PHYSICAL_W_MM;
      var cellPx = GRID_CELL_MM * pxPerMm;
      var cols = Math.floor(ZONE_PHYSICAL_W_MM / GRID_CELL_MM);
      var rows = Math.floor(ZONE_PHYSICAL_H_MM / GRID_CELL_MM);
      return { cols: cols, rows: rows, cellPx: cellPx, pxPerMm: pxPerMm, zoneW: zoneW, zoneH: zoneH };
    }

    function snapToGrid(px, cellPx) {
      return Math.round(px / cellPx) * cellPx;
    }

    function pxToGridCoord(px, cellPx) {
      return Math.round(px / cellPx);
    }

    function gridCoordToPx(gridCoord, cellPx) {
      return gridCoord * cellPx;
    }

    // Reset only a single floor's custom position (keep other floors intact)
    function resetSingleFloorPosition(floorIdx) {
      if (!customPositions) return;
      var cp = customPositions.find(function(c) { return c.index === floorIdx; });
      if (cp) {
        cp.gridX = null;
        cp.gridY = null;
      }
      // If ALL floors are now null, clean up customPositions entirely
      var anyCustom = customPositions.some(function(c) { return c.gridX !== null; });
      if (!anyCustom) customPositions = null;
    }

    // Re-snap a floor in-place after alignment change (stay near current position)
    function reSnapFloorAlignment(floorIdx) {
      if (!customPositions || !currentLayout) return;
      var cp = customPositions.find(function(c) { return c.index === floorIdx; });
      if (!cp || cp.gridX === null) return; // auto-positioned → auto will handle it
      var grid = getGridDimensions();
      // Find this floor's dimensions from current layout
      var pos = null;
      for (var i = 0; i < currentLayout.positions.length; i++) {
        if (currentLayout.positions[i].index === floorIdx) { pos = currentLayout.positions[i]; break; }
      }
      if (!pos) return;
      var currentLeft = cp.gridX * grid.cellPx;
      var currentTop = cp.gridY * grid.cellPx;
      var w = pos.w, h = pos.h;
      var alignX = getFloorAlignX(floorIdx);
      var alignY = getFloorAlignY(floorIdx);
      // Re-snap X: move alignment edge to nearest grid crossing
      var newLeft;
      if (alignX === 'left') {
        newLeft = snapToGrid(currentLeft, grid.cellPx);
      } else if (alignX === 'right') {
        newLeft = snapToGrid(currentLeft + w, grid.cellPx) - w;
      } else {
        newLeft = snapToGrid(currentLeft + w / 2, grid.cellPx) - w / 2;
      }
      // Re-snap Y: move alignment edge to nearest grid crossing
      var newTop;
      if (alignY === 'top') {
        newTop = snapToGrid(currentTop, grid.cellPx);
      } else if (alignY === 'bottom') {
        newTop = snapToGrid(currentTop + h, grid.cellPx) - h;
      } else {
        newTop = snapToGrid(currentTop + h / 2, grid.cellPx) - h / 2;
      }
      // Clamp within zone
      newLeft = Math.max(0, Math.min(grid.zoneW - w, newLeft));
      newTop = Math.max(0, Math.min(grid.zoneH - h, newTop));
      // Store re-snapped position (top-left grid coord)
      cp.gridX = pxToGridCoord(newLeft, grid.cellPx);
      cp.gridY = pxToGridCoord(newTop, grid.cellPx);
    }

    // --- Grid overlay (SVG) ---
    function renderGridOverlay() {
      if (!floorsGrid) return;
      var existing = floorsGrid.querySelector('.grid-overlay');
      if (existing) existing.remove();
      // Respect user toggle
      if (!showGridOverlay) return;
      // Only show grid when edit mode is active
      if (!gridEditMode) return;

      var grid = getGridDimensions();
      var svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('class', 'grid-overlay');
      svg.setAttribute('width', grid.zoneW);
      svg.setAttribute('height', grid.zoneH);
      svg.style.cssText = 'position:absolute;top:0;left:0;pointer-events:none;z-index:1;';

      var actualCols = Math.ceil(grid.zoneW / grid.cellPx);
      var actualRows = Math.ceil(grid.zoneH / grid.cellPx);
      var midCol = Math.floor(actualCols / 2);
      var midRow = Math.floor(actualRows / 2);
      // Skip outermost lines (x=0, x=max, y=0, y=max) to avoid border rectangle
      for (var x = 1; x < actualCols; x++) {
        var line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        var px = x * grid.cellPx;
        line.setAttribute('x1', px); line.setAttribute('y1', 0);
        line.setAttribute('x2', px); line.setAttribute('y2', grid.zoneH);
        var isMid = (x === midCol);
        line.setAttribute('stroke', isMid ? 'rgba(0,0,0,0.25)' : 'rgba(0,0,0,0.07)');
        line.setAttribute('stroke-width', isMid ? '1.5' : '0.5');
        svg.appendChild(line);
      }
      for (var y = 1; y < actualRows; y++) {
        var line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        var py = y * grid.cellPx;
        line.setAttribute('x1', 0); line.setAttribute('y1', py);
        line.setAttribute('x2', grid.zoneW); line.setAttribute('y2', py);
        var isMidH = (y === midRow);
        line.setAttribute('stroke', isMidH ? 'rgba(0,0,0,0.25)' : 'rgba(0,0,0,0.07)');
        line.setAttribute('stroke-width', isMidH ? '1.5' : '0.5');
        svg.appendChild(line);
      }
      floorsGrid.insertBefore(svg, floorsGrid.firstChild);
    }

    // --- Drag handlers ---
    function enableGridDrag() {
      var overlay = document.getElementById('unifiedFloorsOverlay');
      if (overlay) {
        overlay.classList.add('drag-enabled');
        if (showGridOverlay) overlay.classList.add('zone-editing');
      }

      // Don't create customPositions here — auto-layout positions are already
      // grid-snapped by computeFloorLayout(). This prevents ANY visual shift
      // when enabling drag mode. customPositions are created per-floor on
      // first actual drag.

      renderPreviewThumbnails();
      renderGridOverlay();

      // Attach drag handlers + floor buttons to each floor-canvas-wrap
      var wraps = floorsGrid.querySelectorAll('.floor-canvas-wrap');
      for (var wi = 0; wi < wraps.length; wi++) {
        attachDragHandlers(wraps[wi], wi);
      }
      addGridFloorButtons();
    }

    function refreshGridAfterChange() {
      renderPreviewThumbnails();
      renderGridOverlay();
      _dragCleanups.forEach(function(fn) { fn(); });
      _dragCleanups = [];
      var newWraps = floorsGrid.querySelectorAll('.floor-canvas-wrap');
      for (var nw = 0; nw < newWraps.length; nw++) {
        attachDragHandlers(newWraps[nw], nw);
      }
      addGridFloorButtons();
    }

    function addGridFloorButtons() {
      if (!gridEditMode || !floorsGrid || !showGridOverlay) return;
      // Add alignment border lines to each floor thumbnail (only on alignment edges)
      var wraps = floorsGrid.querySelectorAll('.floor-canvas-wrap');
      var grid = getGridDimensions();
      for (var wi = 0; wi < wraps.length; wi++) {
        (function(wrap, posIdx) {
          if (!currentLayout || !currentLayout.positions[posIdx]) return;
          var pos = currentLayout.positions[posIdx];
          var w = pos.w, h = pos.h;

          var svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
          svg.setAttribute('width', w);
          svg.setAttribute('height', h);
          svg.setAttribute('viewBox', '0 0 ' + w + ' ' + h);
          svg.style.cssText = 'position:absolute;top:0;left:0;pointer-events:none;z-index:15;';

          var lineColor = 'rgba(0, 0, 0, 0.55)';
          var lineWidth = '1.5';

          // Per-floor alignment
          var floorIdx = pos.index;
          var thisAlignX = getFloorAlignX(floorIdx);
          var thisAlignY = getFloorAlignY(floorIdx);

          // Compute alignment line positions SNAPPED to actual grid crossings
          // Convert desired position to zone coords, snap to grid, convert back to local
          var rawVx = thisAlignX === 'left' ? 0 : thisAlignX === 'right' ? w : w / 2;
          var zoneVx = parseFloat(wrap.style.left) + rawVx;
          var vx = snapToGrid(zoneVx, grid.cellPx) - parseFloat(wrap.style.left);

          var vLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
          vLine.setAttribute('x1', vx); vLine.setAttribute('y1', 0);
          vLine.setAttribute('x2', vx); vLine.setAttribute('y2', h);
          vLine.setAttribute('stroke', lineColor);
          vLine.setAttribute('stroke-width', lineWidth);
          vLine.setAttribute('stroke-dasharray', '6 4');
          svg.appendChild(vLine);

          var rawHy = thisAlignY === 'top' ? 0 : thisAlignY === 'bottom' ? h : h / 2;
          var zoneHy = parseFloat(wrap.style.top) + rawHy;
          var hy = snapToGrid(zoneHy, grid.cellPx) - parseFloat(wrap.style.top);

          var hLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
          hLine.setAttribute('x1', 0); hLine.setAttribute('y1', hy);
          hLine.setAttribute('x2', w); hLine.setAttribute('y2', hy);
          hLine.setAttribute('stroke', lineColor);
          hLine.setAttribute('stroke-width', lineWidth);
          hLine.setAttribute('stroke-dasharray', '6 4');
          svg.appendChild(hLine);

          wrap.appendChild(svg);
        })(wraps[wi], wi);
      }
    }

    function disableGridDrag() {
      var overlay = document.getElementById('unifiedFloorsOverlay');
      if (overlay) {
        overlay.classList.remove('drag-enabled');
        overlay.classList.remove('zone-editing');
      }

      var gridEl = floorsGrid ? floorsGrid.querySelector('.grid-overlay') : null;
      if (gridEl) gridEl.remove();

      _dragCleanups.forEach(function(fn) { fn(); });
      _dragCleanups = [];
    }

    function attachDragHandlers(wrap, posIndex) {
      var isDragging = false;
      var startX, startY, origLeft, origTop;

      function onStart(e) {
        if (!gridEditMode) return;
        e.preventDefault();
        isDragging = true;
        wrap.classList.add('dragging');

        var point = e.touches ? e.touches[0] : e;
        startX = point.clientX;
        startY = point.clientY;
        origLeft = parseFloat(wrap.style.left) || 0;
        origTop = parseFloat(wrap.style.top) || 0;
      }

      function onMove(e) {
        if (!isDragging) return;
        e.preventDefault();

        var point = e.touches ? e.touches[0] : e;
        var dx = point.clientX - startX;
        var dy = point.clientY - startY;

        var grid = getGridDimensions();
        var wrapW = parseFloat(wrap.style.width) || 0;
        var wrapH = parseFloat(wrap.style.height) || 0;
        var rawLeft = origLeft + dx;
        var rawTop = origTop + dy;

        // Snap by top-left corner (per-floor independent positioning)
        var newLeft = snapToGrid(rawLeft, grid.cellPx);
        var newTop = snapToGrid(rawTop, grid.cellPx);

        // Clamp within zone
        newLeft = Math.max(0, Math.min(grid.zoneW - wrapW, newLeft));
        newTop = Math.max(0, Math.min(grid.zoneH - wrapH, newTop));

        wrap.style.left = newLeft + 'px';
        wrap.style.top = newTop + 'px';
      }

      function onEnd() {
        if (!isDragging) return;
        isDragging = false;
        wrap.classList.remove('dragging');

        // Store top-left grid coords (per-floor independent)
        var grid = getGridDimensions();
        var finalLeft = parseFloat(wrap.style.left) || 0;
        var finalTop = parseFloat(wrap.style.top) || 0;

        if (currentLayout && currentLayout.positions[posIndex]) {
          var floorIdx = currentLayout.positions[posIndex].index;
          // Create customPositions on first drag (null entries = use auto-layout)
          if (!customPositions) {
            customPositions = currentLayout.positions.map(function(pos) {
              return { index: pos.index, gridX: null, gridY: null };
            });
          }
          var cp = customPositions.find(function(c) { return c.index === floorIdx; });
          if (cp) {
            cp.gridX = pxToGridCoord(finalLeft, grid.cellPx);
            cp.gridY = pxToGridCoord(finalTop, grid.cellPx);
          }
          // Check for overlaps after every drag
          checkFloorOverlaps();
          // Show reset icon
          var _resetBtn = document.getElementById('btnResetLayout');
          if (_resetBtn) _resetBtn.style.display = 'inline-flex';
        }
      }

      wrap.addEventListener('mousedown', onStart);
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onEnd);
      wrap.addEventListener('touchstart', onStart, { passive: false });
      document.addEventListener('touchmove', onMove, { passive: false });
      document.addEventListener('touchend', onEnd);

      _dragCleanups.push(function() {
        wrap.removeEventListener('mousedown', onStart);
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onEnd);
        wrap.removeEventListener('touchstart', onStart);
        document.removeEventListener('touchmove', onMove);
        document.removeEventListener('touchend', onEnd);
      });
    }

    // --- Overlap detection ---
    var layoutHasOverlap = false;

    function checkFloorOverlaps() {
      layoutHasOverlap = false;
      if (!currentLayout || currentLayout.positions.length < 2) {
        updateWizardUI();
        return;
      }
      var positions = currentLayout.positions;
      // Get actual rendered positions from the DOM wraps
      var wraps = floorsGrid ? floorsGrid.querySelectorAll('.floor-canvas-wrap') : [];
      var rects = [];
      for (var wi = 0; wi < wraps.length; wi++) {
        var l = parseFloat(wraps[wi].style.left) || 0;
        var t = parseFloat(wraps[wi].style.top) || 0;
        var w = parseFloat(wraps[wi].style.width) || 0;
        var h = parseFloat(wraps[wi].style.height) || 0;
        rects.push({ x: l, y: t, w: w, h: h });
      }
      // Check all pairs for overlap (with 2px tolerance)
      var tolerance = 2;
      for (var a = 0; a < rects.length; a++) {
        for (var b = a + 1; b < rects.length; b++) {
          var ra = rects[a], rb = rects[b];
          if (ra.x + tolerance < rb.x + rb.w && rb.x + tolerance < ra.x + ra.w &&
              ra.y + tolerance < rb.y + rb.h && rb.y + tolerance < ra.y + ra.h) {
            layoutHasOverlap = true;
            break;
          }
        }
        if (layoutHasOverlap) break;
      }
      updateWizardUI();
    }

    // --- Grid position cart properties ---
    function getGridPositionProperties() {
      var props = {};
      // Always include global alignment settings
      props['Uitlijning'] = 'X=' + layoutAlignX + ', Y=' + layoutAlignY;
      // Include per-floor alignment and rotation if any differ from defaults
      for (var fk in floorSettings) {
        if (floorSettings.hasOwnProperty(fk)) {
          var fs = floorSettings[fk];
          var hasCustomAlign = (fs.alignX && fs.alignX !== layoutAlignX) || (fs.alignY && fs.alignY !== layoutAlignY);
          var hasRotation = fs.rotate && fs.rotate !== 0;
          if (hasCustomAlign || hasRotation) {
            var floor = floors[parseInt(fk)];
            var fname = floor ? floor.name : ('Verdieping ' + (parseInt(fk) + 1));
            var parts = [];
            if (hasCustomAlign) {
              var ax = fs.alignX || layoutAlignX;
              var ay = fs.alignY || layoutAlignY;
              parts.push('X=' + ax + ', Y=' + ay);
            }
            if (hasRotation) {
              parts.push('Rotatie=' + fs.rotate + '°');
            }
            props['Instelling ' + fname] = parts.join(', ');
          }
        }
      }
      if (!customPositions) return props;
      for (var i = 0; i < customPositions.length; i++) {
        var cp = customPositions[i];
        var floor = floors[cp.index];
        var floorName = floor ? floor.name : ('Verdieping ' + (cp.index + 1));
        props['Positie ' + floorName] = 'X=' + cp.gridX + ', Y=' + cp.gridY;
      }
      return props;
    }

    // --- Resize handler for grid mode ---
    var _gridResizeTimer = null;
    window.addEventListener('resize', function() {
      if (!gridEditMode) return;
      clearTimeout(_gridResizeTimer);
      _gridResizeTimer = setTimeout(function() {
        renderPreviewThumbnails();
        renderGridOverlay();
        _dragCleanups.forEach(function(fn) { fn(); });
        _dragCleanups = [];
        var wraps = floorsGrid ? floorsGrid.querySelectorAll('.floor-canvas-wrap') : [];
        for (var wi = 0; wi < wraps.length; wi++) {
          attachDragHandlers(wraps[wi], wi);
        }
      }, 200);
    });

    function computeFloorLayout(zoneW, zoneH, includedIndices) {
      // Gather floor dimensions in cm (swap for 90°/270° rotation)
      var items = includedIndices.map(function(i) {
        var f = floors[i];
        var rot = getFloorRotate(i);
        var isSwapped = (rot === 90 || rot === 270);
        return { index: i, w: isSwapped ? f.worldH : f.worldW, h: isSwapped ? f.worldW : f.worldH, name: f.name || '' };
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
        // Two floors — try side-by-side and stacked (no diagonal — causes overlap)
        layouts.push(layoutSideBySide(items));
        layouts.push(layoutStacked(items));
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
      var finalScale = best.scale * 0.88 * layoutScaleFactor; // 12% padding + user scale
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

      // ── Grid-native snap ──
      // Snap all positions to the 5mm physical grid so there is never a
      // visible shift when the user starts dragging.
      var pxPerMm = zoneW / ZONE_PHYSICAL_W_MM;
      var cellPx  = GRID_CELL_MM * pxPerMm;

      for (var ri = 0; ri < result.length; ri++) {
        var r = result[ri];

        // Determine the anchor point per axis based on per-floor alignment
        var floorAlignX = getFloorAlignX(r.index);
        var floorAlignY = getFloorAlignY(r.index);

        // X-axis
        if (floorAlignX === 'left') {
          r.x = snapToGrid(r.x, cellPx);
        } else if (floorAlignX === 'right') {
          // Snap right edge, then derive x
          var rightSnapped = snapToGrid(r.x + r.w, cellPx);
          r.x = rightSnapped - r.w;
        } else {
          // center — snap the center point, then derive x
          var cx = r.x + r.w / 2;
          var cxSnapped = snapToGrid(cx, cellPx);
          r.x = cxSnapped - r.w / 2;
        }

        // Y-axis
        if (floorAlignY === 'top') {
          r.y = snapToGrid(r.y, cellPx);
        } else if (floorAlignY === 'bottom') {
          // Snap bottom edge, then derive y
          var bottomSnapped = snapToGrid(r.y + r.h, cellPx);
          r.y = bottomSnapped - r.h;
        } else {
          // center — snap the center point, then derive y
          var cy = r.y + r.h / 2;
          var cySnapped = snapToGrid(cy, cellPx);
          r.y = cySnapped - r.h / 2;
        }
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

    function getFloorAlignX(floorIndex) {
      var fs = floorSettings[floorIndex];
      return (fs && fs.alignX) ? fs.alignX : layoutAlignX;
    }

    function getFloorAlignY(floorIndex) {
      var fs = floorSettings[floorIndex];
      return (fs && fs.alignY) ? fs.alignY : layoutAlignY;
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
      // In step 4, don't render until "Bereken indeling" is clicked
      if (currentWizardStep === 4 && !layoutCalculated) {
        for (const v of previewViewers) { if (v.renderer) v.renderer.dispose(); }
        previewViewers = [];
        floorsGrid.innerHTML = '';
        return;
      }

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

      // Apply custom positions if user has dragged floors
      // Grid coords store top-left corner (null = use auto-layout position)
      if (customPositions) {
        var _grid = getGridDimensions();
        for (var ci = 0; ci < currentLayout.positions.length; ci++) {
          var _cp = customPositions.find(function(c) { return c.index === currentLayout.positions[ci].index; });
          if (_cp && _cp.gridX !== null) {
            currentLayout.positions[ci].x = _cp.gridX * _grid.cellPx;
            currentLayout.positions[ci].y = _cp.gridY * _grid.cellPx;
          }
        }
      }

      // Always use ortho (flat 2D top-down) for unified preview
      var useOrtho = true;

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
      const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true, preserveDrawingBuffer: true });
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
      const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true, preserveDrawingBuffer: true });
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
      // Reset custom grid positions when layout changes
      if (gridEditMode) { customPositions = null; }
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
      3: 'Controleer de plattegronden',
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
        if (unifiedFloorsOverlay) unifiedFloorsOverlay.classList.toggle('zone-editing', n === 4 && gridEditMode);
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
          var resultSection = document.getElementById('layoutResultSection');
          var noteSection = document.querySelector('.layout-note-section');
          var controlsBar = document.getElementById('layoutControlsBar');

          if (layoutCalculated) {
            // ── Returning from step 5 — restore previous state ──
            renderLayoutView();
            if (resultSection) resultSection.style.display = '';
            if (noteSection) noteSection.style.display = '';
            // Always show tools box
            if (controlsBar) controlsBar.style.display = '';
            var btnCalcBack = document.getElementById('btnCalcLayout');
            if (btnCalcBack) btnCalcBack.style.display = 'none';
            renderPreviewThumbnails();
            updateFloorLabels();
            setTimeout(function() {
              renderGridOverlay();
              // Re-enable drag mode if it was on
              if (gridEditMode) {
                var overlay = document.getElementById('unifiedFloorsOverlay');
                if (overlay) overlay.classList.add('drag-enabled');
                var wraps = floorsGrid ? floorsGrid.querySelectorAll('.floor-canvas-wrap') : [];
                for (var _wi = 0; _wi < wraps.length; _wi++) {
                  attachDragHandlers(wraps[_wi], _wi);
                }
                addGridFloorButtons();
              }
            }, 50);
          } else {
            // ── First visit — smart pre-check based on floor names ──
            excludedFloors = new Set();
            for (var ei = 0; ei < floors.length; ei++) {
              var fn = (floors[ei].name || '').toLowerCase();
              var skipKeywords = ['situatie', 'site', 'tuin', 'garden', 'buitenruimte', 'omgeving',
                'terrein', 'perceel', 'berging', 'garage', 'schuur', 'storage', 'dakterras', 'balkon', 'parkeer'];
              var shouldExclude = skipKeywords.some(function(kw) { return fn.includes(kw); });
              if (shouldExclude) excludedFloors.add(ei);
            }
            layoutCalculated = false;
            customPositions = null;
            gridEditMode = false;

            renderLayoutView();

            // Hide result section + note section until "Bereken indeling"
            if (resultSection) resultSection.style.display = 'none';
            if (noteSection) noteSection.style.display = 'none';
            if (controlsBar) controlsBar.style.display = 'none';

            // Clear preview until calculation (show empty frame)
            if (floorsGrid) floorsGrid.innerHTML = '';

            // Wire up the "Bereken indeling" button
            var btnCalc = document.getElementById('btnCalcLayout');
            if (btnCalc) {
              btnCalc.style.display = '';
              // Enable if any floor is already pre-checked
              var anyPreChecked = false;
              for (var _pc = 0; _pc < floors.length; _pc++) {
                if (!excludedFloors.has(_pc)) { anyPreChecked = true; break; }
              }
              btnCalc.disabled = !anyPreChecked;
              btnCalc.onclick = function() {
                layoutCalculated = true;
                this.style.display = 'none';
                if (resultSection) resultSection.style.display = '';
                if (noteSection) noteSection.style.display = '';
                // Always show tools box (per-floor controls + rotation work for 1+ floors)
                if (controlsBar) controlsBar.style.display = '';
                // Ensure edit mode checkbox starts unchecked, controls disabled
                var chkEM = document.getElementById('chkEditMode');
                var ctrlIn = document.getElementById('layoutControlsInner');
                if (chkEM) chkEM.checked = false;
                if (ctrlIn) ctrlIn.classList.add('disabled');
                gridEditMode = false;
                customPositions = null;
                renderPreviewThumbnails();
                updateFloorLabels();
                renderLayoutView(); // populate per-floor controls
                updateWizardUI(); // show "Volgende" now
                setTimeout(function() { renderGridOverlay(); }, 50);
              };
            }
          }
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
        // Show Funda load button only when no floors loaded yet
        if (btnFunda) {
          if (floors.length > 0 || noFloorsMode) {
            btnFunda.style.display = 'none';
          } else {
            btnFunda.style.display = '';
            btnFunda.disabled = false;
          }
        }
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
      } else if (currentWizardStep === 4) {
        // Step 4: hide "Volgende" until layout is calculated
        btnWizardNext.style.display = layoutCalculated ? '' : 'none';
        // Disable if floors overlap
        if (layoutCalculated && layoutHasOverlap) {
          btnWizardNext.disabled = true;
        }
        // Show/hide overlap warning (static element in HTML)
        var overlapWarn = document.getElementById('overlapWarning');
        if (overlapWarn) overlapWarn.style.display = (layoutHasOverlap && layoutCalculated) ? '' : 'none';
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
      // In step 3: go to previous floor first, then back to step 2
      if (currentWizardStep === 3 && currentFloorReviewIndex > 0) {
        currentFloorReviewIndex--;
        renderFloorReview();
        return;
      }
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

      // Hide thumbstrip when only 1 floor
      if (floors.length <= 1) {
        strip.style.display = 'none';
        return;
      }
      strip.style.display = '';

      // Build thumbnails
      for (let i = 0; i < floors.length; i++) {
        const thumb = document.createElement('div');
        thumb.className = 'floor-thumb';
        if (i === currentFloorReviewIndex) thumb.classList.add('active');
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
        // Remove old status indicator
        var oldDot = t.querySelector('.floor-thumb-status');
        if (oldDot) oldDot.remove();
        // Add status indicator based on review result
        var status = floorReviewStatus[i];
        if (status) {
          var dot = document.createElement('div');
          dot.className = 'floor-thumb-status status-' + status;
          if (status === 'confirmed') {
            dot.innerHTML = '<svg viewBox="0 0 12 12"><polyline points="2.5,6 5,8.5 9.5,3.5"/></svg>';
          } else {
            dot.innerHTML = '<svg viewBox="0 0 12 12"><line x1="3" y1="3" x2="9" y2="9"/><line x1="9" y1="3" x2="3" y2="9"/></svg>';
          }
          t.appendChild(dot);
        }
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

      // Update header: counter + floor name
      var counterEl = document.getElementById('floorReviewCounter');
      var nameEl = document.getElementById('floorReviewName');
      if (counterEl) {
        counterEl.textContent = floors.length > 1 ? 'Plattegrond ' + (currentFloorReviewIndex + 1) + ' van ' + floors.length : '';
      }
      if (nameEl) {
        nameEl.textContent = floors[currentFloorReviewIndex].name || 'Verdieping ' + (currentFloorReviewIndex + 1);
      }

      // Show appropriate panel based on existing review status
      var existingStatus = floorReviewStatus[currentFloorReviewIndex];
      if (existingStatus === 'issue') {
        showFloorIssuePanel();
      } else if (existingStatus === 'major') {
        showFloorMajorPanel();
      } else {
        showFloorDefaultPanel();
      }

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

        // Add subtle interaction hint
        var hint = document.createElement('div');
        hint.className = 'floor-review-hint';
        hint.textContent = 'Sleep om te draaien \u00B7 scroll om te zoomen';
        floorReviewViewerEl.appendChild(hint);
      }, 60);

      // Track viewed floors and update wizard UI
      viewedFloors.add(currentFloorReviewIndex);
      updateWizardUI();

      // Update thumbstrip state (highlight active + status indicators)
      updateThumbstripState();
    }

    function navigateFloorReview(direction) {
      currentFloorReviewIndex += direction;
      if (currentFloorReviewIndex < 0) currentFloorReviewIndex = floors.length - 1;
      if (currentFloorReviewIndex >= floors.length) currentFloorReviewIndex = 0;
      renderFloorReview();
    }

    // ── Step 3: New review handlers ──

    // Helper: advance to next unreviewed floor, or step 4 if all reviewed
    function advanceFloorReview() {
      renderPreviewThumbnails();
      updateFloorLabels();
      // Find next unreviewed floor
      var nextUnreviewed = -1;
      for (var _ur = 0; _ur < floors.length; _ur++) {
        if (!floorReviewStatus[_ur]) { nextUnreviewed = _ur; break; }
      }
      if (nextUnreviewed === -1) {
        // All floors reviewed — proceed to step 4
        showWizardStep(4);
      } else {
        currentFloorReviewIndex = nextUnreviewed;
        renderFloorReview();
      }
    }

    // "Ziet er goed uit →"
    function confirmFloorGood() {
      ensureDomRefs();
      viewedFloors.add(currentFloorReviewIndex);
      floorReviewStatus[currentFloorReviewIndex] = 'confirmed';
      advanceFloorReview();
    }

    // Show issue panel (small difference)
    function showFloorIssuePanel() {
      document.getElementById('floorReviewDefault').style.display = 'none';
      document.getElementById('floorReviewIssue').style.display = '';
      document.getElementById('floorReviewMajor').style.display = 'none';
      // Restore previous note if one was saved for this floor
      var textarea = document.getElementById('floorIssueText');
      if (textarea) textarea.value = floorIssues[currentFloorReviewIndex] || '';
    }

    // "Opslaan & door →"
    function saveFloorIssue() {
      var textarea = document.getElementById('floorIssueText');
      var text = textarea ? textarea.value.trim() : '';
      if (text) {
        floorIssues[currentFloorReviewIndex] = text;
        floorReviewStatus[currentFloorReviewIndex] = 'issue';
      } else {
        // Empty note = treat as confirmed
        delete floorIssues[currentFloorReviewIndex];
        floorReviewStatus[currentFloorReviewIndex] = 'confirmed';
      }
      viewedFloors.add(currentFloorReviewIndex);
      advanceFloorReview();
    }

    // Show major panel (floor is very wrong)
    function showFloorMajorPanel() {
      document.getElementById('floorReviewDefault').style.display = 'none';
      document.getElementById('floorReviewIssue').style.display = 'none';
      document.getElementById('floorReviewMajor').style.display = '';
      // Set up contact email link
      var contactBtn = document.getElementById('btnFloorContact');
      if (contactBtn) {
        var fundaUrl = fundaUrlInput ? fundaUrlInput.value.trim() : '';
        var floorName = floors[currentFloorReviewIndex] ? floors[currentFloorReviewIndex].name : 'Onbekend';
        var subject = encodeURIComponent('Frame³ — plattegrond klopt niet');
        var body = encodeURIComponent('Hoi,\n\nDe plattegrond "' + floorName + '" klopt niet.\n\nFunda link: ' + (fundaUrl || '(niet ingevuld)') + '\n\nKunnen jullie me helpen?\n\nAlvast bedankt!');
        contactBtn.href = 'mailto:vince@mattori.nl?subject=' + subject + '&body=' + body;
      }
    }

    // "Doorgaan →" (major issue, proceed on trust)
    function confirmFloorMajor() {
      viewedFloors.add(currentFloorReviewIndex);
      floorReviewStatus[currentFloorReviewIndex] = 'major';
      advanceFloorReview();
    }

    // "← Terug" from major panel
    function showFloorDefaultPanel() {
      document.getElementById('floorReviewDefault').style.display = '';
      document.getElementById('floorReviewIssue').style.display = 'none';
      document.getElementById('floorReviewMajor').style.display = 'none';
    }

    // ============================================================
    // ADMIN: Toggle frame image (One.png ↔ Two.png)
    // ============================================================
    var FRAME_IMG_ONE = 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Two_zonder_huisje_56d68527-71ff-4333-bfcf-6f2e6eca7d95.png?v=1771605516';
    var FRAME_IMG_TWO = 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/One_2ef26725-1a92-4673-9c0b-3699a5be8e0a.png?v=1771605517';

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
      layoutAlignY = (cb && cb.checked) ? 'bottom' : 'center';
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
      if (gridEditMode) { customPositions = null; }
      updatePreviewWithLoading(function() {
        renderPreviewThumbnails();
        updateFloorLabels();
        if (gridEditMode) enableGridDrag();
      });
    }

    function rotateFloor90(floorIndex) {
      if (!floorSettings[floorIndex]) floorSettings[floorIndex] = {};
      var current = getFloorRotate(floorIndex);
      floorSettings[floorIndex].rotate = (current + 90) % 360;
      // Only reset this floor's position, keep others (dimensions change on rotate)
      resetSingleFloorPosition(floorIndex);
      renderPreviewThumbnails();
      renderGridOverlayIfStep4();
      updateFloorLabels();
      if (gridEditMode) enableGridDrag();
      checkFloorOverlaps();
      renderLayoutView();
    }

    // ============================================================
    // STEP 4: Layout View (grid-native)
    // ============================================================
    var floorOrder = null; // null = default order

    function renderLayoutView() {
      // Cleanup old layout viewers
      for (const v of layoutViewers) {
        if (v.renderer) v.renderer.dispose();
      }
      layoutViewers = [];
      floorLayoutViewer.innerHTML = '';

      // ── Include/exclude checkboxes bar ──
      var includeBar = document.getElementById('floorIncludeBar');
      if (includeBar) {
        includeBar.innerHTML = '';
        for (var fi = 0; fi < floors.length; fi++) {
          (function(floorIdx) {
            var isExcluded = excludedFloors.has(floorIdx);
            var chip = document.createElement('label');
            chip.className = 'floor-include-chip' + (isExcluded ? ' excluded' : '');
            var cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = !isExcluded;
            cb.addEventListener('change', function() {
              // Reset calculation state BEFORE toggling (prevents premature render)
              var wasCalculated = layoutCalculated;
              if (layoutCalculated) {
                layoutCalculated = false;
                gridEditMode = false;
              }
              customPositions = null;

              toggleFloorExclusion(floorIdx);

              // Hide result + note sections after reset
              if (wasCalculated) {
                var resultSection2 = document.getElementById('layoutResultSection');
                if (resultSection2) resultSection2.style.display = 'none';
                var noteSection2 = document.querySelector('.layout-note-section');
                if (noteSection2) noteSection2.style.display = 'none';
                var controlsBar2 = document.getElementById('layoutControlsBar');
                if (controlsBar2) controlsBar2.style.display = 'none';
                // Clear preview (guard in renderPreviewThumbnails handles this now)
                if (floorsGrid) floorsGrid.innerHTML = '';
                updateWizardUI();
              }
              // Update checkbox UI + show "Bereken" button
              var chip2 = this.parentElement;
              if (chip2) chip2.className = 'floor-include-chip' + (excludedFloors.has(floorIdx) ? ' excluded' : '');
              var btnCalc2 = document.getElementById('btnCalcLayout');
              if (btnCalc2) {
                btnCalc2.style.display = '';
                btnCalc2.onclick = function() {
                  layoutCalculated = true;
                  this.style.display = 'none';
                  var rs = document.getElementById('layoutResultSection');
                  if (rs) rs.style.display = '';
                  var ns = document.querySelector('.layout-note-section');
                  if (ns) ns.style.display = '';
                  // Always show tools box
                  var cb2 = document.getElementById('layoutControlsBar');
                  if (cb2) cb2.style.display = '';
                  // Ensure edit mode starts unchecked
                  var chkEM2 = document.getElementById('chkEditMode');
                  var ctrlIn2 = document.getElementById('layoutControlsInner');
                  if (chkEM2) chkEM2.checked = false;
                  if (ctrlIn2) ctrlIn2.classList.add('disabled');
                  gridEditMode = false;
                  customPositions = null;
                  renderPreviewThumbnails();
                  updateFloorLabels();
                  renderLayoutView(); // populate per-floor controls
                  updateWizardUI();
                  setTimeout(function() { renderGridOverlay(); }, 50);
                };
                var anyIncluded = false;
                for (var ci = 0; ci < floors.length; ci++) {
                  if (!excludedFloors.has(ci)) { anyIncluded = true; break; }
                }
                btnCalc2.disabled = !anyIncluded;
              }
            });
            var label = document.createElement('span');
            label.textContent = floors[floorIdx].name || ('Verdieping ' + (floorIdx + 1));
            chip.appendChild(cb);
            chip.appendChild(label);
            includeBar.appendChild(chip);
          })(fi);
        }
      }

      // (Global X/Y alignment buttons removed — per-floor alignment handles this)

      // ── Scale buttons ──
      var scaleGroup = document.getElementById('scaleGroup');
      if (scaleGroup) {
        scaleGroup.innerHTML = '';
        var scaleOptions = [
          { val: 0.82, label: 'S', title: 'Klein' },
          { val: 1.0, label: 'M', title: 'Normaal' },
          { val: 1.1, label: 'L', title: 'Groot' }
        ];
        for (var si = 0; si < scaleOptions.length; si++) {
          (function(opt) {
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'floor-align-btn' + (opt.val === layoutScaleFactor ? ' active' : '');
            btn.textContent = opt.label;
            btn.title = opt.title;
            btn.style.cssText = 'font-size:11px;font-weight:700;min-width:28px;';
            btn.addEventListener('click', function() {
              layoutScaleFactor = opt.val;
              customPositions = null;
              var siblings = scaleGroup.querySelectorAll('.floor-align-btn');
              for (var s = 0; s < siblings.length; s++) siblings[s].classList.remove('active');
              this.classList.add('active');
              updatePreviewWithLoading(function() {
                renderPreviewThumbnails();
                renderGridOverlayIfStep4();
                updateFloorLabels();
                if (gridEditMode) enableGridDrag();
              });
            });
            scaleGroup.appendChild(btn);
          })(scaleOptions[si]);
        }
      }

      // ── Per-floor alignment + rotation section ──
      var perFloorSection = document.getElementById('perFloorAlignSection');
      if (perFloorSection) {
        perFloorSection.innerHTML = '';
        var includedForAlign = [];
        for (var pfi = 0; pfi < floors.length; pfi++) {
          if (!excludedFloors.has(pfi)) includedForAlign.push(pfi);
        }
        // Show per-floor controls for ALL included floors (even with 1 floor for rotation)
        if (includedForAlign.length >= 1) {
          // SVG line icons for alignment (small 14×14)
          var xIcons = {
            left:   '<svg width="14" height="14" viewBox="0 0 14 14"><line x1="1.5" y1="1" x2="1.5" y2="13" stroke="currentColor" stroke-width="1.5"/><rect x="3" y="3" width="8" height="3" rx="0.5" fill="currentColor" opacity="0.7"/><rect x="3" y="8" width="5" height="3" rx="0.5" fill="currentColor" opacity="0.35"/></svg>',
            center: '<svg width="14" height="14" viewBox="0 0 14 14"><line x1="7" y1="1" x2="7" y2="13" stroke="currentColor" stroke-width="1" stroke-dasharray="2 1.5"/><rect x="2" y="3" width="10" height="3" rx="0.5" fill="currentColor" opacity="0.7"/><rect x="3.5" y="8" width="7" height="3" rx="0.5" fill="currentColor" opacity="0.35"/></svg>',
            right:  '<svg width="14" height="14" viewBox="0 0 14 14"><line x1="12.5" y1="1" x2="12.5" y2="13" stroke="currentColor" stroke-width="1.5"/><rect x="3" y="3" width="8" height="3" rx="0.5" fill="currentColor" opacity="0.7"/><rect x="6" y="8" width="5" height="3" rx="0.5" fill="currentColor" opacity="0.35"/></svg>'
          };
          var yIcons = {
            top:    '<svg width="14" height="14" viewBox="0 0 14 14"><line x1="1" y1="1.5" x2="13" y2="1.5" stroke="currentColor" stroke-width="1.5"/><rect x="2" y="3" width="4" height="8" rx="0.5" fill="currentColor" opacity="0.7"/><rect x="8" y="3" width="4" height="5" rx="0.5" fill="currentColor" opacity="0.35"/></svg>',
            center: '<svg width="14" height="14" viewBox="0 0 14 14"><line x1="1" y1="7" x2="13" y2="7" stroke="currentColor" stroke-width="1" stroke-dasharray="2 1.5"/><rect x="2" y="1" width="4" height="12" rx="0.5" fill="currentColor" opacity="0.7"/><rect x="8" y="3" width="4" height="8" rx="0.5" fill="currentColor" opacity="0.35"/></svg>',
            bottom: '<svg width="14" height="14" viewBox="0 0 14 14"><line x1="1" y1="12.5" x2="13" y2="12.5" stroke="currentColor" stroke-width="1.5"/><rect x="2" y="3" width="4" height="8" rx="0.5" fill="currentColor" opacity="0.7"/><rect x="8" y="6" width="4" height="5" rx="0.5" fill="currentColor" opacity="0.35"/></svg>'
          };
          var rotateIcon = '<svg width="13" height="13" viewBox="0 0 16 16" fill="none"><path d="M13 3v4h-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M13 7c-.8-2.2-2.8-3.8-5.2-3.8-3.2 0-5.8 2.6-5.8 5.8s2.6 5.8 5.8 5.8c2.2 0 4-1.3 4.9-3.2" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';

          for (var afi = 0; afi < includedForAlign.length; afi++) {
            (function(floorIdx) {
              var row = document.createElement('div');
              row.className = 'per-floor-row';

              var name = document.createElement('span');
              name.className = 'per-floor-name';
              name.textContent = floors[floorIdx].name || ('Verdieping ' + (floorIdx + 1));
              row.appendChild(name);

              // X alignment buttons (line icons)
              var xBtns = document.createElement('span');
              xBtns.className = 'per-floor-btns';
              var xVals = ['left', 'center', 'right'];
              var currentX = (floorSettings[floorIdx] && floorSettings[floorIdx].alignX) || layoutAlignX;
              for (var bx = 0; bx < xVals.length; bx++) {
                (function(val) {
                  var btn = document.createElement('button');
                  btn.type = 'button';
                  btn.className = 'per-floor-btn' + (val === currentX ? ' active' : '');
                  btn.innerHTML = xIcons[val];
                  btn.addEventListener('click', function() {
                    if (!floorSettings[floorIdx]) floorSettings[floorIdx] = {};
                    floorSettings[floorIdx].alignX = val;
                    reSnapFloorAlignment(floorIdx);
                    renderPreviewThumbnails();
                    renderGridOverlayIfStep4();
                    updateFloorLabels();
                    if (gridEditMode) enableGridDrag();
                    checkFloorOverlaps();
                    renderLayoutView();
                  });
                  xBtns.appendChild(btn);
                })(xVals[bx]);
              }
              row.appendChild(xBtns);

              // Y alignment buttons (line icons)
              var yBtns = document.createElement('span');
              yBtns.className = 'per-floor-btns';
              var yVals = ['top', 'center', 'bottom'];
              var currentY = (floorSettings[floorIdx] && floorSettings[floorIdx].alignY) || layoutAlignY;
              for (var by = 0; by < yVals.length; by++) {
                (function(val) {
                  var btn = document.createElement('button');
                  btn.type = 'button';
                  btn.className = 'per-floor-btn' + (val === currentY ? ' active' : '');
                  btn.innerHTML = yIcons[val];
                  btn.addEventListener('click', function() {
                    if (!floorSettings[floorIdx]) floorSettings[floorIdx] = {};
                    floorSettings[floorIdx].alignY = val;
                    reSnapFloorAlignment(floorIdx);
                    renderPreviewThumbnails();
                    renderGridOverlayIfStep4();
                    updateFloorLabels();
                    if (gridEditMode) enableGridDrag();
                    checkFloorOverlaps();
                    renderLayoutView();
                  });
                  yBtns.appendChild(btn);
                })(yVals[by]);
              }
              row.appendChild(yBtns);

              // Rotate 90° button
              var rotBtn = document.createElement('button');
              rotBtn.type = 'button';
              rotBtn.className = 'per-floor-rotate';
              rotBtn.innerHTML = rotateIcon;
              rotBtn.title = '90° draaien';
              rotBtn.addEventListener('click', function() {
                rotateFloor90(floorIdx);
              });
              row.appendChild(rotBtn);

              perFloorSection.appendChild(row);
            })(includedForAlign[afi]);
          }
        }
      }

      // ── Reset button (icon inside label — stop event so checkbox doesn't toggle) ──
      var btnReset = document.getElementById('btnResetLayout');
      if (btnReset) {
        btnReset.style.display = customPositions ? 'inline-flex' : 'none';
        btnReset.onclick = function(e) {
          e.preventDefault();
          e.stopPropagation();
          customPositions = null;
          layoutHasOverlap = false;
          renderPreviewThumbnails();
          renderGridOverlay();
          updateWizardUI();
          if (gridEditMode) {
            var wraps = floorsGrid ? floorsGrid.querySelectorAll('.floor-canvas-wrap') : [];
            _dragCleanups.forEach(function(fn) { fn(); });
            _dragCleanups = [];
            for (var _rw = 0; _rw < wraps.length; _rw++) attachDragHandlers(wraps[_rw], _rw);
            addGridFloorButtons();
          }
          btnReset.style.display = 'none';
        };
      }

      // ── Edit mode + tools box ──
      var includedCount = 0;
      for (var gi = 0; gi < floors.length; gi++) {
        if (!excludedFloors.has(gi)) includedCount++;
      }
      // Show/hide the tools box based on layout state
      var ctrlBar = document.getElementById('layoutControlsBar');
      if (ctrlBar) ctrlBar.style.display = (layoutCalculated && includedCount >= 2) ? '' : 'none';

      // "Handmatig aanpassen" checkbox toggles edit mode
      var chkEditMode = document.getElementById('chkEditMode');
      var ctrlInner = document.getElementById('layoutControlsInner');
      if (chkEditMode) {
        chkEditMode.checked = gridEditMode;
        // Apply disabled state to controls
        if (ctrlInner) ctrlInner.classList.toggle('disabled', !gridEditMode);
        chkEditMode.onchange = function() {
          gridEditMode = this.checked;
          if (ctrlInner) ctrlInner.classList.toggle('disabled', !gridEditMode);
          if (gridEditMode) {
            enableGridDrag();
          } else {
            disableGridDrag();
            customPositions = null;
            layoutHasOverlap = false;
            renderPreviewThumbnails();
            renderGridOverlay();
            updateWizardUI();
          }
        };
      }

      // ── Grid visibility checkbox ──
      var chkShowGrid = document.getElementById('chkShowGrid');
      if (chkShowGrid) {
        chkShowGrid.onchange = function() {
          showGridOverlay = this.checked;
          // Also toggle the zone-editing dashed border
          var overlay = document.getElementById('unifiedFloorsOverlay');
          if (overlay) overlay.classList.toggle('zone-editing', showGridOverlay);
          if (showGridOverlay) {
            renderGridOverlay();
            addGridFloorButtons();
          } else {
            // Remove grid overlay
            var gridEl = floorsGrid ? floorsGrid.querySelector('.grid-overlay') : null;
            if (gridEl) gridEl.remove();
            // Remove alignment line SVGs from floor wraps
            var wraps = floorsGrid ? floorsGrid.querySelectorAll('.floor-canvas-wrap svg') : [];
            for (var _sv = 0; _sv < wraps.length; _sv++) wraps[_sv].remove();
          }
        };
      }

      // Show grid overlay in step 4 (always visible, subtle)
      renderGridOverlayIfStep4();
    }

    // Helper: show grid overlay when in step 4
    function renderGridOverlayIfStep4() {
      if (currentWizardStep === 4) {
        renderGridOverlay();
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

      // Only show toggle when 2-3 floors; force single for 1 or >3
      var includedCount = floorLabels.length;
      var showToggle = includedCount >= 2 && includedCount <= 3;
      if (!showToggle) {
        labelMode = 'single';
      }

      if (showToggle) {
        // Toggle switch
        var toggle = document.createElement('div');
        toggle.className = 'label-mode-toggle';

        var labelSingle = document.createElement('span');
        labelSingle.className = 'label-mode-label' + (labelMode === 'single' ? ' active' : '');
        labelSingle.textContent = '1 onderschrift';

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
      }

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
        span.textContent = 'Onderschrift';

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

      // Merge hole polygons into outer ring via bridge edges
      // so ear-clipping can triangulate a polygon-with-holes as one simple polygon.
      // outerPts: [{x,y}...] CCW, holePtsList: [ [{x,y}...], ... ] each CW
      function mergeHolesIntoPoly(outerPts, holePtsList) {
        if (!holePtsList || holePtsList.length === 0) return outerPts;
        // Sort holes by rightmost x (descending) — standard for bridge algorithm
        const sorted = holePtsList.slice().sort((a, b) => {
          let maxA = -Infinity, maxB = -Infinity;
          for (const p of a) if (p.x > maxA) maxA = p.x;
          for (const p of b) if (p.x > maxB) maxB = p.x;
          return maxB - maxA;
        });
        let merged = outerPts.slice();
        for (const hole of sorted) {
          if (hole.length < 3) continue;
          // Find rightmost vertex of hole
          let hrIdx = 0;
          for (let i = 1; i < hole.length; i++) {
            if (hole[i].x > hole[hrIdx].x) hrIdx = i;
          }
          const M = hole[hrIdx];
          // Cast ray from M in +x direction, find closest intersecting edge of merged
          let closestDist = Infinity, closestEdgeIdx = -1, closestIntX = 0;
          for (let i = 0; i < merged.length; i++) {
            const j = (i + 1) % merged.length;
            const y1 = merged[i].y, y2 = merged[j].y;
            if ((y1 - M.y) * (y2 - M.y) > 0) continue; // both on same side
            if (Math.abs(y1 - y2) < 1e-10) continue; // horizontal edge
            const t = (M.y - y1) / (y2 - y1);
            if (t < -1e-10 || t > 1 + 1e-10) continue;
            const xi = merged[i].x + t * (merged[j].x - merged[i].x);
            if (xi < M.x - 1e-10) continue; // to the left
            const dist = xi - M.x;
            if (dist < closestDist) {
              closestDist = dist;
              closestEdgeIdx = i;
              closestIntX = xi;
            }
          }
          // Determine bridge vertex on the outer polygon
          let bridgeIdx;
          if (closestEdgeIdx >= 0) {
            const ei = closestEdgeIdx;
            const ej = (ei + 1) % merged.length;
            // If intersection is very close to a vertex, use that vertex
            if (Math.abs(closestIntX - merged[ei].x) < 0.01 &&
                Math.abs(M.y - merged[ei].y) < 0.01) {
              bridgeIdx = ei;
            } else if (Math.abs(closestIntX - merged[ej].x) < 0.01 &&
                       Math.abs(M.y - merged[ej].y) < 0.01) {
              bridgeIdx = ej;
            } else {
              // Use the vertex with larger x as the visible bridge candidate
              bridgeIdx = merged[ei].x >= merged[ej].x ? ei : ej;
            }
          } else {
            // Fallback: closest vertex by distance
            let bestDist = Infinity;
            bridgeIdx = 0;
            for (let i = 0; i < merged.length; i++) {
              const d = Math.hypot(merged[i].x - M.x, merged[i].y - M.y);
              if (d < bestDist) { bestDist = d; bridgeIdx = i; }
            }
          }
          // Build hole path starting from hrIdx (complete loop back to start)
          const holeSeq = [];
          for (let i = 0; i <= hole.length; i++) {
            const src = hole[(hrIdx + i) % hole.length];
            holeSeq.push({ x: src.x, y: src.y });
          }
          // Splice: merged[0..bridgeIdx] → holeSeq → bridge back → merged[bridgeIdx+1..end]
          merged = [
            ...merged.slice(0, bridgeIdx + 1),
            ...holeSeq,
            { x: merged[bridgeIdx].x, y: merged[bridgeIdx].y },
            ...merged.slice(bridgeIdx + 1)
          ];
        }
        return merged;
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
        // polygon = [outerRing, ...holeRings]
        const outerRing = polygon[0];
        const holeRings = polygon.slice(1).filter(r => r.length >= 3);

        if (holeRings.length > 0 && typeof THREE !== 'undefined' && THREE.ShapeUtils) {
          // Polygon WITH holes (e.g. closed perimeter like Vliering)
          // Use THREE.ShapeUtils for proper triangulation with interior cutout
          function cleanRing(ring) {
            const pts = ring.slice();
            if (pts.length > 1) {
              const f = pts[0], l = pts[pts.length - 1];
              if (Math.hypot(f[0] - l[0], f[1] - l[1]) < 0.01) pts.pop();
            }
            return pts;
          }
          const outerPts = cleanRing(outerRing);
          if (outerPts.length < 3) continue;
          const outerObj = outerPts.map(p => ({
            x: (p[0] - centerX) * SCALE,
            y: (p[1] - centerY) * SCALE
          }));
          const holeObjArr = [];
          for (const hRing of holeRings) {
            const hPts = cleanRing(hRing);
            if (hPts.length >= 3) {
              holeObjArr.push(hPts.map(p => ({
                x: (p[0] - centerX) * SCALE,
                y: (p[1] - centerY) * SCALE
              })));
            }
          }

          const contour = outerObj.map(p => new THREE.Vector2(p.x, p.y));
          const holes = holeObjArr.map(h => h.map(p => new THREE.Vector2(p.x, p.y)));
          const tris = THREE.ShapeUtils.triangulateShape(contour, holes);

          // Combined vertex array: outer + all hole vertices
          const allPts = [...outerObj];
          for (const hole of holeObjArr) allPts.push(...hole);

          const baseBot = vertexIndex;
          for (const pt of allPts) {
            vertices.push(`v ${pt.x.toFixed(4)} ${(0).toFixed(4)} ${pt.y.toFixed(4)}`);
          }
          vertexIndex += allPts.length;
          const baseTop = vertexIndex;
          for (const pt of allPts) {
            vertices.push(`v ${pt.x.toFixed(4)} ${WALL_HEIGHT.toFixed(4)} ${pt.y.toFixed(4)}`);
          }
          vertexIndex += allPts.length;
          for (const [a, b, c] of tris) {
            addTriFace(baseBot + a, baseBot + c, baseBot + b);
            addTriFace(baseTop + a, baseTop + b, baseTop + c);
          }

          // Side faces — outer ring
          const nO = outerObj.length;
          const sbO = vertexIndex;
          for (const pt of outerObj) vertices.push(`v ${pt.x.toFixed(4)} ${(0).toFixed(4)} ${pt.y.toFixed(4)}`);
          vertexIndex += nO;
          const stO = vertexIndex;
          for (const pt of outerObj) vertices.push(`v ${pt.x.toFixed(4)} ${WALL_HEIGHT.toFixed(4)} ${pt.y.toFixed(4)}`);
          vertexIndex += nO;
          for (let i = 0; i < nO; i++) {
            const j = (i + 1) % nO;
            addFace(sbO + i, sbO + j, stO + j, stO + i);
          }

          // Side faces — each hole ring (inner wall surfaces)
          for (const holePts of holeObjArr) {
            const nH = holePts.length;
            const sbH = vertexIndex;
            for (const pt of holePts) vertices.push(`v ${pt.x.toFixed(4)} ${(0).toFixed(4)} ${pt.y.toFixed(4)}`);
            vertexIndex += nH;
            const stH = vertexIndex;
            for (const pt of holePts) vertices.push(`v ${pt.x.toFixed(4)} ${WALL_HEIGHT.toFixed(4)} ${pt.y.toFixed(4)}`);
            vertexIndex += nH;
            for (let i = 0; i < nH; i++) {
              const j = (i + 1) % nH;
              addFace(sbH + j, sbH + i, stH + i, stH + j); // reversed winding
            }
          }
        } else {
          // Simple polygon (no holes) — extrude as before
          extrudeWallPoly(outerRing, 0, WALL_HEIGHT);
        }
      }

      // Render walls WITH openings as individual boxes (with L-junction extension)
      for (const wall of openingWalls) {
        // Compute L-junction extension for opening walls.
        // Key fix: at perpendicular junctions with SOLID walls, use NEGATIVE
        // extension to trim the opening wall flush with the solid wall's inner
        // face. This prevents overlapping geometry that breaks slicers.
        const _wdx = wall.b.x - wall.a.x, _wdy = wall.b.y - wall.a.y;
        const _wlen = Math.hypot(_wdx, _wdy);
        const _wIsDiag = _wlen > 0.1 && Math.min(Math.abs(_wdx / _wlen), Math.abs(_wdy / _wlen)) > 0.15;
        let owExtA = 0, owExtB = 0;
        if (!_wIsDiag && _wlen > 0.1) {
          const _ux = _wdx / _wlen, _uy = _wdy / _wlen;
          // Track perpendicular solid walls separately
          let perpSolidA = 0, perpSolidB = 0;
          let hasPerpSolidA = false, hasPerpSolidB = false;

          for (const other of walls) {
            if (other === wall) continue;
            const odx = other.b.x - other.a.x, ody = other.b.y - other.a.y;
            const olen = Math.hypot(odx, ody);
            if (olen < 0.1) continue;
            if (Math.min(Math.abs(odx / olen), Math.abs(ody / olen)) > 0.15) continue; // skip diag
            const otherHt = (other.thickness ?? 20) / 2;
            const otherIsSolid = !(other.openings && other.openings.length > 0);
            // Dot product: 0 = perpendicular, 1 = collinear
            const dot = Math.abs(_ux * (odx / olen) + _uy * (ody / olen));
            const isPerp = dot < 0.3;

            // Endpoint A
            if (Math.hypot(wall.a.x - other.a.x, wall.a.y - other.a.y) < 3 ||
                Math.hypot(wall.a.x - other.b.x, wall.a.y - other.b.y) < 3) {
              if (otherIsSolid && isPerp) {
                hasPerpSolidA = true;
                perpSolidA = Math.max(perpSolidA, otherHt);
              } else {
                owExtA = Math.max(owExtA, otherHt);
              }
            }
            // Endpoint B
            if (Math.hypot(wall.b.x - other.a.x, wall.b.y - other.a.y) < 3 ||
                Math.hypot(wall.b.x - other.b.x, wall.b.y - other.b.y) < 3) {
              if (otherIsSolid && isPerp) {
                hasPerpSolidB = true;
                perpSolidB = Math.max(perpSolidB, otherHt);
              } else {
                owExtB = Math.max(owExtB, otherHt);
              }
            }
          }
          // Perpendicular solid walls: trim to inner face (negative extension)
          if (hasPerpSolidA) owExtA = -perpSolidA;
          if (hasPerpSolidB) owExtB = -perpSolidB;
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


        // Subtract voids (stair openings) — per-polygon to avoid SweepLine crash
        for (const v of floorVoids) {
          if (v.length < 3) continue;
          const vRing = v.map(p => [p.x, p.y]);
          const vf = vRing[0], vl = vRing[vRing.length - 1];
          if (Math.hypot(vf[0] - vl[0], vf[1] - vl[1]) > 0.01) vRing.push([vf[0], vf[1]]);
          const newFloorResult = [];
          for (const poly of floorResult) {
            try {
              const diff = polygonClipping.difference([poly], [[vRing]]);
              for (const d of diff) newFloorResult.push(d);
            } catch (e) {
              newFloorResult.push(poly);
            }
          }
          floorResult = newFloorResult;
        }

        // Extrude each result polygon
        const botY = (-FLOOR_THICKNESS).toFixed(4);
        const topY = (0).toFixed(4);

        for (const polygon of floorResult) {
          // Collect outer ring
          const ring0 = polygon[0].slice();
          if (ring0.length > 1) {
            const ff = ring0[0], ll = ring0[ring0.length - 1];
            if (Math.hypot(ff[0] - ll[0], ff[1] - ll[1]) < 0.01) ring0.pop();
          }
          if (ring0.length < 3) continue;

          const outerObjPts = ring0.map(p => ({
            x: (p[0] - centerX) * SCALE,
            y: (p[1] - centerY) * SCALE
          }));

          // Collect hole rings
          const holeObjPtsArr = [];
          for (let hi = 1; hi < polygon.length; hi++) {
            const hRing = polygon[hi].slice();
            if (hRing.length > 1) {
              const hf = hRing[0], hl = hRing[hRing.length - 1];
              if (Math.hypot(hf[0] - hl[0], hf[1] - hl[1]) < 0.01) hRing.pop();
            }
            if (hRing.length >= 3) {
              holeObjPtsArr.push(hRing.map(p => ({
                x: (p[0] - centerX) * SCALE,
                y: (p[1] - centerY) * SCALE
              })));
            }
          }

          if (holeObjPtsArr.length > 0 && typeof THREE !== 'undefined' && THREE.ShapeUtils) {
            // Polygon WITH holes — use Three.js ShapeUtils for robust triangulation
            const contour = outerObjPts.map(p => new THREE.Vector2(p.x, p.y));
            const holes = holeObjPtsArr.map(h => h.map(p => new THREE.Vector2(p.x, p.y)));
            const tris = THREE.ShapeUtils.triangulateShape(contour, holes);

            // Combined vertex array: outer + all hole vertices
            const allPts = [...outerObjPts];
            for (const hole of holeObjPtsArr) allPts.push(...hole);

            const baseBot = vertexIndex;
            for (const pt of allPts) {
              vertices.push(`v ${pt.x.toFixed(4)} ${botY} ${pt.y.toFixed(4)}`);
            }
            vertexIndex += allPts.length;
            const baseTop = vertexIndex;
            for (const pt of allPts) {
              vertices.push(`v ${pt.x.toFixed(4)} ${topY} ${pt.y.toFixed(4)}`);
            }
            vertexIndex += allPts.length;
            for (const [a, b, c] of tris) {
              addTriFace(baseBot + a, baseBot + c, baseBot + b);
              addTriFace(baseTop + a, baseTop + b, baseTop + c);
            }

            // Side faces — outer ring
            const nOuter = outerObjPts.length;
            const sideBaseBotO = vertexIndex;
            for (const pt of outerObjPts) vertices.push(`v ${pt.x.toFixed(4)} ${botY} ${pt.y.toFixed(4)}`);
            vertexIndex += nOuter;
            const sideBaseTopO = vertexIndex;
            for (const pt of outerObjPts) vertices.push(`v ${pt.x.toFixed(4)} ${topY} ${pt.y.toFixed(4)}`);
            vertexIndex += nOuter;
            for (let i = 0; i < nOuter; i++) {
              const j = (i + 1) % nOuter;
              addFace(sideBaseBotO + i, sideBaseBotO + j, sideBaseTopO + j, sideBaseTopO + i);
            }

            // Side faces — each hole ring (inner walls of void)
            for (const holePts of holeObjPtsArr) {
              const nH = holePts.length;
              const sideBaseBotH = vertexIndex;
              for (const pt of holePts) vertices.push(`v ${pt.x.toFixed(4)} ${botY} ${pt.y.toFixed(4)}`);
              vertexIndex += nH;
              const sideBaseTopH = vertexIndex;
              for (const pt of holePts) vertices.push(`v ${pt.x.toFixed(4)} ${topY} ${pt.y.toFixed(4)}`);
              vertexIndex += nH;
              for (let i = 0; i < nH; i++) {
                const j = (i + 1) % nH;
                // Reverse winding for inner walls
                addFace(sideBaseBotH + j, sideBaseBotH + i, sideBaseTopH + i, sideBaseTopH + j);
              }
            }
          } else {
            // Simple polygon without holes — ear-clip triangulation
            const tris = earClipTriangulate(outerObjPts);
            const baseBot = vertexIndex;
            for (const pt of outerObjPts) {
              vertices.push(`v ${pt.x.toFixed(4)} ${botY} ${pt.y.toFixed(4)}`);
            }
            vertexIndex += outerObjPts.length;
            const baseTop = vertexIndex;
            for (const pt of outerObjPts) {
              vertices.push(`v ${pt.x.toFixed(4)} ${topY} ${pt.y.toFixed(4)}`);
            }
            vertexIndex += outerObjPts.length;
            for (const [a, b, c] of tris) {
              addTriFace(baseBot + a, baseBot + c, baseBot + b);
              addTriFace(baseTop + a, baseTop + b, baseTop + c);
            }

            // Side faces
            const nPts = outerObjPts.length;
            const sideBaseBot = vertexIndex;
            for (const pt of outerObjPts) vertices.push(`v ${pt.x.toFixed(4)} ${botY} ${pt.y.toFixed(4)}`);
            vertexIndex += nPts;
            const sideBaseTop = vertexIndex;
            for (const pt of outerObjPts) vertices.push(`v ${pt.x.toFixed(4)} ${topY} ${pt.y.toFixed(4)}`);
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

    // House icon picker — render options
    (function renderHouseIconPicker() {
      var container = document.getElementById('houseIconOptions');
      if (!container) return;
      container.innerHTML = '';
      houseIconOptions.forEach(function(opt) {
        var btn = document.createElement('div');
        btn.className = 'house-icon-option' + (opt.id === selectedHouseIcon ? ' active' : '');
        btn.dataset.iconId = opt.id;
        var img = document.createElement('img');
        img.src = opt.url;
        img.alt = opt.label;
        btn.appendChild(img);
        btn.addEventListener('click', function() {
          selectedHouseIcon = opt.id;
          // Update active state
          container.querySelectorAll('.house-icon-option').forEach(function(el) {
            el.classList.toggle('active', el.dataset.iconId === opt.id);
          });
          // Update preview
          if (frameHouseIcon) frameHouseIcon.src = opt.url;
        });
        container.appendChild(btn);
      });
    })();

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
      2: 'https://www.funda.nl/detail/koop/amsterdam/appartement-hoofdweg-275-1/89691599/',
      3: 'https://www.funda.nl/detail/koop/verkocht/deventer/huis-veenweg-79/43255889/',
      4: 'https://www.funda.nl/detail/koop/ursem-gem-koggenland/huis-tuinderij-43/89607983/',
      5: 'https://www.funda.nl/detail/koop/oss/huis-baronie-21/89607977/',
      6: 'https://www.funda.nl/detail/koop/heerlen/appartement-van-der-maesenstraat-24/43243330/',
      7: 'https://www.funda.nl/detail/koop/rumpt/huis-raadsteeg-4/89607972/',
      8: 'https://www.funda.nl/detail/koop/amsterdam/appartement-piet-gijzenbrugstraat-29-1/43243333/',
      9: 'https://www.funda.nl/detail/koop/oudorp/huis-esdoornlaan-56/43243326/',
      10: 'https://www.funda.nl/detail/koop/sint-michielsgestel/huis-de-mulder-3/43243308/',
      11: 'https://www.funda.nl/detail/koop/rotterdam/appartement-oleanderstraat-135/89607941/',
      12: 'https://www.funda.nl/detail/koop/purmerend/huis-olympiastraat-28/43243398/',
      13: 'https://www.funda.nl/detail/koop/apeldoorn/huis-tienwoningenweg-39/43243394/',
      14: 'https://www.funda.nl/detail/koop/beinsdorp/huis-venneperweg-576/43243387/',
      15: 'https://www.funda.nl/detail/koop/nijmegen/huis-korhoenstraat-44/89607921/',
      16: 'https://www.funda.nl/detail/koop/oostzaan/huis-dokter-rutgers-van-der-loeffstraat-15/43243370/',
      17: 'https://www.funda.nl/detail/koop/ederveen/huis-hoofdweg-116/43243379/',
      18: 'https://www.funda.nl/detail/koop/berlicum/huis-de-misse-1/43243248/',
      19: 'https://www.funda.nl/detail/koop/breda/huis-oede-van-hoornestraat-12/43243210/',
      20: 'https://www.funda.nl/detail/koop/breda/appartement-haagdijk-141/43132850/',
      21: 'https://www.funda.nl/detail/koop/waarland/huis-veluweweg-32-b/89499833/'
    };
    function pasteTestLink(n) {
      var input = document.getElementById('fundaUrl');
      if (input) input.value = TEST_LINKS[n] || TEST_LINKS[1];
    }
    for (const n of Array.from({length: 21}, (_, i) => i + 1)) {
      const btn = document.getElementById('btnTest' + n);
      if (btn) btn.addEventListener('click', () => pasteTestLink(n));
    }

    // Funda status checker
    function setFundaStatus(state, html) {
      var box = document.getElementById('fundaStatus');
      var icon = document.getElementById('fundaStatusIcon');
      var text = document.getElementById('fundaStatusText');
      if (!box) return;
      // Hide contact email from previous attempt
      var prevEmail = document.getElementById('contactEmailBtn');
      if (prevEmail) prevEmail.style.display = 'none';
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
      btn.textContent = 'Neem contact op';
      statusBox.parentNode.insertBefore(btn, statusBox.nextSibling);
    }

    async function loadFromFunda() {
      ensureDomRefs();
      noFloorsMode = false; // Reset on new attempt
      const url = getFundaUrl();
      clearError();
      if (!url) { setError('Voer een Funda URL in.'); return; }
      if (!url.includes('funda.nl')) {
        setFundaStatus('error', '<strong>Geen Funda link herkend</strong><span>Dit lijkt geen Funda-link te zijn. Controleer de link of neem contact op.</span>');
        showContactEmail(url);
        btnWizardNext.style.display = 'none';
        return;
      }

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
          setFundaStatus('success', '<strong>✓ Funda link herkend</strong><strong class="status-warning">✗ Geen interactieve plattegronden beschikbaar</strong><span>Geen zorgen — we bouwen je Frame\u00B3 handmatig op basis van de Funda-foto\'s.</span>');
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
        var saleStatusLine = data.sale_status ? '<span class="funda-address-line">🏷️ ' + data.sale_status + '</span>' : '';
        setFundaStatus('success', '<strong>✓ Funda link correct</strong><strong>✓ ' + data.floors.length + ' interactieve plattegrond' + (data.floors.length === 1 ? '' : 'en') + ' gevonden</strong><span class="funda-address-line">📍 ' + addrStr + '</span>' + saleStatusLine);

        processFloors(data);

        // Hide load button after successful load
        btnFunda.style.display = 'none';
      } catch (err) {
        if (err.message && (err.message.includes('Load failed') || err.message.includes('Failed to fetch'))) {
          setFundaStatus('error', '<strong>Verbinding mislukt</strong><span>Probeer het zo weer opnieuw.</span>');
        } else {
          setFundaStatus('error', `<strong>Fout</strong><span>Probeer het zo weer opnieuw.</span>`);
        }
        btnWizardNext.style.display = 'none';
      } finally {
        hideLoading();
        if (btnFunda.style.display !== 'none') btnFunda.disabled = false;
      }
    }

    // Order button — adds product to Shopify cart via Cart API
    // Capture screenshot of unified frame preview using html2canvas → localStorage
    // Stores per Funda link so multiple cart items each keep their own preview
    async function capturePreviewToLocalStorage(fundaLink) {
      var previewEl = document.getElementById('unifiedFramePreview');
      if (!previewEl || typeof html2canvas === 'undefined' || !fundaLink) {
        console.warn('[Mattori] Screenshot overgeslagen:', !previewEl ? 'geen preview element' : !fundaLink ? 'geen Funda link' : 'html2canvas niet geladen');
        return;
      }
      try {
        var canvas = await html2canvas(previewEl, {
          useCORS: true,
          allowTaint: false,
          backgroundColor: '#ffffff',
          scale: 0.75
        });
        var dataUrl = canvas.toDataURL('image/jpeg', 0.6);
        try {
          // Clean up old single-preview key from previous versions
          localStorage.removeItem('mattori_preview');
          var previews = JSON.parse(localStorage.getItem('mattori_previews') || '{}');
          previews[fundaLink] = dataUrl;
          localStorage.setItem('mattori_previews', JSON.stringify(previews));
        } catch (e) {
          console.warn('[Mattori] localStorage vol, probeer met alleen deze preview:', e);
          // Quota exceeded — try storing just this one preview (discard older ones)
          try {
            var fresh = {};
            fresh[fundaLink] = dataUrl;
            localStorage.setItem('mattori_previews', JSON.stringify(fresh));
          } catch (e2) { console.warn('[Mattori] localStorage opslag mislukt:', e2); }
        }
      } catch (e) {
        console.warn('[Mattori] Screenshot mislukt:', e);
      }
    }

    async function submitOrder() {
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
      if (orderBtn) { orderBtn.disabled = true; orderBtn.innerHTML = '<span class="btn-spinner"></span> Toevoegen...'; }

      // Force browser repaint so spinner is visible before heavy html2canvas work
      await new Promise(function(r) { requestAnimationFrame(function() { requestAnimationFrame(r); }); });

      var fundaLink = fundaUrlInput ? fundaUrlInput.value.trim() : '';
      var itemProperties = {};
      if (fundaLink) itemProperties['Funda link'] = fundaLink;

      // Address fields
      var street = addressStreet ? addressStreet.value.trim() : '';
      var city = addressCity ? addressCity.value.trim() : '';
      if (street) itemProperties['Adres regel 1'] = street;
      if (city) itemProperties['Adres regel 2'] = city;

      // House icon
      var houseOpt = houseIconOptions.find(function(o) { return o.id === selectedHouseIcon; });
      if (houseOpt) itemProperties['Huisje'] = houseOpt.label;

      if (noFloorsMode) itemProperties['Opmerking'] = 'Geen interactieve plattegronden — handmatig opbouwen';

      // Per-floor review status as individual order properties
      floors.forEach(function(floor, i) {
        var status = floorReviewStatus[i];
        if (!status) return;
        var floorName = floor.name || ('Verdieping ' + (i + 1));
        var key = 'Plattegrond ' + floorName;
        if (status === 'confirmed') {
          itemProperties[key] = '\u2713 Klopt';
        } else if (status === 'issue') {
          var note = floorIssues[i] || '';
          itemProperties[key] = '\u2717 Klopt niet' + (note ? ': \u201C' + note + '\u201D' : '');
        } else if (status === 'major') {
          itemProperties[key] = '\u2717 Klopt helemaal niet';
        }
      });

      // Floor labels (from step 5)
      var currentLabels = getIncludedFloorLabels();
      if (labelMode === 'single') {
        itemProperties['Onderschrift'] = singleLabelText;
      } else if (currentLabels.length > 0) {
        currentLabels.forEach(function(item) {
          itemProperties['Onderschrift ' + (floors[item.index] ? floors[item.index].name : 'Verdieping')] = item.label;
        });
      }

      // Grid positions (from manual layout adjustment in step 4)
      var gridProps = getGridPositionProperties();
      for (var gk in gridProps) {
        if (gridProps.hasOwnProperty(gk)) itemProperties[gk] = gridProps[gk];
      }

      // Layout note (from step 4)
      var layoutNoteEl = document.getElementById('layoutNote');
      var layoutNoteText = layoutNoteEl ? layoutNoteEl.value.trim() : '';
      if (layoutNoteText) itemProperties['Opmerking indeling'] = layoutNoteText;

      // Hide grid overlay before screenshot
      var _gridOverlayEl = floorsGrid ? floorsGrid.querySelector('.grid-overlay') : null;
      if (_gridOverlayEl) _gridOverlayEl.style.display = 'none';

      // Save preview screenshot to localStorage keyed by Funda link (skip in noFloorsMode)
      if (!noFloorsMode && fundaLink) {
        await capturePreviewToLocalStorage(fundaLink);
      }

      // Restore grid overlay
      if (_gridOverlayEl) _gridOverlayEl.style.display = '';

      try {
        var res = await fetch('/cart/add.js', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ items: [{ id: parseInt(variantId), quantity: 1, properties: itemProperties }] })
        });
        if (!res.ok) throw new Error('Status ' + res.status);
        window.location.href = '/cart';
      } catch (err) {
        showToast('Kon niet toevoegen aan winkelwagen.');
        if (orderBtn) { orderBtn.disabled = false; orderBtn.textContent = originalText; }
      }
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
      // Enable button now that JS is loaded and startConfigurator exists
      btnStartConfigurator.disabled = false;
    }
