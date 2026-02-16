// Minimal OrbitControls implementation (rotate + zoom + pan)
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

      // Initialize spherical from current camera position
      const offset = new THREE.Vector3().subVectors(camera.position, this.target);
      this._spherical.setFromVector3(offset);

      this._onPointerDown = this._onPointerDown.bind(this);
      this._onPointerMove = this._onPointerMove.bind(this);
      this._onPointerUp = this._onPointerUp.bind(this);
      this._onWheel = this._onWheel.bind(this);
      this._onContextMenu = e => e.preventDefault();

      domElement.addEventListener('pointerdown', this._onPointerDown);
      domElement.addEventListener('wheel', this._onWheel, { passive: false });
      domElement.addEventListener('contextmenu', this._onContextMenu);
    }

    _onPointerDown(e) {
      this.domElement.setPointerCapture(e.pointerId);
      // Left button = rotate, Right/Middle = pan
      if (e.button === 0) {
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
        this._panStart.set(e.clientX, e.clientY);
      }
    }

    _onPointerUp(e) {
      this.domElement.releasePointerCapture(e.pointerId);
      this.domElement.removeEventListener('pointermove', this._onPointerMove);
      this.domElement.removeEventListener('pointerup', this._onPointerUp);
      this._pointerType = null;
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

    const CANVAS_W = 260;
    const CANVAS_H = 400;

    // Wizard state
    let currentWizardStep = 1;
    const TOTAL_WIZARD_STEPS = 5;
    let currentFloorReviewIndex = 0;
    var viewedFloors = new Set();

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
        unifiedFloorsOverlay, unifiedLabelsOverlay,
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
    }

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

        const parts = streetSlug.split('-');
        const titleCased = parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
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
    // WALL ENDPOINT EXTENSION
    // ============================================================
    function extendWalls(walls) {
      const TOLERANCE = 3;

      function pointsNear(px, py, qx, qy) {
        return Math.hypot(px - qx, py - qy) < TOLERANCE;
      }

      function pointOnSegmentInterior(px, py, wallAx, wallAy, wallBx, wallBy) {
        const wdx = wallBx - wallAx, wdy = wallBy - wallAy;
        const wlen2 = wdx * wdx + wdy * wdy;
        if (wlen2 < 0.001) return false;
        const t = ((px - wallAx) * wdx + (py - wallAy) * wdy) / wlen2;
        if (t < 0.02 || t > 0.98) return false;
        const projX = wallAx + wdx * t;
        const projY = wallAy + wdy * t;
        return Math.hypot(px - projX, py - projY) < TOLERANCE;
      }

      const extended = walls.map(w => ({
        a: { x: w.a.x, y: w.a.y },
        b: { x: w.b.x, y: w.b.y },
        thickness: w.thickness ?? 20,
        openings: w.openings ?? [],
        _arcSeg: w._arcSeg ?? false,
        _heightA: w._heightA ?? 265,
        _heightB: w._heightB ?? 265
      }));

      for (let i = 0; i < extended.length; i++) {
        const wall = extended[i];
        if (wall._arcSeg) continue;
        const origAx = walls[i].a.x, origAy = walls[i].a.y;
        const origBx = walls[i].b.x, origBy = walls[i].b.y;
        const wdx = origBx - origAx;
        const wdy = origBy - origAy;
        const wlen = Math.hypot(wdx, wdy);
        if (wlen < 0.1) continue;
        const ux = wdx / wlen;
        const uy = wdy / wlen;

        let extendA = 0;
        let extendB = 0;

        for (let j = 0; j < walls.length; j++) {
          if (i === j) continue;
          const other = walls[j];
          const otherHalfThick = (other.thickness ?? 20) / 2;
          const otherDx = other.b.x - other.a.x;
          const otherDy = other.b.y - other.a.y;
          const otherLen = Math.hypot(otherDx, otherDy);
          if (otherLen < 0.1) continue;

          const sinAngle = Math.abs(ux * (otherDy / otherLen) - uy * (otherDx / otherLen));
          if (sinAngle < 0.1) continue;

          const ext = Math.min(otherHalfThick / sinAngle, otherHalfThick * 3);

          const aShares = pointsNear(origAx, origAy, other.a.x, other.a.y) ||
                          pointsNear(origAx, origAy, other.b.x, other.b.y);
          const aOnInt = pointOnSegmentInterior(origAx, origAy, other.a.x, other.a.y, other.b.x, other.b.y);
          if (aShares || aOnInt) extendA = Math.max(extendA, ext);

          const bShares = pointsNear(origBx, origBy, other.a.x, other.a.y) ||
                          pointsNear(origBx, origBy, other.b.x, other.b.y);
          const bOnInt = pointOnSegmentInterior(origBx, origBy, other.a.x, other.a.y, other.b.x, other.b.y);
          if (bShares || bOnInt) extendB = Math.max(extendB, ext);
        }

        if (extendA > 0) {
          wall.a.x -= ux * extendA;
          wall.a.y -= uy * extendA;
        }
        if (extendB > 0) {
          wall.b.x += ux * extendB;
          wall.b.y += uy * extendB;
        }
      }

      return extended;
    }

    // ============================================================
    // BALUSTRADE ENDPOINT EXTENSION
    // ============================================================
    function extendBalustrades(balustrades) {
      return balustrades.map(bal => {
        const ax = bal.a.x, ay = bal.a.y;
        const bx = bal.b.x, by = bal.b.y;
        const dx = bx - ax, dy = by - ay;
        const len = Math.hypot(dx, dy);
        if (len < 0.1) return { ...bal };
        const ux = dx / len, uy = dy / len;
        const ext = (bal.thickness ?? 10) * 1.0;
        return {
          a: { x: ax - ux * ext, y: ay - uy * ext },
          b: { x: bx + ux * ext, y: by + uy * ext },
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
          if ((surface.role ?? -1) !== 14) continue;
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
      const btnTest = document.getElementById('btnTest');
      if (btnTest) btnTest.style.display = 'none';
      var adminFrameToggle = document.getElementById('adminFrameToggle');
      if (adminFrameToggle) adminFrameToggle.style.display = '';

      // Switch to unified preview
      if (productHeroImage) productHeroImage.style.display = 'none';
      unifiedFramePreview.style.display = '';

      // Update address in preview
      updateFrameAddress();

      // Render thumbnails (hidden until step 5)
      if (unifiedFloorsOverlay) unifiedFloorsOverlay.style.display = 'none';
      if (floorsLoading) floorsLoading.classList.remove('hidden');
      setTimeout(() => {
        renderPreviewThumbnails();
        updateFloorLabels();
        if (floorsLoading) floorsLoading.classList.add('hidden');
      }, 50);

      // Stay on step 1, update UI to show "Volgende" button
      updateWizardUI();
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
    function buildFloorScene(floorIndex) {
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
        color: 0xA89478,
        flatShading: true,
        side: THREE.DoubleSide,
        shininess: 10
      });
      const floorMaterial = new THREE.MeshPhongMaterial({
        color: 0xA89478,
        flatShading: true,
        side: THREE.DoubleSide,
        shininess: 5
      });

      if (groups.walls) scene.add(new THREE.Mesh(groups.walls, wallMaterial));
      if (groups.floor) scene.add(new THREE.Mesh(groups.floor, floorMaterial));

      scene.add(new THREE.AmbientLight(0xffffff, 0.6));
      const dirLight = new THREE.DirectionalLight(0xffffff, 0.5);
      dirLight.position.set(0, 8, 5);
      scene.add(dirLight);
      const fillLight = new THREE.DirectionalLight(0xffffff, 0.2);
      fillLight.position.set(-4, 6, -1);
      scene.add(fillLight);

      return { scene, size, center };
    }

    // ============================================================
    // STATIC THUMBNAIL RENDER (for unified preview)
    // ============================================================
    function renderStaticThumbnail(floorIndex, container) {
      const result = buildFloorScene(floorIndex);
      if (!result) return;
      const { scene, size, center } = result;

      const rect = container.getBoundingClientRect();
      const width = Math.round(rect.width) || 200;
      const height = Math.round(rect.height) || 260;

      const maxDim = Math.max(size.x, size.y, size.z);
      const padding = 1.25;
      const halfW = (size.x * padding) / 2;
      const halfZ = (size.z * padding) / 2;
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

      // Single frame render — no animation loop
      renderer.render(scene, camera);
      container.appendChild(renderer.domElement);

      previewViewers.push({ renderer });
    }

    function renderPreviewThumbnails() {
      // Cleanup old preview viewers
      for (const v of previewViewers) {
        if (v.renderer) v.renderer.dispose();
      }
      previewViewers = [];

      floorsGrid.innerHTML = '';

      // Count included floors and determine single-floor mode
      var includedIndices = [];
      for (let i = 0; i < floors.length; i++) {
        if (!excludedFloors.has(i)) includedIndices.push(i);
      }
      // Respect custom floor order from layout step
      if (floorOrder && floorOrder.length === includedIndices.length) {
        includedIndices = floorOrder.slice();
      }
      const regularCount = includedIndices.filter(i => {
        const n = (floors[i].name || '').toLowerCase().trim();
        return !(/^(kelder|zolder|berging|garage|dak|tuin)/.test(n));
      }).length;
      const isSingleFloor = regularCount <= 1;

      for (const i of includedIndices) {
        const card = document.createElement('div');
        card.className = 'floor-card';

        const viewerWrap = document.createElement('div');
        viewerWrap.className = 'floor-canvas-wrap';

        card.appendChild(viewerWrap);
        floorsGrid.appendChild(card);

        renderStaticThumbnail(i, viewerWrap);
      }
    }

    // ============================================================
    // INTERACTIVE 3D VIEWER (for step 4 floor review)
    // ============================================================
    function renderInteractiveViewer(floorIndex, container) {
      const result = buildFloorScene(floorIndex);
      if (!result) return null;
      const { scene, size, center } = result;

      const rect = container.getBoundingClientRect();
      const width = Math.round(rect.width) || 400;
      const height = Math.round(rect.height) || 500;

      const padding = 1.25;
      const halfW = (size.x * padding) / 2;
      const halfZ = (size.z * padding) / 2;
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
      const { scene, size, center } = result;

      const rect = container.getBoundingClientRect();
      const width = Math.round(rect.width) || 200;
      const height = Math.round(rect.height) || 260;

      const padding = 1.3;
      const halfW = (size.x * padding) / 2;
      const halfZ = (size.z * padding) / 2;
      const aspect = width / height;

      let camHalfW, camHalfH;
      if (halfW / halfZ > aspect) {
        camHalfW = halfW;
        camHalfH = halfW / aspect;
      } else {
        camHalfH = halfZ;
        camHalfW = halfZ * aspect;
      }

      const camera = new THREE.OrthographicCamera(-camHalfW, camHalfW, camHalfH, -camHalfH, 0.01, 1000);
      camera.position.set(0, 50, 0);
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

    function showWizardStep(n) {
      if (n < 1 || n > TOTAL_WIZARD_STEPS) return;
      currentWizardStep = n;

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

        // Step 3: show floor review viewer in left column, hide unified preview
        if (n === 3) {
          if (unifiedFramePreview) unifiedFramePreview.style.display = 'none';
          if (floorReviewViewerEl) floorReviewViewerEl.style.display = '';
        } else {
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
          currentFloorReviewIndex = 0;
          viewedFloors.clear();
          buildThumbstrip();
          renderFloorReview();
        } else if (n === 4) {
          renderLayoutView();
          // Re-render unified preview thumbnails so they show during layout step
          renderPreviewThumbnails();
          updateFloorLabels();
        } else if (n === 5) {
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

      // Floor nav button (step 3 only — cycles through floors, hidden on last)
      var btnFN = document.getElementById('btnFloorNext');
      if (btnFN) {
        var showFloorNav = currentWizardStep === 3 && currentFloorReviewIndex < floors.length - 1;
        btnFN.style.display = showFloorNav ? '' : 'none';
      }

      if (currentWizardStep === TOTAL_WIZARD_STEPS) {
        // Last step: hide next, show order
        btnWizardNext.style.display = 'none';
      } else if (currentWizardStep === 1) {
        // Step 1: only show next if data is loaded
        btnWizardNext.style.display = floors.length > 0 ? '' : 'none';
        btnWizardNext.disabled = false;
      } else if (currentWizardStep === 3) {
        // Step 3 (floor review): disable next until all floors viewed
        btnWizardNext.style.display = '';
        btnWizardNext.disabled = viewedFloors.size < floors.length;
      } else {
        btnWizardNext.style.display = '';
        btnWizardNext.disabled = false;
      }
    }

    function nextWizardStep() {
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

      // Render interactive viewer
      floorReviewViewer = renderInteractiveViewer(currentFloorReviewIndex, floorReviewViewerEl);

      // Add zoom/rotate hint label
      var hint = document.createElement('div');
      hint.className = 'floor-review-hint';
      hint.textContent = 'Klik en sleep om te draaien \u00B7 scroll om te zoomen';
      floorReviewViewerEl.appendChild(hint);

      // Track viewed floors and update wizard UI
      viewedFloors.add(currentFloorReviewIndex);
      updateWizardUI();

      // Update thumbstrip state (highlight active, don't rebuild)
      updateThumbstripState();

      // Update checkbox
      floorIncludeCb.checked = !excludedFloors.has(currentFloorReviewIndex);

      // Show/hide excluded overlay
      const existingOverlay = floorReviewViewerEl.querySelector('.floor-review-excluded');
      if (existingOverlay) existingOverlay.remove();
      if (excludedFloors.has(currentFloorReviewIndex)) {
        const overlay = document.createElement('div');
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

    // "Er klopt iets niet" toggle
    function toggleFloorIssue() {
      var details = document.getElementById('floorIssueDetails');
      var cb = document.getElementById('floorIssueCb');
      if (details) details.style.display = cb && cb.checked ? '' : 'none';
    }

    // ============================================================
    // ADMIN: Toggle frame image (One.png ↔ Two.png)
    // ============================================================
    var FRAME_IMG_ONE = 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/One.png?v=1771252893';
    var FRAME_IMG_TWO = 'https://cdn.shopify.com/s/files/1/0958/8614/7958/files/Two.png?v=1771252896';

    function toggleAdminFrame() {
      ensureDomRefs();
      var cb = document.getElementById('adminUseAltFrame');
      var img = document.getElementById('unifiedFrameImage');
      if (!cb || !img) return;
      img.src = cb.checked ? FRAME_IMG_TWO : FRAME_IMG_ONE;
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
      // Show loading spinner while re-rendering
      floorLayoutViewer.innerHTML = '';
      var loader = document.createElement('div');
      loader.className = 'wizard-step-loading';
      loader.innerHTML = '<div class="step-spinner"></div>';
      floorLayoutViewer.appendChild(loader);
      setTimeout(function() {
        renderLayoutView();
        // Highlight moved card briefly
        var cards = floorLayoutViewer.querySelectorAll('.floor-layout-card');
        if (cards[toIdx]) {
          cards[toIdx].classList.add('just-moved');
          setTimeout(function() { cards[toIdx].classList.remove('just-moved'); }, 600);
        }
        // Also update unified preview labels + thumbnails
        renderPreviewThumbnails();
        updateFloorLabels();
      }, 120);
    }

    function renderLayoutView() {
      // Cleanup old layout viewers
      for (const v of layoutViewers) {
        if (v.renderer) v.renderer.dispose();
      }
      layoutViewers = [];
      floorLayoutViewer.innerHTML = '';

      var includedIndices = [];
      for (var i = 0; i < floors.length; i++) {
        if (!excludedFloors.has(i)) includedIndices.push(i);
      }

      // Use custom order if set, otherwise default
      if (floorOrder && floorOrder.length === includedIndices.length) {
        includedIndices = floorOrder.slice();
      } else {
        floorOrder = includedIndices.slice();
      }

      for (var idx = 0; idx < includedIndices.length; idx++) {
        var floorIdx = includedIndices[idx];
        var card = document.createElement('div');
        card.className = 'floor-layout-card';

        // Number
        var numEl = document.createElement('div');
        numEl.className = 'floor-layout-number';
        numEl.textContent = (idx + 1);

        // Canvas
        var canvasWrap = document.createElement('div');
        canvasWrap.className = 'floor-layout-canvas-wrap';

        // Up/down arrows
        var arrows = document.createElement('div');
        arrows.className = 'floor-layout-arrows';

        var btnUp = document.createElement('button');
        btnUp.type = 'button';
        btnUp.innerHTML = '&#9650;';
        btnUp.disabled = idx === 0;

        var btnDown = document.createElement('button');
        btnDown.type = 'button';
        btnDown.innerHTML = '&#9660;';
        btnDown.disabled = idx === includedIndices.length - 1;

        // Attach click handlers via closure
        (function(currentIdx) {
          btnUp.addEventListener('click', function() { moveFloorInOrder(currentIdx, currentIdx - 1); });
          btnDown.addEventListener('click', function() { moveFloorInOrder(currentIdx, currentIdx + 1); });
        })(idx);

        arrows.appendChild(btnUp);
        arrows.appendChild(btnDown);

        // Floor name label
        var nameEl = document.createElement('div');
        nameEl.className = 'floor-layout-name';
        nameEl.textContent = floors[floorIdx].name || 'Verdieping ' + (idx + 1);

        card.appendChild(numEl);
        card.appendChild(canvasWrap);
        card.appendChild(nameEl);
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

      function createWallBox(x1, y1, x2, y2, bottomZ, topZ, halfThickness, normalX, normalY) {
        if (topZ <= bottomZ) return;
        const corners = [
          { x: x1 + normalX * halfThickness, y: y1 + normalY * halfThickness },
          { x: x2 + normalX * halfThickness, y: y2 + normalY * halfThickness },
          { x: x2 - normalX * halfThickness, y: y2 - normalY * halfThickness },
          { x: x1 - normalX * halfThickness, y: y1 - normalY * halfThickness }
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

      const extendedWalls = extendWalls(walls);

      faces.push('g walls');

      for (const wall of extendedWalls) {
        const ax = (wall.a.x - centerX) * SCALE;
        const ay = (wall.a.y - centerY) * SCALE;
        const bx = (wall.b.x - centerX) * SCALE;
        const by = (wall.b.y - centerY) * SCALE;
        const thickness = wall.thickness * SCALE;
        const wdx = bx - ax, wdy = by - ay;
        const wlen = Math.hypot(wdx, wdy);
        if (wlen < 0.001) continue;
        const wnx = -wdy / wlen;
        const wny = wdx / wlen;
        const halfThick = thickness / 2;

        const openings = wall.openings;
        const wallWorldLen = Math.hypot(wall.b.x - wall.a.x, wall.b.y - wall.a.y);

        if (openings.length === 0) {
          createWallBox(ax, ay, bx, by, 0, WALL_HEIGHT, halfThick, wnx, wny);
        } else {
          const sortedOpenings = openings.map(op => {
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
              const segAx = ax + wdx * currentT, segAy = ay + wdy * currentT;
              const segBx = ax + wdx * op.startT, segBy = ay + wdy * op.startT;
              createWallBox(segAx, segAy, segBx, segBy, 0, WALL_HEIGHT, halfThick, wnx, wny);
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
            createWallBox(ax + wdx * currentT, ay + wdy * currentT, bx, by, 0, WALL_HEIGHT, halfThick, wnx, wny);
          }
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
        const CELL = 3;

        const floorSources = [];

        for (const area of design.areas ?? []) {
          const tessellated = tessellateSurfacePoly(area.poly ?? []);
          if (tessellated.length >= 3) floorSources.push(tessellated);
        }

        for (const surface of design.surfaces ?? []) {
          if (isSurfaceOutsideWalls(surface, wallBBox)) continue;
          const sName = (surface.name ?? "").trim();
          if (!sName) continue;
          const cName = (surface.customName ?? "").trim();
          if (cName && cName.toLowerCase() !== sName.toLowerCase()) continue;
          const tessellated = tessellateSurfacePoly(surface.poly ?? []);
          if (tessellated.length >= 3) floorSources.push(tessellated);
        }

        for (const wall of extendedWalls) {
          const dx = wall.b.x - wall.a.x, dy = wall.b.y - wall.a.y;
          const len = Math.hypot(dx, dy);
          if (len < 0.1) continue;
          const nx = -dy / len, ny = dx / len;
          const ht = (wall.thickness ?? 20) / 2;
          floorSources.push([
            { x: wall.a.x + nx * ht, y: wall.a.y + ny * ht },
            { x: wall.b.x + nx * ht, y: wall.b.y + ny * ht },
            { x: wall.b.x - nx * ht, y: wall.b.y - ny * ht },
            { x: wall.a.x - nx * ht, y: wall.a.y - ny * ht }
          ]);
        }

        const balStripsOBJ = mergeBalustradeStrips(design.balustrades ?? []);
        for (const strip of balStripsOBJ) {
          if (strip.length >= 3) floorSources.push(strip);
        }
        const balFillsOBJ = buildBalustradeFillPolygons(design.balustrades ?? []);
        for (const fill of balFillsOBJ) {
          if (fill.length >= 3) floorSources.push(fill);
        }

        const EXPAND = 1;
        for (let si = 0; si < floorSources.length; si++) {
          const poly = floorSources[si];
          if (poly.length < 3) continue;
          let cx = 0, cy = 0;
          for (const p of poly) { cx += p.x; cy += p.y; }
          cx /= poly.length; cy /= poly.length;
          floorSources[si] = poly.map(p => {
            const dx = p.x - cx, dy = p.y - cy;
            const dist = Math.hypot(dx, dy);
            if (dist < 0.001) return { x: p.x, y: p.y };
            return { x: p.x + (dx / dist) * EXPAND, y: p.y + (dy / dist) * EXPAND };
          });
        }

        let gMinX = Infinity, gMaxX = -Infinity, gMinY = Infinity, gMaxY = -Infinity;
        for (const src of floorSources) {
          for (const p of src) {
            gMinX = Math.min(gMinX, p.x); gMaxX = Math.max(gMaxX, p.x);
            gMinY = Math.min(gMinY, p.y); gMaxY = Math.max(gMaxY, p.y);
          }
        }
        gMinX = Math.floor(gMinX / CELL) * CELL;
        gMinY = Math.floor(gMinY / CELL) * CELL;
        gMaxX = Math.ceil(gMaxX / CELL) * CELL;
        gMaxY = Math.ceil(gMaxY / CELL) * CELL;

        for (let wx = gMinX; wx < gMaxX; wx += CELL) {
          for (let wy = gMinY; wy < gMaxY; wy += CELL) {
            const wcx = wx + CELL / 2;
            const wcy = wy + CELL / 2;

            let inVoid = false;
            for (const v of floorVoids) {
              if (pointInPolygon(wcx, wcy, v)) { inVoid = true; break; }
            }
            if (inVoid) continue;

            let inFloor = false;
            for (const src of floorSources) {
              if (pointInPolygon(wcx, wcy, src)) { inFloor = true; break; }
            }
            if (!inFloor) continue;

            const x0 = (wx - centerX) * SCALE;
            const y0 = (wy - centerY) * SCALE;
            const x1 = (wx + CELL - centerX) * SCALE;
            const y1 = (wy + CELL - centerY) * SCALE;

            const base = vertexIndex;
            var botY = (-FLOOR_THICKNESS).toFixed(4);
            var topY = (0).toFixed(4);
            // Bottom face vertices
            vertices.push(`v ${x0.toFixed(4)} ${botY} ${y0.toFixed(4)}`);
            vertices.push(`v ${x1.toFixed(4)} ${botY} ${y0.toFixed(4)}`);
            vertices.push(`v ${x1.toFixed(4)} ${botY} ${y1.toFixed(4)}`);
            vertices.push(`v ${x0.toFixed(4)} ${botY} ${y1.toFixed(4)}`);
            // Top face vertices
            vertices.push(`v ${x0.toFixed(4)} ${topY} ${y0.toFixed(4)}`);
            vertices.push(`v ${x1.toFixed(4)} ${topY} ${y0.toFixed(4)}`);
            vertices.push(`v ${x1.toFixed(4)} ${topY} ${y1.toFixed(4)}`);
            vertices.push(`v ${x0.toFixed(4)} ${topY} ${y1.toFixed(4)}`);
            vertexIndex += 8;

            // Bottom face
            addFace(base + 3, base + 2, base + 1, base + 0);
            // Top face
            addFace(base + 4, base + 5, base + 6, base + 7);
            // Side faces
            addFace(base + 0, base + 1, base + 5, base + 4);
            addFace(base + 2, base + 3, base + 7, base + 6);
            addFace(base + 3, base + 0, base + 4, base + 7);
            addFace(base + 1, base + 2, base + 6, base + 5);
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

    // Floor review include/exclude checkbox
    floorIncludeCb.addEventListener('change', () => {
      toggleFloorExclusion(currentFloorReviewIndex);
      renderFloorReview(); // refresh overlay
    });

    // Funda URL loading (refs in ensureDomRefs)

    function getFundaUrl() {
      return fundaUrlInput.value.trim();
    }

    fundaUrlInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') loadFromFunda();
    });
    btnFunda.addEventListener('click', () => loadFromFunda());

    // Function declaration so it hoists (Shopify addEventListener issue)
    function pasteTestLink() {
      var input = document.getElementById('fundaUrl');
      if (input) input.value = 'https://www.funda.nl/detail/koop/arnhem/huis-madelievenstraat-61/43269652/';
    }
    const btnTest = document.getElementById('btnTest');
    if (btnTest) {
      btnTest.addEventListener('click', () => {
        pasteTestLink();
      });
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
        icon.textContent = '✓';
      } else if (state === 'error') {
        icon.textContent = '✕';
      }
      text.innerHTML = html;
    }

    function hideFundaStatus() {
      var box = document.getElementById('fundaStatus');
      if (box) box.className = 'funda-status';
    }

    async function loadFromFunda() {
      ensureDomRefs();
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

        if (data.error) {
          setFundaStatus('error', `<strong>Geen plattegrond gevonden</strong><span>${data.error}</span>`);
          return;
        }

        if (!data?.floors?.length) {
          setFundaStatus('error', '<strong>Geen verdiepingen gevonden</strong><span>Deze woning heeft geen plattegrond.</span>');
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

        const addrStr = addr ? `${addr.street}, ${addr.city}` : 'Adres niet gevonden';
        setFundaStatus('success', `<strong>${data.floors.length} interactieve plattegrond${data.floors.length === 1 ? '' : 'en'} gevonden</strong><span>${addrStr}</span>`);

        processFloors(data);
      } catch (err) {
        if (err.message && (err.message.includes('Load failed') || err.message.includes('Failed to fetch'))) {
          setFundaStatus('error', '<strong>Verbinding mislukt</strong><span>Draait de Flask server?</span>');
        } else {
          setFundaStatus('error', `<strong>Fout</strong><span>${err.message}</span>`);
        }
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
      var orderBtn = document.getElementById('btnOrder');
      if (orderBtn) { orderBtn.disabled = true; orderBtn.textContent = 'Toevoegen…'; }
      fetch('/cart/add.js', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items: [{ id: parseInt(variantId), quantity: 1 }] })
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
        if (orderBtn) { orderBtn.disabled = false; orderBtn.textContent = 'Afronden & bestellen \u2726'; }
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
