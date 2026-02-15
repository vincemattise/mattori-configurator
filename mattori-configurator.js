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
        // Pan in camera-local XY plane
        const offset = new THREE.Vector3().subVectors(this.camera.position, this.target);
        let targetDist = offset.length();
        // Half the fov in radians
        targetDist *= Math.tan((this.camera.fov / 2) * Math.PI / 180);
        // Pan left/right
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

      // Apply rotation delta (always full amount)
      this._spherical.theta += this._sphericalDelta.theta;
      this._spherical.phi += this._sphericalDelta.phi;

      // Clamp polar angle
      this._spherical.phi = Math.max(this.minPolarAngle, Math.min(this.maxPolarAngle, this._spherical.phi));
      this._spherical.makeSafe();

      // Apply zoom — handle orthographic camera separately
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

      // Apply pan
      this.target.add(this._panOffset);

      // Convert back to position
      offset.setFromSpherical(this._spherical);
      this.camera.position.copy(this.target).add(offset);
      this.camera.lookAt(this.target);

      // Decay deltas
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
    let originalFmlData = null; // Store original FML data for download
    let originalFileName = '';  // Store uploaded filename

    const CANVAS_W = 260;
    const CANVAS_H = 400;

    // ============================================================
    // DOM REFERENCES
    // ============================================================
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const errorMsg = document.getElementById('errorMsg');
    const loadingOverlay = document.getElementById('loadingOverlay');
    const mainWrapper = document.getElementById('mainWrapper');
    const orderFlow = document.getElementById('orderFlow');
    const floorsGrid = document.getElementById('floorsGrid');
    const floorCheckboxes = document.getElementById('floorCheckboxes');
    const btnExport = document.getElementById('btnExport');
    const btnDownloadFml = document.getElementById('btnDownloadFml');
    const toast = document.getElementById('toast');
    const fileLabel = document.getElementById('fileLabel');
    const addressStreet = document.getElementById('addressStreet');
    const addressCity = document.getElementById('addressCity');
    const addressEditToggle = document.getElementById('addressEditToggle');
    const addressFields = document.getElementById('addressFields');
    const framePreview = document.getElementById('framePreview');
    const frameStreet = document.getElementById('frameStreet');
    const frameCity = document.getElementById('frameCity');
    const stepViewers = document.getElementById('stepViewers');
    const stepLabelsPreview = document.getElementById('stepLabelsPreview');
    const labelsPreview = document.getElementById('labelsPreview');
    const labelsOverlay = document.getElementById('labelsOverlay');
    const labelsEditToggle = document.getElementById('labelsEditToggle');
    const labelsFields = document.getElementById('labelsFields');
    const stepRemarks = document.getElementById('stepRemarks');
    const stepOrder = document.getElementById('stepOrder');
    const stepDisclaimer = document.getElementById('stepDisclaimer');

    // ============================================================
    // UTILITY FUNCTIONS
    // ============================================================
    function clearError() { errorMsg.textContent = ''; }
    function setError(msg) { errorMsg.textContent = msg; }
    function showLoading() { loadingOverlay.classList.add('active'); }
    function hideLoading() { loadingOverlay.classList.remove('active'); }

    let toastTimer = null;
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
        // segments: ["detail", "koop", "arnhem", "huis-madelievenstraat-61", "43269652", ...]
        if (segments.length < 4) return null;

        const city = segments[2]; // "arnhem"
        const slug = segments[3]; // "huis-madelievenstraat-61"

        // Strip property type prefix
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

        // Convert kebab-case to Title Case
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
    function computeBoundingBox(design) {
      const points = [];
      for (const wall of design.walls ?? []) {
        points.push({ x: wall.a.x, y: wall.a.y }, { x: wall.b.x, y: wall.b.y });
        // For arc walls, sample the curve to get accurate bounding box
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
    // WALL ENDPOINT EXTENSION (shared by 2D rendering + OBJ export)
    // ============================================================
    // Extends wall endpoints so they meet at outer edges instead of
    // stopping at the centerline of crossing walls. Fixes gaps at
    // L-corners, T-junctions, and angled joints.
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
        // Skip extension for arc sub-segments (they connect to each other, not wall junctions)
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
          if (sinAngle < 0.1) continue; // walls are parallel, skip

          const ext = Math.min(otherHalfThick / sinAngle, otherHalfThick * 3);

          // Check point A
          const aShares = pointsNear(origAx, origAy, other.a.x, other.a.y) ||
                          pointsNear(origAx, origAy, other.b.x, other.b.y);
          const aOnInt = pointOnSegmentInterior(origAx, origAy, other.a.x, other.a.y, other.b.x, other.b.y);
          if (aShares || aOnInt) extendA = Math.max(extendA, ext);

          // Check point B
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
    // Extend every balustrade at both ends by 0.75× its thickness.
    // The FML data shows balustrade endpoints are offset ~7 units
    // from their connecting wall/balustrade (≈ thickness × √2 / 2).
    // Using 0.75× thickness (= 7.5 for thick=10) covers this gap.
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
    // Groups adjacent balustrade segments into ordered chains.
    // Each chain is a sequence of {bal, flipped} entries where the
    // a→b direction is consistent along the chain.
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

        // Extend from B end
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

        // Extend from A end (prepend)
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

      // console.log(`buildBalustradeChains: ${balustrades.length} balustrades → ${chains.length} chains`);
      return chains;
    }

    // Extract the OUTER EDGE of a balustrade chain as an array of points.
    // For a chain of N segments, the outer edge has N+1 points.
    // Returns { outerEdge, innerEdge } where each is an array of {x,y}.
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

    // Build STRIP polygons (thin footprint of the balustrade itself)
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

    // Build FILL polygons — the area enclosed by the OUTER edge of a
    // balustrade chain, closed by a straight line from end back to start.
    // This creates a D-shaped polygon that covers the balcony floor area
    // beyond the Balkon surface polygon (which only covers the inner part).
    // Together with the Balkon surface, this fills the entire balcony.
    // Only for chains with 3+ segments (curved chains).
    function buildBalustradeFillPolygons(balustrades) {
      const chains = buildBalustradeChains(balustrades);
      const fills = [];

      for (const chain of chains) {
        if (chain.length < 3) continue; // Only for curved chains

        const { leftEdge, rightEdge } = getChainEdges(chain);

        // Determine which edge is "outer" (further from center of house)
        // by comparing bounding box areas — the outer edge sweeps a larger area.
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

        // D-shape fill: outer edge forward, then straight close back to start.
        // This covers the area from the outer railing to the chord line.
        if (outerEdge.length >= 3) {
          fills.push([...outerEdge]);
        }

        // Also add inner edge as D-shape — fills the gap between the
        // Balkon surface polygon and the balustrade railing.
        if (innerEdge.length >= 3) {
          fills.push([...innerEdge]);
        }
      }

      // console.log(`buildBalustradeFillPolygons: ${fills.length} fill polygons`);
      return fills;
    }

    // ============================================================
    // ARC WALL TESSELLATION (quadratic Bézier → straight segments)
    // ============================================================
    // FML walls with a `c` property are curved: a quadratic Bézier
    // from point `a` through control point `c` to point `b`.
    // This function splits such walls into N straight sub-segments.
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
        // Quadratic Bézier: P(t) = (1-t)²·A + 2(1-t)t·C + t²·B
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

    // Flatten all walls: arc walls → many straight segments, straight walls pass through.
    // Openings on arc walls are ignored (they're very rare on curved walls).
    function flattenWalls(walls) {
      const ARC_SEGMENTS = 16;
      const result = [];
      let arcCount = 0;
      for (const wall of walls) {
        if (wall.c && wall.c.x != null && wall.c.y != null) {
          result.push(...tessellateArcWall(wall, ARC_SEGMENTS));
          arcCount++;
        } else {
          // Carry wall height through for straight walls
          result.push({
            ...wall,
            _heightA: wall.az?.h ?? wall.bz?.h ?? 265,
            _heightB: wall.bz?.h ?? wall.az?.h ?? 265
          });
        }
      }
      // arcCount tracked but not logged (clean output)
      return result;
    }

    // Tessellate a single arc balustrade (same Bézier logic but keeps thickness/height)
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

    // Flatten balustrades: arc → segments, straight → pass through
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
    // FML surface polygons can have cx/cy/cz control points on vertices.
    // When vertex[i] has cx/cy, the edge FROM vertex[i-1] TO vertex[i]
    // is a quadratic Bézier curve using (cx,cy) as the control point.
    // This function tessellates those curved edges into straight segments.
    function tessellateSurfacePoly(poly) {
      if (!poly || poly.length < 3) return poly;
      const ARC_SEGMENTS = 24;
      const result = [];
      for (let i = 0; i < poly.length; i++) {
        const curr = poly[i];
        const prev = poly[(i - 1 + poly.length) % poly.length];
        if (curr.cx != null && curr.cy != null) {
          // Curved edge from prev to curr via control point (cx, cy)
          const ax = prev.x, ay = prev.y;
          const bx = curr.x, by = curr.y;
          const cx = curr.cx, cy = curr.cy;
          // Don't add prev (it was already added in previous iteration)
          // Add intermediate points + endpoint
          for (let s = 1; s <= ARC_SEGMENTS; s++) {
            const t = s / ARC_SEGMENTS;
            const px = (1-t)*(1-t)*ax + 2*(1-t)*t*cx + t*t*bx;
            const py = (1-t)*(1-t)*ay + 2*(1-t)*t*cy + t*t*by;
            result.push({ x: px, y: py, z: curr.z ?? 0 });
          }
        } else {
          // Straight edge — just add this vertex
          result.push({ x: curr.x, y: curr.y, z: curr.z ?? 0 });
        }
      }
      return result;
    }

    // ============================================================
    // STAIR VOID DETECTION
    // ============================================================
    // Detects floor openings (trapgaten) using two methods:
    // 1. Surfaces with role=14 (explicit void markers)
    // 2. Cross-floor item matching: items with same refid at same
    //    position on adjacent floors are stairs → void on upper floor
    function detectStairVoids(allFloorDesigns) {
      const POSITION_TOL = 5; // tolerance for matching positions
      const voidsByFloor = allFloorDesigns.map(() => []);

      for (let fi = 0; fi < allFloorDesigns.length; fi++) {
        const design = allFloorDesigns[fi];

        // Method 1: role=14 surfaces
        for (const surface of design.surfaces ?? []) {
          if ((surface.role ?? -1) !== 14) continue;
          const poly = tessellateSurfacePoly(surface.poly ?? []);
          if (poly.length >= 3) {
            voidsByFloor[fi].push(poly.map(p => ({ x: p.x, y: p.y })));
          }
        }

        // Method 2: cross-floor stair items (void appears on the UPPER floor)
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

              // Found a stair: create void rectangle from item bounding box
              // Shrink by VOID_MARGIN to avoid clipping adjacent walls/doors
              const VOID_MARGIN = 10;
              const w = Math.max(0, (curr.width ?? 0) - VOID_MARGIN * 2);
              const h = Math.max(0, (curr.height ?? 0) - VOID_MARGIN * 2);
              if (w < 30 || h < 30) continue; // skip tiny items
              const cx = curr.x ?? 0;
              const cy = curr.y ?? 0;
              const rot = (curr.rotation ?? 0) * Math.PI / 180;
              const cosR = Math.cos(rot);
              const sinR = Math.sin(rot);
              const hw = w / 2, hh = h / 2;

              // Rotated rectangle corners
              const corners = [
                { x: cx + cosR * (-hw) - sinR * (-hh), y: cy + sinR * (-hw) + cosR * (-hh) },
                { x: cx + cosR * ( hw) - sinR * (-hh), y: cy + sinR * ( hw) + cosR * (-hh) },
                { x: cx + cosR * ( hw) - sinR * ( hh), y: cy + sinR * ( hw) + cosR * ( hh) },
                { x: cx + cosR * (-hw) - sinR * ( hh), y: cy + sinR * (-hw) + cosR * ( hh) }
              ];
              voidsByFloor[fi].push(corners);
              break; // one match per item is enough
            }
          }
        }
      }
      return voidsByFloor;
    }

    // Check if a point is inside a polygon (ray casting)
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

    // Check if a polygon's centroid falls inside any void
    function polygonOverlapsVoid(poly, voids) {
      if (!voids || voids.length === 0) return false;
      // Use centroid for a quick overlap check
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

      // Detect stair voids across all floors
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
        // Flatten arc balustrades into straight segments
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

      renderUI();
      setTimeout(() => {
        renderFramePreview();
        renderAll3DViewers();
        renderLabelsPreview();
      }, 50);
    }

    // ============================================================
    // 3D VIEWER (Three.js)
    // ============================================================

    // Parse OBJ string into Three.js BufferGeometry
    /**
     * Parse OBJ string into grouped geometries based on 'g' directives.
     * Returns { walls: BufferGeometry, floor: BufferGeometry, all: BufferGeometry }
     */
    function parseOBJToGroups(objString) {
      const allVertices = [];   // flat array [x,y,z, x,y,z, ...]
      const groups = {};        // groupName → [faceIndex, ...]
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

      // Combine wall groups (use concat to avoid call stack overflow with large arrays)
      const wallFaces = (groups['walls'] || [])
        .concat(groups['walls_balustrades'] || [])
        .concat(groups['default'] || []);
      // All faces combined
      const allFaces = Object.values(groups).reduce((acc, g) => acc.concat(g), []);

      return {
        walls: makeGeometry(wallFaces),
        floor: makeGeometry(groups['floor'] || []),
        all: makeGeometry(allFaces)
      };
    }

    // Active viewers for cleanup
    let activeViewers = [];  // { renderer, controls, animId }

    function renderSingle3DViewer(index) {
      const floor = floors[index];
      const container = viewerContainers[index];
      if (!container || !floor) return;

      // Get actual container dimensions
      const rect = container.getBoundingClientRect();
      const width = Math.round(rect.width) || 400;
      const height = Math.round(rect.height) || 300;

      // Generate OBJ and parse to grouped geometries
      const objString = generateFloorOBJ(floor);
      const groups = parseOBJToGroups(objString);

      // Compute bounding box from combined geometry for camera positioning
      const allGeo = groups.all;
      allGeo.computeBoundingBox();
      const box = allGeo.boundingBox;
      const center = new THREE.Vector3();
      box.getCenter(center);
      const size = new THREE.Vector3();
      box.getSize(size);

      // Center all geometry at origin so it renders centered in the viewer
      const offsetX = -center.x;
      const offsetY = -center.y;
      const offsetZ = -center.z;
      if (groups.walls) groups.walls.translate(offsetX, offsetY, offsetZ);
      if (groups.floor) groups.floor.translate(offsetX, offsetY, offsetZ);
      // Update center to origin after translation
      center.set(0, 0, 0);

      // Scene setup
      const scene = new THREE.Scene();
      scene.background = new THREE.Color(0xffffff);

      // Materials — walls slightly darker than floor for top-down contrast
      const wallMaterial = new THREE.MeshPhongMaterial({
        color: 0xA89478,  // Darker warm beige for walls
        flatShading: true,
        side: THREE.DoubleSide,
        shininess: 10
      });
      const floorMaterial = new THREE.MeshPhongMaterial({
        color: 0xC2AD91,  // Lighter warm beige for floor
        flatShading: true,
        side: THREE.DoubleSide,
        shininess: 5
      });

      // Add wall mesh
      if (groups.walls) {
        scene.add(new THREE.Mesh(groups.walls, wallMaterial));
      }
      // Add floor mesh
      if (groups.floor) {
        scene.add(new THREE.Mesh(groups.floor, floorMaterial));
      }

      // Lighting — optimized for near-top-down perspective view
      // Ambient provides base visibility; directional lights create
      // subtle shading on wall inner sides visible through perspective
      scene.add(new THREE.AmbientLight(0xffffff, 0.6));

      // Main light from front-top to illuminate wall inner faces
      const dirLight = new THREE.DirectionalLight(0xffffff, 0.5);
      dirLight.position.set(0, 8, 5);
      scene.add(dirLight);

      // Fill from side to add depth to wall corners
      const fillLight = new THREE.DirectionalLight(0xffffff, 0.2);
      fillLight.position.set(-4, 6, -1);
      scene.add(fillLight);

      // Camera — near-orthographic perspective (narrow FOV, high up)
      // Uses a tight FOV so it looks almost top-down but with enough
      // perspective to reveal the inner sides of walls — like looking
      // at a real frame up close with your face near the glass.
      const maxDim = Math.max(size.x, size.y, size.z);
      const padding = 1.25;
      const halfW = (size.x * padding) / 2;
      const halfZ = (size.z * padding) / 2;
      const halfExtent = Math.max(halfW, halfZ);

      const FOV = 12; // very narrow = near-orthographic, but perspective reveals wall depth
      const aspect = width / height;
      const camera = new THREE.PerspectiveCamera(FOV, aspect, 0.01, halfExtent * 100);

      // Position high above, with a tiny Z-offset for subtle depth cue
      const camDist = halfExtent / Math.tan((FOV / 2) * Math.PI / 180);
      camera.position.set(0, camDist, camDist * 0.14);
      camera.lookAt(center);

      // Renderer
      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
      renderer.setClearColor(0x000000, 0);
      renderer.setSize(width, height);
      renderer.setPixelRatio(dpr);
      container.appendChild(renderer.domElement);

      // OrbitControls — rotate, zoom, pan
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

      // Prevent card click (OBJ download) when interacting with 3D viewer
      renderer.domElement.addEventListener('mousedown', e => e.stopPropagation());
      renderer.domElement.addEventListener('click', e => e.stopPropagation());
      renderer.domElement.addEventListener('touchstart', e => e.stopPropagation());

      // Animation loop for OrbitControls
      let animId;
      function animate() {
        animId = requestAnimationFrame(animate);
        controls.update();
        renderer.render(scene, camera);
      }
      animate();

      activeViewers.push({ renderer, controls, animId });
    }

    function renderAll3DViewers() {
      for (let i = 0; i < floors.length; i++) {
        renderSingle3DViewer(i);
      }
    }

    // ============================================================
    // FRAME PREVIEW — Text only (address on frame)
    // ============================================================
    function renderFramePreview() {
      if (!floors.length) return;
      framePreview.classList.add('active');
      updateFrameAddress();
    }

    function updateFrameAddress() {
      frameStreet.textContent = addressStreet.value || '';
      frameCity.textContent = addressCity.value || '';
    }

    // ============================================================
    // LABELS PREVIEW (floor labels on bottom frame)
    // ============================================================
    let floorLabels = []; // current labels array: [{index, label}]

    function renderLabelsPreview() {
      if (!floors.length) return;
      stepLabelsPreview.style.display = '';
      labelsPreview.classList.add('active');
      updateFloorLabels();
    }

    // Translate a Dutch floor name to an English label (US numbering).
    // US convention: begane grond = 1st floor, eerste verdieping = 2nd floor, etc.
    // Special areas use lowercase: storage, basement, attic, etc.
    // If singleFloor is true, returns "floor plan" for any regular floor.
    function translateFloorName(name, singleFloor) {
      const lower = name.toLowerCase().trim();
      // Special areas — lowercase for elegant typography
      if (/^kelder/.test(lower)) return 'basement';
      if (/^zolder/.test(lower)) return 'attic';
      if (/^berging/.test(lower)) return 'storage';
      if (/^garage/.test(lower)) return 'garage';
      if (/^dak/.test(lower)) return 'roof';
      if (/^tuin/.test(lower)) return 'garden';
      // Single floor → "floor plan"
      if (singleFloor) return 'floor plan';
      // US numbering: begane grond = 1st floor
      if (/^begane\s*grond/.test(lower)) return '1st floor';
      // Ordinal verdieping — shifted +1 for US convention
      const ordMap = [
        [/^eerste\b/, '2nd'], [/^tweede\b/, '3rd'], [/^derde\b/, '4th'],
        [/^vierde\b/, '5th'], [/^vijfde\b/, '6th'], [/^zesde\b/, '7th'],
        [/^zevende\b/, '8th'], [/^achtste\b/, '9th'], [/^negende\b/, '10th'],
        [/^tiende\b/, '11th'],
      ];
      for (const [regex, ord] of ordMap) {
        if (regex.test(lower)) return ord + ' floor';
      }
      // Numeric patterns: "1e verdieping" → +1 for US
      const numMatch = lower.match(/^(\d+)e?\s+verdieping/);
      if (numMatch) {
        const n = parseInt(numMatch[1]) + 1;
        const suf = n === 1 ? 'st' : n === 2 ? 'nd' : n === 3 ? 'rd' : 'th';
        return n + suf + ' floor';
      }
      // No match — lowercase the original
      return name.charAt(0).toLowerCase() + name.slice(1).toLowerCase();
    }

    function getIncludedFloorLabels() {
      // Build label list with English translated floor names
      const labels = [];
      const includedIndices = [];
      for (let i = 0; i < floors.length; i++) {
        if (excludedFloors.has(i)) continue;
        includedIndices.push(i);
      }
      // Count only regular floors (not special areas) to determine if single floor
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

      // Update overlay — clean labels only, no badges (those are only on the floor cards)
      labelsOverlay.innerHTML = '';
      for (let li = 0; li < floorLabels.length; li++) {
        const item = floorLabels[li];
        const el = document.createElement('div');
        el.className = 'label-item';
        el.textContent = item.label;
        labelsOverlay.appendChild(el);
      }

      // Update edit fields
      updateLabelsFields();
    }

    function updateLabelsFields() {
      labelsFields.innerHTML = '';
      for (let li = 0; li < floorLabels.length; li++) {
        const item = floorLabels[li];
        const row = document.createElement('div');
        row.className = 'label-field-row';

        const span = document.createElement('span');
        span.textContent = floors[item.index].name;

        const inp = document.createElement('input');
        inp.type = 'text';
        inp.value = item.label;
        inp.placeholder = item.label;
        inp.addEventListener('input', () => {
          floorLabels[li].label = inp.value;
          updateLabelsOverlayOnly();
        });

        row.appendChild(span);
        row.appendChild(inp);
        labelsFields.appendChild(row);
      }
    }

    function updateLabelsOverlayOnly() {
      const items = labelsOverlay.querySelectorAll('.label-item');
      items.forEach((el, i) => {
        if (floorLabels[i]) {
          el.textContent = floorLabels[i].label;
        }
      });
      // Sync sublabels on floor cards
      syncCardSublabels();
    }

    function syncCardSublabels() {
      // Sync edited labels back to floor card names in preview 2
      for (const item of floorLabels) {
        const nameEl = floorsGrid.querySelector(`.floor-name[data-floor-index="${item.index}"]`);
        if (nameEl) nameEl.textContent = item.label;
      }
    }

    // ============================================================
    // FLOOR EXCLUSION (situatietekening / tuin detection)
    // ============================================================
    let excludedFloors = new Set();  // indices of excluded floors

    function isLikelySituatie(floor) {
      const name = (floor.name || '').toLowerCase();
      const excludeKeywords = ['situatie', 'site', 'tuin', 'garden', 'buitenruimte', 'omgeving', 'terrein', 'perceel'];
      return excludeKeywords.some(kw => name.includes(kw));
    }

    // Check if a floor is an "extra" floor (berging, garage, schuur, zolder)
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
      updateFloorCardStates();
      updateFloorLabels(); // Update labels preview when floors change
    }

    function updateFloorCardStates() {
      // Update 3D viewer cards
      const cards = floorsGrid.querySelectorAll('.floor-card');
      cards.forEach((card, i) => {
        if (excludedFloors.has(i)) {
          card.classList.add('excluded');
        } else {
          card.classList.remove('excluded');
        }
      });

      // Update checkbox list
      const checkItems = floorCheckboxes.querySelectorAll('.floor-check-item');
      checkItems.forEach((item) => {
        const idx = parseInt(item.dataset.index);
        const cb = item.querySelector('input[type="checkbox"]');
        if (excludedFloors.has(idx)) {
          item.classList.add('excluded');
          if (cb) cb.checked = false;
        } else {
          item.classList.remove('excluded');
          if (cb) cb.checked = true;
        }
      });
    }

    // ============================================================
    // UI RENDERING (Three.js)
    // ============================================================
    let viewerContainers = [];

    function renderUI() {
      // Show admin export buttons
      document.getElementById('uploadActions').classList.add('active');

      // Hide test button after successful load
      const btnTest = document.getElementById('btnTest');
      if (btnTest) btnTest.style.display = 'none';

      // Show order flow steps with staggered fade-in
      const stepsToShow = [stepViewers, stepLabelsPreview, stepRemarks, stepOrder, stepDisclaimer];
      stepsToShow.forEach((step, i) => {
        step.style.display = '';
        step.classList.remove('animate-in');
        void step.offsetWidth; // force reflow
        step.style.animationDelay = `${i * 0.1}s`;
        step.classList.add('animate-in');
      });

      // Enable order button
      document.getElementById('btnOrder').disabled = false;

      // Cleanup old viewers (stop animation loops, dispose renderers)
      for (const v of activeViewers) {
        if (v.animId) cancelAnimationFrame(v.animId);
        if (v.controls) v.controls.dispose();
        if (v.renderer) v.renderer.dispose();
      }
      activeViewers = [];

      // Auto-detect excluded floors (situatietekening, tuin, etc.)
      excludedFloors = new Set();
      for (let i = 0; i < floors.length; i++) {
        if (isLikelySituatie(floors[i])) {
          excludedFloors.add(i);
        }
      }

      // Build floor viewer cards (3D viewers only, no checkboxes here)
      floorsGrid.innerHTML = '';
      viewerContainers = [];

      // Track included floor numbering for badge coupling
      let includedCount = 0;
      const floorBadgeMap = new Map(); // floorIndex → badge number
      for (let i = 0; i < floors.length; i++) {
        if (!excludedFloors.has(i)) {
          includedCount++;
          floorBadgeMap.set(i, includedCount);
        }
      }

      // Count regular floors (not special areas) to determine single-floor mode
      const regularFloorCount = Array.from(floorBadgeMap.keys()).filter(i => {
        const n = (floors[i].name || '').toLowerCase().trim();
        return !(/^(kelder|zolder|berging|garage|dak|tuin)/.test(n));
      }).length;
      const isSingleFloor = regularFloorCount <= 1;

      for (let i = 0; i < floors.length; i++) {
        const floor = floors[i];
        const card = document.createElement('div');
        card.className = 'floor-card';
        if (excludedFloors.has(i)) card.classList.add('excluded');

        // Add badge number for non-excluded floors
        const badgeNum = floorBadgeMap.get(i);
        if (badgeNum) {
          const badge = document.createElement('div');
          badge.className = 'floor-badge';
          badge.textContent = badgeNum;
          card.appendChild(badge);
        }

        // Floor name label — translated to English (US numbering) to match labels preview
        const nameEl = document.createElement('div');
        nameEl.className = 'floor-name';
        nameEl.dataset.floorIndex = i;
        nameEl.textContent = translateFloorName(floor.name, isSingleFloor);

        const viewerWrap = document.createElement('div');
        viewerWrap.className = 'floor-canvas-wrap';
        viewerContainers.push(viewerWrap);

        card.appendChild(nameEl);
        card.appendChild(viewerWrap);
        floorsGrid.appendChild(card);
      }

      // Build checkbox list for extra floors (berging, garage, etc.)
      floorCheckboxes.innerHTML = '';
      const hasExtraFloors = floors.some((f, i) => isExtraFloor(f) && !isLikelySituatie(f));

      for (let i = 0; i < floors.length; i++) {
        const floor = floors[i];
        // Only show checkboxes for extra floors or auto-excluded floors
        if (!isExtraFloor(floor) && !isLikelySituatie(floor)) continue;

        const item = document.createElement('div');
        item.className = 'floor-check-item';
        item.dataset.index = i;
        if (excludedFloors.has(i)) item.classList.add('excluded');

        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.id = `floor-cb-${i}`;
        cb.checked = !excludedFloors.has(i);
        cb.addEventListener('change', () => toggleFloorExclusion(i));

        const lbl = document.createElement('label');
        lbl.htmlFor = `floor-cb-${i}`;
        lbl.textContent = `${floor.name} meenemen`;

        item.appendChild(cb);
        item.appendChild(lbl);
        floorCheckboxes.appendChild(item);
      }
    }

    // Download a single floor as OBJ file
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
        // Try to parse address from FML labels
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
    // OBJ EXPORT — generates a SEPARATE OBJ per floor
    // ============================================================

    // Helper: generate OBJ content for a single floor (centered at origin)
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
      const walls = flattenWalls(design.walls ?? []);
      const bbox = floor.bbox;
      const centerX = (bbox.minX + bbox.maxX) / 2;
      const centerY = (bbox.minY + bbox.maxY) / 2;

      // Use shared wall extension function
      const extendedWalls = extendWalls(walls);

      // Group marker for walls (used by parseOBJToGroups for separate materials)
      faces.push('g walls');

      // WALLS (using extended endpoints)
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

      // Group marker for floor slab (used by parseOBJToGroups for separate materials)
      faces.push('g floor');

      // FLOOR SLAB — extrude each area/surface polygon + wall footprints
      // This respects the actual shape of the house (no convex hull overshoot).
      const FLOOR_THICKNESS = 0.30; // 30cm thick floor slab

      // Extrude a polygon into a solid slab (top at y=0, bottom at y=-FLOOR_THICKNESS).
      // Uses centroid-fan triangulation: adds a center vertex and fans triangles
      // from center to each edge. Works for any simple polygon (convex or concave).
      function extrudePolygon(poly) {
        if (poly.length < 3) return;
        const n = poly.length;

        // Compute centroid
        let cx = 0, cy = 0;
        for (const pt of poly) { cx += pt.x; cy += pt.y; }
        cx /= n; cy /= n;

        // Bottom ring + center = n+1 vertices, then top ring + center = n+1
        const baseBot = vertexIndex;
        // Bottom ring vertices
        for (const pt of poly) {
          vertices.push(`v ${pt.x.toFixed(4)} ${(-FLOOR_THICKNESS).toFixed(4)} ${pt.y.toFixed(4)}`);
          vertexIndex++;
        }
        // Bottom center vertex
        vertices.push(`v ${cx.toFixed(4)} ${(-FLOOR_THICKNESS).toFixed(4)} ${cy.toFixed(4)}`);
        const botCenter = vertexIndex;
        vertexIndex++;

        const baseTop = vertexIndex;
        // Top ring vertices
        for (const pt of poly) {
          vertices.push(`v ${pt.x.toFixed(4)} ${(0).toFixed(4)} ${pt.y.toFixed(4)}`);
          vertexIndex++;
        }
        // Top center vertex
        vertices.push(`v ${cx.toFixed(4)} ${(0).toFixed(4)} ${cy.toFixed(4)}`);
        const topCenter = vertexIndex;
        vertexIndex++;

        // Bottom face triangles (fan from center, reversed winding)
        for (let i = 0; i < n; i++) {
          const j = (i + 1) % n;
          addTriFace(botCenter, baseBot + j, baseBot + i);
        }
        // Top face triangles (fan from center)
        for (let i = 0; i < n; i++) {
          const j = (i + 1) % n;
          addTriFace(topCenter, baseTop + i, baseTop + j);
        }
        // Side faces
        for (let i = 0; i < n; i++) {
          const j = (i + 1) % n;
          addFace(baseBot + i, baseBot + j, baseTop + j, baseTop + i);
        }
      }

      // Collect void polygons in world coordinates
      const floorVoids = floor.voids ?? [];

      // ---- FLOOR SLAB: grid rasterization approach ----
      // Instead of extruding each area/surface/wall separately,
      // rasterize the ENTIRE floor as one grid. A cell gets floor if
      // it's inside ANY area, surface, or wall footprint, and NOT in a void.
      {
        const CELL = 3; // ~3cm grid cells (smooth curves)

        // Collect all "floor source" polygons in world coords
        const floorSources = [];

        // Areas
        for (const area of design.areas ?? []) {
          const tessellated = tessellateSurfacePoly(area.poly ?? []);
          if (tessellated.length >= 3) floorSources.push(tessellated);
        }

        // Named surfaces (skip sub-zones)
        for (const surface of design.surfaces ?? []) {
          const sName = (surface.name ?? "").trim();
          if (!sName) continue;
          const cName = (surface.customName ?? "").trim();
          if (cName && cName.toLowerCase() !== sName.toLowerCase()) continue;
          // Balkon surfaces ARE included — they cover the inner area.
          // The balustrade fill polygons cover the outer area beyond the surface.
          const tessellated = tessellateSurfacePoly(surface.poly ?? []);
          if (tessellated.length >= 3) floorSources.push(tessellated);
        }

        // Wall footprints (in world coords, using extended walls)
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

        // Balustrade-derived floor polygons:
        // 1. Strip polygons (thin footprint)
        const balStripsOBJ = mergeBalustradeStrips(design.balustrades ?? []);
        for (const strip of balStripsOBJ) {
          if (strip.length >= 3) floorSources.push(strip);
        }
        // 2. Fill polygons (entire balcony area enclosed by curved chains)
        const balFillsOBJ = buildBalustradeFillPolygons(design.balustrades ?? []);
        for (const fill of balFillsOBJ) {
          if (fill.length >= 3) floorSources.push(fill);
        }

        // Expand each floor source polygon outward by a tiny amount (1 unit = 1cm).
        // This eliminates seam gaps where adjacent polygons share an exact boundary
        // (e.g., area polygon edge at x=7.5 meets wall footprint edge at x=7.5).
        // The ray-casting pointInPolygon test is unreliable for points exactly ON
        // a polygon boundary, so a tiny overlap ensures full coverage.
        const EXPAND = 1; // 1 unit = 1cm in world coords
        for (let si = 0; si < floorSources.length; si++) {
          const poly = floorSources[si];
          if (poly.length < 3) continue;
          // Compute centroid
          let cx = 0, cy = 0;
          for (const p of poly) { cx += p.x; cy += p.y; }
          cx /= poly.length; cy /= poly.length;
          // Push each vertex outward from centroid by EXPAND
          floorSources[si] = poly.map(p => {
            const dx = p.x - cx, dy = p.y - cy;
            const dist = Math.hypot(dx, dy);
            if (dist < 0.001) return { x: p.x, y: p.y };
            return { x: p.x + (dx / dist) * EXPAND, y: p.y + (dy / dist) * EXPAND };
          });
        }

        // Compute overall bounding box of all sources
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

        // Rasterize: for each cell, check if center is inside any source polygon
        for (let wx = gMinX; wx < gMaxX; wx += CELL) {
          for (let wy = gMinY; wy < gMaxY; wy += CELL) {
            const wcx = wx + CELL / 2;
            const wcy = wy + CELL / 2;

            // Check if inside any void → skip
            let inVoid = false;
            for (const v of floorVoids) {
              if (pointInPolygon(wcx, wcy, v)) { inVoid = true; break; }
            }
            if (inVoid) continue;

            // Check if inside any floor source polygon
            let inFloor = false;
            for (const src of floorSources) {
              if (pointInPolygon(wcx, wcy, src)) { inFloor = true; break; }
            }
            if (!inFloor) continue;

            // Emit only the top face of this cell (no sides/bottom).
            // This prevents visible seams between adjacent cells in perspective view.
            const x0 = (wx - centerX) * SCALE;
            const y0 = (wy - centerY) * SCALE;
            const x1 = (wx + CELL - centerX) * SCALE;
            const y1 = (wy + CELL - centerY) * SCALE;

            const base = vertexIndex;
            vertices.push(`v ${x0.toFixed(4)} ${(0).toFixed(4)} ${y0.toFixed(4)}`);
            vertices.push(`v ${x1.toFixed(4)} ${(0).toFixed(4)} ${y0.toFixed(4)}`);
            vertices.push(`v ${x1.toFixed(4)} ${(0).toFixed(4)} ${y1.toFixed(4)}`);
            vertices.push(`v ${x0.toFixed(4)} ${(0).toFixed(4)} ${y1.toFixed(4)}`);
            vertexIndex += 4;

            addFace(base + 0, base + 1, base + 2, base + 3); // top face only
          }
        }
      }

      // Group marker for balustrades (same material as walls)
      faces.push('g walls_balustrades');

      // BALUSTRADES — solid low wall, extended at both ends
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

        // One solid box from floor to balustrade height
        createWallBox(bax, bay, bbx, bby, 0, bheight, bhalfThick, bnx, bny);
        // (balustrade floor footprint is included in the grid rasterization above)
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

    // Sanitize floor name for use as filename
    function sanitizeFilename(name) {
      return name.replace(/[^a-zA-Z0-9\-_ ]/g, '').replace(/\s+/g, '_').toLowerCase() || 'verdieping';
    }

    // Main export function — separate OBJ per floor
    async function exportOBJ() {
      if (floors.length === 0) {
        setError('Geen plattegronden beschikbaar om te exporteren.');
        return;
      }

      // Generate OBJ content per floor (always export all)
      const objFiles = floors.map(floor => ({
        name: sanitizeFilename(floor.name),
        content: generateFloorOBJ(floor)
      }));

      // If only 1 floor selected: download as single OBJ
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

      // Multiple floors: bundle as ZIP
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

    // Address edit toggle — show/hide address fields
    addressEditToggle.addEventListener('change', () => {
      if (addressEditToggle.checked) {
        addressFields.classList.add('visible');
      } else {
        addressFields.classList.remove('visible');
      }
    });

    // Labels edit toggle — show/hide label fields
    labelsEditToggle.addEventListener('change', () => {
      if (labelsEditToggle.checked) {
        labelsFields.classList.add('visible');
      } else {
        labelsFields.classList.remove('visible');
      }
    });

    // Address fields — live update frame preview
    addressStreet.addEventListener('input', () => updateFrameAddress());
    addressCity.addEventListener('input', () => updateFrameAddress());

    // Funda URL loading
    const fundaUrlInput = document.getElementById('fundaUrl');
    const btnFunda = document.getElementById('btnFunda');

    function getFundaUrl() {
      return fundaUrlInput.value.trim();
    }

    fundaUrlInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') loadFromFunda();
    });
    btnFunda.addEventListener('click', () => loadFromFunda());

    const btnTest = document.getElementById('btnTest');
    if (btnTest) {
      btnTest.addEventListener('click', () => {
        fundaUrlInput.value = 'https://www.funda.nl/detail/koop/arnhem/huis-madelievenstraat-61/43269652/';
        loadFromFunda();
      });
    }

    // Funda status checker
    const statusEls = [
      { box: document.getElementById('fundaStatus'), icon: document.getElementById('fundaStatusIcon'), text: document.getElementById('fundaStatusText') }
    ];

    function setFundaStatus(state, html) {
      for (const el of statusEls) {
        if (!el.box) continue;
        el.box.className = 'funda-status visible ' + state;
        if (state === 'loading') {
          el.icon.innerHTML = '<div class="mini-spinner"></div>';
        } else if (state === 'success') {
          el.icon.textContent = '✓';
        } else if (state === 'error') {
          el.icon.textContent = '✕';
        }
        el.text.innerHTML = html;
      }
    }

    function hideFundaStatus() {
      for (const el of statusEls) {
        if (el.box) el.box.className = 'funda-status';
      }
    }

    async function loadFromFunda() {
      const url = getFundaUrl();
      clearError();
      if (!url) { setError('Voer een Funda URL in.'); return; }
      if (!url.includes('funda.nl')) { setError('Dit is geen Funda URL.'); return; }

      // Show loading state in status checker + full-page spinner
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
          setFundaStatus('error', `<strong>Geen FML gevonden</strong><span>${data.error}</span>`);
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

        // Parse address from Funda URL first, fallback to FML labels
        const addr = parseFundaAddress(url) || parseAddressFromFML(data);
        if (addr) {
          addressStreet.value = addr.street;
          addressCity.value = addr.city;
          currentAddress = addr;
        } else {
          addressStreet.value = '';
          addressCity.value = '';
        }

        // Show success in status checker
        const addrStr = addr ? `${addr.street}, ${addr.city}` : 'Adres niet gevonden';
        setFundaStatus('success', `<strong>FML gevonden — ${data.floors.length} verdiepingen</strong><span>${addrStr}</span>`);

        processFloors(data);

        // Smooth scroll to the frame preview after loading
        setTimeout(() => {
          const preview = document.getElementById('framePreview');
          if (preview) preview.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 300);
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

    // Order button (placeholder for now)
    document.getElementById('btnOrder').addEventListener('click', () => {
      showToast('Bestelfunctie komt binnenkort!');
    });

    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
      // Ctrl+O / Cmd+O — open file
      if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
        e.preventDefault();
        fileInput.click();
      }
      // Ctrl+E / Cmd+E — export OBJ
      if ((e.ctrlKey || e.metaKey) && e.key === 'e') {
        e.preventDefault();
        if (floors.length > 0) exportOBJ();
      }
      // Ctrl+D / Cmd+D — download FML
      if ((e.ctrlKey || e.metaKey) && e.key === 'd') {
        e.preventDefault();
        if (originalFmlData) btnDownloadFml.click();
      }
    });

    // v16 — admin panel is always visible (no toggle)
