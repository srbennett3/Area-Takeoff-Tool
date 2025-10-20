/* eslint-disable no-undef */
/**
 * Floorplan Space & Wall Analyzer
 * - Fabric.js for drawing
 * - State persisted to localStorage
 * - Scale per floor
 * - Excel export via SheetJS
 *
 * Key UX:
 * - Draw Scale: click two points to create a reference line. Enter real length + unit.
 * - Draw Space: click to add vertices; click near first vertex to finish. Drag vertices to edit.
 * - Select Space: click polygon. Select Edge: click near an edge.
 * - Edit properties in sidebar. All derived values auto-recompute.
 */

(function () {
  if (typeof window.fabric === "undefined") {
    const statusEl = document.getElementById("statusText");
    if (statusEl) statusEl.textContent = "Error: Fabric.js failed to load. Check network/CDN.";
    return;
  }
  // --------------------------
  // Constants and helpers
  // --------------------------
  const STORAGE_KEY = "fp_floorplan_app_state_v1";
  const DEFAULT_CANVAS_WIDTH = 1200;
  const DEFAULT_CANVAS_HEIGHT = 800;

  // Visual constants
  const COLOR_SPACE = "rgba(16, 185, 129, 0.18)"; // greenish
  const COLOR_SPACE_STROKE = "#10b981";
  const COLOR_SPACE_SELECTED = "rgba(37, 99, 235, 0.18)"; // blueish
  const COLOR_SPACE_SELECTED_STROKE = "#2563eb";
  const COLOR_EDGE_SELECTED = "rgba(255,221,87,0.95)"; // yellow highlight
  const COLOR_VERTEX_SELECTED_FILL = COLOR_EDGE_SELECTED; // match selected edge color
  const COLOR_VERTEX_SELECTED_STROKE = COLOR_EDGE_SELECTED; // match selected edge color
  const COLOR_EDGE_EXTERIOR = "#f59e0b"; // orange for exterior edges (persistent)
  // Configurable edge hover/selection buffer (in pixels in canvas space)
  const EDGE_HIT_BUFFER_PX = 6;
  // Configurable vertex drag/hover radius and handle size
  const VERTEX_DRAG_RADIUS_PX = 12;
  const VERTEX_HANDLE_SIZE_PX = 10;
  const EDGE_HIGHLIGHT_THICKNESS_PX = 6; // thickness of selected edge overlay (filled rect)
  const EDGE_OVERLAY_THICKNESS_PX = 2; // base edge visual thickness (filled rects replacing stroke)
  const TEMP_EDGE_THICKNESS_PX = 2; // thickness for temp draw segments
  // Configurable proximity (in pixels) to close polygon by clicking near first vertex
  const SPACE_CLOSE_THRESHOLD_PX = 12;

  // Units: Keep internal values in feet; convert for UI/export
  const METERS_PER_FOOT = 0.3048;
  function unitAbbrev() { return (AppState.displayUnit === "meters") ? "m" : "ft"; }
  function feetToDisplayLength(feet) {
    return (AppState.displayUnit === "meters") ? feet * METERS_PER_FOOT : feet;
  }
  function displayLengthToFeet(val) {
    return (AppState.displayUnit === "meters") ? val / METERS_PER_FOOT : val;
  }
  function feet2ToDisplayArea(feet2) {
    return (AppState.displayUnit === "meters") ? feet2 * (METERS_PER_FOOT * METERS_PER_FOOT) : feet2;
  }

  // IDs
  const dom = {
    canvasEl: document.getElementById("floorCanvas"),
    canvasHolder: document.getElementById("canvasHolder"),
    statusText: document.getElementById("statusText"),
    // Project
    projectName: document.getElementById("projectName"),

    // Floors
    floorSelect: document.getElementById("floorSelect"),
    btnAddFloor: document.getElementById("btnAddFloor"),
    btnDeleteFloor: document.getElementById("btnDeleteFloor"),
    fileFloorImage: document.getElementById("fileFloorImage"),

    // Drawing
    btnDrawSpace: document.getElementById("btnDrawSpace"),
    btnDeleteSpace: document.getElementById("btnDeleteSpace"),
    btnInsertVertex: document.getElementById("btnInsertVertex"),
    btnDeleteVertex: document.getElementById("btnDeleteVertex"),
    btnScaleDraw: document.getElementById("btnScaleDraw"),
    btnScaleToggle: document.getElementById("btnScaleToggle"),

    // Scale
    scaleLength: document.getElementById("scaleLength"),
    scaleUnit: document.getElementById("scaleUnit"),

    // Space props
    spaceName: document.getElementById("spaceName"),
    spaceCeiling: document.getElementById("spaceCeiling"),
    spaceArea: document.getElementById("spaceArea"),
    spaceExteriorPerim: document.getElementById("spaceExteriorPerim"),
    spaceCeilingUnit: document.getElementById("spaceCeilingUnit"),

    // Edge props
    edgeIsExterior: document.getElementById("edgeIsExterior"),
    edgeHeight: document.getElementById("edgeHeight"),
    edgeWinWidth: document.getElementById("edgeWinWidth"),
    edgeWinHeight: document.getElementById("edgeWinHeight"),
    edgeDirection: document.getElementById("edgeDirection"),
    edgeLength: document.getElementById("edgeLength"),
    edgeWindowArea: document.getElementById("edgeWindowArea"),
    edgeHeightUnit: document.getElementById("edgeHeightUnit"),
    edgeWinWidthUnit: document.getElementById("edgeWinWidthUnit"),
    edgeWinHeightUnit: document.getElementById("edgeWinHeightUnit"),

    // Export
    btnExportExcel: document.getElementById("btnExportExcel"),
  };

  function uid(prefix = "id") {
    return `${prefix}_${Math.random().toString(36).slice(2, 10)}`;
  }

  function clampNum(n) {
    if (typeof n !== "number" || isNaN(n)) return 0;
    return n;
  }

  function distance(p1, p2) {
    return Math.sqrt((p1.x - p2.x) ** 2 + (p1.y - p2.y) ** 2);
  }

  function segmentDistance(point, a, b) {
    // Min distance from point to line segment ab (symmetric inside/outside)
    const A = { x: a.x, y: a.y };
    const B = { x: b.x, y: b.y };
    const P = { x: point.x, y: point.y };
    const ABx = B.x - A.x, ABy = B.y - A.y;
    const APx = P.x - A.x, APy = P.y - A.y;
    const ab2 = ABx * ABx + ABy * ABy;
    if (ab2 === 0) return Math.sqrt(APx * APx + APy * APy);
    let t = (APx * ABx + APy * ABy) / ab2;
    t = Math.max(0, Math.min(1, t));
    const C = { x: A.x + ABx * t, y: A.y + ABy * t };
    return distance(P, C);
  }

  function toFixedSmart(n, digits = 1) {
    if (!isFinite(n)) return "-";
    return Number(n.toFixed(digits)).toString();
  }
  function roundToTenth(n) { return Math.round((isFinite(n) ? n : 0) * 10) / 10; }

  function isPointInPolygon(point, vertices) {
    let inside = false;
    for (let i = 0, j = vertices.length - 1; i < vertices.length; j = i++) {
      const xi = vertices[i].x, yi = vertices[i].y;
      const xj = vertices[j].x, yj = vertices[j].y;
      const intersect = ((yi > point.y) !== (yj > point.y)) &&
        (point.x < (xj - xi) * (point.y - yi) / ((yj - yi) || 1e-9) + xi);
      if (intersect) inside = !inside;
    }
    return inside;
  }

  // Shoelace area
  function polygonArea(pts) {
    let area = 0;
    const n = pts.length;
    for (let i = 0; i < n; i++) {
      const j = (i + 1) % n;
      area += pts[i].x * pts[j].y - pts[j].x * pts[i].y;
    }
    return Math.abs(area) / 2;
  }

  // --------------------------
  // App State
  // --------------------------
  const AppState = {
    projectName: "",
    displayUnit: "feet",
    floors: [], // [{ id, name, imageSrc, backgroundFit, scale: { realLen, pixelLen, unit, line }, spaces: [Space] }]
    activeFloorId: null,
  };

  // Fabric canvas
  const canvas = new fabric.Canvas(dom.canvasEl, {
    selection: true,
    preserveObjectStacking: true,
  });
  // Ensure hover cursor defaults to system arrow when hovering targets
  canvas.hoverCursor = "default";

  // Current interaction mode flags
  let isDrawingSpace = false;
  let tempDrawPoints = []; // for polygon drawing
  let tempDrawCircles = [];
  let tempDrawLines = [];

  let isDrawingScale = false;
  let tempScalePoints = []; // 0 or 1 or 2 points during scale draw
  const SCALE_LINE_WIDTH = 8;

  let isInsertingVertex = false; // insert vertex mode

  // Selected objects
  let selectedSpaceId = null;
  let selectedEdgeIndex = null; // index within selected space polygon edges
  let hoverEdgeIndex = null;    // edge index under cursor for the selected space (or null)
  let canDragSelectedSpace = false; // true when inside selected space and not near an edge
  let suppressDeselectUntilMouseUp = false; // guards background deselect while pointer is active near edge
  let lastSelectedSpaceId = null; // to restore selection if Fabric clears while pointer active
  let lastPointerCanvas = { x: 0, y: 0 };
  // Vertex selection state
  let selectedVertexIndex = null; // currently selected vertex index within selected space
  let selectedVertexVisual = null; // highlight circle for selected vertex

  // Mapping from polygon object to space id
  const polygonIdToSpaceId = new Map(); // fabric object id -> spaceId
  const spaceIdToPolygon = new Map();   // spaceId -> fabric.Polygon

  // --------------------------
  // Persistence
  // --------------------------
  function saveState() {
    try {
      const stateToSave = JSON.parse(JSON.stringify(AppState));
      localStorage.setItem(STORAGE_KEY, JSON.stringify(stateToSave));
    } catch (e) {
      console.warn("Failed to save state", e);
    }
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && parsed.floors) {
          AppState.floors = parsed.floors;
          AppState.activeFloorId = parsed.activeFloorId || (parsed.floors[0]?.id ?? null);
          if (parsed.projectName) AppState.projectName = parsed.projectName;
          if (parsed.displayUnit) AppState.displayUnit = parsed.displayUnit;
          // migrate scale to internal feet
          AppState.floors.forEach(f => {
            if (!f.scale) return;
            if (typeof f.scale.realLenFeet !== 'number') {
              const unit = f.scale.unit || 'feet';
              const real = clampNum(f.scale.realLen);
              f.scale.realLenFeet = unit === 'meters' ? (real / METERS_PER_FOOT) : real;
            }
          });
        }
      }
    } catch (e) {
      console.warn("Failed to load state", e);
    }
  }

  // --------------------------
  // UI helpers
  // --------------------------
  function setStatus(text) {
    dom.statusText.textContent = text;
  }

  function setProjectNameUI(name) {
    // Preserve spaces exactly as typed in the input; only trim for title fallback.
    const val = typeof name === "string" ? name : "";
    if (dom.projectName) dom.projectName.value = val;
    // Also reflect in document title (trim only for emptiness check)
    const titleVal = val && val.trim() ? val.trim() : "Project Title Click To Enter";
    document.title = `${titleVal} – Area Takeoff Tool`;
  }

  function confirmAction(message) {
    return window.confirm(message);
  }

  function promptText(message, defaultValue = "") {
    // Return string or null
    const r = window.prompt(message, defaultValue);
    if (r === null) return null;
    const trimmed = r.trim();
    if (!trimmed) return null;
    return trimmed;
  }

  function updateFloorSelectOptions() {
    dom.floorSelect.innerHTML = "";
    AppState.floors.forEach(f => {
      const opt = document.createElement("option");
      opt.value = f.id;
      opt.textContent = f.name;
      dom.floorSelect.appendChild(opt);
    });
    dom.floorSelect.value = AppState.activeFloorId || "";
  }

  function activeFloor() {
    return AppState.floors.find(f => f.id === AppState.activeFloorId) || null;
  }

  function ensureEdgeArrayForSpace(space) {
    const n = space.vertices.length;
    if (!Array.isArray(space.edges)) space.edges = [];
    if (space.edges.length !== n) {
      const existing = space.edges.slice(0);
      const newEdges = [];
      for (let i = 0; i < n; i++) {
        newEdges[i] = existing[i] || {
          id: uid("edge"),
          isExterior: false,
          height: undefined,
          winWidth: undefined,
          winHeight: undefined,
          direction: "N",
          length: 0,
          winArea: 0,
        };
      }
      space.edges = newEdges;
    }
  }

  function getScaleFactorForFloor(floor) {
    if (!floor?.scale) return 0;
    const pixelLen = clampNum(floor.scale.pixelLen);
    const realLenFeet = clampNum(floor.scale.realLenFeet);
    if (pixelLen <= 0 || realLenFeet <= 0) return 0;
    return realLenFeet / pixelLen; // feet per pixel
  }

  function setScaleInputsFromFloor(floor) {
    if (!floor || !floor.scale) {
      dom.scaleLength.value = "";
      dom.scaleUnit.value = AppState.displayUnit || "feet";
      return;
    }
    const feetLen = clampNum(floor.scale.realLenFeet || 0);
    dom.scaleLength.value = feetToDisplayLength(feetLen) || "";
    dom.scaleUnit.value = AppState.displayUnit || "feet";
  }

  // --------------------------
  // Fabric helpers
  // --------------------------
  function fitBackgroundImageToCanvas(img, floor) {
    // Compute scale to fit within canvas dimensions
    const canvasW = canvas.getWidth();
    const canvasH = canvas.getHeight();
    const imgW = img.width;
    const imgH = img.height;
    const scale = Math.min(canvasW / imgW, canvasH / imgH);
    img.scaleX = scale;
    img.scaleY = scale;
    img.originX = "left";
    img.originY = "top";
    img.left = (canvasW - imgW * scale) / 2;
    img.top = (canvasH - imgH * scale) / 2;

    floor.backgroundFit = {
      left: img.left,
      top: img.top,
      scaleX: img.scaleX,
      scaleY: img.scaleY,
    };
  }

  function setBackgroundFromFloor(floor) {
    return new Promise((resolve) => {
      if (!floor?.imageSrc) {
        canvas.setBackgroundImage(null, () => {
          canvas.renderAll();
          resolve();
        });
        return;
      }
      fabric.Image.fromURL(floor.imageSrc, (img) => {
        if (!floor.backgroundFit) {
          fitBackgroundImageToCanvas(img, floor);
        } else {
          img.scaleX = floor.backgroundFit.scaleX;
          img.scaleY = floor.backgroundFit.scaleY;
          img.left = floor.backgroundFit.left;
          img.top = floor.backgroundFit.top;
          img.originX = "left";
          img.originY = "top";
        }
        canvas.setBackgroundImage(img, () => {
          canvas.renderAll();
          resolve();
        });
      }, { crossOrigin: "anonymous" });
    });
  }

  function clearCanvasOverlays() {
    // Remove all polygon, scale lines, temp drawing visuals (not background)
    const keep = [];
    canvas.getObjects().forEach(obj => {
      if (obj === canvas.backgroundImage) return;
      keep.push(obj);
    });
    keep.forEach(obj => canvas.remove(obj));
  }

  function removeScaleVisuals() {
    const toRemove = canvas.getObjects().filter(o => {
      const t = o.get("fpType");
      return t === "scaleLine" || t === "scaleLabel" || t === "scaleLabelLeader";
    });
    toRemove.forEach(o => canvas.remove(o));
  }

  function drawScaleLineForFloor(floor) {
    if (!floor?.scale?.line) return;
    const visible = floor.scale.visible !== false; // default visible
    const { x1, y1, x2, y2 } = floor.scale.line;
    const line = new fabric.Line([x1, y1, x2, y2], {
      stroke: "#f59e0b",
      strokeWidth: SCALE_LINE_WIDTH,
      selectable: false,
      evented: false,
      hoverCursor: "default",
      visible
    });
    line.set("fpType", "scaleLine");
    canvas.add(line);
    line.bringToFront();
    canvas.renderAll();
  }

  function addPolygonForSpace(space) {
    // Points stored as absolute canvas coords. Convert to polygon points relative to left/top.
    const pts = space.vertices.map(p => ({ x: p.x, y: p.y }));
    const minX = Math.min(...pts.map(p => p.x));
    const minY = Math.min(...pts.map(p => p.y));
    const rel = pts.map(p => ({ x: p.x - minX, y: p.y - minY }));

    const poly = new fabric.Polygon(rel, {
      left: minX,
      top: minY,
      fill: COLOR_SPACE,
      stroke: null, // disable stroke; we render edges as precise overlays
      strokeWidth: 0,
      objectCaching: false,
      hasControls: true,
      hasBorders: false,
      selectable: true,
      perPixelTargetFind: true,
      hoverCursor: "default",
    });
    poly.set("fpType", "space");
    poly.set("spaceId", space.id);
    canvas.add(poly);

    polygonIdToSpaceId.set(poly.__uid || poly.owningCursor || poly.id || uid("poly"), space.id);
    spaceIdToPolygon.set(space.id, poly);

    enablePolygonVertexEditing(poly);
    // Draw edge overlays replacing polygon stroke
    updateEdgeOverlaysForSpace(space.id);
    return poly;
  }

  function getPolygonAbsolutePoints(poly) {
    // Convert polygon relative points to absolute canvas coords accounting for transform
    const matrix = poly.calcTransformMatrix();
    const points = poly.get("points");
    return points.map(p =>
      fabric.util.transformPoint(new fabric.Point(p.x - poly.pathOffset.x, p.y - poly.pathOffset.y), matrix)
    );
  }

  function refreshAllPolygonsForFloor(floor) {
    // Clear existing polygons
    const toRemove = canvas.getObjects().filter(o => o.get("fpType") === "space");
    toRemove.forEach(o => canvas.remove(o));
    // Remove existing edge overlays
    const overlays = canvas.getObjects().filter(o => o.get && o.get("fpType") === "edgeOverlay");
    overlays.forEach(o => canvas.remove(o));
    polygonIdToSpaceId.clear();
    spaceIdToPolygon.clear();

    floor.spaces.forEach(space => {
      addPolygonForSpace(space);
      updateEdgeOverlaysForSpace(space.id);
    });
    // After polygons, draw scale line overlay
    drawScaleLineForFloor(floor);
    canvas.renderAll();
  }

  // --------------------------
  // Edge overlays (replace polygon stroke with precise rectangles)
  // --------------------------
  function removeEdgeOverlaysForSpace(spaceId) {
    const toRemove = canvas.getObjects().filter(o => o.get && o.get("fpType") === "edgeOverlay" && o.get("spaceId") === spaceId);
    toRemove.forEach(o => canvas.remove(o));
  }

  function updateEdgeOverlaysForSpace(spaceId) {
    const floor = activeFloor();
    if (!floor) return;
    const space = floor.spaces.find(s => s.id === spaceId);
    if (!space) return;
    const poly = spaceIdToPolygon.get(spaceId);
    if (!poly) return;

    removeEdgeOverlaysForSpace(spaceId);
    const pts = getPolygonAbsolutePoints(poly);
    ensureEdgeArrayForSpace(space);
    for (let i = 0; i < pts.length; i++) {
      const a = pts[i];
      const b = pts[(i + 1) % pts.length];
      const dx = b.x - a.x;
      const dy = b.y - a.y;
      const len = Math.sqrt(dx * dx + dy * dy);
      const cx = (a.x + b.x) / 2;
      const cy = (a.y + b.y) / 2;
      const angleDeg = Math.atan2(dy, dx) * 180 / Math.PI;
      const edge = space.edges[i];
      const isExterior = !!(edge && edge.isExterior);
      const baseColor = isExterior ? COLOR_EDGE_EXTERIOR : ((selectedSpaceId === spaceId) ? COLOR_SPACE_SELECTED_STROKE : COLOR_SPACE_STROKE);
      const rect = new fabric.Rect({
        left: cx,
        top: cy,
        originX: "center",
        originY: "center",
        width: len,
        height: EDGE_OVERLAY_THICKNESS_PX,
        angle: angleDeg,
        fill: baseColor,
        stroke: null,
        selectable: false,
        evented: false,
        objectCaching: false,
      });
      rect.set("fpType", "edgeOverlay");
      rect.set("spaceId", spaceId);
      rect.set("edgeIndex", i);
      canvas.add(rect);
    }
    // Ensure selection visuals stay on top of overlays
    if (highlightedEdgeVisual) {
      highlightedEdgeVisual.bringToFront();
    }
    if (selectedVertexVisual) {
      selectedVertexVisual.bringToFront();
    }
    canvas.renderAll();
  }

  // --------------------------
  // Polygon vertex editing (Fabric custom controls)
  // --------------------------
  function enablePolygonVertexEditing(polygon) {
    // Based on Fabric polygon editing example
    polygon.edit = true;
    polygon.objectCaching = false;
    polygon.hasBorders = false;
    polygon.cornerColor = "#93c5fd"; // blue-300
    polygon.cornerStyle = "circle";
    polygon.cornerSize = VERTEX_HANDLE_SIZE_PX;
    polygon.touchCornerSize = Math.max(VERTEX_HANDLE_SIZE_PX, 24);
    polygon.transparentCorners = false;

    const lastControl = polygon.points.length - 1;

    polygon.cornerStyle = "circle";
    polygon.controls = polygon.points.reduce(function (acc, point, index) {
      acc["p" + index] = new fabric.Control({
        positionHandler: polygonPositionHandler(index),
        actionHandler: anchorWrapper(index, actionHandler),
        actionName: "modifyPolygon",
        pointIndex: index
      });
      return acc;
    }, {});

    polygon.hasControls = true;
    polygon.on("modified", () => {
      onPolygonModified(polygon);
    });
    polygon.on("mousedown", () => {
      selectSpaceByPolygon(polygon);
    });
  }

  function polygonPositionHandler(pointIndex) {
    return function (dim, finalMatrix, fabricObject) {
      const x = (fabricObject.points[pointIndex].x - fabricObject.pathOffset.x);
      const y = (fabricObject.points[pointIndex].y - fabricObject.pathOffset.y);
      return fabric.util.transformPoint(
        new fabric.Point(x, y),
        fabric.util.multiplyTransformMatrices(
          fabricObject.canvas.viewportTransform,
          fabricObject.calcTransformMatrix()
        )
      );
    };
  }

  function actionHandler(eventData, transform, x, y) {
    const polygon = transform.target;
    const currentControl = polygon.controls[polygon.__corner];
    const canvas = polygon.canvas;
    const vpt = canvas && canvas.viewportTransform ? canvas.viewportTransform : fabric.iMatrix;
    const invVpt = fabric.util.invertTransform(vpt);
    // Pointer in canvas coords
    const pointerCanvas = fabric.util.transformPoint(new fabric.Point(x, y), invVpt);
    // Convert pointer to polygon local space (before pathOffset subtraction)
    const invMat = fabric.util.invertTransform(polygon.calcTransformMatrix());
    const pointerLocal = fabric.util.transformPoint(pointerCanvas, invMat);
    const finalPoint = {
      x: pointerLocal.x + polygon.pathOffset.x,
      y: pointerLocal.y + polygon.pathOffset.y,
    };
    polygon.points[currentControl.pointIndex] = finalPoint;
    // Deselect any selected edge while dragging a vertex
    if (selectedEdgeIndex != null) {
      selectedEdgeIndex = null;
      hoverEdgeIndex = null;
      clearEdgeHighlight();
      updateEdgePanelFromSelection();
      if (canvas && canvas.upperCanvasEl && canvas.upperCanvasEl.style) {
        canvas.defaultCursor = "default";
        canvas.upperCanvasEl.style.cursor = "default";
      }
    }
    polygon.dirty = true;
    polygon.setCoords();
    // Keep vertex highlight synced if dragging the selected vertex
    if (selectedSpaceId && selectedVertexIndex != null && currentControl && typeof currentControl.pointIndex === 'number') {
      const spaceId = polygon.get("spaceId");
      if (spaceId === selectedSpaceId && currentControl.pointIndex === selectedVertexIndex) {
        const absPts = getPolygonAbsolutePoints(polygon);
        if (absPts[selectedVertexIndex]) updateVertexHighlightPosition(absPts[selectedVertexIndex]);
      }
    }
    return true;
  }

  function anchorWrapper(pointIndex, fn) {
    return function (eventData, transform, x, y) {
      const polygon = transform.target;
      const actionPerformed = fn(eventData, transform, x, y);
      // Update corner coordinates without forcing Fabric to recompute dimensions/left/top
      polygon.setCoords();
      return actionPerformed;
    };
  }

  // --------------------------
  // Interactions
  // --------------------------
  function enterDrawSpaceMode() {
    cancelAllModes();
    isDrawingSpace = true;
    tempDrawPoints = [];
    tempDrawCircles.forEach(c => canvas.remove(c));
    tempDrawLines.forEach(l => canvas.remove(l));
    tempDrawCircles = [];
    tempDrawLines = [];
    setStatus("Drawing space: click to add vertices, click near first point to finish.");
    // Crosshair while drawing spaces
    canvas.defaultCursor = "crosshair";
    // Deselect any selected space when starting a new draw
    canvas.discardActiveObject();
    onCanvasSelectionCleared();
    // Reset fills to default for all spaces
    canvas.getObjects().forEach(o => {
      if (o.get("fpType") === "space") {
        o.set("fill", COLOR_SPACE);
        o.set("stroke", COLOR_SPACE_STROKE);
      }
    });
    canvas.requestRenderAll();
  }

  function enterScaleMode() {
    cancelAllModes();
    isDrawingScale = true;
    tempScalePoints = [];
    setStatus("Scale: click two points to create reference line.");
    // Crosshair while drawing scale line
    canvas.defaultCursor = "crosshair";
    const floor = activeFloor();
    removeScaleVisuals();
    if (floor?.scale) {
      floor.scale.line = null;
      floor.scale.visible = true;
      saveState();
    }
  }

  function cancelAllModes() {
    isDrawingSpace = false;
    isDrawingScale = false;
    isInsertingVertex = false;
    // Reset cursor when leaving draw modes
    canvas.defaultCursor = "default";
    // Unhighlight insert vertex button
    if (dom.btnInsertVertex) {
      dom.btnInsertVertex.classList.remove('active');
    }
    // Turn off scale line editing
    canvas.getObjects().forEach(o => {
      if (o.get("fpType") === "scaleLine") {
        o.selectable = false;
        o.evented = false;
        o.hasControls = false;
        o.off("modified", onScaleLineModified);
      }
    });
  }

  function endDrawSpace() {
    if (!isDrawingSpace) return;
    if (tempDrawPoints.length < 3) {
      setStatus("Need at least 3 points for a polygon.");
      return;
    }
    const floor = activeFloor();
    if (!floor) return;

    // Build space
    const space = {
      id: uid("space"),
      name: "Room",
      ceilingHeight: null,
      vertices: tempDrawPoints.map(p => ({ x: p.x, y: p.y })),
      edges: [], // will be ensured
      area: 0,
      exteriorPerimeter: 0,
    };
    ensureEdgeArrayForSpace(space);
    floor.spaces.push(space);
    saveState();

    // Cleanup temp visuals
    tempDrawCircles.forEach(c => canvas.remove(c));
    tempDrawLines.forEach(l => canvas.remove(l));
    tempDrawPoints = [];
    tempDrawCircles = [];
    tempDrawLines = [];

    // Add polygon and immediately select it so vertex controls are visible
    const poly = addPolygonForSpace(space);
    recalcSpaceDerived(space);
    if (poly) {
      canvas.setActiveObject(poly);
    }
    selectSpace(space.id);
    updateEdgeOverlaysForSpace(space.id);
    canvas.requestRenderAll();
    isDrawingSpace = false;
    setStatus("Space created. Select edges by clicking near them.");
    saveState();
  }

  function onPolygonModified(poly) {
    // Update the space vertices based on polygon absolute points
    const spaceId = poly.get("spaceId");
    const floor = activeFloor();
    if (!floor || !spaceId) return;
    const space = floor.spaces.find(s => s.id === spaceId);
    if (!space) return;

    // Capture absolute positions of ALL vertices before container update
    const absPtsBefore = getPolygonAbsolutePoints(poly);
    
    // Recompute container dimensions to encompass all vertices
    if (typeof poly._setPositionDimensions === 'function') {
      poly._setPositionDimensions({});
    }
    
    // Get absolute positions after container recalculation
    const absPtsAfter = getPolygonAbsolutePoints(poly);
    
    // Calculate how much the vertices shifted (use first vertex as reference)
    if (absPtsBefore.length > 0 && absPtsAfter.length > 0) {
      const deltaX = absPtsBefore[0].x - absPtsAfter[0].x;
      const deltaY = absPtsBefore[0].y - absPtsAfter[0].y;
      
      // Compensate by adjusting the polygon's position
      poly.left += deltaX;
      poly.top += deltaY;
    }
    
    poly.setCoords();
    
    // Now save the corrected absolute positions
    const absPts = getPolygonAbsolutePoints(poly);
    space.vertices = absPts.map(p => ({ x: p.x, y: p.y }));
    ensureEdgeArrayForSpace(space);
    recalcSpaceDerived(space);
    updateSpacePanel(space);
    // Refresh overlays after any geometry change
    updateEdgeOverlaysForSpace(space.id);
    saveState();
  }

  function onCanvasMouseDown(opt) {
    const floor = activeFloor();
    if (!floor) {
      setStatus("Add a floor first.");
      return;
    }

    const pointer = canvas.getPointer(opt.e, false);
    lastPointerCanvas = { x: pointer.x, y: pointer.y };
    // Edge hover/selection gating: pointer cursor appears over edges only when a space is selected
    if (!selectedSpaceId) {
      canvas.defaultCursor = (isDrawingSpace || isDrawingScale) ? "crosshair" : "default";
    }
    if (isDrawingSpace) {
      // Close polygon if clicking near the first vertex (without adding a new vertex)
      if (tempDrawPoints.length > 0) {
        const first = tempDrawPoints[0];
        const dToFirst = distance(pointer, first);
        if (dToFirst <= SPACE_CLOSE_THRESHOLD_PX) {
          // Prevent the background mouse:down handler from clearing selection in this cycle
          suppressDeselectUntilMouseUp = true;
          if (opt && opt.e) { try { opt.e.preventDefault(); opt.e.stopPropagation(); } catch(_){} }
          endDrawSpace();
          return;
        }
      }
      // Add point + temp visuals
      const circ = new fabric.Circle({
        radius: 3,
        fill: "#93c5fd",
        left: pointer.x, // center at cursor
        top: pointer.y,  // center at cursor
        originX: "center",
        originY: "center",
        selectable: false,
        evented: false,
      });
      canvas.add(circ);
      tempDrawCircles.push(circ);

      if (tempDrawPoints.length > 0) {
        const prev = tempDrawPoints[tempDrawPoints.length - 1];
        const dx = pointer.x - prev.x;
        const dy = pointer.y - prev.y;
        const len = Math.sqrt(dx * dx + dy * dy);
        const cx = (prev.x + pointer.x) / 2;
        const cy = (prev.y + pointer.y) / 2;
        const angleDeg = Math.atan2(dy, dx) * 180 / Math.PI;
        const seg = new fabric.Rect({
          left: cx,
          top: cy,
          originX: "center",
          originY: "center",
          width: len,
          height: TEMP_EDGE_THICKNESS_PX,
          angle: angleDeg,
          fill: "#60a5fa",
          stroke: null,
          selectable: false,
          evented: false,
          objectCaching: false,
        });
        canvas.add(seg);
        tempDrawLines.push(seg);
      }
      tempDrawPoints.push({ x: pointer.x, y: pointer.y });
      canvas.renderAll();
      return;
    }

    if (isDrawingScale) {
      tempScalePoints.push({ x: pointer.x, y: pointer.y });
      if (tempScalePoints.length === 2) {
        // Compute scale from the two clicked points and override any existing scale line
        const [p1, p2] = tempScalePoints;
        tempScalePoints = [];
        const pixelLen = distance(p1, p2);
        floor.scale = floor.scale || { realLen: 0, pixelLen: 0, unit: dom.scaleUnit.value, line: null };
        floor.scale.pixelLen = pixelLen;
        floor.scale.unit = dom.scaleUnit.value;
        floor.scale.line = { x1: p1.x, y1: p1.y, x2: p2.x, y2: p2.y };
        floor.scale.visible = true;
        // Ask user to enter realLen if blank
        if (!floor.scale.realLen || floor.scale.realLen <= 0) {
          const entered = window.prompt("Enter real-world length for the drawn scale (in selected unit):", "10");
          const num = parseFloat(entered);
          if (!isNaN(num) && num > 0) {
            floor.scale.realLen = num;
            dom.scaleLength.value = num;
          }
        }
        saveState();
        recalcAllSpacesForFloor(floor);
        // Draw visuals
        drawScaleLineForFloor(floor);
        updateScaleToggleLabel();
        cancelAllModes();
        setStatus("Scale set. Reference line shown.");
        canvas.renderAll();
        return;
      }
    }

    // Insert vertex mode: only insert if crosshair cursor is visible (near an edge)
    if (isInsertingVertex && selectedSpaceId) {
      const isCrosshairVisible = (canvas.defaultCursor === "crosshair") || (canvas.upperCanvasEl && canvas.upperCanvasEl.style && canvas.upperCanvasEl.style.cursor === "crosshair");
      if (isCrosshairVisible && hoverEdgeIndex !== null) {
        insertVertexAtEdge(selectedSpaceId, hoverEdgeIndex, pointer);
        // Exit insert vertex mode and unhighlight button
        isInsertingVertex = false;
        if (dom.btnInsertVertex) {
          dom.btnInsertVertex.classList.remove('active');
        }
        setStatus("Vertex inserted.");
        canvas.defaultCursor = "default";
        return;
      } else if (!isCrosshairVisible) {
        // Clicked outside edge buffer - don't insert vertex
        return;
      }
    }

    // Selection & dragging: only allow dragging when move cursor is visible
    if (selectedSpaceId) {
      const poly = spaceIdToPolygon.get(selectedSpaceId);
      if (poly) {
        // Vertex selection: click near a vertex toggles selected vertex
        const absPtsPre = getPolygonAbsolutePoints(poly);
        const pointerForVertex = canvas.getPointer(opt.e, false);
        const vIdx = findClosestVertexIndex(absPtsPre, pointerForVertex, VERTEX_DRAG_RADIUS_PX);
        if (vIdx != null) {
          // Selecting a vertex clears edge selection
          selectedEdgeIndex = null;
          clearEdgeHighlight();
          updateEdgePanelFromSelection();
          setSelectedVertex(vIdx);
          setStatus(`Vertex ${vIdx + 1} selected.`);
        } else {
          // Clicking outside vertices clears vertex selection
          clearSelectedVertex();
        }
      // If the pointer cursor is visible and we already have a hover edge, select it immediately
      const pointerVisibleEarly = (canvas.defaultCursor === "pointer") || (canvas.upperCanvasEl && canvas.upperCanvasEl.style && canvas.upperCanvasEl.style.cursor === "pointer");
      if (pointerVisibleEarly && hoverEdgeIndex != null) {
        const absPtsHover = getPolygonAbsolutePoints(poly);
        clearSelectedVertex();
        selectedEdgeIndex = hoverEdgeIndex;
        highlightSelectedEdge(absPtsHover, selectedEdgeIndex);
        canvas.setActiveObject(poly);
        updateEdgePanelFromSelection();
        setStatus(`Edge ${selectedEdgeIndex + 1} selected.`);
        if (opt && opt.e) { try { opt.e.preventDefault(); opt.e.stopPropagation(); } catch(_){} }
        return;
      }
        const absPts = absPtsPre;
        // Avoid edge selection when near a vertex control
        const nearVertex = absPts.some(p => distance(pointer, p) <= Math.max(4, EDGE_HIT_BUFFER_PX * 0.6));
        let idx = nearVertex ? null : findClosestEdgeIndex(absPts, pointer, EDGE_HIT_BUFFER_PX);
        if (idx === null && hoverEdgeIndex != null) {
          // Use the last hover edge if the click hit-test missed (e.g., jitter or canvas rounding)
          idx = hoverEdgeIndex;
        } else if (idx !== null) {
          hoverEdgeIndex = idx; // update hover edge when we have a positive hit
        }
        const isPointerVisible = (canvas.defaultCursor === "pointer") || (canvas.upperCanvasEl && canvas.upperCanvasEl.style && canvas.upperCanvasEl.style.cursor === "pointer");
        // Only allow selection if the cursor is visibly a pointer (hover state) over the edge
        if (idx !== null || (isPointerVisible && hoverEdgeIndex != null)) {
          if (idx === null) idx = hoverEdgeIndex;
          if (!isPointerVisible) {
            canvas.defaultCursor = "pointer";
            if (canvas.upperCanvasEl && canvas.upperCanvasEl.style) canvas.upperCanvasEl.style.cursor = "pointer";
            if (opt && opt.e) { try { opt.e.preventDefault(); opt.e.stopPropagation(); } catch(_){} }
            return; // require visible pointer prior to selection
          }
          clearSelectedVertex();
          selectedEdgeIndex = idx;
          highlightSelectedEdge(absPts, idx);
          // Keep the polygon as the active object so Fabric doesn't fire selection:cleared
          canvas.setActiveObject(poly);
          updateEdgePanelFromSelection();
          setStatus(`Edge ${idx + 1} selected.`);
          canvas.defaultCursor = "pointer";
          if (opt && opt.e) { try { opt.e.preventDefault(); opt.e.stopPropagation(); } catch(_){} }
          return;
        } else {
          // Not near an edge: only allow dragging if move cursor is active
          const isMoveCursor = (canvas.defaultCursor === "move") || (canvas.upperCanvasEl && canvas.upperCanvasEl.style && canvas.upperCanvasEl.style.cursor === "move");
          if (!isMoveCursor) {
            // prevent inadvertent deselect/drag; keep selection and do nothing
            if (opt && opt.e) { try { opt.e.preventDefault(); opt.e.stopPropagation(); } catch(_){} }
            return;
          }
          clearEdgeHighlight();
          selectedEdgeIndex = null;
          updateEdgePanelFromSelection();
          canvas.defaultCursor = "default";
        }
      }
    }

    // If click is not inside any space polygon, deselect all
    const f = activeFloor();
    if (f && Array.isArray(f.spaces)) {
      let insideAny = false;
      for (const sp of f.spaces) {
        if (isPointInPolygon(pointer, sp.vertices)) { insideAny = true; break; }
      }
      // Keep selection if we're near an edge of the currently selected space
      if (!insideAny) {
        if (selectedSpaceId) {
          const poly = spaceIdToPolygon.get(selectedSpaceId);
          if (poly) {
            const absPts = getPolygonAbsolutePoints(poly);
            const idx = findClosestEdgeIndex(absPts, pointer, EDGE_HIT_BUFFER_PX);
            const pointerVisible = (canvas.defaultCursor === "pointer") || (canvas.upperCanvasEl && canvas.upperCanvasEl.style && canvas.upperCanvasEl.style.cursor === "pointer");
            if (idx !== null || hoverEdgeIndex != null || pointerVisible) {
              // do not clear selection when near an edge
              return;
            }
          }
        }
        canvas.discardActiveObject();
        onCanvasSelectionCleared();
        // Reset fills to default
        canvas.getObjects().forEach(o => {
          if (o.get("fpType") === "space") {
            o.set("fill", COLOR_SPACE);
            o.set("stroke", COLOR_SPACE_STROKE);
          }
        });
        canvas.renderAll();
      }
    }
  }

  function onCanvasSelectionCreated(e) {
    const target = e.selected?.[0];
    if (!target) return;
    if (target.get("fpType") === "space") {
      selectSpaceByPolygon(target);
      if (dom.btnDeleteSpace) {
        const show = canvas.getActiveObjects().length === 1;
        dom.btnDeleteSpace.style.display = show ? '' : 'none';
      }
    }
  }

  function onCanvasSelectionUpdated(e) {
    onCanvasSelectionCreated(e);
  }

  function onCanvasSelectionCleared() {
    // If pointer is active near an edge, restore previous selection immediately
    const pointerVisible = (canvas.defaultCursor === "pointer") || (canvas.upperCanvasEl && canvas.upperCanvasEl.style && canvas.upperCanvasEl.style.cursor === "pointer");
    if ((hoverEdgeIndex != null || pointerVisible) && lastSelectedSpaceId) {
      const poly = spaceIdToPolygon.get(lastSelectedSpaceId);
      if (poly) {
        canvas.setActiveObject(poly);
        selectSpace(lastSelectedSpaceId);
        canvas.requestRenderAll();
        return;
      }
    }
    // Deselect space
    selectedSpaceId = null;
    selectedEdgeIndex = null;
    clearSelectedVertex();
    hoverEdgeIndex = null;
    clearEdgeHighlight();
    // Keep edge overlays visible; refresh to base color when nothing is selected
    const floor = activeFloor();
    if (floor && Array.isArray(floor.spaces)) {
      floor.spaces.forEach(s => updateEdgeOverlaysForSpace(s.id));
    }
    updateSpacePanel();
    updateEdgePanelFromSelection();
    setSpaceInputsEnabled(false);
    setEdgeInputsEnabled(false);
    if (dom.btnDeleteSpace) dom.btnDeleteSpace.style.display = 'none';
    if (dom.btnInsertVertex) dom.btnInsertVertex.style.display = 'none';
    // Cancel insert vertex mode if active
    if (isInsertingVertex) {
      isInsertingVertex = false;
      if (dom.btnInsertVertex) {
        dom.btnInsertVertex.classList.remove('active');
      }
      canvas.defaultCursor = "default";
    }
  }

  function findClosestEdgeIndex(points, clickPoint, tolerancePx) {
    if (!Array.isArray(points) || points.length < 2) return null;
    let bestIdx = null;
    let bestDist = Infinity;
    for (let i = 0; i < points.length; i++) {
      const j = (i + 1) % points.length;
      const d = segmentDistance(clickPoint, points[i], points[j]);
      if (d < bestDist) {
        bestDist = d;
        bestIdx = i;
      }
    }
    if (bestDist <= tolerancePx) return bestIdx;
    return null;
  }

  function findClosestVertexIndex(points, clickPoint, tolerancePx) {
    if (!Array.isArray(points) || points.length === 0) return null;
    let bestIdx = null;
    let bestDist = Infinity;
    for (let i = 0; i < points.length; i++) {
      const d = distance(clickPoint, points[i]);
      if (d < bestDist) {
        bestDist = d;
        bestIdx = i;
      }
    }
    if (bestDist <= tolerancePx) return bestIdx;
    return null;
  }

  let highlightedEdgeVisual = null;
  function highlightSelectedEdge(points, idx) {
    clearEdgeHighlight();
    const a = points[idx];
    const b = points[(idx + 1) % points.length];
    const dx = b.x - a.x;
    const dy = b.y - a.y;
    const len = Math.sqrt(dx * dx + dy * dy);
    const cx = (a.x + b.x) / 2;
    const cy = (a.y + b.y) / 2;
    const angleDeg = Math.atan2(dy, dx) * 180 / Math.PI;
    // Use a filled rotated rectangle so coverage is symmetric and fully occludes the underlying stroke
    highlightedEdgeVisual = new fabric.Rect({
      left: cx,
      top: cy,
      originX: "center",
      originY: "center",
      width: len,
      height: EDGE_HIGHLIGHT_THICKNESS_PX,
      rx: Math.min(EDGE_HIGHLIGHT_THICKNESS_PX / 2, 4),
      ry: Math.min(EDGE_HIGHLIGHT_THICKNESS_PX / 2, 4),
      angle: angleDeg,
      fill: COLOR_EDGE_SELECTED,
      stroke: null,
      selectable: false,
      evented: false,
      objectCaching: false,
    });
    highlightedEdgeVisual.set("fpType", "edgeHighlight");
    canvas.add(highlightedEdgeVisual);
    highlightedEdgeVisual.bringToFront();
    // Keep vertex highlight above selected edge as well
    if (selectedVertexVisual) selectedVertexVisual.bringToFront();
    canvas.renderAll();
  }
  function clearEdgeHighlight() {
    if (highlightedEdgeVisual) {
      canvas.remove(highlightedEdgeVisual);
      highlightedEdgeVisual = null;
      canvas.renderAll();
    }
  }

  // --------------------------
  // Vertex selection visuals
  // --------------------------
  function clearVertexHighlight() {
    if (selectedVertexVisual) {
      canvas.remove(selectedVertexVisual);
      selectedVertexVisual = null;
      canvas.renderAll();
    }
  }

  function drawVertexHighlightAt(point) {
    clearVertexHighlight();
    selectedVertexVisual = new fabric.Circle({
      radius: Math.max(6, VERTEX_HANDLE_SIZE_PX * 0.6),
      fill: COLOR_VERTEX_SELECTED_FILL,
      stroke: COLOR_VERTEX_SELECTED_STROKE,
      strokeWidth: 2,
      left: point.x,
      top: point.y,
      originX: "center",
      originY: "center",
      selectable: false,
      evented: false,
    });
    selectedVertexVisual.set("fpType", "vertexHighlight");
    canvas.add(selectedVertexVisual);
    selectedVertexVisual.bringToFront();
    canvas.renderAll();
  }

  function updateVertexHighlightPosition(point) {
    if (!selectedVertexVisual) return;
    selectedVertexVisual.set({ left: point.x, top: point.y });
    selectedVertexVisual.setCoords();
    selectedVertexVisual.bringToFront();
    canvas.renderAll();
  }

  function setSelectedVertex(index) {
    selectedVertexIndex = index;
    if (dom.btnDeleteVertex) dom.btnDeleteVertex.style.display = '';
    const poly = selectedSpaceId ? spaceIdToPolygon.get(selectedSpaceId) : null;
    if (poly) {
      const pts = getPolygonAbsolutePoints(poly);
      const p = pts[selectedVertexIndex];
      if (p) drawVertexHighlightAt(p);
    }
  }

  function clearSelectedVertex() {
    selectedVertexIndex = null;
    clearVertexHighlight();
    if (dom.btnDeleteVertex) dom.btnDeleteVertex.style.display = 'none';
  }

  function setSpaceInputsEnabled(enabled) {
    dom.spaceName.disabled = !enabled;
    dom.spaceCeiling.disabled = !enabled;
  }

  function setEdgeInputsEnabled(enabled) {
    dom.edgeIsExterior.disabled = !enabled;
    dom.edgeHeight.disabled = !enabled;
    dom.edgeWinWidth.disabled = !enabled;
    dom.edgeWinHeight.disabled = !enabled;
    dom.edgeDirection.disabled = !enabled;
  }

  function selectSpaceByPolygon(poly) {
    const sid = poly.get("spaceId");
    selectSpace(sid);
  }

  function selectSpace(spaceId) {
    const changedSpace = selectedSpaceId !== spaceId;
    selectedSpaceId = spaceId;
    lastSelectedSpaceId = spaceId;
    if (changedSpace) {
      selectedEdgeIndex = null;
      clearEdgeHighlight();
      hoverEdgeIndex = null;
      clearSelectedVertex();
    }
    // Update fills
    canvas.getObjects().forEach(o => {
      if (o.get("fpType") === "space") {
        if (o.get("spaceId") === spaceId) {
          o.set("fill", COLOR_SPACE_SELECTED);
        } else {
          o.set("fill", COLOR_SPACE);
        }
      }
    });
    canvas.renderAll();

    // Refresh edge overlays colors/geometry for all spaces
    const floorForOverlays = activeFloor();
    if (floorForOverlays && Array.isArray(floorForOverlays.spaces)) {
      floorForOverlays.spaces.forEach(s => updateEdgeOverlaysForSpace(s.id));
    }

    // Update panel
    const floor = activeFloor();
    const space = floor?.spaces.find(s => s.id === spaceId);
    updateSpacePanel(space);
    updateEdgePanelFromSelection();
    setSpaceInputsEnabled(!!space);
    // Keep edge inputs enabled when an edge is currently selected
    setEdgeInputsEnabled(selectedEdgeIndex != null);
    // Toggle Delete Space button visibility: only when exactly one space is selected
    if (dom.btnDeleteSpace) {
      const show = !!space && canvas.getActiveObjects().length === 1;
      dom.btnDeleteSpace.style.display = show ? '' : 'none';
    }
    // Toggle Insert Vertex button visibility: only when a space is selected
    if (dom.btnInsertVertex) {
      const show = !!space;
      dom.btnInsertVertex.style.display = show ? '' : 'none';
    }
    if (dom.btnDeleteVertex) {
      const showDel = !!space && selectedVertexIndex != null;
      dom.btnDeleteVertex.style.display = showDel ? '' : 'none';
    }
  }

  // --------------------------
  // Scale line modification
  // --------------------------
  function onScaleLineModified(opt) {
    const line = opt.target;
    const floor = activeFloor();
    if (!floor) return;
    const [x1, y1, x2, y2] = [line.x1, line.y1, line.x2, line.y2];
    const pxLen = distance({ x: x1, y: y1 }, { x: x2, y: y2 });
    floor.scale = floor.scale || { realLen: 0, pixelLen: 0, unit: dom.scaleUnit.value, line: null };
    floor.scale.pixelLen = pxLen;
    floor.scale.line = { x1, y1, x2, y2 };
    saveState();
    recalcAllSpacesForFloor(floor);
    setStatus("Scale line updated.");
  }

  // --------------------------
  // Panels update
  // --------------------------
  function updateSpacePanel(space = null) {
    if (!space) {
      dom.spaceName.value = "";
      dom.spaceCeiling.value = "";
      dom.spaceArea.textContent = "-";
      dom.spaceExteriorPerim.textContent = "-";
      if (dom.spaceCeiling) dom.spaceCeiling.classList.remove('input-error');
      return;
    }
    dom.spaceName.value = space.name || "";
    dom.spaceCeiling.value = space.ceilingHeight ?? "";
    dom.spaceArea.textContent = formatWithUnit(space.area, true);
    dom.spaceExteriorPerim.textContent = formatWithUnit(space.exteriorPerimeter, false);
    // Mark ceiling input error if selected space and empty
    if (selectedSpaceId && space.id === selectedSpaceId) {
      const empty = !(space.ceilingHeight || space.ceilingHeight === 0) ? true : (dom.spaceCeiling.value === "");
      if (empty) dom.spaceCeiling.classList.add('input-error'); else dom.spaceCeiling.classList.remove('input-error');
    } else {
      dom.spaceCeiling.classList.remove('input-error');
    }
  }

  function updateEdgePanelFromSelection() {
    const floor = activeFloor();
    if (!floor || !selectedSpaceId) {
      dom.edgeIsExterior.checked = false;
      dom.edgeHeight.value = "";
      dom.edgeWinWidth.value = "";
      dom.edgeWinHeight.value = "";
      dom.edgeDirection.value = "N";
      dom.edgeLength.textContent = "-";
      dom.edgeWindowArea.textContent = "-";
      setEdgeInputsEnabled(false);
      if (dom.edgeHeight) dom.edgeHeight.classList.remove('input-error');
      if (dom.edgeWinWidth) dom.edgeWinWidth.classList.remove('input-error');
      if (dom.edgeWinHeight) dom.edgeWinHeight.classList.remove('input-error');
      return;
    }
    // If a space is selected but no explicit edge index, keep inputs disabled until an edge is selected
    if (selectedEdgeIndex == null) {
    // Reset to defaults when no edge is selected
      dom.edgeIsExterior.checked = false;
      dom.edgeHeight.value = "";
      dom.edgeWinWidth.value = "";
      dom.edgeWinHeight.value = "";
      dom.edgeDirection.value = "N";
      dom.edgeLength.textContent = "-";
      dom.edgeWindowArea.textContent = "-";
      setEdgeInputsEnabled(false);
    if (dom.edgeHeight) dom.edgeHeight.classList.remove('input-error');
    if (dom.edgeWinWidth) dom.edgeWinWidth.classList.remove('input-error');
    if (dom.edgeWinHeight) dom.edgeWinHeight.classList.remove('input-error');
      return;
    }
    const space = floor.spaces.find(s => s.id === selectedSpaceId);
    ensureEdgeArrayForSpace(space);
    const edge = space.edges[selectedEdgeIndex];
    dom.edgeIsExterior.checked = !!edge.isExterior;
    dom.edgeHeight.value = (edge.height ?? "");
    dom.edgeWinWidth.value = (edge.winWidth ?? "");
    dom.edgeWinHeight.value = (edge.winHeight ?? "");
    dom.edgeDirection.value = edge.direction || "N";
    dom.edgeLength.textContent = formatWithUnit(edge.length, false);
    dom.edgeWindowArea.textContent = formatWithUnit(edge.winArea, true, true);
    // Ensure the inputs are enabled when an edge is selected
    setEdgeInputsEnabled(true);
    // Mark empty edge textboxes as error when an edge is selected
    if (dom.edgeHeight) (dom.edgeHeight.value === "" ? dom.edgeHeight.classList.add('input-error') : dom.edgeHeight.classList.remove('input-error'));
    if (dom.edgeWinWidth) (dom.edgeWinWidth.value === "" ? dom.edgeWinWidth.classList.add('input-error') : dom.edgeWinWidth.classList.remove('input-error'));
    if (dom.edgeWinHeight) (dom.edgeWinHeight.value === "" ? dom.edgeWinHeight.classList.add('input-error') : dom.edgeWinHeight.classList.remove('input-error'));
  }

  function formatWithUnit(value, isArea, showZero = false) {
    const unit = unitAbbrev();
    const displayValRaw = isArea ? feet2ToDisplayArea(value) : feetToDisplayLength(value);
    const displayVal = roundToTenth(displayValRaw);
    const suffix = isArea ? ` ${unit}²` : ` ${unit}`;
    if (!isFinite(displayVal)) return "-";
    if (displayVal > 0) return `${toFixedSmart(displayVal)}${suffix}`;
    if (showZero && displayVal === 0) return `0${suffix}`;
    return "-";
  }

  function updateUnitSuffixes() {
    const unit = unitAbbrev();
    if (dom.spaceCeilingUnit) dom.spaceCeilingUnit.textContent = unit;
    if (dom.edgeHeightUnit) dom.edgeHeightUnit.textContent = unit;
    if (dom.edgeWinWidthUnit) dom.edgeWinWidthUnit.textContent = unit;
    if (dom.edgeWinHeightUnit) dom.edgeWinHeightUnit.textContent = unit;
  }

  // --------------------------
  // Recalculations
  // --------------------------
  function recalcSpaceDerived(space) {
    const floor = activeFloor();
    if (!floor) return;
    const scale = getScaleFactorForFloor(floor);
    if (scale <= 0) {
      space.area = 0;
      space.exteriorPerimeter = 0;
      space.edges.forEach(e => {
        e.length = 0;
        // Keep inputs blank when undefined; winArea should reflect only defined numbers
        const w = clampNum(e.winWidth);
        const h = clampNum(e.winHeight);
        e.winArea = (isFinite(w) && isFinite(h) && w > 0 && h > 0) ? (w * h) : 0;
      });
      return;
    }

    const pts = space.vertices;
    const pxArea = polygonArea(pts);
    space.area = pxArea * scale * scale;

    ensureEdgeArrayForSpace(space);
    let exteriorPerim = 0;
    for (let i = 0; i < pts.length; i++) {
      const a = pts[i];
      const b = pts[(i + 1) % pts.length];
      const pxLen = distance(a, b);
      const edge = space.edges[i];
      edge.length = pxLen * scale;
      {
        const w = clampNum(edge.winWidth);
        const h = clampNum(edge.winHeight);
        edge.winArea = (isFinite(w) && isFinite(h) && w > 0 && h > 0) ? (w * h) : 0;
      }
      if (edge.isExterior) {
        exteriorPerim += edge.length;
      }
    }
    space.exteriorPerimeter = exteriorPerim;
  }

  function recalcAllSpacesForFloor(floor) {
    floor.spaces.forEach(s => recalcSpaceDerived(s));
    updatePanelsIfSelectionActive();
    saveState();
  }

  function updatePanelsIfSelectionActive() {
    if (selectedSpaceId) {
      const floor = activeFloor();
      const space = floor?.spaces.find(s => s.id === selectedSpaceId);
      updateSpacePanel(space);
      updateEdgePanelFromSelection();
    }
  }

  // --------------------------
  // Floors management
  // --------------------------
  function addFloorWithImage(imageSrc, name) {
    const floor = {
      id: uid("floor"),
      name,
      imageSrc,
      backgroundFit: null,
      scale: { realLen: 0, pixelLen: 0, unit: dom.scaleUnit.value, line: null, visible: true },
      spaces: [],
    };
    AppState.floors.push(floor);
    AppState.activeFloorId = floor.id;
    saveState();
    updateFloorSelectOptions();
    loadFloorIntoCanvas(floor);
  }

  async function loadFloorIntoCanvas(floor) {
    clearCanvasOverlays();
    await setBackgroundFromFloor(floor);
    refreshAllPolygonsForFloor(floor);
    setScaleInputsFromFloor(floor);
    updateScaleToggleLabel();
    updateUnitSuffixes();
    selectedSpaceId = null;
    selectedEdgeIndex = null;
    setStatus(`Loaded floor "${floor.name}".`);
  }

  function deleteActiveFloor() {
    const floor = activeFloor();
    if (!floor) return;
    if (!confirmAction(`Delete floor "${floor.name}" and all its spaces? This cannot be undone.`)) return;
    AppState.floors = AppState.floors.filter(f => f.id !== floor.id);
    if (AppState.floors.length > 0) {
      AppState.activeFloorId = AppState.floors[0].id;
    } else {
      AppState.activeFloorId = null;
    }
    saveState();
    updateFloorSelectOptions();
    const newFloor = activeFloor();
    if (newFloor) {
      loadFloorIntoCanvas(newFloor);
    } else {
      clearCanvasOverlays();
      canvas.setBackgroundImage(null, () => canvas.renderAll());
      setStatus("No floor selected.");
      setScaleInputsFromFloor(null);
    }
  }

  // --------------------------
  // Space operations
  // --------------------------
  function insertVertexAtEdge(spaceId, edgeIdx, clickPoint) {
    const floor = activeFloor();
    if (!floor) return;
    const space = floor.spaces.find(s => s.id === spaceId);
    if (!space) return;
    
    const poly = spaceIdToPolygon.get(spaceId);
    if (!poly) return;
    
    // Get the two vertices of the edge
    const v1 = space.vertices[edgeIdx];
    const v2 = space.vertices[(edgeIdx + 1) % space.vertices.length];
    
    // Find the closest point on the edge to the click point
    const A = { x: v1.x, y: v1.y };
    const B = { x: v2.x, y: v2.y };
    const P = { x: clickPoint.x, y: clickPoint.y };
    const ABx = B.x - A.x, ABy = B.y - A.y;
    const APx = P.x - A.x, APy = P.y - A.y;
    const ab2 = ABx * ABx + ABy * ABy;
    let t = 0.5; // default to midpoint
    if (ab2 > 0) {
      t = (APx * ABx + APy * ABy) / ab2;
      t = Math.max(0, Math.min(1, t));
    }
    const newVertex = {
      x: A.x + ABx * t,
      y: A.y + ABy * t
    };
    
    // Save the original edge properties before modifying the edges array
    ensureEdgeArrayForSpace(space);
    const originalEdge = space.edges[edgeIdx];
    const edgePropertiesToCopy = {
      isExterior: originalEdge.isExterior,
      height: originalEdge.height,
      winWidth: originalEdge.winWidth,
      winHeight: originalEdge.winHeight,
      direction: originalEdge.direction
    };
    
    // Insert the new vertex after edgeIdx
    space.vertices.splice(edgeIdx + 1, 0, newVertex);
    
    // Rebuild edge array (this will create a new edge at edgeIdx+1)
    ensureEdgeArrayForSpace(space);
    
    // Copy properties from original edge to the new edge created by the split
    if (space.edges[edgeIdx + 1]) {
      space.edges[edgeIdx + 1].isExterior = edgePropertiesToCopy.isExterior;
      space.edges[edgeIdx + 1].height = edgePropertiesToCopy.height;
      space.edges[edgeIdx + 1].winWidth = edgePropertiesToCopy.winWidth;
      space.edges[edgeIdx + 1].winHeight = edgePropertiesToCopy.winHeight;
      space.edges[edgeIdx + 1].direction = edgePropertiesToCopy.direction;
    }
    
    // Update polygon points without changing left/top: convert absolute vertices to polygon local space
    const invMat = fabric.util.invertTransform(poly.calcTransformMatrix());
    const newPoints = space.vertices.map(v => {
      const local = fabric.util.transformPoint(new fabric.Point(v.x, v.y), invMat);
      return { x: local.x + poly.pathOffset.x, y: local.y + poly.pathOffset.y };
    });
    poly.set({ points: newPoints });
    
    // Rebuild the vertex controls with new points array
    poly.controls = poly.points.reduce(function (acc, point, index) {
      acc["p" + index] = new fabric.Control({
        positionHandler: polygonPositionHandler(index),
        actionHandler: anchorWrapper(index, actionHandler),
        actionName: "modifyPolygon",
        pointIndex: index
      });
      return acc;
    }, {});
    
    // Capture absolute positions before dimension update
    const absPtsBefore = getPolygonAbsolutePoints(poly);
    
    // Update polygon dimensions and compensate for any shift (anchor by first absolute point)
    if (typeof poly._setPositionDimensions === 'function') {
      poly._setPositionDimensions({});
    }
    const absPtsAfter = getPolygonAbsolutePoints(poly);
    if (absPtsBefore.length > 0 && absPtsAfter.length > 0) {
      const deltaX = absPtsBefore[0].x - absPtsAfter[0].x;
      const deltaY = absPtsBefore[0].y - absPtsAfter[0].y;
      poly.left += deltaX;
      poly.top += deltaY;
    }
    
    poly.setCoords();
    poly.dirty = true;
    
    // Update space vertices with corrected absolute positions
    const finalAbsPts = getPolygonAbsolutePoints(poly);
    space.vertices = finalAbsPts.map(p => ({ x: p.x, y: p.y }));
    
    // Recalculate derived values
    recalcSpaceDerived(space);
    // Update edge overlays for this space
    updateEdgeOverlaysForSpace(spaceId);
    
    // Ensure the polygon is still selected and active
    canvas.setActiveObject(poly);
    canvas.renderAll();
    
    saveState();
  }

  function deleteSelectedSpace() {
    if (!selectedSpaceId) return;
    const floor = activeFloor();
    if (!floor) return;
    const space = floor.spaces.find(s => s.id === selectedSpaceId);
    if (!space) return;
    if (!confirmAction(`Delete space "${space.name || "Room"}"?`)) return;

    // Remove polygon from canvas
    const poly = spaceIdToPolygon.get(space.id);
    if (poly) {
      canvas.remove(poly);
      polygonIdToSpaceId.delete(poly.__uid || poly.owningCursor || poly.id);
      spaceIdToPolygon.delete(space.id);
    }
    floor.spaces = floor.spaces.filter(s => s.id !== space.id);
    // Remove overlays for this space
    removeEdgeOverlaysForSpace(space.id);
    selectedSpaceId = null;
    selectedEdgeIndex = null;
    clearEdgeHighlight();
    updateSpacePanel();
    updateEdgePanelFromSelection();
    saveState();
    setStatus("Space deleted.");
  }

  // --------------------------
  // Export to Excel
  // --------------------------
  function exportToExcel() {
    if (AppState.floors.length === 0) {
      alert("No floors to export.");
      return;
    }

    const wb = XLSX.utils.book_new();

    AppState.floors.forEach(floor => {
      const unit = unitAbbrev();
      // Build rows per space with required columns
      const rows = floor.spaces.map(space => {
        // Aggregate wall area and window area by direction (exterior only)
        const dirKeys = ["N","NW","NE","S","SW","SE","E","W"]; // keep ordering consistent with requirements
        const wallAreaByDir = Object.fromEntries(dirKeys.map(k => [k, 0]));
        const winAreaByDir = Object.fromEntries(dirKeys.map(k => [k, 0]));
        ensureEdgeArrayForSpace(space);
        for (const edge of space.edges) {
          if (!edge.isExterior) continue;
          const wallArea = clampNum(edge.length) * clampNum(edge.height);
          const winArea = clampNum(edge.winArea);
          const dir = edge.direction || "N";
          if (wallAreaByDir[dir] != null) wallAreaByDir[dir] += wallArea;
          if (winAreaByDir[dir] != null) winAreaByDir[dir] += winArea;
        }

        return {
          "Room Name": space.name || "",
          [`Average Ceiling Height (${unit})`]: roundToTenth(feetToDisplayLength(space.ceilingHeight || 0)),
          [`Exterior Perimeter Length (${unit})`]: roundToTenth(feetToDisplayLength(space.exteriorPerimeter || 0)),
          [`Floor Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(space.area || 0)),

          [`N Wall Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(wallAreaByDir["N"])) ,
          [`NW Wall Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(wallAreaByDir["NW"])) ,
          [`NE Wall Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(wallAreaByDir["NE"])) ,
          [`S Wall Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(wallAreaByDir["S"])) ,
          [`SW Wall Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(wallAreaByDir["SW"])) ,
          [`SE Wall Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(wallAreaByDir["SE"])) ,
          [`E Wall Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(wallAreaByDir["E"])) ,
          [`W Wall Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(wallAreaByDir["W"])) ,

          [`N Window Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(winAreaByDir["N"])) ,
          [`NW Window Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(winAreaByDir["NW"])) ,
          [`NE Window Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(winAreaByDir["NE"])) ,
          [`S Window Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(winAreaByDir["S"])) ,
          [`SW Window Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(winAreaByDir["SW"])) ,
          [`SE Window Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(winAreaByDir["SE"])) ,
          [`E Window Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(winAreaByDir["E"])) ,
          [`W Window Area (${unit}²)`]: roundToTenth(feet2ToDisplayArea(winAreaByDir["W"])) ,

          "_units_note": unit
        };
      });

      // Define column order
      const header = Object.keys(rows[0] || {});

      const ws = XLSX.utils.json_to_sheet(rows, { header });
      // Add units in header row 1 (optional)
      // Freeze header
      ws["!freeze"] = { xSplit: 0, ySplit: 1 };

      XLSX.utils.book_append_sheet(wb, ws, floor.name.substring(0, 31) || "Floor");
    });

    XLSX.writeFile(wb, "Floorplan_Export.xlsx");
  }

  // --------------------------
  // Event handlers: Panels
  // --------------------------
  dom.scaleLength.addEventListener("change", () => {
    const floor = activeFloor();
    if (!floor) return;
    const realLenDisplay = parseFloat(dom.scaleLength.value);
    if (!(realLenDisplay > 0)) {
      alert("Real length must be a positive number.");
      const existingFeet = clampNum(floor.scale?.realLenFeet || 0);
      dom.scaleLength.value = existingFeet ? feetToDisplayLength(existingFeet) : "";
      return;
    }
    const realLenFeet = displayLengthToFeet(realLenDisplay);
    floor.scale = floor.scale || { realLenFeet: 0, pixelLen: 0, unit: dom.scaleUnit.value, line: null };
    floor.scale.realLenFeet = realLenFeet;
    saveState();
    recalcAllSpacesForFloor(floor);
    setStatus("Scale real length updated.");
  });

  dom.scaleUnit.addEventListener("change", () => {
    AppState.displayUnit = dom.scaleUnit.value;
    saveState();
    updatePanelsIfSelectionActive();
    updateUnitSuffixes();
    setStatus("Scale unit updated.");
  });

  dom.spaceName.addEventListener("input", () => {
    if (!selectedSpaceId) return;
    const floor = activeFloor();
    const space = floor?.spaces.find(s => s.id === selectedSpaceId);
    if (!space) return;
    space.name = dom.spaceName.value.trim();
    saveState();
  });

  dom.spaceCeiling.addEventListener("change", () => {
    if (!selectedSpaceId) return;
    const val = parseFloat(dom.spaceCeiling.value);
    if (!(val >= 0)) {
      // revert to blank and mark error immediately
      dom.spaceCeiling.value = "";
      const floor0 = activeFloor();
      const space0 = floor0?.spaces.find(s => s.id === selectedSpaceId);
      if (space0) space0.ceilingHeight = undefined;
      if (dom.spaceCeiling) dom.spaceCeiling.classList.add('input-error');
      saveState();
      return;
    }
    const floor = activeFloor();
    const space = floor?.spaces.find(s => s.id === selectedSpaceId);
    if (!space) return;
    space.ceilingHeight = val;
    if (dom.spaceCeiling) dom.spaceCeiling.classList.remove('input-error');
    saveState();
  });

  dom.spaceCeiling.addEventListener("input", () => {
    if (!selectedSpaceId) return;
    const floor = activeFloor();
    const space = floor?.spaces.find(s => s.id === selectedSpaceId);
    if (!space) return;
    const raw = dom.spaceCeiling.value;
    const val = parseFloat(raw);
    if (raw === "" || !isFinite(val) || val < 0) {
      space.ceilingHeight = undefined;
      if (dom.spaceCeiling) dom.spaceCeiling.classList.add('input-error');
    } else {
      space.ceilingHeight = val;
      if (dom.spaceCeiling) dom.spaceCeiling.classList.remove('input-error');
    }
    saveState();
  });

  dom.edgeIsExterior.addEventListener("change", () => {
    const edge = getSelectedEdge();
    if (!edge) return;
    edge.isExterior = !!dom.edgeIsExterior.checked;
    recalcSelectedSpaceAndRefresh();
    // Refresh overlays so persistent exterior color applies immediately
    const floor = activeFloor();
    if (floor && selectedSpaceId) updateEdgeOverlaysForSpace(selectedSpaceId);
  });

  dom.edgeHeight.addEventListener("change", () => {
    const edge = getSelectedEdge();
    if (!edge) return;
    const val = parseFloat(dom.edgeHeight.value);
    if (!(val >= 0)) {
      // Revert to blank on invalid
      dom.edgeHeight.value = "";
      edge.height = undefined;
      recalcSelectedSpaceAndRefresh();
      return;
    }
    edge.height = val;
    recalcSelectedSpaceAndRefresh();
  });

  // Live updates while typing
  dom.edgeHeight.addEventListener("input", () => {
    const edge = getSelectedEdge();
    if (!edge) return;
    const val = parseFloat(dom.edgeHeight.value);
    edge.height = isFinite(val) && val >= 0 ? val : undefined;
    recalcSelectedSpaceAndRefresh();
  });

  dom.edgeWinWidth.addEventListener("change", () => {
    const edge = getSelectedEdge();
    if (!edge) return;
    const val = parseFloat(dom.edgeWinWidth.value);
    if (!(val >= 0)) {
      dom.edgeWinWidth.value = "";
      edge.winWidth = undefined;
      recalcSelectedSpaceAndRefresh();
      return;
    }
    edge.winWidth = val;
    recalcSelectedSpaceAndRefresh();
  });

  dom.edgeWinWidth.addEventListener("input", () => {
    const edge = getSelectedEdge();
    if (!edge) return;
    const val = parseFloat(dom.edgeWinWidth.value);
    edge.winWidth = isFinite(val) && val >= 0 ? val : undefined;
    recalcSelectedSpaceAndRefresh();
  });

  dom.edgeWinHeight.addEventListener("change", () => {
    const edge = getSelectedEdge();
    if (!edge) return;
    const val = parseFloat(dom.edgeWinHeight.value);
    if (!(val >= 0)) {
      dom.edgeWinHeight.value = "";
      edge.winHeight = undefined;
      recalcSelectedSpaceAndRefresh();
      return;
    }
    edge.winHeight = val;
    recalcSelectedSpaceAndRefresh();
  });

  dom.edgeWinHeight.addEventListener("input", () => {
    const edge = getSelectedEdge();
    if (!edge) return;
    const val = parseFloat(dom.edgeWinHeight.value);
    edge.winHeight = isFinite(val) && val >= 0 ? val : undefined;
    recalcSelectedSpaceAndRefresh();
  });

  dom.edgeDirection.addEventListener("change", () => {
    const edge = getSelectedEdge();
    if (!edge) return;
    edge.direction = dom.edgeDirection.value;
    saveState();
  });

  function getSelectedEdge() {
    const floor = activeFloor();
    if (!floor || !selectedSpaceId || selectedEdgeIndex == null) return null;
    const space = floor.spaces.find(s => s.id === selectedSpaceId);
    if (!space) return null;
    ensureEdgeArrayForSpace(space);
    return space.edges[selectedEdgeIndex] || null;
  }

  function recalcSelectedSpaceAndRefresh() {
    const floor = activeFloor();
    if (!floor || !selectedSpaceId) return;
    const space = floor.spaces.find(s => s.id === selectedSpaceId);
    if (!space) return;
    recalcSpaceDerived(space);
    updateSpacePanel(space);
    updateEdgePanelFromSelection();
    saveState();
  }

  // --------------------------
  // Event handlers: Toolbar buttons
  // --------------------------
  dom.btnDrawSpace.addEventListener("click", () => {
    const floor = activeFloor();
    if (!floor) {
      alert("Add a floor first.");
      return;
    }
    enterDrawSpaceMode();
  });

  if (dom.btnInsertVertex) {
    dom.btnInsertVertex.addEventListener("click", () => {
      if (!selectedSpaceId) {
        setStatus("Select a space first.");
        return;
      }
      // Toggle insert vertex mode
      if (isInsertingVertex) {
        // Cancel insert vertex mode
        isInsertingVertex = false;
        dom.btnInsertVertex.classList.remove('active');
        canvas.defaultCursor = "default";
        setStatus("Insert vertex mode cancelled.");
      } else {
        // Enter insert vertex mode
        isDrawingSpace = false;
        isDrawingScale = false;
        canvas.defaultCursor = "default"; // Will change to crosshair when near an edge
        isInsertingVertex = true;
        dom.btnInsertVertex.classList.add('active');
        // Deselect any currently selected edge or vertex
        selectedEdgeIndex = null;
        hoverEdgeIndex = null;
        clearEdgeHighlight();
        clearSelectedVertex();
        updateEdgePanelFromSelection();
        setStatus("Hover near an edge to insert a vertex.");
      }
    });
  }

  if (dom.btnScaleDraw) {
    dom.btnScaleDraw.addEventListener("click", () => {
      const floor = activeFloor();
      if (!floor) {
        alert("Add a floor first.");
        return;
      }
      enterScaleMode();
    });
  }

  function updateScaleToggleLabel() {
    const floor = activeFloor();
    if (!dom.btnScaleToggle || !floor) return;
    const isVisible = floor.scale?.visible !== false && !!floor.scale?.line;
    dom.btnScaleToggle.textContent = isVisible ? "Hide Scale" : "Show Scale";
  }

  if (dom.btnScaleToggle) {
    dom.btnScaleToggle.addEventListener("click", () => {
      const floor = activeFloor();
      if (!floor) {
        alert("Add a floor first.");
        return;
      }
      if (!floor.scale?.line) {
        alert("No scale line to show or hide. Use Draw Scale first.");
        return;
      }
      const newVisible = !(floor.scale.visible !== false);
      floor.scale.visible = newVisible;
      // Toggle visibility of existing visuals or redraw
      removeScaleVisuals();
      if (newVisible) {
        drawScaleLineForFloor(floor);
      }
      updateScaleToggleLabel();
      saveState();
      canvas.renderAll();
    });
  }

  dom.btnDeleteSpace.addEventListener("click", () => {
    deleteSelectedSpace();
  });

  // Delete Vertex button
  if (dom.btnDeleteVertex) {
    dom.btnDeleteVertex.addEventListener("click", () => {
      deleteSelectedVertex();
    });
  }

  dom.btnExportExcel.addEventListener("click", () => {
    exportToExcel();
  });

  dom.btnAddFloor.addEventListener("click", () => {
    dom.fileFloorImage.value = "";
    dom.fileFloorImage.click();
  });

  dom.fileFloorImage.addEventListener("change", (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const floorName = promptText("Enter floor name:", "First Floor");
    if (!floorName) return;
    const reader = new FileReader();
    reader.onload = () => {
      addFloorWithImage(reader.result, floorName);
    };
    reader.readAsDataURL(file);
  });

  dom.btnDeleteFloor.addEventListener("click", () => {
    deleteActiveFloor();
  });

  dom.floorSelect.addEventListener("change", async () => {
    const floorId = dom.floorSelect.value;
    AppState.activeFloorId = floorId;
    saveState();
    const floor = activeFloor();
    if (floor) {
      await loadFloorIntoCanvas(floor);
    }
  });

  // --------------------------
  // Canvas events
  // --------------------------
  canvas.on("mouse:down", onCanvasMouseDown);
  canvas.on("selection:created", onCanvasSelectionCreated);
  canvas.on("selection:updated", onCanvasSelectionUpdated);
  canvas.on("selection:cleared", onCanvasSelectionCleared);
  canvas.on("mouse:up", function(){ suppressDeselectUntilMouseUp = false; });
  // Update cursor on hover to show pointer only when a space is selected and near an edge
  canvas.on("mouse:move", function(opt) {
    if (isDrawingSpace || isDrawingScale) {
      canvas.defaultCursor = "crosshair";
      return;
    }
    if (!selectedSpaceId) {
      // Before selection: show pointer when hovering over any space polygon
      const pointer = canvas.getPointer(opt.e, false);
      const f = activeFloor();
      let overAny = false;
      if (f && Array.isArray(f.spaces)) {
        for (const sp of f.spaces) {
          if (isPointInPolygon(pointer, sp.vertices)) { overAny = true; break; }
        }
      }
      canvas.defaultCursor = overAny ? "pointer" : "default";
      if (canvas.upperCanvasEl && canvas.upperCanvasEl.style) {
        canvas.upperCanvasEl.style.cursor = canvas.defaultCursor;
      }
      return;
    }
    const poly = spaceIdToPolygon.get(selectedSpaceId);
    if (!poly) { canvas.defaultCursor = "default"; return; }
    const pointer = canvas.getPointer(opt.e, false); // canvas-space coords
    const absPts = getPolygonAbsolutePoints(poly);
    const nearVertex = absPts.some(p => distance(pointer, p) <= VERTEX_DRAG_RADIUS_PX);
    // Keep edge overlays perfectly aligned during hover/move
    updateEdgeOverlaysForSpace(selectedSpaceId);
    if (nearVertex && !isInsertingVertex) { hoverEdgeIndex = null; canDragSelectedSpace = false; canvas.defaultCursor = "default"; return; }
    const idx = findClosestEdgeIndex(absPts, pointer, EDGE_HIT_BUFFER_PX);
    hoverEdgeIndex = (idx !== null) ? idx : null;
    // Keep selected vertex highlight in sync
    if (selectedVertexIndex != null && Array.isArray(absPts) && absPts[selectedVertexIndex]) {
      updateVertexHighlightPosition(absPts[selectedVertexIndex]);
    }
    
    // Insert vertex mode: show crosshair when near edge, default otherwise
    if (isInsertingVertex) {
      if (hoverEdgeIndex !== null) {
        canvas.defaultCursor = "crosshair";
      } else {
        canvas.defaultCursor = "default";
      }
      if (canvas.upperCanvasEl && canvas.upperCanvasEl.style) {
        canvas.upperCanvasEl.style.cursor = canvas.defaultCursor;
      }
      return;
    }
    
    if (hoverEdgeIndex !== null) {
      canDragSelectedSpace = false;
      canvas.defaultCursor = "pointer";
      suppressDeselectUntilMouseUp = true;
    } else {
      // If pointer is inside the polygon (not near an edge), show move and allow dragging
      const floor = activeFloor();
      const space = floor?.spaces.find(s => s.id === selectedSpaceId);
      const inside = space ? isPointInPolygon(pointer, space.vertices) : false;
      canDragSelectedSpace = !!inside;
      if (inside) {
        canvas.defaultCursor = "move";
      } else {
        // Not near selected edge and not inside selected space: show pointer if over any other space
        let overOther = false;
        const f = activeFloor();
        if (f && Array.isArray(f.spaces)) {
          for (const sp of f.spaces) {
            if (sp.id === selectedSpaceId) continue;
            if (isPointInPolygon(pointer, sp.vertices)) { overOther = true; break; }
          }
        }
        canvas.defaultCursor = overOther ? "pointer" : "default";
      }
      if (!inside) suppressDeselectUntilMouseUp = false;
    }
    // Force the visible cursor to match our logic even when hovering Fabric targets
    if (canvas.upperCanvasEl && canvas.upperCanvasEl.style) {
      canvas.upperCanvasEl.style.cursor = canvas.defaultCursor;
    }
  });

  // Deselect when clicking empty background within the canvas area (but not outside app)
  canvas.on("mouse:down", function(opt) {
    if (isDrawingSpace || isDrawingScale) return;
    if (opt.target) return; // clicking on object
    // Do not clear selection if an edge is currently selected via custom logic
    if (selectedEdgeIndex != null) return;
    // Keep selection if hovering a selectable edge (pointer cursor active)
    const pointerVisible = (canvas.defaultCursor === "pointer") || (canvas.upperCanvasEl && canvas.upperCanvasEl.style && canvas.upperCanvasEl.style.cursor === "pointer");
    if (hoverEdgeIndex != null || pointerVisible || suppressDeselectUntilMouseUp) {
      if (opt && opt.e) { try { opt.e.preventDefault(); opt.e.stopPropagation(); } catch(_){} }
      return;
    }
    canvas.discardActiveObject();
    canvas.requestRenderAll();
    onCanvasSelectionCleared();
    canvas.defaultCursor = "default";
  });

  // Keep vertex highlight synced after polygon/object modifications
  canvas.on("object:modified", function(opt){
    const t = opt && opt.target;
    if (!t || (t.get && t.get("fpType") !== "space")) return;
    if (!selectedSpaceId || selectedVertexIndex == null) return;
    const poly = spaceIdToPolygon.get(selectedSpaceId);
    if (!poly) return;
    const pts = getPolygonAbsolutePoints(poly);
    if (Array.isArray(pts) && pts[selectedVertexIndex]) {
      updateVertexHighlightPosition(pts[selectedVertexIndex]);
    }
  });

  // --------------------------
  // Delete selected vertex
  // --------------------------
  function deleteSelectedVertex() {
    if (!selectedSpaceId || selectedVertexIndex == null) return;
    const floor = activeFloor();
    if (!floor) return;
    const space = floor.spaces.find(s => s.id === selectedSpaceId);
    if (!space) return;
    if (!Array.isArray(space.vertices) || space.vertices.length <= 3) {
      alert("A space must have at least 3 vertices.");
      return;
    }
    const poly = spaceIdToPolygon.get(selectedSpaceId);
    if (!poly) return;

    // Preserve edges prior to mutation
    ensureEdgeArrayForSpace(space);
    const prevEdges = space.edges.slice();
    const prevVertexCount = space.vertices.length;

    // Remove the vertex
    space.vertices.splice(selectedVertexIndex, 1);

    // Rebuild polygon geometry from updated vertices without changing left/top
    // Convert absolute vertices into polygon local space using inverse transform
    const invMat = fabric.util.invertTransform(poly.calcTransformMatrix());
    const newPoints = space.vertices.map(v => {
      const local = fabric.util.transformPoint(new fabric.Point(v.x, v.y), invMat);
      return {
        x: local.x + poly.pathOffset.x,
        y: local.y + poly.pathOffset.y,
      };
    });
    poly.set({ points: newPoints });

    // Rebuild controls
    poly.controls = poly.points.reduce(function (acc, point, index) {
      acc["p" + index] = new fabric.Control({
        positionHandler: polygonPositionHandler(index),
        actionHandler: anchorWrapper(index, actionHandler),
        actionName: "modifyPolygon",
        pointIndex: index
      });
      return acc;
    }, {});

    // Compensate for container resize like in vertex drag: anchor on pre-update absolute points
    const absPtsBefore = getPolygonAbsolutePoints(poly);
    if (typeof poly._setPositionDimensions === 'function') {
      poly._setPositionDimensions({});
    }
    const absPtsAfter = getPolygonAbsolutePoints(poly);
    if (absPtsBefore.length > 0 && absPtsAfter.length > 0) {
      const deltaX = absPtsBefore[0].x - absPtsAfter[0].x;
      const deltaY = absPtsBefore[0].y - absPtsAfter[0].y;
      poly.left += deltaX;
      poly.top += deltaY;
    }
    poly.setCoords();
    poly.dirty = true;

    // Persist corrected absolute positions back to space
    const finalAbsPts = getPolygonAbsolutePoints(poly);
    space.vertices = finalAbsPts.map(p => ({ x: p.x, y: p.y }));

    // Rebuild edges and try to preserve merged edge properties
    ensureEdgeArrayForSpace(space);
    const newN = space.vertices.length;
    const mergedIdx = (selectedVertexIndex - 1 + newN) % newN;
    const leftOldIdx = (selectedVertexIndex - 1 + prevVertexCount) % prevVertexCount;
    const candidate = prevEdges[leftOldIdx] || prevEdges[selectedVertexIndex % prevVertexCount];
    if (space.edges[mergedIdx] && candidate) {
      space.edges[mergedIdx].isExterior = !!candidate.isExterior;
      space.edges[mergedIdx].height = clampNum(candidate.height);
      space.edges[mergedIdx].winWidth = clampNum(candidate.winWidth);
      space.edges[mergedIdx].winHeight = clampNum(candidate.winHeight);
      space.edges[mergedIdx].direction = candidate.direction || "N";
    }

    // Recalculate derived values
    recalcSpaceDerived(space);
    // Update edge overlays for this space
    updateEdgeOverlaysForSpace(selectedSpaceId);

    // Clear selected vertex state after deletion
    clearSelectedVertex();

    // Keep polygon active
    canvas.setActiveObject(poly);
    canvas.renderAll();
    saveState();
    setStatus("Vertex deleted.");
  }

  // --------------------------
  // Initialization
  // --------------------------
  function init() {
    // Canvas base size
    canvas.setWidth(DEFAULT_CANVAS_WIDTH);
    canvas.setHeight(DEFAULT_CANVAS_HEIGHT);

    loadState();
    // Initialize project header
    setProjectNameUI(AppState.projectName);
    if (dom.projectName) {
      // Focus: if empty, show placeholder and clear value for easy typing
      dom.projectName.addEventListener('focus', () => {
        if (!AppState.projectName || !AppState.projectName.trim()) {
          dom.projectName.value = "";
        }
      });
      // Input: live update state and title
      dom.projectName.addEventListener('input', () => {
        const raw = dom.projectName.value;
        AppState.projectName = raw;
        setProjectNameUI(AppState.projectName);
        saveState();
      });
      // Blur: if empty, revert to placeholder behavior
      dom.projectName.addEventListener('blur', () => {
        if (!dom.projectName.value.trim()) {
          AppState.projectName = "";
          setProjectNameUI("");
          saveState();
        }
      });
    }
    updateFloorSelectOptions();
    if (AppState.activeFloorId) {
      const floor = activeFloor();
      loadFloorIntoCanvas(floor);
    } else {
      setStatus("Add a floor to begin.");
    }
    // Disable inputs until a selection is made
    setSpaceInputsEnabled(false);
    setEdgeInputsEnabled(false);
    // Delete Space button visibility on load
    if (dom.btnDeleteSpace) dom.btnDeleteSpace.style.display = 'none';
    if (dom.btnDeleteVertex) dom.btnDeleteVertex.style.display = 'none';

    // Keyboard panning with arrow keys (scroll the canvas holder)
    if (dom.canvasHolder) {
      const PAN_STEP = 40; // pixels per keypress
      document.addEventListener('keydown', (e) => {
        // Avoid intercepting typing in inputs/selects/textarea
        const tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : '';
        const isTyping = tag === 'input' || tag === 'textarea' || tag === 'select';
        if (isTyping) return;
        if (e.key === 'ArrowLeft') { dom.canvasHolder.scrollLeft -= PAN_STEP; }
        else if (e.key === 'ArrowRight') { dom.canvasHolder.scrollLeft += PAN_STEP; }
        else if (e.key === 'ArrowUp') { dom.canvasHolder.scrollTop -= PAN_STEP; }
        else if (e.key === 'ArrowDown') { dom.canvasHolder.scrollTop += PAN_STEP; }
      });
    }
  }

  init();
})();


