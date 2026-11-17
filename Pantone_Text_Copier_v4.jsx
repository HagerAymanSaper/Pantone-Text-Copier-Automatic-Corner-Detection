#target photoshop
app.bringToFront();

/*
  PANTONE TEXT COPIER v4 - automatic corner detection
  ---------------------------------------------------
  Workflow:
    1) Ctrl+Click ONLY the Pantone piece layers in the Layers panel.
    2) Run this script.
    3) Choose the text layer, the distance (default 0.7 cm) and, optionally, the
       preferred corner (only used to break ties between equally good corners).
    4) For every piece the script:
         - traces the real outline from the layer's transparency,
         - splits the outline into straight edges and finds every convex corner,
         - measures each corner's interior angle and picks the one closest to 90 degrees
           (if several are about 90, the one nearest your preferred corner),
         - places a copy of the text inside that corner, the chosen distance away from
           BOTH edges that form it (measured from the real outline), rotated to follow
           the corner's edges, and keeps it fully inside the piece where it fits.

  The original text layer and the piece layers are never resized or edited.
  A temporary "Work Path" is created and deleted (an existing Work Path is replaced).
*/

(function () {
    if (!app.documents.length) {
        alert("Open your Pantone document first.");
        return;
    }

    var doc = app.activeDocument;
    var oldUnits = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;

    // ---------- settings ----------
    var PATH_TOLERANCE = 1.0;   // px, work-path fitting tolerance
    var CURVE_STEPS    = 12;    // samples per bezier segment
    var RUN_DEV        = 2.0;   // px, max deviation for a "straight" edge run

    function cID(s) { return charIDToTypeID(s); }
    function sID(s) { return stringIDToTypeID(s); }

    // ---------- layer selection helpers ----------
    function getActiveLayerID() {
        var r = new ActionReference();
        r.putProperty(cID("Prpr"), cID("LyrI"));
        r.putEnumerated(cID("Lyr "), cID("Ordn"), cID("Trgt"));
        return executeActionGet(r).getInteger(cID("LyrI"));
    }

    function layerIDByIndex(idx) {
        var r = new ActionReference();
        r.putProperty(cID("Prpr"), cID("LyrI"));
        r.putIndex(cID("Lyr "), idx);
        return executeActionGet(r).getInteger(cID("LyrI"));
    }

    function selectLayerByID(id, add) {
        var d = new ActionDescriptor();
        var r = new ActionReference();
        r.putIdentifier(cID("Lyr "), id);
        d.putReference(cID("null"), r);
        if (add) {
            d.putEnumerated(sID("selectionModifier"), sID("selectionModifierType"), sID("addToSelection"));
        }
        d.putBoolean(cID("MkVs"), false);
        executeAction(cID("slct"), d, DialogModes.NO);
    }

    function getSelectedLayerIDs() {
        var ids = [];
        var hasBackground = false;
        try { var bg = doc.backgroundLayer; hasBackground = true; } catch (eBg) {}

        try {
            var ref = new ActionReference();
            ref.putEnumerated(cID("Dcmn"), cID("Ordn"), cID("Trgt"));
            var desc = executeActionGet(ref);
            if (desc.hasKey(sID("targetLayers"))) {
                var list = desc.getList(sID("targetLayers"));
                for (var i = 0; i < list.count; i++) {
                    var idx = list.getReference(i).getIndex();
                    if (!hasBackground) idx += 1;
                    ids.push(layerIDByIndex(idx));
                }
            }
        } catch (e) {
            ids = [];
        }

        if (!ids.length) {
            try { ids.push(getActiveLayerID()); } catch (e2) {}
        }
        return ids;
    }

    function collectTextLayers(container, prefix, output) {
        for (var i = 0; i < container.layers.length; i++) {
            var layer = container.layers[i];
            var fullName = prefix ? prefix + " / " + layer.name : layer.name;

            if (layer.typename === "ArtLayer" && layer.kind === LayerKind.TEXT) {
                try {
                    doc.activeLayer = layer;
                    output.push({ id: getActiveLayerID(), name: fullName, layer: layer });
                } catch (e) {}
            } else if (layer.typename === "LayerSet") {
                collectTextLayers(layer, fullName, output);
            }
        }
    }

    // ---------- dialog ----------
    function chooseSettings(items) {
        var w = new Window("dialog", "Pantone Text Copier");
        w.orientation = "column";
        w.alignChildren = "fill";
        w.spacing = 10;
        w.margins = 16;

        w.add("statictext", undefined, "Text layer to copy onto the selected pieces:");
        var names = [];
        for (var i = 0; i < items.length; i++) names.push(items[i].name);
        var dropdown = w.add("dropdownlist", undefined, names);
        dropdown.selection = 0;
        dropdown.preferredSize.width = 380;

        var g = w.add("group");
        g.add("statictext", undefined, "Distance from both edges of the corner (cm):");
        var et = g.add("edittext", undefined, "0.7");
        et.characters = 6;

        var gc = w.add("group");
        gc.add("statictext", undefined, "Preferred corner (tie-break between ~90\u00B0 corners):");
        var cornerDD = gc.add("dropdownlist", undefined,
            ["Bottom-left", "Bottom-right", "Top-left", "Top-right"]);
        cornerDD.selection = 0;

        var buttons = w.add("group");
        buttons.alignment = "right";
        buttons.add("button", undefined, "Cancel", { name: "cancel" });
        buttons.add("button", undefined, "Place Labels", { name: "ok" });

        if (w.show() !== 1 || !dropdown.selection) return null;

        var v = parseFloat(String(et.text).replace(",", "."));
        if (isNaN(v) || v < 0) {
            alert("Invalid distance value.");
            return null;
        }
        var ci = cornerDD.selection ? cornerDD.selection.index : 0;
        return { item: items[dropdown.selection.index], insetCm: v,
                 right: (ci === 1 || ci === 3), top: (ci === 2 || ci === 3) };
    }

    // ---------- outline extraction ----------
    function selectLayerTransparency() {
        // Standard Script-Listener form: acts on the currently targeted layer.
        var d = new ActionDescriptor();
        var rSel = new ActionReference();
        rSel.putProperty(cID("Chnl"), cID("fsel"));
        d.putReference(cID("null"), rSel);

        var rTrsp = new ActionReference();
        rTrsp.putEnumerated(cID("Chnl"), cID("Chnl"), cID("Trsp"));
        d.putReference(cID("T   "), rTrsp);

        executeAction(cID("setd"), d, DialogModes.NO);
    }

    function findWorkPath() {
        for (var i = 0; i < doc.pathItems.length; i++) {
            try {
                if (doc.pathItems[i].kind === PathKind.WORKPATH) return doc.pathItems[i];
            } catch (e) {}
        }
        return null;
    }

    function removeWorkPath() {
        var wp = findWorkPath();
        if (wp) { try { wp.remove(); } catch (e) {} }
    }

    function pt(arr) { return { x: Number(arr[0]), y: Number(arr[1]) }; }
    function dist(a, b) { var dx = a.x - b.x, dy = a.y - b.y; return Math.sqrt(dx * dx + dy * dy); }

    function flattenSubPath(sub) {
        var pts = sub.pathPoints;
        var n = pts.length;
        var out = [];
        if (n < 2) return out;
        var closed = sub.closed;
        var segCount = closed ? n : n - 1;

        for (var i = 0; i < segCount; i++) {
            var p0 = pts[i], p1 = pts[(i + 1) % n];
            var a0 = pt(p0.anchor), c0 = pt(p0.rightDirection);
            var c1 = pt(p1.leftDirection), a1 = pt(p1.anchor);

            if (dist(a0, c0) < 0.01 && dist(a1, c1) < 0.01) {
                out.push(a0);
            } else {
                for (var s = 0; s < CURVE_STEPS; s++) {
                    var t = s / CURVE_STEPS, mt = 1 - t;
                    out.push({
                        x: mt * mt * mt * a0.x + 3 * mt * mt * t * c0.x + 3 * mt * t * t * c1.x + t * t * t * a1.x,
                        y: mt * mt * mt * a0.y + 3 * mt * mt * t * c0.y + 3 * mt * t * t * c1.y + t * t * t * a1.y
                    });
                }
            }
        }
        if (!closed) out.push(pt(pts[n - 1].anchor));

        // drop consecutive duplicates (including wrap-around)
        var clean = [];
        for (var k = 0; k < out.length; k++) {
            if (!clean.length || dist(clean[clean.length - 1], out[k]) > 0.01) clean.push(out[k]);
        }
        while (clean.length > 1 && dist(clean[0], clean[clean.length - 1]) <= 0.01) clean.pop();
        return clean;
    }

    function polyArea2(P) {
        var a = 0;
        for (var i = 0; i < P.length; i++) {
            var j = (i + 1) % P.length;
            a += P[i].x * P[j].y - P[j].x * P[i].y;
        }
        return a;
    }

    // Trace the active layer's transparency and return its OUTER outline as a polygon.
    function getOutline(id) {
        selectLayerByID(id);
        removeWorkPath();
        selectLayerTransparency();

        var hasSel = true;
        try { var sb = doc.selection.bounds; } catch (eSel) { hasSel = false; }
        if (!hasSel) {
            throw new Error("The layer has no visible pixels (empty transparency selection).");
        }

        try {
            doc.selection.makeWorkPath(PATH_TOLERANCE);   // returns nothing in the DOM
        } catch (eMk) {
            try { doc.selection.deselect(); } catch (e1) {}
            throw new Error("Photoshop could not create a work path: " + eMk.message);
        }
        try { doc.selection.deselect(); } catch (e2) {}

        var wp = findWorkPath();
        if (!wp) throw new Error("Work path was not created.");

        var best = null, bestArea = 0;
        try {
            for (var s = 0; s < wp.subPathItems.length; s++) {
                var poly = flattenSubPath(wp.subPathItems[s]);
                if (poly.length < 3) continue;
                var area = Math.abs(polyArea2(poly));
                if (area > bestArea) { bestArea = area; best = poly; }
            }
        } finally {
            try { wp.remove(); } catch (e3) {}
        }

        if (!best) throw new Error("Traced outline was empty.");
        return best;
    }

    // ---------- edge detection ----------
    function turnAngle(ax, ay, bx, by) {
        var cross = ax * by - ay * bx;
        var dot = ax * bx + ay * by;
        return Math.abs(Math.atan2(cross, dot)) * 180 / Math.PI;
    }

    function maxDeviation(verts, a, b) {
        var dx = b.x - a.x, dy = b.y - a.y;
        var len = Math.sqrt(dx * dx + dy * dy);
        if (len < 1e-6) return 1e9;
        var m = 0;
        for (var i = 0; i < verts.length; i++) {
            var d = Math.abs((verts[i].x - a.x) * dy - (verts[i].y - a.y) * dx) / len;
            if (d > m) m = d;
        }
        return m;
    }

    // Split the closed polygon into straight-ish runs (curves become short chords).
    function buildRuns(P) {
        var n = P.length;
        var dirs = [];
        for (var i = 0; i < n; i++) {
            var j = (i + 1) % n;
            dirs.push({ x: P[j].x - P[i].x, y: P[j].y - P[i].y });
        }

        // start at the sharpest corner so a run never needs to wrap around
        var start = 0, bestTurn = -1;
        for (var k = 0; k < n; k++) {
            var prev = (k - 1 + n) % n;
            var t = turnAngle(dirs[prev].x, dirs[prev].y, dirs[k].x, dirs[k].y);
            if (t > bestTurn) { bestTurn = t; start = k; }
        }

        var runs = [];
        var cur = null;
        for (var m = 0; m < n; m++) {
            var idx = (start + m) % n;
            var nxt = (idx + 1) % n;
            if (!cur) {
                cur = { verts: [P[idx], P[nxt]] };
                continue;
            }
            var testVerts = cur.verts;
            var candidateEnd = P[nxt];
            if (maxDeviation(testVerts, testVerts[0], candidateEnd) <= RUN_DEV) {
                cur.verts.push(candidateEnd);
            } else {
                runs.push(cur);
                cur = { verts: [P[idx], P[nxt]] };
            }
        }
        if (cur) runs.push(cur);
        return runs;
    }

    function contains(arr, v) {
        for (var i = 0; i < arr.length; i++) if (arr[i] === v) return true;
        return false;
    }

    function angDiff(a, b) {
        var d = Math.abs(a - b);
        return d > 180 ? 360 - d : d;
    }

    function pointAlong(pts, target) {
        var acc = 0;
        for (var i = 0; i < pts.length - 1; i++) {
            var seg = dist(pts[i], pts[i + 1]);
            if (seg > 0 && acc + seg >= target) {
                var t = (target - acc) / seg;
                return { x: pts[i].x + (pts[i + 1].x - pts[i].x) * t,
                         y: pts[i].y + (pts[i + 1].y - pts[i].y) * t };
            }
            acc += seg;
        }
        return pts[pts.length - 1];
    }

    function pointInPoly(p, P) {
        var inside = false;
        for (var i = 0, j = P.length - 1; i < P.length; j = i++) {
            var xi = P[i].x, yi = P[i].y, xj = P[j].x, yj = P[j].y;
            if (((yi > p.y) !== (yj > p.y)) &&
                (p.x < (xj - xi) * (p.y - yi) / (yj - yi) + xi)) {
                inside = !inside;
            }
        }
        return inside;
    }

    function distToPolyBoundary(p, P) {
        var best = 1e18;
        for (var i = 0, j = P.length - 1; i < P.length; j = i++) {
            var ax = P[j].x, ay = P[j].y, bx = P[i].x, by = P[i].y;
            var dx = bx - ax, dy = by - ay;
            var l2 = dx * dx + dy * dy;
            var t = l2 > 0 ? ((p.x - ax) * dx + (p.y - ay) * dy) / l2 : 0;
            if (t < 0) t = 0; else if (t > 1) t = 1;
            var ex = ax + t * dx - p.x, ey = ay + t * dy - p.y;
            var d2 = ex * ex + ey * ey;
            if (d2 < best) best = d2;
        }
        return Math.sqrt(best);
    }

    function boundsPx(layer) {
        var b = layer.bounds;
        return { left: b[0].as("px"), top: b[1].as("px"), right: b[2].as("px"), bottom: b[3].as("px") };
    }

    // ---------- corner detection ----------
    function reversedCopy(a) {
        var o = [];
        for (var i = a.length - 1; i >= 0; i--) o.push(a[i]);
        return o;
    }

    function unitVec(dx, dy) {
        var l = Math.sqrt(dx * dx + dy * dy);
        return l > 0 ? { x: dx / l, y: dy / l } : { x: 0, y: 0 };
    }

    // Merge consecutive straight runs into longer edges (a slightly bowed edge stays one edge).
    function buildChains(P) {
        var runs = buildRuns(P);
        var chains = [];
        var cur = null, firstAng = 0, lastAng = 0;
        for (var r = 0; r < runs.length; r++) {
            var v = runs[r].verts;
            var a = v[0], b = v[v.length - 1];
            var ang = Math.atan2(b.y - a.y, b.x - a.x) * 180 / Math.PI;
            if (cur && angDiff(ang, lastAng) <= 8 && angDiff(ang, firstAng) <= 20) {
                for (var q = 1; q < v.length; q++) cur.pts.push(v[q]);
                lastAng = ang;
            } else {
                if (cur) chains.push(cur);
                cur = { pts: v.slice(0) };
                firstAng = ang; lastAng = ang;
            }
        }
        if (cur) chains.push(cur);
        for (var c = 0; c < chains.length; c++) {
            var len = 0;
            for (var z = 0; z < chains[c].pts.length - 1; z++) len += dist(chains[c].pts[z], chains[c].pts[z + 1]);
            chains[c].len = len;
        }
        return chains;
    }

    // All convex corners formed by two long straight edges (rounded corners are handled by
    // intersecting the two edges' lines). Returns {corners, minX, maxX, minY, maxY, diag}.
    function findCorners(P, insetPx, Dfit) {
        var sgn = polyArea2(P) > 0 ? 1 : -1;
        var minX = 1e18, maxX = -1e18, minY = 1e18, maxY = -1e18;
        for (var i = 0; i < P.length; i++) {
            if (P[i].x < minX) minX = P[i].x;
            if (P[i].x > maxX) maxX = P[i].x;
            if (P[i].y < minY) minY = P[i].y;
            if (P[i].y > maxY) maxY = P[i].y;
        }
        var diag = Math.sqrt((maxX - minX) * (maxX - minX) + (maxY - minY) * (maxY - minY));
        var res = { corners: [], minX: minX, maxX: maxX, minY: minY, maxY: maxY, diag: diag };

        var chains = buildChains(P);
        var n = chains.length;
        var Lmin = Math.max(3 * insetPx, 0.02 * diag);    // what counts as a real edge
        var Cmax = Math.max(4 * insetPx, 0.02 * diag, 25); // max rounded-corner / connector length

        var longIdx = [];
        for (var k = 0; k < n; k++) if (chains[k].len >= Lmin) longIdx.push(k);
        var Lc = longIdx.length;
        if (Lc < 2) return res;

        for (var m = 0; m < Lc; m++) {
            var ia = longIdx[m], ib = longIdx[(m + 1) % Lc];
            if (ia === ib) continue;

            var conn = 0, j = (ia + 1) % n;
            while (j !== ib) { conn += chains[j].len; j = (j + 1) % n; }
            if (conn > Cmax) continue;       // joined by a real curve, not a corner

            var A = chains[ia], B = chains[ib];
            var aEnd = A.pts[A.pts.length - 1], bStart = B.pts[0];
            var pA = pointAlong(reversedCopy(A.pts), Math.min(Dfit, A.len));
            var pB = pointAlong(B.pts, Math.min(Dfit, B.len));
            var dA = unitVec(aEnd.x - pA.x, aEnd.y - pA.y);   // travelling INTO the corner
            var dB = unitVec(pB.x - bStart.x, pB.y - bStart.y); // travelling OUT of the corner

            var cross = dA.x * dB.y - dA.y * dB.x;
            if (Math.abs(cross) < 0.17) continue;   // (almost) straight: not a corner
            if (sgn * cross <= 0) continue;          // reflex (inward-pointing) corner

            var wx = bStart.x - aEnd.x, wy = bStart.y - aEnd.y;
            var s = (wx * dB.y - wy * dB.x) / cross;
            var V = { x: aEnd.x + s * dA.x, y: aEnd.y + s * dA.y };
            if (dist(V, aEnd) > Cmax * 1.5 || dist(V, bStart) > Cmax * 1.5) continue;

            var eA = { x: -dA.x, y: -dA.y };   // from the corner, along edge A
            var eB = dB;                        // from the corner, along edge B
            var dot = eA.x * eB.x + eA.y * eB.y;
            var theta = Math.acos(Math.max(-1, Math.min(1, dot))) * 180 / Math.PI;
            res.corners.push({ V: V, eA: eA, eB: eB, theta: theta, dev: Math.abs(theta - 90) });
        }
        return res;
    }

    // Closest to 90 degrees wins; corners within 6 degrees of the best are treated as equal
    // and the one nearest the preferred bounding-box corner is taken.
    function chooseCorner(info, wantRight, wantTop) {
        var minDev = 1e18, i;
        for (i = 0; i < info.corners.length; i++) if (info.corners[i].dev < minDev) minDev = info.corners[i].dev;
        var tx = wantRight ? info.maxX : info.minX;
        var ty = wantTop ? info.minY : info.maxY;
        var best = null, bestD = 1e18;
        for (i = 0; i < info.corners.length; i++) {
            var c = info.corners[i];
            if (c.dev > minDev + 6) continue;
            var d = dist(c.V, { x: tx, y: ty });
            if (d < bestD) { bestD = d; best = c; }
        }
        return best;
    }

    // Place a w x h label inside the corner, `inset` px from both of its edges.
    function placeLabel(corner, w, h, inset, outline, diag) {
        var eA = corner.eA, eB = corner.eB, V = corner.V;

        // inward normals of the two edges
        var dotAB = eA.x * eB.x + eA.y * eB.y;
        var mA = unitVec(eB.x - dotAB * eA.x, eB.y - dotAB * eA.y);
        var mB = unitVec(eA.x - dotAB * eB.x, eA.y - dotAB * eB.y);
        var det = mA.x * mB.y - mA.y * mB.x;
        if (Math.abs(det) < 1e-6) throw new Error("Corner edges are parallel.");
        var K = { x: V.x + (inset * mB.y - mA.y * inset) / det,
                  y: V.y + (mA.x * inset - mB.x * inset) / det };   // inset from both edge lines

        // reading direction: the edge direction closest to horizontal, pointing rightwards
        var rA = eA.x >= 0 ? eA : { x: -eA.x, y: -eA.y };
        var rB = eB.x >= 0 ? eB : { x: -eB.x, y: -eB.y };
        var r = Math.abs(Math.atan2(rA.y, rA.x)) <= Math.abs(Math.atan2(rB.y, rB.x)) ? rA : rB;
        var q = { x: r.y, y: -r.x };    // "up" for the text

        // which label corner points into the polygon corner
        var bis = unitVec(eA.x + eB.x, eA.y + eB.y);
        var bestSc = -1e18, sr = 1, sq = 1, cvec = null;
        for (var a = -1; a <= 1; a += 2) {
            for (var b = -1; b <= 1; b += 2) {
                var cx = a * r.x * w / 2 + b * q.x * h / 2;
                var cy = a * r.y * w / 2 + b * q.y * h / 2;
                var sc = -(cx * bis.x + cy * bis.y);
                if (sc > bestSc) { bestSc = sc; sr = a; sq = b; cvec = { x: cx, y: cy }; }
            }
        }
        var c0 = { x: K.x - cvec.x, y: K.y - cvec.y };
        var slide = { x: -sr * r.x, y: -sr * r.y };    // away from the corner along the reading axis

        var samples = [[-1, -1], [1, -1], [1, 1], [-1, 1], [-0.5, -1], [0.5, -1], [1, -0.5], [1, 0.5],
                       [0.5, 1], [-0.5, 1], [-1, 0.5], [-1, -0.5]];
        var centerAt = function (t) { return { x: c0.x + slide.x * t, y: c0.y + slide.y * t }; };
        var sampleAt = function (t, smp) {
            var c = centerAt(t);
            return { x: c.x + r.x * smp[0] * w / 2 + q.x * smp[1] * h / 2,
                     y: c.y + r.y * smp[0] * w / 2 + q.y * smp[1] * h / 2 };
        };
        var tol = Math.max(3, 0.04 * inset);
        var clearAt = function (t, thresh) {
            for (var si = 0; si < samples.length; si++) {
                var p = sampleAt(t, samples[si]);
                if (!pointInPoly(p, outline)) return false;
                if (distToPolyBoundary(p, outline) < thresh) return false;
            }
            return true;
        };
        var findT = function (thresh) {
            var step = inset / 3;
            if (step < 1) step = 1;
            var tMax = Math.min(diag, Math.max(6 * w, 20 * inset));
            for (var t = 0; t <= tMax; t += step) {
                if (clearAt(t, thresh)) {
                    if (t === 0) return 0;
                    var lo = t - step, hi = t;
                    for (var it = 0; it < 7; it++) {
                        var mid = (lo + hi) / 2;
                        if (clearAt(mid, thresh)) hi = mid; else lo = mid;
                    }
                    return hi;
                }
            }
            return null;
        };

        var T = null, usedFactor = 1, factors = [1.0, 0.75, 0.5];
        for (var fi = 0; fi < factors.length && T === null; fi++) {
            T = findT(Math.max(0, inset * factors[fi] - tol));
            usedFactor = factors[fi];
        }
        var fit = true;
        if (T === null) { T = 0; usedFactor = 0; fit = false; }

        var minClear = 1e18;
        for (var mi = 0; mi < samples.length; mi++) {
            var d = distToPolyBoundary(sampleAt(T, samples[mi]), outline);
            if (d < minClear) minClear = d;
        }
        var cen = centerAt(T);
        return { cx: cen.x, cy: cen.y, angle: Math.atan2(r.y, r.x) * 180 / Math.PI,
                 usedFactor: usedFactor, fit: fit, minClear: minClear };
    }

    // ---------- setup ----------
    var selectedIDs = getSelectedLayerIDs();
    if (!selectedIDs.length) {
        app.preferences.rulerUnits = oldUnits;
        alert("Select the Pantone piece layers first, then run the script.");
        return;
    }

    var textLayers = [];
    collectTextLayers(doc, "", textLayers);
    try { selectLayerByID(selectedIDs[0]); for (var q0 = 1; q0 < selectedIDs.length; q0++) selectLayerByID(selectedIDs[q0], true); } catch (e0) {}

    if (!textLayers.length) {
        app.preferences.rulerUnits = oldUnits;
        alert("No text layers were found in this document.");
        return;
    }

    var settings = chooseSettings(textLayers);
    if (!settings) {
        app.preferences.rulerUnits = oldUnits;
        return;
    }

    var chosen = settings.item;
    var pieceIDs = [];
    for (var k0 = 0; k0 < selectedIDs.length; k0++) {
        if (selectedIDs[k0] !== chosen.id) pieceIDs.push(selectedIDs[k0]);
    }
    if (!pieceIDs.length) {
        app.preferences.rulerUnits = oldUnits;
        alert("No piece layers are selected.\n\nSelect the Pantone piece layers only, then run the script again.");
        return;
    }

    var sourceLayer = chosen.layer;
    var insetPx = (settings.insetCm / 2.54) * doc.resolution;

    var created = 0;
    var skipped = [];
    var notes = [];

    // ---------- main loop ----------
    try {
        for (var n = 0; n < pieceIDs.length; n++) {
            var pieceName = "Layer ID " + pieceIDs[n];
            var copy = null;
            try {
                selectLayerByID(pieceIDs[n]);
                var lyr = doc.activeLayer;
                pieceName = lyr.name;

                if (lyr.typename === "LayerSet") {
                    throw new Error("is a group. Select the piece layers inside it, not the group.");
                }
                if (lyr.kind === LayerKind.TEXT) {
                    throw new Error("is a text layer, not a piece.");
                }
                var okKinds = [LayerKind.NORMAL, LayerKind.SMARTOBJECT, LayerKind.SOLIDFILL,
                               LayerKind.GRADIENTFILL, LayerKind.PATTERNFILL];
                var kindOK = false;
                for (var kk = 0; kk < okKinds.length; kk++) if (lyr.kind === okKinds[kk]) kindOK = true;
                if (!kindOK) throw new Error("is an unsupported layer type.");

                var outline = getOutline(pieceIDs[n]);

                // Duplicate the label (original untouched) and measure it unrotated.
                selectLayerByID(pieceIDs[n]);
                copy = sourceLayer.duplicate(doc, ElementPlacement.PLACEATBEGINNING);
                copy.name = sourceLayer.name + " - " + pieceName;
                copy.visible = true;

                var b0 = boundsPx(copy);
                var w = b0.right - b0.left;
                var h = b0.bottom - b0.top;
                if (w < 1 || h < 1) throw new Error("The text layer is empty.");

                var info = findCorners(outline, insetPx, Math.max(w + 4 * insetPx, 40));
                if (!info.corners.length) {
                    throw new Error("no usable corner found (the outline has no two straight edges meeting at a corner).");
                }
                var corner = chooseCorner(info, settings.right, settings.top);
                if (!corner) throw new Error("no usable corner found.");

                var pl = placeLabel(corner, w, h, insetPx, outline, info.diag);
                if (!pl.fit) {
                    notes.push(pieceName + ": WARNING - label does not fit inside the piece at this distance. Check it.");
                } else if (pl.usedFactor < 1) {
                    notes.push(pieceName + ": label placed closer than the full distance (" +
                               Math.round(pl.usedFactor * 100) + "%) because the corner is tight.");
                }

                // Centre the unrotated label on the computed centre, then rotate about it.
                copy.translate(pl.cx - (b0.left + b0.right) / 2, pl.cy - (b0.top + b0.bottom) / 2);
                if (Math.abs(pl.angle) > 0.05) {
                    copy.rotate(pl.angle, AnchorPosition.MIDDLECENTER);
                }

                created++;
                notes.push(pieceName + ": corner " + corner.theta.toFixed(1) + " deg (" +
                           info.corners.length + " corners found), text angle " + pl.angle.toFixed(1) +
                           " deg, clearance " + (pl.minClear / doc.resolution * 2.54).toFixed(2) + " cm.");
            } catch (pieceError) {
                if (copy) { try { copy.remove(); } catch (eRm) {} }
                skipped.push(pieceName + ": " + pieceError.message);
            }
        }
    } finally {
        removeWorkPath();
        try {
            selectLayerByID(pieceIDs[0]);
            for (var r2 = 1; r2 < pieceIDs.length; r2++) selectLayerByID(pieceIDs[r2], true);
        } catch (eSel2) {}
        app.preferences.rulerUnits = oldUnits;
    }

    // ---------- report ----------
    var msg;
    if (created === 0) {
        msg = "NO LABELS WERE CREATED.\n\nPieces processed: " + pieceIDs.length +
              "\nSkipped: " + skipped.length + "\n\n" + skipped.join("\n");
    } else {
        msg = "Labels created: " + created + " of " + pieceIDs.length;
        if (skipped.length) msg += "\nSkipped: " + skipped.length + "\n\n" + skipped.join("\n");
        if (notes.length) {
            var shown = notes.slice(0, 20);
            msg += "\n\nDetails:\n" + shown.join("\n");
            if (notes.length > 20) msg += "\n... (" + (notes.length - 20) + " more)";
        }
    }
    alert(msg);
})();
