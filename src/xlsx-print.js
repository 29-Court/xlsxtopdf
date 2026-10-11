/* 29Court Convert: pixel-faithful XLSX -> PDF renderer (pdf-lib).
 * Reads the rich cells from xlsx-rich.js (theme-resolved colors, borders,
 * real column widths / row heights, fonts, alignment, merges, number formats)
 * and draws the sheet the way Excel prints it: one sheet after another.
 * 100% client-side. Plus a 29Court cover page and a footer on every page.
 * DOM-free; testable in Node with injected env.
 */
(function(){
"use strict";

/* Inlined at build time by build.py (site/logo-29court.png). */
var LOGO_B64 = "__LOGO_B64__";

var DEF_COL_W = 8.43;   /* Excel width units */
var DEF_ROW_H = 15;     /* points */
var DEF_MARGIN = 36;     /* default page margin (all sides) */
var FOOT = 30;          /* footer reserve */
var NAVY = [0.043, 0.055, 0.102];
var GOLD = [0.961, 0.773, 0.094];
var INK = [0.13, 0.12, 0.10];
var GRAY = [0.45, 0.44, 0.42];

function rgb(c){ return [c[0], c[1], c[2]]; }

/* Excel column-width units -> points. MDW = max digit width of Calibri 11. */
function colWidthToPt(w){
  var mdw = 7;
  var px = Math.floor((256 * w + Math.floor(128 / mdw)) / 256 * mdw);
  return px * 72 / 96;
}

var BSTYLE = {
  hair:   { w: 0.5 },
  thin:   { w: 0.75 },
  dotted: { w: 0.75, d: [1, 1.6] },
  dashed: { w: 0.75, d: [3.5, 2] },
  dashDot:{ w: 0.75, d: [3.5, 1.8, 1.2, 1.8] },
  dashDotDot: { w: 0.75, d: [3.5, 1.8, 1.2, 1.8, 1.2, 1.8] },
  medium: { w: 1.75 },
  mediumDashed: { w: 1.75, d: [4.5, 2.5] },
  mediumDashDot: { w: 1.75, d: [4.5, 2.5, 1.6, 2.5] },
  mediumDashDotDot: { w: 1.75, d: [4.5, 2.5, 1.6, 2.5, 1.6, 2.5] },
  thick:  { w: 2.5 },
  double: { w: 0.75, dbl: true },
  slantDashDot: { w: 0.75, d: [4.5, 1.8, 1.2, 1.8] }
};

function b64ToBytes(b64){
  /* Library-safe: atob in browsers, Buffer in Node. No DOM dependency. */
  if (typeof Buffer !== 'undefined'){
    var b = Buffer.from(b64, 'base64');
    return new Uint8Array(b.buffer, b.byteOffset, b.length);
  }
  var s = atob(b64), n = s.length, out = new Uint8Array(n);
  for (var i = 0; i < n; i++) out[i] = s.charCodeAt(i);
  return out;
}

/* Never crash on text the embedded font can't encode (CJK, RTL, emoji…):
 * Helvetica is WinAnsi-only, so degrade honestly with □ placeholders. */
function drawTextSafe(pg, PDFLib, text, opts){
  try {
    pg.drawText(text, opts);
  } catch(e){
    try {
      pg.drawText(String(text).replace(/[^\x00-\xFF]/g, '□'), opts);
    } catch(e2){ /* skip the run rather than failing the document */ }
  }
}

function pickFont(F, f){
  if (f.bold && f.italic) return F.bi;
  if (f.bold) return F.bold;
  if (f.italic) return F.italic;
  return F.reg;
}

/* ---- CJK text runs ----
 * Helvetica can't encode CJK; when F.cjk (Noto Sans JP, subset-embedded) is
 * loaded, text is split into Latin/CJK runs and each run draws in its own
 * font. Bold/italic CJK falls back to the regular CJK weight (documented).
 * When no CJK font is available, CJK runs hit the □-degradation, as before. */
var CJK_RUN_RE = /[\u3040-\u30FF\u4E00-\u9FFF]/;
function splitRuns(t){
  t = String(t);
  var runs = [], cur = '', cjk = false, started = false;
  for (var i = 0; i < t.length; i++){
    var cp = t.charCodeAt(i), ch = t[i];
    if (cp >= 0xD800 && cp <= 0xDBFF && i + 1 < t.length){
      var lo = t.charCodeAt(i + 1);
      if (lo >= 0xDC00 && lo <= 0xDFFF){
        ch = t.substr(i, 2); cp = 0x10000 + ((cp - 0xD800) << 10) + (lo - 0xDC00); i++;
      }
    }
    var isC = (cp >= 0x3040 && cp <= 0x30FF) || (cp >= 0x4E00 && cp <= 0x9FFF);
    if (!started){ cjk = isC; started = true; }
    if (isC !== cjk){ runs.push({ t: cur, cjk: cjk }); cur = ch; cjk = isC; }
    else cur += ch;
  }
  if (cur || !runs.length) runs.push({ t: cur, cjk: cjk });
  return runs;
}
function runFont(F, f, isCJK){
  if (isCJK && F.cjk) return F.cjk;
  return pickFont(F, f);
}
function runWidth(F, f, text, size){
  var runs = splitRuns(text), w = 0;
  for (var i = 0; i < runs.length; i++){
    var font = runFont(F, f, runs[i].cjk);
    try { w += font.widthOfTextAtSize(runs[i].t, size); }
    catch(e){ w += runs[i].t.length * size * 0.55; }
  }
  return w;
}
/* Draw text with per-run fonts; returns the total width. Never throws. */
function drawTextMixed(page, PDFLib, F, f, text, x, y, size, color){
  var runs = splitRuns(text), xx = x;
  for (var i = 0; i < runs.length; i++){
    var font = runFont(F, f, runs[i].cjk);
    try {
      page.drawText(runs[i].t, { x: xx, y: y, size: size, font: font, color: color });
    } catch(e){
      try {
        page.drawText(runs[i].t.replace(/[^\x00-\xFF]/g, '□'),
          { x: xx, y: y, size: size, font: font, color: color });
      } catch(e2){}
    }
    try { xx += font.widthOfTextAtSize(runs[i].t, size); }
    catch(e){ xx += runs[i].t.length * size * 0.55; }
  }
  return xx - x;
}

/* Conditional-formatting icon sets: small vector glyphs (arrows, traffic
 * lights). idx 0 = worst, top = best; colors follow Excel's defaults. */
function drawIconSet(page, PDFLib, spec, x, yMid, size){
  var n = spec.total || 3, idx = Math.max(0, Math.min(n - 1, spec.index || 0));
  var frac = n > 1 ? idx / (n - 1) : 0.5;
  var col = frac >= 0.66 ? [0.18, 0.65, 0.28]
    : frac >= 0.34 ? [0.95, 0.72, 0.15] : [0.88, 0.28, 0.22];
  var color = PDFLib.rgb(col[0], col[1], col[2]);
  var s = Math.max(2.5, size / 2), cx = x + s, cy = yMid;
  function tri(x1, y1, x2, y2, x3, y3){
    page.drawSvgPath('M ' + x1 + ' ' + y1 + ' L ' + x2 + ' ' + y2 +
      ' L ' + x3 + ' ' + y3 + ' Z', { color: color });
  }
  try {
    if (/TrafficLights/i.test(String(spec.set || ''))){
      page.drawCircle({ x: cx, y: cy, size: s * 0.9, color: color });
    } else if (frac >= 0.66){ tri(cx, cy + s, cx + s, cy - s, cx - s, cy - s); }
    else if (frac >= 0.34){ tri(cx + s, cy, cx - s, cy + s, cx - s, cy - s); }
    else { tri(cx, cy - s, cx + s, cy + s, cx - s, cy + s); }
  } catch(e){}
}

/* Draw one border side. x,yBot,w,h describe the cell rect. */
function drawBorderSide(page, PDFLib, side, spec, x, yBot, w, h){
  var st = BSTYLE[spec.style] || BSTYLE.thin;
  var col = spec.color || [0, 0, 0];
  var color = PDFLib.rgb(col[0], col[1], col[2]);
  function line(x1, y1, x2, y2){
    var o = { start: { x: x1, y: y1 }, end: { x: x2, y: y2 },
      thickness: st.w, color: color };
    if (st.d) o.dashArray = st.d;
    page.drawLine(o);
  }
  var off = st.dbl ? 1.6 : 0;
  function side2(x1, y1, x2, y2, dx, dy){
    line(x1, y1, x2, y2);
    if (st.dbl) line(x1 + dx * off, y1 + dy * off, x2 + dx * off, y2 + dy * off);
  }
  if (side === 'top') side2(x, yBot + h, x + w, yBot + h, 0, -1);
  else if (side === 'bottom') side2(x, yBot, x + w, yBot, 0, 1);
  else if (side === 'left') side2(x, yBot, x, yBot + h, 1, 0);
  else if (side === 'right') side2(x + w, yBot, x + w, yBot + h, -1, 0);
}

/* Cell text is clipped so it can never draw over a neighboring cell's
 * content. Excel-faithful rule: text overflows into adjacent EMPTY cells
 * (even across the tile's empty columns) but clips at the first occupied
 * cell. Returns { x, w } (an extended clip rect) or null for plain cell clip. */
function overflowClip(rows, ri, cj, tile, colWpt, k, rc, textW, baseX, baseW){
  var hAlign = (rc.align && rc.align.h) || 'general';
  if (hAlign === 'general') hAlign = rc.num ? 'right' : 'left';
  var rot = (rc.align && rc.align.rotation) || 0;
  if (rot || (rc.align && rc.align.wrap) || textW <= baseW) return null;
  var x0 = baseX, x1 = baseX + baseW, c, n, cw;
  var spanEnd = cj + (rc.span || 1) - 1;
  if (hAlign === 'right' || hAlign === 'center'){
    var needL = textW - baseW;
    for (c = cj - 1; c >= tile.c0 && needL > 0; c--){
      n = rows[ri][c];
      if (n && (n.t !== '' || n.skip)) break;
      cw = colWpt[c] * k; x0 -= cw; needL -= cw;
    }
  }
  if (hAlign === 'left' || hAlign === 'center'){
    var needR = textW - (x1 - x0);
    for (c = spanEnd + 1; c <= tile.c1 && needR > 0; c++){
      n = rows[ri][c];
      if (n && (n.t !== '' || n.skip)) break;
      cw = colWpt[c] * k; x1 += cw; needR -= cw;
    }
  }
  if (x0 === baseX && x1 === baseX + baseW) return null;
  return { x: x0, w: x1 - x0 };
}

/* Cell text is ALWAYS clipped to the cell rect. Excel clips text at the cell
 * edge when the neighbor is occupied; we clip unconditionally so a long
 * value can never draw on top of a neighboring cell. */
function drawCellText(page, PDFLib, PDF, F, rc, x, yBot, w, h, k, clip){
  var f = rc.font || {};
  var t = String(rc.t == null ? '' : rc.t);
  if (!t) return;
  var cx = clip ? clip.x : x, cw = clip ? clip.w : w;
  page.pushOperators(PDFLib.pushGraphicsState(),
    PDFLib.rectangle(cx, yBot, cw, h), PDFLib.clip(), PDFLib.endPath());
  try {
    drawCellTextInner(page, PDFLib, PDF, F, rc, f, t, x, yBot, w, h, k, cw);
  } finally {
    page.pushOperators(PDFLib.popGraphicsState());
  }
}

function drawCellTextInner(page, PDFLib, PDF, F, rc, f, t, x, yBot, w, h, k, effW){
  var font = pickFont(F, f);
  var size = (f.size || 11) * k;
  if (size < 4) size = 4;
  var col = f.color || INK;
  var color = PDFLib.rgb(col[0], col[1], col[2]);
  var pad = 3.5 * k;
  var ipad = (rc.iconPadU || 0) * k; /* icon-set glyphs reserve the left edge */
  var maxW = Math.max((effW || w) - pad * 2 - ipad, 8);
  var rot = (rc.align && rc.align.rotation) || 0;
  if (rot){
    /* Rotated text: pivot around the cell center so the text block stays
     * inside the cell, the way Excel centers rotated text. With CJK runs,
     * each run draws in its own font along the rotated baseline. */
    var rad = rot * Math.PI / 180, cos = Math.cos(rad), sin = Math.sin(rad);
    var totalW = runWidth(F, f, t, size);
    var runs = splitRuns(t);
    var ccx = x + w / 2, ccy = yBot + h / 2;
    var offY = size * 0.2; /* baseline drop below the text center, unrotated */
    var acc = 0, ri2;
    for (ri2 = 0; ri2 < runs.length; ri2++){
      var run = runs[ri2], rfont = runFont(F, f, run.cjk), rw;
      try { rw = rfont.widthOfTextAtSize(run.t, size); }
      catch(e){ rw = run.t.length * size * 0.55; }
      /* run start relative to the text center, then rotated into place */
      var sx = acc - totalW / 2;
      var dx = ccx + sx * cos + offY * sin, dy = ccy + sx * sin - offY * cos;
      var ro = { x: dx, y: dy, size: size, font: rfont,
        color: color, rotate: PDFLib.degrees(rot) };
      try { page.drawText(run.t, ro); }
      catch(e){
        try { page.drawText(run.t.replace(/[^\x00-\xFF]/g, '□'), ro); }
        catch(e2){}
      }
      acc += rw;
    }
    return;
  }
  var measFn = function(s){ return runWidth(F, f, s, size); };
  var lines = (rc.align && rc.align.wrap)
    ? PDF.wrap(t, font, size, maxW, measFn) : [t];
  /* Bounded shrink-to-fit: text a little too wide for its cell scales down
   * a touch so labels render whole — but never far enough to look like a
   * different font size than its siblings. Beyond SHRINK_MIN the text keeps
   * its true per-cell size and clips at the rect, the way Excel clips it
   * (faithful size always wins over fitting). Numbers keep hard clipping
   * (Excel shows ### for overflowing numbers). Wrapped text keeps its line
   * breaks (row height was measured for them); only the size shrinks. */
  var SHRINK_MIN = 0.8;
  if (!rc.num){
    var maxLine = 0, li;
    for (li = 0; li < lines.length; li++){
      var mlw = runWidth(F, f, lines[li], size);
      if (mlw > maxLine) maxLine = mlw;
    }
    if (maxLine > maxW && maxW > 4)
      size = Math.max(size * maxW / maxLine, size * SHRINK_MIN, 5);
  }
  var lh = size * 1.2;
  var blockH = lines.length * lh;
  var v = (rc.align && rc.align.v) || 'bottom';
  var blockTop;
  if (v === 'middle') blockTop = yBot + (h + blockH) / 2;
  else if (v === 'bottom') blockTop = yBot + pad + blockH;
  else blockTop = yBot + h - pad;
  var hAlign = (rc.align && rc.align.h) || 'general';
  /* Excel "General": numbers right, text left. */
  if (hAlign === 'general') hAlign = rc.num ? 'right' : 'left';
  var indent = ((rc.align && rc.align.indent) || 0) * 6 * k;
  lines.forEach(function(ln, li){
    var base = blockTop - size * 0.72 - li * lh;
    var lw = runWidth(F, f, ln, size);
    var x0;
    if (hAlign === 'center') x0 = x + (w - lw) / 2;
    else if (hAlign === 'right') x0 = x + w - pad - lw;
    else x0 = x + pad + indent + ipad;
    var tw = drawTextMixed(page, PDFLib, F, f, ln, x0, base, size, color);
    if (f.underline){
      page.drawLine({ start: { x: x0, y: base - 1.6 * k }, end: { x: x0 + tw, y: base - 1.6 * k },
        thickness: Math.max(0.5, size / 16), color: color });
    }
    if (f.strike){
      page.drawLine({ start: { x: x0, y: base + size * 0.3 }, end: { x: x0 + tw, y: base + size * 0.3 },
        thickness: Math.max(0.5, size / 18), color: color });
    }
  });
}

/* Row heights: explicit file heights win; otherwise fit wrapped content.
 * Rotated text sizes from its rotated bounding box (like Excel autofit).
 * Cells spanning multiple rows don't drive row height (the span covers them).
 * k = the tile's render scale: wrap widths and font sizes are measured at the
 * RENDERED size so text can never outgrow the row it was measured for. */
function computeRowHeights(sheet, colWpt, k, PDF, F){
  k = k || 1;
  var rows = sheet.rows, ncols = colWpt.length;
  var out = [];
  for (var r = 0; r < rows.length; r++){
    var explicit = sheet.rowHeights && sheet.rowHeights[r];
    if (typeof explicit === 'number' && explicit > 0){ out.push(explicit * k); continue; }
    var need = DEF_ROW_H * k;
    var row = rows[r];
    for (var c = 0; c < ncols; c++){
      var rc = row[c];
      if (!rc || rc.skip || !rc.t) continue;
      if ((rc.spanR || 1) > 1) continue; /* vertical span: rows size from other content */
      var size = (((rc.font && rc.font.size) || 11)) * k;
      var font = pickFont(F, rc.font || {});
      var sc = Math.min(rc.span || 1, ncols - c), mw = 0;
      for (var a = 0; a < sc; a++) mw += colWpt[c + a];
      mw = Math.max(mw * k - 7 * k, 8);
      var rot = (rc.align && rc.align.rotation) || 0;
      var measH = function(s){ return runWidth(F, rc.font || {}, s, size); };
      if (rot){
        var rad = rot * Math.PI / 180, lw = measH(String(rc.t));
        need = Math.max(need, Math.abs(lw * Math.sin(rad)) + size * Math.abs(Math.cos(rad)) + 7 * k);
      } else {
        var lines = (rc.align && rc.align.wrap)
          ? PDF.wrap(String(rc.t), font, size, mw, measH).length : 1;
        need = Math.max(need, lines * size * 1.2 + 7 * k);
      }
    }
    out.push(need);
  }
  return out;
}

/* Column index (0-based) -> Excel letters. */
function colLetter(i){
  var s = '';
  i++;
  while (i > 0){ var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
  return s;
}

/* Wide-sheet layout. Fit-to-width with sane, width-balanced horizontal tiles.
 * - Single tile when everything fits; scale UP to K_UP so short tables use
 *   the page instead of floating in whitespace.
 * - Otherwise split into the fewest tiles that keep every tile at >= K_FLOOR
 *   scale, with boundaries balanced by column width (never a stranded orphan).
 * - Boundaries snap off merged-cell interiors.
 * Each tile is { c0, c1, k } (0-based column range + scale).
 * Hard contract: every column lands in exactly one tile. Nothing is trimmed. */
/* Readability floor for fit-to-page-width: a sheet only goes on one page
 * when the fit scale stays at or above this. At FIT_FLOOR an 11pt body
 * renders at ~5.5pt — the smallest honestly readable print size. Below
 * it, the sheet tiles horizontally (Excel print-tiling) at >= K_FLOOR
 * scale instead of shrinking to ants. */
var FIT_FLOOR = 0.5;
var K_FLOOR = 0.75;  /* min readable scale; below this, tile instead of shrink */
var K_UP = 1.5;      /* max upscale so small tables fill the page */
function layoutTiles(colWpt, usableW, mergeIvs){
  var n = colWpt.length;
  function wsum(a, b){ var s = 0; for (var i = a; i <= b; i++) s += colWpt[i]; return s; }
  var total = wsum(0, n - 1);
  if (total <= usableW)
    return [{ c0: 0, c1: n - 1, k: Math.min(usableW / total, K_UP) }];
  var nTiles = Math.ceil(total / (usableW / K_FLOOR));
  if (nTiles < 1) nTiles = 1;  /* fit-to-width in [K_FLOOR, 1): single tile */
  if (nTiles > n) nTiles = n;
  function cutMerge(b){
    for (var i = 0; i < (mergeIvs || []).length; i++){
      var m0 = mergeIvs[i][0], m1 = mergeIvs[i][1];
      if (m0 < b && b <= m1) return [m0, m1];
    }
    return null;
  }
  /* prefix sums for straddle-optimal boundary picks */
  var pref = [0];
  for (var pc = 0; pc < n; pc++) pref.push(pref[pc] + colWpt[pc]);
  function split(nT){
    var bounds = [0], t, c, b;
    for (t = 1; t < nT; t++){
      var target = total * t / nT;
      var maxC = n - (nT - t);  /* last feasible boundary: leave 1 col/tile */
      var best = bounds[t - 1] + 1, bestD = Infinity;
      for (c = bounds[t - 1] + 1; c <= maxC; c++){
        var dd = Math.abs(pref[c] - target);
        if (dd < bestD){ bestD = dd; best = c; }
      }
      b = best;
      var hit = cutMerge(b);  /* snap off merge interiors -> nearest edge */
      if (hit) b = (b - hit[0] <= hit[1] + 1 - b) ? hit[0] : hit[1] + 1;
      if (b < bounds[t - 1] + 1) b = bounds[t - 1] + 1;
      if (b > maxC) b = maxC;
      bounds.push(b);
    }
    bounds.push(n);
    var clean = [bounds[0]];  /* drop any collapsed (empty) tiles */
    for (t = 1; t < bounds.length; t++)
      if (bounds[t] > clean[clean.length - 1]) clean.push(bounds[t]);
    var out = [];
    for (t = 0; t < clean.length - 1; t++){
      var c0 = clean[t], c1 = clean[t + 1] - 1;
      out.push({ c0: c0, c1: c1, k: Math.min(usableW / wsum(c0, c1), K_UP) });
    }
    return out;
  }
  /* repair: if the discrete straddle leaves a tile under the readability
   * floor, add a tile and re-split (bounded: nTiles only grows). */
  var tiles = split(nTiles);
  for (var att = 0; att < n; att++){
    var minK = Math.min.apply(null, tiles.map(function(x){ return x.k; }));
    if (minK >= K_FLOOR - 1e-9 || nTiles >= n) break;
    nTiles++;
    tiles = split(nTiles);
  }
  /* coverage fallback: if anything got lost, single tile (never trim) */
  var seen = {}, ok = tiles.length > 0;
  tiles.forEach(function(tl){ for (var q = tl.c0; q <= tl.c1; q++) seen[q] = 1; });
  for (var q2 = 0; q2 < n; q2++) if (!seen[q2]) ok = false;
  if (!ok){
    if (typeof console !== 'undefined' && console.warn)
      console.warn('layoutTiles: coverage fallback, single tile');
    return [{ c0: 0, c1: n - 1, k: usableW / total }];
  }
  return tiles;
}

/* Column intervals occupied by horizontal merges, for tile-boundary snapping. */
function mergeColIntervals(rows){
  var ivs = [];
  for (var r = 0; r < rows.length; r++){
    var row = rows[r];
    for (var c = 0; c < row.length; c++){
      var rc = row[c];
      if (rc && !rc.skip && (rc.span || 1) > 1) ivs.push([c, c + (rc.span || 1) - 1]);
    }
  }
  return ivs;
}

function drawSheetHeader(page, PDFLib, F, name, rows, cols, y, PW, mg){
  mg = mg || { l: DEF_MARGIN, r: DEF_MARGIN };
  drawTextMixed(page, PDFLib, F, { bold: true }, String(name), mg.l, y - 13, 13,
    PDFLib.rgb(NAVY[0], NAVY[1], NAVY[2]));
  var sub = rows + ' rows x ' + cols + ' columns';
  var w = runWidth(F, {}, sub, 9);
  drawTextMixed(page, PDFLib, F, {}, sub, PW - mg.r - w, y - 12, 9,
    PDFLib.rgb(GRAY[0], GRAY[1], GRAY[2]));
  page.drawLine({ start: { x: mg.l, y: y - 20 }, end: { x: PW - mg.r, y: y - 20 },
    thickness: 1, color: PDFLib.rgb(GOLD[0], GOLD[1], GOLD[2]) });
  return y - 30;
}

/* ---- Print headers/footers (Excel oddHeader/oddFooter et al.) ----
 * Parses Excel header/footer code strings (&L &C &R sections; &"Font,Bold"
 * and &size; &B &I &U &S toggles; &KRRGGBB color; &P page; &N pages; &D date;
 * &T time; &F filename; &Z path; &A sheet name; &[Picture]; && literal &)
 * into styled runs, and draws them in the page header/footer zones.
 * &[Picture] indices count in canonical key order (oddHeader, oddFooter,
 * evenHeader, evenFooter, firstHeader, firstFooter) against the sheet's
 * VML-ordered printHF.images. Unknown codes are kept as literal text. */
var HF_KEYS = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter',
  'firstHeader', 'firstFooter'];

function hfClamp01(x){ return x < 0 ? 0 : x > 1 ? 1 : x; }

function hfHexColor(hex){
  var v = parseInt(hex, 16);
  return [((v >> 16) & 255) / 255, ((v >> 8) & 255) / 255, (v & 255) / 255];
}
function hfCountPics(str){
  var m = String(str || '').match(/&\[Picture\]/g);
  return m ? m.length : 0;
}
function hfPicStart(hf, key){
  var n = 0;
  for (var i = 0; i < HF_KEYS.length; i++){
    if (HF_KEYS[i] === key) break;
    n += hfCountPics(hf[HF_KEYS[i]]);
  }
  return n;
}

/* Parse one header/footer string into {left:[],center:[],right:[]} of runs.
 * run = {t,b,i,u,s,size,color,pic,dyn}; pic = picture index (-1 = none).
 * lockSection: &L/&C/&R codes are skipped (for custom {left,center,right}). */
function parseHF(str, picStart, lockSection){
  var secs = { left: [], center: [], right: [] };
  var cur = 'center', picN = picStart || 0;
  var st = { b: false, i: false, u: false, s: false, size: 0, color: null };
  function newRun(){
    return { t: '', b: st.b, i: st.i, u: st.u, s: st.s, size: st.size,
      color: st.color, pic: -1, dyn: null };
  }
  var run = newRun();
  function push(){
    if (run.t || run.pic >= 0 || run.dyn) secs[cur].push(run);
    run = newRun();
  }
  str = String(str || '');
  var re = /&(&|"([^"]*)"|\[Picture\]|K([0-9A-Fa-f]{6})|(\d{1,3})|([A-Za-z])|(.))/g;
  var last = 0, m;
  while ((m = re.exec(str))){
    if (m.index > last) run.t += str.slice(last, m.index);
    var full = m[0];
    if (full === '&&'){ run.t += '&'; }
    else if (m[2] !== undefined){
      push();
      var rest = (m[2].split(',')[1] || '');
      if (/bold/i.test(rest)) st.b = true;
      if (/italic/i.test(rest)) st.i = true;
      if (/regular/i.test(rest)){ st.b = false; st.i = false; }
      var dm = rest.match(/(\d{1,3})/);
      if (dm) st.size = parseInt(dm[1], 10);
      run = newRun();
    }
    else if (full === '&[Picture]'){ push(); run.pic = picN++; push(); }
    else if (m[3] !== undefined){ push(); st.color = hfHexColor(m[3]); run = newRun(); }
    else if (m[4] !== undefined){ push(); st.size = parseInt(m[4], 10); run = newRun(); }
    else if (m[5] !== undefined){
      var L = m[5].toUpperCase();
      if (L === 'L' || L === 'C' || L === 'R'){
        if (!lockSection){
          push(); cur = L === 'L' ? 'left' : (L === 'C' ? 'center' : 'right');
          run = newRun();
        }
        /* locked: section codes inside a custom section string are skipped */
      }
      else if (L === 'B'){ push(); st.b = !st.b; run = newRun(); }
      else if (L === 'I'){ push(); st.i = !st.i; run = newRun(); }
      else if (L === 'U'){ push(); st.u = !st.u; run = newRun(); }
      else if (L === 'S'){ push(); st.s = !st.s; run = newRun(); }
      else if ('PNDTFZA'.indexOf(L) >= 0){ push(); run.dyn = L; push(); }
      else run.t += full;
    }
    else run.t += full;
    last = re.lastIndex;
  }
  if (last < str.length) run.t += str.slice(last);
  push();
  return secs;
}

function hfHasContent(secs){
  return ['left', 'center', 'right'].some(function(k){
    return (secs[k] || []).some(function(r){ return r.t || r.pic >= 0 || r.dyn; });
  });
}

function hfDynText(d, ctx){
  switch (d){
    case 'P': return String(ctx.page);
    case 'N': return String(ctx.pages);
    case 'D': return ctx.dateStr || '';
    case 'T': return ctx.timeStr || '';
    case 'F': return ctx.filename || '';
    case 'Z': return '';
    case 'A': return ctx.sheetName || '';
  }
  return '';
}

/* Resolve the header/footer spec for one page.
 * Returns null (draw nothing), 'DEFAULT' (the built-in footer), or
 * { secs, images } with images = embedded PDFImage objects (or nulls).
 * ctx: {page, pages, dateStr, timeStr, filename, sheetName, isFirstOfSheet,
 *   mL, mR, PW}. opts: {header, footer, useWorkbookHeaders}. */
function pickHF(sheet, kind, ctx, opts){
  var opt = kind === 'header' ? opts.header : opts.footer;
  if (opt == null) opt = 'auto';
  if (opt === 'none') return null;
  if (typeof opt === 'object'){
    var secs = { left: [], center: [], right: [] };
    ['left', 'center', 'right'].forEach(function(k){
      var p = parseHF(String(opt[k] || ''), 0, true);
      secs[k] = p.left.concat(p.center, p.right);
    });
    return { secs: secs, images: [] };
  }
  if (opt === 'default') return kind === 'footer' ? 'DEFAULT' : null;
  /* 'auto': the workbook's own print header/footer wins when the file
   * defines one; otherwise the built-in default footer (header: nothing). */
  var hf = sheet && sheet.printHF;
  if (opts.useWorkbookHeaders !== false && hf){
    var key;
    if (kind === 'header'){
      key = (hf.differentFirst && ctx.isFirstOfSheet && hf.firstHeader) ? 'firstHeader'
        : (hf.differentOddEven && ctx.page % 2 === 0 && hf.evenHeader) ? 'evenHeader'
        : 'oddHeader';
    } else {
      key = (hf.differentFirst && ctx.isFirstOfSheet && hf.firstFooter) ? 'firstFooter'
        : (hf.differentOddEven && ctx.page % 2 === 0 && hf.evenFooter) ? 'evenFooter'
        : 'oddFooter';
    }
    if (hf[key]){
      var parsed = parseHF(hf[key], hfPicStart(hf, key), false);
      if (hfHasContent(parsed))
        return { secs: parsed, images: (sheet._hfEmbedded || []) };
    }
  }
  return kind === 'footer' ? 'DEFAULT' : null;
}

/* Split parsed sections into per-section line lists of drawable items.
 * item = {text, run, pic:-1} | {pic, run}. */
function hfLines(secs, ctx){
  var out = {};
  ['left', 'center', 'right'].forEach(function(k){
    var lines = [], cur = [];
    (secs[k] || []).forEach(function(r){
      if (r.pic >= 0){ cur.push({ pic: r.pic, run: r }); return; }
      var t = r.dyn ? hfDynText(r.dyn, ctx) : r.t;
      String(t).split('\n').forEach(function(part, li, arr){
        if (part) cur.push({ text: part, run: r, pic: -1 });
        if (li < arr.length - 1){ lines.push(cur); cur = []; }
      });
    });
    lines.push(cur);
    while (lines.length > 1 && !lines[lines.length - 1].length) lines.pop();
    out[k] = lines;
  });
  return out;
}

function hfRunFlags(r, style){
  var sf = String(style.font || '');
  return { bold: /bold/.test(sf) || !!r.b, italic: /italic/.test(sf) || !!r.i };
}
function hfRunFont(F, r, style){
  if (/mono/.test(String(style.font || '')) && F.mono) return F.mono;
  return pickFont(F, hfRunFlags(r, style));
}
/* CJK-aware width using the same font choice the draw pass uses. */
function hfTextWidth(F, it, size){
  var runs = splitRuns(it.text), w = 0;
  for (var i = 0; i < runs.length; i++){
    var font = (runs[i].cjk && F.cjk) ? F.cjk : it._font;
    try { w += font.widthOfTextAtSize(runs[i].t, size); }
    catch(e){ w += runs[i].t.length * size * 0.55; }
  }
  return w;
}
function hfDrawText(page, F, it, x, y, size, color){
  var runs = splitRuns(it.text), xx = x, w = 0;
  for (var i = 0; i < runs.length; i++){
    var font = (runs[i].cjk && F.cjk) ? F.cjk : it._font;
    try {
      page.drawText(runs[i].t, { x: xx, y: y, size: size, font: font, color: color });
    } catch(e){
      try {
        page.drawText(runs[i].t.replace(/[^\x00-\xFF]/g, '□'),
          { x: xx, y: y, size: size, font: font, color: color });
      } catch(e2){}
    }
    var adv = 0;
    try { adv = font.widthOfTextAtSize(runs[i].t, size); }
    catch(e){ adv = runs[i].t.length * size * 0.55; }
    xx += adv; w += adv;
  }
  return w;
}
function hfSumW(items){
  return items.reduce(function(a, it){ return a + (it._w || 0); }, 0);
}
function hfDrawItems(page, PDFLib, F, items, x, baseY, style){
  var xx = x;
  items.forEach(function(it){
    var r = it.run;
    if (it.pic >= 0){
      if (it._emb && it._w > 0){
        var ph = style.pictureHeight;
        page.drawImage(it._emb, { x: xx, y: baseY - (ph - it._size) / 2,
          width: it._w, height: ph });
      }
      xx += it._w + (it._w ? 4 : 0);
      return;
    }
    var col = r.color || style.color;
    var color = PDFLib.rgb(col[0], col[1], col[2]);
    var w = hfDrawText(page, F, it, xx, baseY, it._size, color);
    if (r.u || r.s){
      var ly = baseY + (r.s ? it._size * 0.32 : -1);
      page.drawLine({ start: { x: xx, y: ly }, end: { x: xx + w, y: ly },
        thickness: Math.max(0.5, it._size / 16), color: color });
    }
    xx += w;
  });
}

/* Height a header/footer block needs (for the layout reserve). */
function hfBlockHeight(spec, style){
  var lines = hfLines(spec.secs, { page: 1, pages: 1, dateStr: '', timeStr: '',
    filename: '', sheetName: '', isFirstOfSheet: false });
  var n = Math.max(lines.left.length, lines.center.length, lines.right.length);
  var h = 0;
  for (var li = 0; li < n; li++){
    var lh = 0;
    [lines.left[li] || [], lines.center[li] || [], lines.right[li] || []]
      .forEach(function(items){
        items.forEach(function(it){
          var s = it.pic >= 0 ? style.pictureHeight
            : ((it.run.size || style.size) * 1.25);
          if (s > lh) lh = s;
        });
      });
    h += lh || style.size * 1.25;
  }
  return h > 0 ? h + 6 : 0;
}

/* Draw a header/footer block top-down within [yBot, yTop]. */
function drawHFBlock(page, PDFLib, F, spec, yTop, yBot, style, ctx){
  var lines = hfLines(spec.secs, ctx);
  var images = spec.images || [];
  var n = Math.max(lines.left.length, lines.center.length, lines.right.length);
  var y = yTop;
  for (var li = 0; li < n; li++){
    var L = lines.left[li] || [], C = lines.center[li] || [], R = lines.right[li] || [];
    [L, C, R].forEach(function(items){
      items.forEach(function(it){
        var r = it.run;
        it._size = r.size || style.size;
        if (it.pic >= 0){
          var emb = images[it.pic] || null;
          it._emb = emb;
          var ph = style.pictureHeight;
          it._w = (emb && emb.width) ? ph * emb.width / emb.height : 0;
        } else {
          it._font = hfRunFont(F, r, style);
          it._w = hfTextWidth(F, it, it._size);
        }
      });
    });
    var lh = 0;
    [L, C, R].forEach(function(items){
      items.forEach(function(it){
        var h = it.pic >= 0 ? style.pictureHeight : it._size * 1.25;
        if (h > lh) lh = h;
      });
    });
    if (!lh) lh = style.size * 1.25;
    if (y - lh < yBot - 0.5 && li > 0) break; /* never draw below the zone */
    var baseY = y - lh * 0.82;
    hfDrawItems(page, PDFLib, F, L, ctx.mL, baseY, style);
    hfDrawItems(page, PDFLib, F, C, ctx.PW / 2 - hfSumW(C) / 2, baseY, style);
    hfDrawItems(page, PDFLib, F, R, ctx.PW - ctx.mR - hfSumW(R), baseY, style);
    y -= lh;
  }
}

function normalizeHFColor(c, fb){
  if (Array.isArray(c) && c.length >= 3)
    return [hfClamp01(+c[0]), hfClamp01(+c[1]), hfClamp01(+c[2])];
  if (typeof c === 'string'){
    var m = c.replace(/^#/, '').match(/^([0-9a-fA-F]{6})$/);
    if (m) return hfHexColor(m[1]);
  }
  return fb;
}
function normalizeHFStyle(s, fb){
  s = s || {};
  return {
    font: String(s.font != null ? s.font : fb.font).toLowerCase(),
    size: s.size != null ? +s.size : fb.size,
    color: normalizeHFColor(s.color, fb.color),
    rule: s.rule != null ? !!s.rule : fb.rule,
    ruleColor: normalizeHFColor(s.ruleColor, fb.ruleColor),
    pictureHeight: s.pictureHeight != null ? +s.pictureHeight : fb.pictureHeight
  };
}
function normalizeMargins(m){
  var o = { t: DEF_MARGIN, r: DEF_MARGIN, b: DEF_MARGIN, l: DEF_MARGIN };
  if (typeof m === 'number' && isFinite(m)){
    o = { t: m, r: m, b: m, l: m };
  } else if (m){
    if (m.top != null) o.t = +m.top;
    if (m.right != null) o.r = +m.right;
    if (m.bottom != null) o.b = +m.bottom;
    if (m.left != null) o.l = +m.left;
  }
  return o;
}

/* The built-in default footer: filename, date, Page N of M. Unchanged from
 * the original design; used when the workbook defines no print footer. */
function drawDefaultFooter(pg, PDFLib, F, i, total, pW, mg, logoImg, showLogo,
    filename, dateStr){
  var fy = mg.b + 6;
  pg.drawLine({ start: { x: mg.l, y: fy + 18 }, end: { x: pW - mg.r, y: fy + 18 },
    thickness: 0.75, color: PDFLib.rgb(GOLD[0], GOLD[1], GOLD[2]) });
  var tx = mg.l;
  if (showLogo && logoImg){
    pg.drawImage(logoImg, { x: mg.l, y: fy - 2, width: 18, height: 18 });
    tx = mg.l + 24;
  }
  var left = filename + '  ·  ' + dateStr;
  drawTextMixed(pg, PDFLib, F, {}, left.slice(0, 90), tx, fy + 2, 8,
    PDFLib.rgb(GRAY[0], GRAY[1], GRAY[2]));
  var num = 'Page ' + (i + 1) + ' of ' + total;
  var nw = runWidth(F, {}, num, 8);
  drawTextMixed(pg, PDFLib, F, {}, num, pW - mg.r - nw, fy + 2, 8,
    PDFLib.rgb(GRAY[0], GRAY[1], GRAY[2]));
}

/* Max header height over the applicable variants, so the layout reserve
 * fits the tallest header this sheet can show (odd/even/first). */
function headerReserve(s, meta, headerStyle, base){
  var combos;
  if ((meta.header == null || meta.header === 'auto') &&
      meta.useWorkbookHeaders !== false && s.printHF){
    combos = [];
    [true, false].forEach(function(first){
      [1, 2].forEach(function(p){ combos.push({ isFirstOfSheet: first, page: p }); });
    });
  } else {
    combos = [{ isFirstOfSheet: true, page: 1 }];
  }
  var h = 0;
  combos.forEach(function(cb){
    var ctx = Object.assign({}, base, cb);
    var spec = pickHF(s, 'header', ctx, meta);
    if (spec && spec !== 'DEFAULT')
      h = Math.max(h, hfBlockHeight(spec, headerStyle));
  });
  return h;
}

async function sheetsToPdf(sheets, meta, env){
  var PDFLib = env.PDFLib, PDF = env.PDF, XlsxCharts = env.XlsxCharts;
  meta = meta || {};
  /* ---- sheet selection + ordering ----
   * includeSheets filters; sheetOrder reorders (unlisted sheets keep their
   * relative order after the listed ones). */
  if (meta.includeSheets && meta.includeSheets.length)
    sheets = sheets.filter(function(s){ return meta.includeSheets.indexOf(s.name) !== -1; });
  if (meta.sheetOrder && meta.sheetOrder.length){
    var sRank = {};
    meta.sheetOrder.forEach(function(n, i){ if (!(n in sRank)) sRank[n] = i; });
    sheets = sheets.slice().sort(function(a, b){
      var ra = (a.name in sRank) ? sRank[a.name] : 1e9;
      var rb = (b.name in sRank) ? sRank[b.name] : 1e9;
      return ra - rb;
    });
  }
  /* ---- styling options: every visual decision is overridable ----
   * All defaults reproduce the site's exact behavior (Victor's BOMA
   * screenshots): pageSize 'a4'|'letter'|[w,h], orientation
   * 'auto'|'portrait'|'landscape', margins {top,right,bottom,left} in
   * points, cover false skips the branded cover, header/footer 'auto'
   * honors the workbook's own print headers/footers when defined. */
  var PAGE = (function(){
    var ps = meta.pageSize;
    if (Array.isArray(ps) && ps.length === 2 && +ps[0] > 0 && +ps[1] > 0)
      return [+ps[0], +ps[1]];
    return (String(ps || 'a4').toLowerCase() === 'letter' && PDF.LETTER)
      ? PDF.LETTER : PDF.A4;
  })();
  var MG = normalizeMargins(meta.margins);
  var wantCover = meta.cover !== false;
  var coverTitle = meta.coverTitle == null ? '29COURT' : String(meta.coverTitle);
  var coverTagline = meta.coverTagline == null
    ? 'Made with 29Court Convert. Your file never left this device.'
    : String(meta.coverTagline);
  var showCoverLogo = meta.logoCover !== false;
  var showFooterLogo = meta.logoFooter !== false;
  var headerStyle = normalizeHFStyle(meta.headerStyle,
    { font: 'reg', size: 10, color: [0.13, 0.12, 0.10], rule: false,
      ruleColor: GOLD, pictureHeight: 22 });
  var footerStyle = normalizeHFStyle(meta.footerStyle,
    { font: 'reg', size: 8, color: GRAY, rule: true,
      ruleColor: GOLD, pictureHeight: 16 });
  var FOOTR = (meta.footer === 'none') ? 0 : FOOT;
  var forceOrient = String(meta.orientation || 'auto').toLowerCase();
  if (forceOrient !== 'portrait' && forceOrient !== 'landscape') forceOrient = 'auto';
  var d = await PDF.doc(), F0 = await PDF.loadFonts(d);
  /* Per-document font set: loadFonts caches at module level, so clone
   * before attaching the CJK font — otherwise cjk leaks into the cache
   * and later documents falsely report cjkEmbedded. */
  var F = Object.assign({}, F0);
  /* CJK: if any cell text needs it, subset-embed Noto Sans JP. Never fatal:
   * without it, CJK degrades to □ placeholders as before. */
  try {
    if (PDF.hasCJKText && PDF.hasCJKText(sheets)) F.cjk = await PDF.loadCJK(d, env);
  } catch(e){ F.cjk = null; }
  if (meta.stats) meta.stats.cjkEmbedded = !!F.cjk;
  var logoImg = null;
  try {
    /* Branding is caller-supplied (meta.logoB64), never baked into the
     * library: the core renders a clean document when no logo is given. */
    var logoB64 = meta.logoB64 || LOGO_B64;
    if (logoB64 && logoB64.indexOf('__LOGO') !== 0)
      logoImg = await d.embedPng(b64ToBytes(logoB64));
  } catch(e){ logoImg = null; }
  var filename = meta.filename || 'workbook.xlsx';
  var dateStr = meta.dateStr || '';
  var timeStr = '';
  try { timeStr = new Date().toLocaleTimeString(); } catch(e){}

  /* ---------- cover page ---------- */
  var PW = PAGE[0], PH = PAGE[1];
  if (wantCover){
    var cover = d.addPage([PW, PH]);
    cover.drawRectangle({ x: 0, y: 0, width: PW, height: PH,
      color: PDFLib.rgb(NAVY[0], NAVY[1], NAVY[2]) });
    var cy = PH - 120;
    if (logoImg && showCoverLogo){
      var lw = 150, lh = 150 * logoImg.height / logoImg.width;
      cover.drawImage(logoImg, { x: (PW - lw) / 2, y: cy - lh, width: lw, height: lh });
      cy -= lh + 34;
    }
    function centered(text, size, font, color, yy){
      var w;
      try { w = font.widthOfTextAtSize(text, size); } catch(e){ w = text.length * size * 0.6; }
      drawTextSafe(cover, PDFLib, text, { x: (PW - w) / 2, y: yy, size: size, font: font, color: color });
      return yy;
    }
    if (coverTitle)
      cy = centered(coverTitle, 30, F.bold, PDFLib.rgb(GOLD[0], GOLD[1], GOLD[2]), cy) - 44;
    cy = centered(String(filename), 19, F.bold, PDFLib.rgb(1, 1, 1), cy) - 30;
    cy = centered('Converted ' + dateStr, 11, F.reg, PDFLib.rgb(0.62, 0.65, 0.75), cy) - 44;
    cover.drawLine({ start: { x: PW / 2 - 60, y: cy }, end: { x: PW / 2 + 60, y: cy },
      thickness: 1.5, color: PDFLib.rgb(GOLD[0], GOLD[1], GOLD[2]) });
    cy -= 40;
    cy = centered('SHEETS', 10, F.bold, PDFLib.rgb(GOLD[0], GOLD[1], GOLD[2]), cy) - 26;
    sheets.forEach(function(s){
      var ncols = s.rows.reduce(function(a, r){ return Math.max(a, r.length); }, 0);
      var line = s.name + '  -  ' + s.rows.length + ' rows x ' + ncols + ' columns';
      cy = centered(line, 11, F.reg, PDFLib.rgb(0.92, 0.93, 0.96), cy) - 22;
    });
    cy -= 30;
    if (coverTagline)
      centered(coverTagline, 9,
        F.reg, PDFLib.rgb(0.55, 0.58, 0.68), 70);
  }

  /* ---------- sheets, one after another ---------- */
  /* Pre-embed sheet images (async) so the synchronous draw loops below can
   * place them. Images that fail to embed are skipped, never fatal. */
  for (var si = 0; si < sheets.length; si++){
    var simgs = sheets[si].images || [];
    for (var ii = 0; ii < simgs.length; ii++){
      var sim = simgs[ii];
      try {
        sim.embedded = (sim.ext === 'jpg' || sim.ext === 'jpeg')
          ? await d.embedJpg(sim.buf) : await d.embedPng(sim.buf);
      } catch(e){ sim.embedded = null; }
    }
    /* Print header/footer &[Picture] images (VML drawings), in the same
     * canonical order the header/footer parser assigns indices. */
    var hfImgs = (sheets[si].printHF && sheets[si].printHF.images) || [];
    sheets[si]._hfEmbedded = [];
    for (var hi = 0; hi < hfImgs.length; hi++){
      try {
        sheets[si]._hfEmbedded[hi] = (hfImgs[hi].ext === 'jpg' || hfImgs[hi].ext === 'jpeg')
          ? await d.embedJpg(hfImgs[hi].data) : await d.embedPng(hfImgs[hi].data);
      } catch(e){ sheets[si]._hfEmbedded[hi] = null; }
    }
  }
  sheets.forEach(function(s){
    var rows = s.rows;
    var ncols = rows.reduce(function(a, r){ return Math.max(a, r.length); }, 0);
    if (!ncols) return;
    var colWpt = [];
    for (var c = 0; c < ncols; c++)
      colWpt.push(colWidthToPt((s.colWidths && s.colWidths[c]) || DEF_COL_W));
    var totalW = colWpt.reduce(function(a, b){ return a + b; }, 0);
    var cW = PAGE[0] - MG.l - MG.r, cWL = PAGE[1] - MG.l - MG.r;
    /* Fit-to-page-width (Excel "Fit to 1 page wide"): a wide sheet goes on
     * ONE page, in the orientation that renders the largest type, as long
     * as the fit scale stays at or above FIT_FLOOR. Horizontal tiling
     * ("cols A-K, 1 of 2") only kicks in when even landscape fit-to-width
     * would drop below the floor. meta.orientation forces one orientation. */
    var tiles, landscape;
    var kP = cW / totalW, kL = cWL / totalW;
    if (forceOrient === 'portrait' || forceOrient === 'landscape'){
      landscape = (forceOrient === 'landscape');
      var kf = landscape ? kL : kP, uw = landscape ? cWL : cW;
      tiles = (kf >= FIT_FLOOR)
        ? [{ c0: 0, c1: ncols - 1, k: Math.min(kf, K_UP) }]
        : layoutTiles(colWpt, uw, mergeColIntervals(rows));
    } else if (kP >= 1){
      landscape = false;
      tiles = [{ c0: 0, c1: ncols - 1, k: Math.min(kP, K_UP) }];
    } else if (kL >= FIT_FLOOR){
      landscape = true;
      tiles = [{ c0: 0, c1: ncols - 1, k: kL }];
    } else if (kP >= FIT_FLOOR){
      landscape = false;
      tiles = [{ c0: 0, c1: ncols - 1, k: kP }];
    } else {
      /* Below the readability floor: width-balanced horizontal tiles
       * (Excel print-tiling), each at a readable scale, instead of
       * shrinking the sheet into illegibility. */
      landscape = totalW > cW;
      var usableW = landscape ? cWL : cW;
      tiles = layoutTiles(colWpt, usableW, mergeColIntervals(rows));
    }
    PW = landscape ? PAGE[1] : PAGE[0];
    PH = landscape ? PAGE[0] : PAGE[1];
    /* Print header (the workbook's own, or meta.header): reserve its height
     * above the sheet content on every page of this sheet. Sized from the
     * tallest applicable variant (odd/even/first) so nothing overlaps. */
    var headH = headerReserve(s, meta, headerStyle, { pages: 1,
      dateStr: dateStr, timeStr: timeStr, filename: filename,
      sheetName: s.name, mL: MG.l, mR: MG.r, PW: PW });
    var topY = PH - MG.t - headH;
    /* Row heights per tile at the tile's own scale (wrap is measured at the
     * rendered width, so text can never outgrow its row). */
    tiles.forEach(function(tl){
      tl.rowH = computeRowHeights({ rows: rows, rowHeights: s.rowHeights }, colWpt, tl.k, PDF, F);
    });
    /* Stack short tiles vertically on one page when they all fit, so an
     * 8-row table doesn't float in a sea of whitespace. */
    function blockH(tl){
      return 30 + tl.rowH.reduce(function(a, b){ return a + b; }, 0);
    }
    var availH = (PH - MG.t) - (MG.b + FOOTR);
    var stackNeed = 0;
    tiles.forEach(function(tl){ stackNeed += blockH(tl); });
    var stacked = tiles.length > 1 && stackNeed + (tiles.length - 1) * 22 <= availH + 1; /* +1: float safety */
    var page = null, y = 0, sheetFirstDone = false;
    tiles.forEach(function(tile, ti){
      var k = tile.k;
      var rowH = tile.rowH;
      var colX = [], x = MG.l;
      for (var ci = tile.c0; ci <= tile.c1; ci++){ colX.push(x); x += colWpt[ci] * k; }
      tile.colX = colX; /* retained for chart anchor placement */
      tile.rowPos = tile.rowPos || {};
      var tileName = s.name + (tiles.length > 1
        ? ' (cols ' + colLetter(tile.c0) + '-' + colLetter(tile.c1) + ', ' + (ti + 1) + ' of ' + tiles.length + ')'
        : '');
      var ncolsThis = tile.c1 - tile.c0 + 1;
      if (!page || !stacked || y - 22 - blockH(tile) < MG.b + FOOTR - 2){
        page = d.addPage([PW, PH]); y = topY;
        /* Tag the page for the header/footer pass: which sheet it belongs
         * to, and whether it is the sheet's first page (differentFirst). */
        page._hfSheet = s;
        page._hfFirst = !sheetFirstDone;
        sheetFirstDone = true;
      } else {
        y -= 22; /* stacked tile gap */
      }
      y = drawSheetHeader(page, PDFLib, F, tileName, rows.length, ncolsThis, y, PW, MG);
      if (!rows.length){
        drawTextSafe(page, PDFLib, '(empty sheet)', { x: MG.l, y: y - 12, size: 10,
          font: F.italic, color: PDFLib.rgb(GRAY[0], GRAY[1], GRAY[2]) });
        return;
      }
      /* merged rect, clamped to this tile's column range */
      function mergedRect(ri, ci){
        var rc = rows[ri][ci];
        if (!rc || rc.skip) return null;
        var w = 0, h = 0, sc = rc.span || 1, sr = rc.spanR || 1;
        var cEnd = Math.min(ci + sc - 1, tile.c1);
        for (var a = ci; a <= cEnd; a++) w += colWpt[a] * k;
        for (var b = 0; b < sr && ri + b < rowH.length; b++) h += rowH[ri + b];
        return { w: w, h: h };
      }
      for (var ri = 0; ri < rows.length; ri++){
        var rh = rowH[ri];
        if (y - rh < MG.b + FOOTR){
          page = d.addPage([PW, PH]); y = topY;
          page._hfSheet = s; page._hfFirst = false;
          y = drawSheetHeader(page, PDFLib, F, tileName + ' (continued)', rows.length, (tile.c1 - tile.c0 + 1), y, PW, MG);
        }
        tile.rowPos[ri] = { page: page, yTop: y };
        /* Embedded images anchored at this row, inside this tile's columns. */
        (s.images || []).forEach(function(img){
          if (!img.embedded || img.r0 !== ri || img.c0 < tile.c0 || img.c0 > tile.c1) return;
          var ix = colX[img.c0 - tile.c0], iw = 0;
          for (var ic = img.c0; ic < img.c1 && ic < colWpt.length; ic++){
            var frac = Math.min(ic + 1, img.c1) - Math.max(ic, img.c0);
            if (frac > 0) iw += colWpt[ic] * k * frac;
          }
          var ih = 0;
          for (var ir = img.r0; ir < img.r1 && ir < rowH.length; ir++){
            var rfrac = Math.min(ir + 1, img.r1) - Math.max(ir, img.r0);
            if (rfrac > 0) ih += rowH[ir] * rfrac;
          }
          if (iw > 2 && ih > 2) page.drawImage(img.embedded, { x: ix, y: y - ih, width: iw, height: ih });
        });
        var row = rows[ri];
        for (var cj = tile.c0; cj <= tile.c1; cj++){
          var cell = row[cj];
          if (!cell || cell.skip) continue;
          var mr = mergedRect(ri, cj);
          var cw = mr.w, ch = mr.h, cx = colX[cj - tile.c0], yTop = y;
          var yBot = yTop - ch;
          if (cell.fill){
            page.drawRectangle({ x: cx, y: yBot, width: cw, height: ch,
              color: PDFLib.rgb(cell.fill[0], cell.fill[1], cell.fill[2]) });
          }
          /* Conditional formatting: data bars draw behind the text (after
           * the fill); icon sets reserve space at the cell's left edge. */
          if (cell.dataBar && cell.dataBar.frac > 0){
            var dbc = cell.dataBar.color || [0.39, 0.56, 0.78];
            var dbw = cw * cell.dataBar.frac - 2 * k;
            if (dbw > 1) page.drawRectangle({ x: cx + k, y: yBot + 1.5 * k,
              width: dbw, height: Math.max(1, ch - 3 * k),
              color: PDFLib.rgb(dbc[0], dbc[1], dbc[2]) });
          }
          if (cell.iconSet){
            drawIconSet(page, PDFLib, cell.iconSet, cx + 2 * k, yBot + ch / 2, 9 * k);
          }
          var bd = cell.border || {};
          ['top','right','bottom','left'].forEach(function(side){
            if (bd[side]) drawBorderSide(page, PDFLib, side, bd[side], cx, yBot, cw, ch);
          });
          /* Excel-faithful overflow: measure the text; let it spill into
           * adjacent empty cells, clipped at the first occupied one. */
          var clip = null, cf = cell.font || {};
          var ct = String(cell.t == null ? '' : cell.t);
          if (ct && !(cell.align && cell.align.wrap) && !((cell.align && cell.align.rotation) || 0)){
            var tw = runWidth(F, cf, ct, (cf.size || 11) * k);
            if (tw > cw) clip = overflowClip(rows, ri, cj, tile, colWpt, k, cell, tw, cx, cw);
          }
          drawCellText(page, PDFLib, PDF, F, cell, cx, yBot, cw, ch, k, clip);
        }
        y -= rh;
      }
    });
    /* Charts: drawn after the grid, anchored by twoCellAnchor. The tile
     * containing the anchor's left column owns the draw; a chart whose
     * anchor column was tiled away is dropped with a stat (honest out). */
    if (XlsxCharts && s.charts && s.charts.length){
      s.charts.forEach(function(entry){
        try {
          var ch = entry.chart, a = entry.anchor;
          if (!ch || ch.kind === 'unsupported' || !a) return;
          var owner = null;
          tiles.forEach(function(tl){ if (!owner && a.c0 >= tl.c0 && a.c0 <= tl.c1) owner = tl; });
          if (!owner){ (meta.stats || {}).chartsTiledAway = ((meta.stats || {}).chartsTiledAway || 0) + 1; return; }
          var rp = owner.rowPos[a.r0];
          if (!rp) return;
          var kk = owner.k, rH = owner.rowH, cX = owner.colX;
          var chx = cX[a.c0 - owner.c0];
          var chw = 0;
          for (var cc = a.c0; cc < a.c1 && cc < colWpt.length; cc++)
            chw += colWpt[cc] * kk;
          chw = Math.min(chw, (PW - MG.r) - chx); /* clamp to page */
          var chh = 0;
          for (var cr = a.r0; cr < a.r1 && cr < rH.length; cr++) chh += rH[cr];
          chh = Math.min(chh, rp.yTop - (MG.b + FOOTR)); /* clamp to page */
          if (chw > 40 && chh > 40)
            XlsxCharts.drawChart(rp.page, PDFLib, F, ch, chx, rp.yTop, chw, chh);
        } catch(e){}
      });
    }
  });

  /* ---------- headers + footers on every page except the cover ----------
   * Numbering counts ALL pages including the cover, so the footer always
   * matches the viewer's "Page N of M". With cover:false the footer starts
   * on page 1. footer 'auto' (default): the workbook's own oddFooter wins
   * when the file defines one; otherwise the built-in filename/date/Page N
   * of M footer. header 'auto': the workbook's own oddHeader when defined,
   * else nothing. 'none' disables; {left,center,right} gives custom text
   * (same & codes as Excel). differentFirst/differentOddEven variants from
   * the workbook are honored. */
  var pages = d.getPages();
  var firstContent = wantCover ? 1 : 0;
  for (var i = firstContent; i < pages.length; i++){
    var pg = pages[i], sz = pg.getSize();
    var pW = sz.width, pH = sz.height;
    var sheet = pg._hfSheet || null;
    var ctx = { page: i + 1, pages: pages.length, dateStr: dateStr,
      timeStr: timeStr, filename: filename, sheetName: sheet ? sheet.name : '',
      isFirstOfSheet: !!pg._hfFirst, mL: MG.l, mR: MG.r, PW: pW };
    /* header, drawn above the sheet content */
    var hs = pickHF(sheet, 'header', ctx, meta);
    if (hs && hs !== 'DEFAULT'){
      var hh = hfBlockHeight(hs, headerStyle);
      if (hh > 0){
        drawHFBlock(pg, PDFLib, F, hs, pH - MG.t, pH - MG.t - hh, headerStyle, ctx);
        if (headerStyle.rule){
          var hrc = headerStyle.ruleColor;
          pg.drawLine({ start: { x: MG.l, y: pH - MG.t - hh - 3 },
            end: { x: pW - MG.r, y: pH - MG.t - hh - 3 }, thickness: 0.75,
            color: PDFLib.rgb(hrc[0], hrc[1], hrc[2]) });
        }
      }
    }
    /* footer */
    var fs = pickHF(sheet, 'footer', ctx, meta);
    if (fs === 'DEFAULT' || !sheet){
      drawDefaultFooter(pg, PDFLib, F, i, pages.length, pW, MG, logoImg,
        showFooterLogo, filename, dateStr);
    } else if (fs){
      var fTop = MG.b + 24, fBot = MG.b;
      drawHFBlock(pg, PDFLib, F, fs, fTop, fBot, footerStyle, ctx);
      if (footerStyle.rule){
        var frc = footerStyle.ruleColor;
        pg.drawLine({ start: { x: MG.l, y: fTop + 2 },
          end: { x: pW - MG.r, y: fTop + 2 }, thickness: 0.75,
          color: PDFLib.rgb(frc[0], frc[1], frc[2]) });
      }
      /* The 29Court logo only joins the built-in default footer; a
       * workbook-defined or custom footer carries its own &[Picture]. */
    }
  }

  var bytes = await d.save();
  /* Library API: raw PDF bytes out. Callers that need a Blob (browsers)
   * wrap it themselves; Node callers write the bytes straight to disk. */
  return bytes;
}

var api = { sheetsToPdf: sheetsToPdf, colWidthToPt: colWidthToPt,
  layoutTiles: layoutTiles, mergeColIntervals: mergeColIntervals,
  computeRowHeights: computeRowHeights, parseHF: parseHF, pickHF: pickHF };

if (typeof module !== 'undefined' && module.exports){ module.exports = api; }
else if (typeof globalThis !== 'undefined'){ globalThis.XlsxPrint = api; }
})();
