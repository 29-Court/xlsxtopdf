/* 29Court Convert: DrawingML chart parser + vector renderer.
 * Parses xl/charts/chartN.xml (namespace-agnostic: Excel writes <c:barChart>,
 * openpyxl writes <barChart>), resolves series data against the EVALUATED
 * value grid (computed formula values, not cached — the same advantage as
 * the main pipeline), and renders bar/column, line, area and pie charts as
 * vector PDF graphics via injected pdf-lib.
 * DOM-free; testable in Node. Anchors come from xdr:twoCellAnchor.
 * Honest outs (v1): 3D, stock, surface, radar, bubble, scatter, trendlines,
 * error bars, secondary axes, rich data labels, external-workbook refs.
 */
(function(){
"use strict";

/* Strip namespace prefixes from tag names: <c:barChart> -> <barChart>.
 * Attribute values are untouched (only tag opens are rewritten). */
function stripNS(xml){
  return String(xml).replace(/<(\/?)([A-Za-z_][\w.-]*):/g, '<$1');
}

function colL2I(s){
  var n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}
/* "'Charts'!$B$1:$B$4" -> {sheet, r1,c1,r2,c2} 0-based, or null. */
function parseRangeRef(ref){
  var m = /^'?(.*?)'?!(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?$/.exec(String(ref || ''));
  if (!m) return null;
  function cell(a){
    var x = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/.exec(a);
    if (!x) return null;
    return { r: +x[4] - 1, c: colL2I(x[2]) };
  }
  var A = cell(m[2]); if (!A) return null;
  var B = m[3] ? cell(m[3]) : A; if (!B) return null;
  return { sheet: m[1], r1: Math.min(A.r, B.r), c1: Math.min(A.c, B.c),
           r2: Math.max(A.r, B.r), c2: Math.max(A.c, B.c) };
}

function hexRgb(hex){
  var m = /^([0-9A-Fa-f]{6})$/.exec(hex || '');
  if (!m) return null;
  return [parseInt(m[1].slice(0,2),16)/255, parseInt(m[1].slice(2,4),16)/255, parseInt(m[1].slice(4,6),16)/255];
}
/* <a:solidFill> -> [r,g,b] or null. Handles srgbClr; schemeClr falls back
 * to null (caller substitutes the palette). */
function parseSolidFill(s){
  if (!s) return null;
  var m = /<solidFill>([\s\S]*?)<\/solidFill>/.exec(s);
  if (!m) return null;
  var srgb = /<srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(m[1]);
  if (srgb) return hexRgb(srgb[1]);
  return null;
}

/* Resolve a <cat>/<val>/<tx> block to values via getVal, falling back to
 * inline caches (<numCache>/<strCache>) when there is no formula ref. */
function resolveDataRef(block, getVal){
  if (!block) return [];
  var fm = /<f>([^<]+)<\/f>/.exec(block);
  if (fm){
    var r = parseRangeRef(fm[1]);
    if (!r) return [];
    var vals = [];
    var horiz = r.r1 === r.r2;
    var n = horiz ? (r.c2 - r.c1 + 1) : (r.r2 - r.r1 + 1);
    for (var i = 0; i < n; i++){
      var rr = horiz ? r.r1 : r.r1 + i, cc = horiz ? r.c1 + i : r.c1;
      vals.push(getVal(r.sheet, rr + 1, cc + 1));
    }
    return vals;
  }
  /* inline cache fallback */
  var out = [], re = /<pt[^>]*>[\s\S]*?<v>([^<]*)<\/v>[\s\S]*?<\/pt>/g, m;
  while ((m = re.exec(block))) out.push(m[1]);
  return out;
}

/* One <ser> block -> {title, cats[], vals[], fill, sliceFills[]}. */
function parseSeries(s, getVal){
  var out = { title: '', cats: [], vals: [], fill: null, sliceFills: [] };
  var txm = /<tx>([\s\S]*?)<\/tx>/.exec(s);
  if (txm){
    var v = /<v>([^<]*)<\/v>/.exec(txm[1]);
    if (v) out.title = v[1];
    else {
      var vals = resolveDataRef(txm[1], getVal);
      if (vals.length) out.title = String(vals[0]);
    }
  }
  var catm = /<cat>([\s\S]*?)<\/cat>/.exec(s);
  var valm = /<val>([\s\S]*?)<\/val>/.exec(s);
  out.cats = resolveDataRef(catm, getVal).map(String);
  out.vals = resolveDataRef(valm, getVal).map(function(v){
    var n = Number(v);
    return isNaN(n) ? NaN : n;
  });
  out.fill = parseSolidFill(s);
  /* per-slice colors for pie (<dPt><spPr><solidFill>) */
  var re = /<dPt>[\s\S]*?<spPr>([\s\S]*?)<\/spPr>[\s\S]*?<\/dPt>/g, m;
  while ((m = re.exec(s))){
    var f = parseSolidFill(m[1]);
    if (f) out.sliceFills.push(f);
  }
  return out;
}

function chartTitle(xml){
  var m = /<title>[\s\S]*?<\/title>/.exec(xml);
  if (!m) return '';
  var texts = m[0].match(/<t>([^<]*)<\/t>/g) || [];
  return texts.map(function(t){ return t.replace(/<\/?t>/g, ''); }).join('');
}

/* Parse one chart XML -> model or {kind:'unsupported'}. */
function parseChart(xml, getVal){
  xml = stripNS(xml);
  var kind = null, barDir = 'col', grouping = 'clustered';
  if (/<barChart>/.test(xml)){
    kind = 'bar';
    var bd = /<barDir\s+val="([^"]+)"/.exec(xml);
    if (bd) barDir = bd[1];
    var gp = /<grouping\s+val="([^"]+)"/.exec(xml);
    if (gp) grouping = gp[1];
  }
  else if (/<lineChart>/.test(xml)) kind = 'line';
  else if (/<areaChart>/.test(xml)) kind = 'area';
  else if (/<pieChart>/.test(xml) || /<doughnutChart>/.test(xml)) kind = 'pie';
  else if (/<scatterChart>/.test(xml)) kind = 'scatter';
  if (!kind) return { kind: 'unsupported', why: 'no recognized plot type' };
  if (kind === 'scatter') return { kind: 'unsupported', why: 'scatter' };
  /* 3D charts: honest out */
  if (/<scene3d>/.test(xml)) return { kind: 'unsupported', why: '3d' };
  var series = [];
  var re = /<ser>([\s\S]*?)<\/ser>/g, m;
  while ((m = re.exec(xml))) series.push(parseSeries(m[1], getVal));
  if (!series.length) return { kind: 'unsupported', why: 'no series' };
  return {
    kind: kind, barDir: barDir, grouping: grouping,
    title: chartTitle(xml),
    legend: /<legend>/.test(xml),
    series: series
  };
}

/* <Relationship Id="rId1" Target="../charts/chart1.xml"/> -> {rId1: target}
 * Attribute order is not significant. */
function parseRels(xml){
  var out = {}, re = /<Relationship\s+[^>]*>/g, m;
  while ((m = re.exec(xml))){
    var id = /Id="([^"]+)"/.exec(m[0]), tgt = /Target="([^"]+)"/.exec(m[0]);
    if (id && tgt) out[id[1]] = tgt[1];
  }
  return out;
}

/* drawing.xml -> [{c0,r0,c1,r1, chartRelId}] (0-based grid units).
 * Supports twoCellAnchor and oneCellAnchor (openpyxl); the latter gives an
 * <ext> size in EMUs, converted to grid units via default cell metrics. */
function parseAnchors(xml){
  xml = stripNS(xml);
  var out = [];
  function parseBlock(b){
    function cell(tag){
      var mm = new RegExp('<' + tag + '>([\\s\\S]*?)</' + tag + '>').exec(b);
      if (!mm) return null;
      var c = /<col>(\d+)<\/col>/.exec(mm[1]), r = /<row>(\d+)<\/row>/.exec(mm[1]);
      if (!c || !r) return null;
      return { c: +c[1], r: +r[1] };
    }
    var from = cell('from');
    var cm = /<chart[^>]*r:id="([^"]+)"/.exec(b) || /<chart[^>]*id="([^"]+)"/.exec(b);
    if (!from || !cm) return;
    var to = cell('to');
    if (to){
      out.push({ c0: from.c, r0: from.r, c1: to.c, r1: to.r, relId: cm[1] });
      return;
    }
    /* oneCellAnchor: <ext cx cy> in EMUs (914400 = 1 inch) */
    var em = /<ext\s+[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(b) ||
             /<ext\s+[^>]*cy="(\d+)"[^>]*cx="(\d+)"/.exec(b);
    if (em){
      var wPt = (+em[1]) * 72 / 914400, hPt = (+em[2]) * 72 / 914400;
      out.push({ c0: from.c, r0: from.r,
        c1: from.c + Math.max(1, Math.round(wPt / 64)),
        r1: from.r + Math.max(1, Math.round(hPt / 15)), relId: cm[1] });
    }
  }
  var re = /<twoCellAnchor>([\s\S]*?)<\/twoCellAnchor>/g, m;
  while ((m = re.exec(xml))) parseBlock(m[1]);
  re = /<oneCellAnchor>([\s\S]*?)<\/oneCellAnchor>/g;
  while ((m = re.exec(xml))) parseBlock(m[1]);
  return out;
}

function baseDir(p){ var i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i + 1); }
function resolveTarget(base, target){
  /* Absolute targets ("/xl/worksheets/sheet1.xml") are package-rooted. */
  if (target.charAt(0) === '/') target = target.slice(1);
  else target = base + target;
  var parts = target.split('/');
  var out = [];
  for (var i = 0; i < parts.length; i++){
    if (parts[i] === '..') out.pop();
    else if (parts[i] !== '.') out.push(parts[i]);
  }
  return out.join('/');
}

/* Full chain: workbook -> sheet files -> drawing files -> chart files.
 * Returns { sheetName: [{ chart, anchor }] }. Never throws. */
async function parseCharts(JSZip, buf, getVal){
  var out = {};
  try {
    var zip = await JSZip.loadAsync(buf);
    async function read(path){
      var f = zip.file(path);
      return f ? await f.async('string') : null;
    }
    var wbRels = parseRels(await read('xl/_rels/workbook.xml.rels') || '');
    var wbXml = stripNS(await read('xl/workbook.xml') || '');
    /* sheet name -> sheet file */
    var sheetFiles = {};
    var re = /<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/>/g, m;
    while ((m = re.exec(wbXml))){
      var tgt = wbRels[m[2]];
      if (tgt) sheetFiles[m[1]] = resolveTarget('xl/', tgt);
    }
    for (var name in sheetFiles){
      var sheetPath = sheetFiles[name];
      var sRels = parseRels(await read(baseDir(sheetPath) + '_rels/' +
        sheetPath.slice(sheetPath.lastIndexOf('/') + 1) + '.rels') || '');
      for (var rid in sRels){
        var dtgt = sRels[rid];
        if (!/drawing/.test(dtgt)) continue;
        var drawingPath = resolveTarget(baseDir(sheetPath), dtgt);
        var drawingXml = await read(drawingPath);
        if (!drawingXml) continue;
        var anchors = parseAnchors(drawingXml);
        if (!anchors.length) continue;
        var dRels = parseRels(await read(baseDir(drawingPath) + '_rels/' +
          drawingPath.slice(drawingPath.lastIndexOf('/') + 1) + '.rels') || '');
        anchors.forEach(function(a){
          var ctgt = dRels[a.relId];
          if (!ctgt || !/chart/.test(ctgt)) return;
          out[name] = out[name] || [];
          out[name].push({ chartPath: resolveTarget(baseDir(drawingPath), ctgt), anchor: a });
        });
      }
    }
    /* now parse the chart XMLs */
    for (var sn in out){
      var list = [];
      for (var i = 0; i < out[sn].length; i++){
        var entry = out[sn][i];
        var cxml = await read(entry.chartPath);
        if (!cxml) continue;
        var chart = parseChart(cxml, getVal);
        list.push({ chart: chart, anchor: entry.anchor });
      }
      out[sn] = list;
    }
  } catch(e){}
  return out;
}

/* ---------------- renderers (vector, pdf-lib primitives) ---------------- */

var PALETTE = [
  [0.29, 0.53, 0.91], [0.91, 0.30, 0.24], [0.18, 0.80, 0.44], [0.99, 0.75, 0.21],
  [0.61, 0.35, 0.71], [0.20, 0.74, 0.73], [0.95, 0.45, 0.66], [0.55, 0.55, 0.55]
];
var INKC = [0.13, 0.12, 0.10];

function seriesColor(s, i){
  return s.fill || PALETTE[i % PALETTE.length];
}
function textW(font, t, size){
  try { return font.widthOfTextAtSize(t, size); } catch(e){ return t.length * size * 0.55; }
}

/* Shared cartesian frame: title, axes, gridlines, category labels, legend.
 * Returns plot geometry for the series renderer. */
function cartesianFrame(page, PDFLib, F, chart, x, yTop, w, h, cats, allVals){
  var INK = PDFLib.rgb(INKC[0], INKC[1], INKC[2]);
  var GRAY = PDFLib.rgb(0.78, 0.77, 0.75);
  var ty = yTop;
  if (chart.title){
    var tw = textW(F.bold, chart.title, 11);
    page.drawText(chart.title, { x: x + Math.max(0, (w - tw) / 2), y: ty - 12,
      size: 11, font: F.bold, color: INK });
    ty -= 24;
  }
  var legW = chart.legend ? 110 : 8;
  var L = x + 44, R = x + w - legW, T = ty - 6, B = yTop - h + 22;
  if (R - L < 60 || T - B < 60) return null;
  var maxV = Math.max.apply(null, allVals.concat([0]));
  var minV = Math.min.apply(null, allVals.concat([0]));
  if (!(maxV > minV)){ maxV = 1; minV = 0; }
  var yOf = function(v){ return B + (T - B) * (v - minV) / (maxV - minV); };
  var i, v, yy, lab;
  for (i = 0; i <= 4; i++){
    v = minV + (maxV - minV) * i / 4; yy = yOf(v);
    page.drawLine({ start: { x: L, y: yy }, end: { x: R, y: yy },
      thickness: 0.5, color: GRAY });
    lab = String(Math.round(v * 100) / 100);
    page.drawText(lab, { x: L - 6 - textW(F.reg, lab, 7), y: yy - 2.5,
      size: 7, font: F.reg, color: INK });
  }
  page.drawLine({ start: { x: L, y: B }, end: { x: R, y: B }, thickness: 1, color: INK });
  page.drawLine({ start: { x: L, y: B }, end: { x: L, y: T }, thickness: 1, color: INK });
  return { L: L, R: R, T: T, B: B, yOf: yOf, minV: minV, maxV: maxV,
           legX: x + w - legW + 12, legW: legW };
}

function drawLegend(page, PDFLib, F, chart, frame, yTop){
  if (!chart.legend) return;
  var INK = PDFLib.rgb(INKC[0], INKC[1], INKC[2]);
  chart.series.forEach(function(s, si){
    var y = yTop - 30 - si * 16;
    var col = seriesColor(s, si);
    page.drawRectangle({ x: frame.legX, y: y - 2, width: 11, height: 8,
      color: PDFLib.rgb(col[0], col[1], col[2]), borderWidth: 0.5, borderColor: INK });
    page.drawText(s.title || ('Series ' + (si + 1)),
      { x: frame.legX + 15, y: y - 2, size: 8, font: F.reg, color: INK });
  });
}

function drawBarChart(page, PDFLib, F, chart, x, yTop, w, h){
  var horizontal = chart.barDir === 'bar';
  var cats = chart.series[0].cats.length ? chart.series[0].cats
    : chart.series[0].vals.map(function(_, i){ return '' + (i + 1); });
  var n = cats.length, ns = chart.series.length;
  var allVals = [];
  chart.series.forEach(function(s){ s.vals.forEach(function(v){ if (!isNaN(v)) allVals.push(v); }); });
  if (!n || !allVals.length) return;
  var INK = PDFLib.rgb(INKC[0], INKC[1], INKC[2]);
  /* stacked values for axis scaling */
  var stackMax = 0, stackMin = 0;
  if (chart.grouping === 'stacked' || chart.grouping === 'percentStacked'){
    for (var ci = 0; ci < n; ci++){
      var pos = 0, neg = 0;
      chart.series.forEach(function(s){
        var v = s.vals[ci];
        if (isNaN(v)) return;
        if (chart.grouping === 'percentStacked') return;
        if (v >= 0) pos += v; else neg += v;
      });
      if (pos > stackMax) stackMax = pos;
      if (neg < stackMin) stackMin = neg;
    }
    if (chart.grouping === 'percentStacked'){ stackMax = 1; stackMin = 0; }
    allVals = [stackMax, stackMin];
  }
  var fr = cartesianFrame(page, PDFLib, F, chart, x, yTop, w, h, cats, allVals);
  if (!fr) return;
  var pw = fr.R - fr.L, slotW = pw / n;
  var si, vi, v;
  if (!horizontal){
    var barW = Math.min(40, slotW / (ns + 0.6) * 0.62);
    chart.series.forEach(function(s, six){
      var col = seriesColor(s, six);
      var c = PDFLib.rgb(col[0], col[1], col[2]);
      s.vals.forEach(function(vv, ci2){
        if (isNaN(vv)) return;
        var vv2 = vv;
        if (chart.grouping === 'percentStacked'){
          var tot = 0;
          chart.series.forEach(function(o){ var q = o.vals[ci2]; if (!isNaN(q) && q > 0) tot += q; });
          vv2 = tot ? vv / tot : 0;
        }
        var gx = fr.L + slotW * ci2 + slotW / 2;
        var bx = gx - (ns * barW) / 2 + six * barW;
        var base = 0, top = vv2;
        if (chart.grouping === 'stacked' || chart.grouping === 'percentStacked'){
          base = 0;
          for (var k = 0; k < six; k++){
            var pv = chart.series[k].vals[ci2];
            if (chart.grouping === 'percentStacked'){
              var pt = 0;
              chart.series.forEach(function(o){ var q = o.vals[ci2]; if (!isNaN(q) && q > 0) pt += q; });
              pv = (!isNaN(pv) && pv > 0 && pt) ? pv / pt : 0;
            }
            if (!isNaN(pv) && pv > 0) base += pv;
          }
          top = base + (vv2 > 0 ? vv2 : 0);
          bx = gx - barW / 2; barW = Math.min(40, slotW * 0.62);
        }
        var y0 = fr.yOf(chart.grouping === 'stacked' || chart.grouping === 'percentStacked' ? base : Math.min(0, top));
        var y1 = fr.yOf(top);
        /* negative stacked values: draw below */
        if (top < base){ y0 = fr.yOf(top); y1 = fr.yOf(base); }
        page.drawRectangle({ x: bx, y: y0, width: barW, height: Math.max(1, y1 - y0),
          color: c, borderWidth: 0.5, borderColor: INK });
      });
    });
  } else {
    /* horizontal bars: categories on y, values on x */
    var bw2 = Math.min(30, slotW / (ns + 0.6) * 0.62);
    chart.series.forEach(function(s, six){
      var col2 = seriesColor(s, six);
      var c2 = PDFLib.rgb(col2[0], col2[1], col2[2]);
      s.vals.forEach(function(vv, ci2){
        if (isNaN(vv)) return;
        var gy = fr.B + slotW * ci2 + slotW / 2;
        var by = gy - (ns * bw2) / 2 + six * bw2;
        /* x maps value->position using the frame's yOf on swapped axes */
        var span = fr.maxV - fr.minV || 1;
        var xOf = function(val){ return fr.L + (fr.R - fr.L) * (val - fr.minV) / span; };
        var x0 = xOf(Math.min(0, vv)), x1 = xOf(Math.max(0, vv));
        page.drawRectangle({ x: x0, y: by, width: Math.max(1, x1 - x0), height: bw2,
          color: c2, borderWidth: 0.5, borderColor: INK });
      });
    });
  }
  /* category labels */
  cats.forEach(function(cname, ci3){
    var lab = String(cname), tw = textW(F.reg, lab, 7);
    if (tw > slotW - 4) lab = lab.slice(0, Math.max(1, Math.floor(lab.length * (slotW - 4) / tw))) + '…';
    page.drawText(lab, { x: fr.L + slotW * ci3 + slotW / 2 - textW(F.reg, lab, 7) / 2,
      y: fr.B - 12, size: 7, font: F.reg, color: INK });
  });
  drawLegend(page, PDFLib, F, chart, fr, yTop);
}

function drawLineChart(page, PDFLib, F, chart, x, yTop, w, h, filled){
  var cats = chart.series[0].cats.length ? chart.series[0].cats
    : chart.series[0].vals.map(function(_, i){ return '' + (i + 1); });
  var n = cats.length;
  var allVals = [];
  chart.series.forEach(function(s){ s.vals.forEach(function(v){ if (!isNaN(v)) allVals.push(v); }); });
  if (!n || !allVals.length) return;
  var INK = PDFLib.rgb(INKC[0], INKC[1], INKC[2]);
  var fr = cartesianFrame(page, PDFLib, F, chart, x, yTop, w, h, cats, allVals);
  if (!fr) return;
  var pw = fr.R - fr.L, stepX = n > 1 ? pw / (n - 1) : 0;
  var xOf = function(i){ return fr.L + stepX * i; };
  chart.series.forEach(function(s, si){
    var col = seriesColor(s, si);
    var c = PDFLib.rgb(col[0], col[1], col[2]);
    var pts = [];
    s.vals.forEach(function(v, i){
      if (!isNaN(v)) pts.push({ x: xOf(i), y: fr.yOf(v) });
    });
    if (filled && pts.length > 1){
      /* area fill under the line */
      var fillPath = 'M ' + pts[0].x + ' ' + fr.yOf(0);
      pts.forEach(function(p){ fillPath += ' L ' + p.x + ' ' + p.y; });
      fillPath += ' L ' + pts[pts.length - 1].x + ' ' + fr.yOf(0) + ' Z';
      try {
        page.drawSvgPath(fillPath, { color: c, opacity: 0.35, borderWidth: 0 });
      } catch(e){}
    }
    for (var i = 1; i < pts.length; i++){
      page.drawLine({ start: pts[i - 1], end: pts[i], thickness: 1.5, color: c });
    }
    pts.forEach(function(p){
      page.drawCircle({ x: p.x, y: p.y, size: 3, color: c, borderWidth: 0.75, borderColor: INK });
    });
  });
  cats.forEach(function(cname, ci){
    var lab = String(cname);
    page.drawText(lab, { x: xOf(ci) - textW(F.reg, lab, 7) / 2, y: fr.B - 12,
      size: 7, font: F.reg, color: INK });
  });
  drawLegend(page, PDFLib, F, chart, fr, yTop);
}

function drawPieChart(page, PDFLib, F, chart, x, yTop, w, h){
  var s = chart.series[0];
  var vals = s.vals.filter(function(v){ return !isNaN(v) && v > 0; });
  var cats = s.cats;
  if (!vals.length) return;
  var INK = PDFLib.rgb(INKC[0], INKC[1], INKC[2]);
  if (chart.title){
    var tw = textW(F.bold, chart.title, 11);
    page.drawText(chart.title, { x: x + Math.max(0, (w - tw) / 2), y: yTop - 12,
      size: 11, font: F.bold, color: INK });
  }
  var legW = chart.legend ? 120 : 10;
  var cx = x + (w - legW) / 2, cy = yTop - h / 2 - 6;
  var rad = Math.min((w - legW) / 2 - 8, h / 2 - 30);
  if (rad < 20) return;
  var total = vals.reduce(function(a, b){ return a + b; }, 0) || 1;
  var ang = Math.PI / 2, i;
  for (i = 0; i < vals.length; i++){
    var sweep = vals[i] / total * Math.PI * 2;
    var col = s.sliceFills[i] || PALETTE[i % PALETTE.length];
    var c = PDFLib.rgb(col[0], col[1], col[2]);
    /* wedge as a many-sided polygon (pdf-lib has no arc fill) */
    var steps = Math.max(8, Math.ceil(sweep / (Math.PI / 24)));
    var path = 'M ' + cx + ' ' + cy;
    for (var k = 0; k <= steps; k++){
      var a = ang - sweep * k / steps;
      path += ' L ' + (cx + rad * Math.cos(a)).toFixed(2) + ' ' + (cy + rad * Math.sin(a)).toFixed(2);
    }
    path += ' Z';
    try { page.drawSvgPath(path, { color: c, borderWidth: 0.75, borderColor: INK }); }
    catch(e){}
    /* percent label near the wedge middle */
    var mid = ang - sweep / 2;
    var lx = cx + rad * 0.62 * Math.cos(mid), ly = cy + rad * 0.62 * Math.sin(mid);
    var pct = Math.round(vals[i] / total * 100) + '%';
    page.drawText(pct, { x: lx - textW(F.reg, pct, 7) / 2, y: ly - 3,
      size: 7, font: F.reg, color: INK });
    ang -= sweep;
  }
  if (chart.legend){
    var ly0 = yTop - 34;
    vals.forEach(function(v, li){
      var col2 = s.sliceFills[li] || PALETTE[li % PALETTE.length];
      var y = ly0 - li * 16;
      if (y < yTop - h + 14) return;
      page.drawRectangle({ x: x + w - legW + 10, y: y - 2, width: 11, height: 8,
        color: PDFLib.rgb(col2[0], col2[1], col2[2]), borderWidth: 0.5, borderColor: INK });
      var lab = (cats[li] || ('Slice ' + (li + 1))) + ' ' + Math.round(v / total * 100) + '%';
      page.drawText(lab, { x: x + w - legW + 25, y: y - 2, size: 8, font: F.reg, color: INK });
    });
  }
}

/* Draw one chart into the rect (x, yTop, w, h). Never throws. */
function drawChart(page, PDFLib, F, chart, x, yTop, w, h){
  try {
    if (w < 40 || h < 40) return;
    if (chart.kind === 'bar') drawBarChart(page, PDFLib, F, chart, x, yTop, w, h);
    else if (chart.kind === 'line') drawLineChart(page, PDFLib, F, chart, x, yTop, w, h, false);
    else if (chart.kind === 'area') drawLineChart(page, PDFLib, F, chart, x, yTop, w, h, true);
    else if (chart.kind === 'pie') drawPieChart(page, PDFLib, F, chart, x, yTop, w, h);
  } catch(e){}
}

var api = {
  parseCharts: parseCharts,
  parseChart: parseChart,
  parseRangeRef: parseRangeRef,
  parseAnchors: parseAnchors,
  drawChart: drawChart,
  stripNS: stripNS
};

if (typeof module !== 'undefined' && module.exports){ module.exports = api; }
else if (typeof globalThis !== 'undefined'){ globalThis.XlsxCharts = api; }
})();

