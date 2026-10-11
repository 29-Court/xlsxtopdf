/* 29Court Convert: rich XLSX reader (ExcelJS first, SheetJS fallback).
 * DOM-free on purpose: pure data layer, testable in Node.
 * Produces sheets of rich cells:
 *   { t, span, spanR, skip,
 *     bold, italic, fill:[r,g,b]|null, color:[r,g,b]|null,   (legacy top-level)
 *     font: { name, size, bold, italic, underline, strike, color },
 *     border: { top:{style,color}, right:{...}, bottom:{...}, left:{...} },
 *     align: { h, v, wrap, rotation, indent } }
 * Colors are [r,g,b] in 0..1. Theme colors (the big fidelity win) are resolved
 * from the workbook theme (xl/theme/theme1.xml, exposed by ExcelJS as
 * wb.model.themes.theme1) with the MS linear tint model.
 * Sheets also carry colWidths (Excel width units) and rowHeights (points).
 * formulaMode 'final' (default): formula cells render their computed value.
 * formulaMode 'raw': formula cells render as "=FORMULA".
 * In final mode every formula is evaluated in the browser with an Excel
 * formula engine (fast-formula-parser, MIT), so files whose formulas were
 * never recalculated still render numbers. A formula that genuinely cannot
 * be evaluated renders blank and increments stats.noCache (honest log line),
 * never as raw "=FORMULA" text.
 */
(function(){
"use strict";

function clamp01(x){ return x < 0 ? 0 : x > 1 ? 1 : x; }

function argbToRgb(argb){
  var s = String(argb == null ? '' : argb);
  if (s.length === 8) s = s.slice(2); /* drop alpha */
  if (!/^[0-9A-Fa-f]{6}$/.test(s)) return null;
  return [parseInt(s.slice(0,2),16)/255, parseInt(s.slice(2,4),16)/255, parseInt(s.slice(4,6),16)/255];
}
function isWhite(rgb){ return !!rgb && rgb[0] > 0.96 && rgb[1] > 0.96 && rgb[2] > 0.96; }

/* ---- theme colors ---- */
/* OOXML theme index order: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink */
var THEME_NAMES = ['lt1','dk1','lt2','dk2','accent1','accent2','accent3',
  'accent4','accent5','accent6','hlink','folHlink'];

function themeColorHex(xml, name){
  var sec = new RegExp('<a:' + name + '\\b[\\s\\S]*?</a:' + name + '>', 'i').exec(xml);
  if (!sec) return null;
  var s = sec[0], m;
  m = /<a:srgbClr[^>]*\bval="([0-9A-Fa-f]{6})"/i.exec(s);
  if (m) return m[1];
  m = /<a:sysClr[^>]*\blastClr="([0-9A-Fa-f]{6})"/i.exec(s);
  if (m) return m[1];
  return null;
}

/* Parse theme XML -> array of 12 [r,g,b] (0..1), or null. */
function parseThemeColors(xml){
  if (typeof xml !== 'string' || xml.indexOf('clrScheme') === -1) return null;
  var out = [];
  for (var i = 0; i < THEME_NAMES.length; i++){
    var hex = themeColorHex(xml, THEME_NAMES[i]);
    if (!hex) return null;
    out.push([parseInt(hex.slice(0,2),16)/255, parseInt(hex.slice(2,4),16)/255, parseInt(hex.slice(4,6),16)/255]);
  }
  return out;
}

/* MS linear tint model for styles.xml theme colors:
 * tint in [-1,1]; negative shades toward black, positive tints toward white. */
function applyTint(rgb, tint){
  var r = rgb[0], g = rgb[1], b = rgb[2];
  if (tint < 0){ var f = 1 + tint; r *= f; g *= f; b *= f; }
  else { r += (1 - r) * tint; g += (1 - g) * tint; b += (1 - b) * tint; }
  return [clamp01(r), clamp01(g), clamp01(b)];
}

/* ExcelJS color ({argb} | {theme, tint} | {indexed}) -> [r,g,b] 0..1, or null. */
function resolveColor(c, theme){
  if (!c) return null;
  if (c.argb) return argbToRgb(c.argb);
  if (c.theme !== undefined && c.theme !== null && theme){
    var base = theme[c.theme];
    if (!base) return null;
    return applyTint(base, c.tint || 0);
  }
  return null;
}

function themeFromWorkbook(wb){
  try {
    var th = wb.model && wb.model.themes;
    if (!th) return null;
    var keys = Object.keys(th);
    for (var i = 0; i < keys.length; i++){
      var xml = th[keys[i]];
      if (typeof xml === 'string'){
        var colors = parseThemeColors(xml);
        if (colors) return colors;
      }
    }
  } catch(e){}
  return null;
}

function colLettersToIndex(s){
  var n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}
/* "A8:F8" -> {r1,c1,r2,c2} zero-based, or null */
function parseMergeRef(ref){
  var m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(String(ref || ''));
  if (!m) return null;
  return { r1: +m[2] - 1, c1: colLettersToIndex(m[1]), r2: +m[4] - 1, c2: colLettersToIndex(m[3]) };
}

function emptyRich(){
  return {
    t:'', span:1, spanR:1, skip:false, num:false,
    bold:false, italic:false, fill:null, color:null,
    font:{ name:null, size:11, bold:false, italic:false, underline:false, strike:false, color:null },
    border:{ top:null, right:null, bottom:null, left:null },
    align:{ h:'general', v:'bottom', wrap:false, rotation:0, indent:0 }
  };
}

/* ---- formula-engine polyfills ----
 * fast-formula-parser ships several widely-used functions as empty stubs
 * (MAX, MIN, MATCH, SUMIFS, ...). These override them with Excel-faithful
 * implementations via the parser's `functions` option. Custom-function args
 * arrive wrapped as { value, isArray, isRangeRef, isCellRef }. */
function fpUnwrap(arg){
  return (arg && typeof arg === 'object' && 'value' in arg) ? arg.value : arg;
}
function fpIsErr(v){ return !!(v && typeof v === 'object' && typeof v._error === 'string'); }
function fpFlatten(v, out){
  out = out || [];
  if (Array.isArray(v)){ for (var i = 0; i < v.length; i++) fpFlatten(v[i], out); }
  else out.push(v);
  return out;
}
/* numbers for MAX/MIN-style aggregation: refs contribute numbers only;
 * literals also accept numeric text and booleans (Excel semantics). */
function fpNumbers(args){
  var nums = [];
  for (var i = 0; i < args.length; i++){
    var a = args[i], literal = !(a && typeof a === 'object' && a.isRangeRef);
    var vals = fpFlatten(fpUnwrap(a));
    for (var j = 0; j < vals.length; j++){
      var v = vals[j];
      if (fpIsErr(v)) return { err: v._error };
      if (typeof v === 'number') nums.push(v);
      else if (literal){
        if (typeof v === 'boolean') nums.push(v ? 1 : 0);
        else if (typeof v === 'string' && v !== '' && !isNaN(Number(v))) nums.push(Number(v));
      }
    }
  }
  return { nums: nums };
}
/* criteria for SUMIFS/COUNTIFS/...: number, boolean, "text", "wild*card",
 * ">5", "<=3", "<>x". Returns a predicate over a cell value. */
function fpCriteriaFn(crit){
  crit = fpUnwrap(crit);
  if (crit === undefined || crit === null || crit === '') return function(v){ return v === undefined || v === null || v === ''; };
  if (typeof crit === 'number' || typeof crit === 'boolean')
    return function(v){ return v === crit; };
  var s = String(crit), m = /^(>=|<=|<>|>|<|=)([\s\S]*)$/.exec(s);
  var op = null, operand = s;
  if (m){ op = m[1]; operand = m[2]; }
  var num = Number(operand), operandIsNum = operand !== '' && !isNaN(num);
  function eq(a, b){
    if (typeof a === 'number' && operandIsNum) return a === num;
    return String(a == null ? '' : a).toLowerCase() === String(b).toLowerCase();
  }
  if (op === '>' || op === '<' || op === '>=' || op === '<='){
    return function(v){
      if (typeof v !== 'number' || !operandIsNum) return false;
      return op === '>' ? v > num : op === '<' ? v < num : op === '>=' ? v >= num : v <= num;
    };
  }
  if (op === '<>') return function(v){ return !eq(v, operand); };
  if (/[*?]/.test(operand)){
    var rx = new RegExp('^' + operand.replace(/[.+^${}()|[\]\\]/g, '\\$&')
      .replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
    return function(v){ return typeof v === 'string' && rx.test(v); };
  }
  return function(v){ return eq(v, operand); };
}
/* parallel walk of criteria ranges; returns the flat indexes (into the
 * first range) where every criterion matches. */
function fpIfsRows(critRanges, crits){
  var lists = critRanges.map(function(r){
    var v = fpUnwrap(r);
    return fpFlatten(Array.isArray(v) ? v : [[v]]);
  });
  var n = lists[0].length, hits = [];
  var fns = crits.map(fpCriteriaFn);
  for (var i = 0; i < n; i++){
    var okAll = true;
    for (var k = 0; k < fns.length; k++){
      if (!fns[k](lists[k][i])){ okAll = false; break; }
    }
    if (okAll) hits.push(i);
  }
  return hits;
}
function fpIfsArgs(args){
  var critRanges = [], crits = [];
  for (var i = 0; i < args.length; i += 2){ critRanges.push(args[i]); crits.push(args[i + 1]); }
  return { critRanges: critRanges, crits: crits };
}
function fpErrObj(code){ return { _error: code }; }

var FORMULA_POLYFILLS = {
  MAX: function(){
    var r = fpNumbers(arguments);
    if (r.err) return fpErrObj(r.err);
    return r.nums.length ? Math.max.apply(null, r.nums) : 0;
  },
  MIN: function(){
    var r = fpNumbers(arguments);
    if (r.err) return fpErrObj(r.err);
    return r.nums.length ? Math.min.apply(null, r.nums) : 0;
  },
  MAXA: function(){
    var nums = [];
    for (var i = 0; i < arguments.length; i++){
      var vals = fpFlatten(fpUnwrap(arguments[i]));
      for (var j = 0; j < vals.length; j++){
        var v = vals[j];
        if (fpIsErr(v)) return fpErrObj(v._error);
        if (typeof v === 'number') nums.push(v);
        else if (typeof v === 'boolean') nums.push(v ? 1 : 0);
        else nums.push(0);
      }
    }
    return nums.length ? Math.max.apply(null, nums) : 0;
  },
  MINA: function(){
    var nums = [];
    for (var i = 0; i < arguments.length; i++){
      var vals = fpFlatten(fpUnwrap(arguments[i]));
      for (var j = 0; j < vals.length; j++){
        var v = vals[j];
        if (fpIsErr(v)) return fpErrObj(v._error);
        if (typeof v === 'number') nums.push(v);
        else if (typeof v === 'boolean') nums.push(v ? 1 : 0);
        else nums.push(0);
      }
    }
    return nums.length ? Math.min.apply(null, nums) : 0;
  },
  LARGE: function(array, k){
    var vals = fpFlatten(fpUnwrap(array)).filter(function(v){ return typeof v === 'number'; });
    k = Math.floor(Number(fpUnwrap(k)));
    if (!(k >= 1) || k > vals.length) return fpErrObj('#NUM!');
    vals.sort(function(a, b){ return b - a; });
    return vals[k - 1];
  },
  SMALL: function(array, k){
    var vals = fpFlatten(fpUnwrap(array)).filter(function(v){ return typeof v === 'number'; });
    k = Math.floor(Number(fpUnwrap(k)));
    if (!(k >= 1) || k > vals.length) return fpErrObj('#NUM!');
    vals.sort(function(a, b){ return a - b; });
    return vals[k - 1];
  },
  MEDIAN: function(){
    var r = fpNumbers(arguments);
    if (r.err) return fpErrObj(r.err);
    if (!r.nums.length) return fpErrObj('#NUM!');
    r.nums.sort(function(a, b){ return a - b; });
    var mid = Math.floor(r.nums.length / 2);
    return r.nums.length % 2 ? r.nums[mid] : (r.nums[mid - 1] + r.nums[mid]) / 2;
  },
  COUNTBLANK: function(range){
    var vals = fpFlatten(fpUnwrap(range)), n = 0;
    for (var i = 0; i < vals.length; i++){
      var v = vals[i];
      if (v === undefined || v === null || v === '') n++;
    }
    return n;
  },
  SUMIFS: function(sumRange){
    var p = fpIfsArgs(Array.prototype.slice.call(arguments, 1));
    if (arguments.length % 2 === 0) return fpErrObj('#VALUE!');
    var sumVals = fpFlatten(fpUnwrap(sumRange)), total = 0;
    var hits = fpIfsRows(p.critRanges, p.crits);
    for (var h = 0; h < hits.length; h++){
      var v = sumVals[hits[h]];
      if (fpIsErr(v)) return fpErrObj(v._error);
      if (typeof v === 'number') total += v;
    }
    return total;
  },
  COUNTIFS: function(){
    var args = Array.prototype.slice.call(arguments);
    if (!args.length || args.length % 2) return fpErrObj('#VALUE!');
    var p = fpIfsArgs(args);
    return fpIfsRows(p.critRanges, p.crits).length;
  },
  AVERAGEIFS: function(avgRange){
    var p = fpIfsArgs(Array.prototype.slice.call(arguments, 1));
    if (arguments.length % 2 === 0) return fpErrObj('#VALUE!');
    var vals = fpFlatten(fpUnwrap(avgRange));
    var hits = fpIfsRows(p.critRanges, p.crits), sum = 0, n = 0;
    for (var h = 0; h < hits.length; h++){
      var v = vals[hits[h]];
      if (fpIsErr(v)) return fpErrObj(v._error);
      if (typeof v === 'number'){ sum += v; n++; }
    }
    if (!n) return fpErrObj('#DIV/0!');
    return sum / n;
  },
  MAXIFS: function(maxRange){
    var p = fpIfsArgs(Array.prototype.slice.call(arguments, 1));
    if (arguments.length % 2 === 0) return fpErrObj('#VALUE!');
    var vals = fpFlatten(fpUnwrap(maxRange));
    var hits = fpIfsRows(p.critRanges, p.crits), best = null;
    for (var h = 0; h < hits.length; h++){
      var v = vals[hits[h]];
      if (fpIsErr(v)) return fpErrObj(v._error);
      if (typeof v === 'number' && (best === null || v > best)) best = v;
    }
    return best === null ? 0 : best;
  },
  MINIFS: function(minRange){
    var p = fpIfsArgs(Array.prototype.slice.call(arguments, 1));
    if (arguments.length % 2 === 0) return fpErrObj('#VALUE!');
    var vals = fpFlatten(fpUnwrap(minRange));
    var hits = fpIfsRows(p.critRanges, p.crits), best = null;
    for (var h = 0; h < hits.length; h++){
      var v = vals[hits[h]];
      if (fpIsErr(v)) return fpErrObj(v._error);
      if (typeof v === 'number' && (best === null || v < best)) best = v;
    }
    return best === null ? 0 : best;
  },
  INDEX: function(ctx, array, rowNum, colNum){
    /* INDEX is a "needs context, no data retrieve" function: the array arg
     * arrives as a raw { ref } and must be resolved via the parser utils. */
    function numOf(v){
      if (v && typeof v === 'object'){
        if (v.ref !== undefined && ctx && ctx.utils){ v = ctx.utils.extractRefValue(v).val; }
        else if ('value' in v){ v = v.value; }
      }
      return Math.floor(Number(v));
    }
    var r = numOf(rowNum);
    var c = (colNum === undefined || colNum === null) ? 1 : numOf(colNum);
    var vals;
    if (array && typeof array === 'object' && array.ref !== undefined && ctx && ctx.utils){
      vals = ctx.utils.extractRefValue(array).val;
    } else if (array && typeof array === 'object' && 'value' in array){
      vals = array.value;
    } else {
      vals = array;
    }
    if (!Array.isArray(vals)) vals = [[vals]];
    if (r === 0 || c === 0) return vals; /* whole row/column: collapse downstream */
    if (!(r >= 1) || r > vals.length) return fpErrObj('#REF!');
    var row = vals[r - 1];
    if (!Array.isArray(row)) row = [row];
    if (!(c >= 1) || c > row.length) return fpErrObj('#REF!');
    var out = row[c - 1];
    if (fpIsErr(out)) return out;
    return out === undefined ? 0 : out;
  },
  MATCH: function(lookup, array, type){
    lookup = fpUnwrap(lookup);
    var vals = fpFlatten(fpUnwrap(array));
    var t = type === undefined ? 1 : Math.floor(Number(fpUnwrap(type)));
    if (t !== 0 && t !== 1 && t !== -1) return fpErrObj('#N/A');
    if (t === 0){
      var rx = null;
      if (typeof lookup === 'string' && /[*?]/.test(lookup))
        rx = new RegExp('^' + lookup.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
      for (var i = 0; i < vals.length; i++){
        var v = vals[i];
        if (rx){ if (typeof v === 'string' && rx.test(v)) return i + 1; }
        else if (typeof v === typeof lookup &&
          (typeof v === 'string' ? v.toLowerCase() === String(lookup).toLowerCase() : v === lookup)) return i + 1;
      }
      return fpErrObj('#N/A');
    }
    var pos = -1;
    for (var j = 0; j < vals.length; j++){
      var w = vals[j];
      if (typeof w !== 'number' || typeof lookup !== 'number') continue;
      if (t === 1 && w <= lookup) pos = j;
      if (t === -1 && w >= lookup) pos = j;
    }
    return pos < 0 ? fpErrObj('#N/A') : pos + 1;
  }
};

/* ---- formula evaluation (fast-formula-parser, MIT licensed) ----
 * NOTE on the engine choice: HyperFormula is GPL-3.0-only, which would
 * copyleft the whole product. fast-formula-parser is MIT and evaluates
 * the full Excel function set we need (SUM, IF, VLOOKUP, arithmetic...).
 *
 * rawSheets: [{ name, cells }] where cells[ri][ci] (0-based) is
 *   { k:'v', v } a plain value | { k:'f', f } a formula WITHOUT leading '='
 * Returns per-sheet evals[ri][ci] = { ok:true, v } | { ok:false } for
 * formula cells that had no cached value, and fills stats:
 *   evaluated (ok), evalFailed (threw / unsupported), engineMissing.
 * Never throws. Values flow through the parser's own dependency graph via
 * onCell/onRange, so chains like R6/G6 (formula referencing a formula)
 * resolve correctly, with a cycle guard. Ranges are clamped to the used
 * area so full-column refs like SUM(A:A) cannot hang the tab. */
function evaluateFormulas(rawSheets, FPClass, stats){
  stats = stats || {};
  if (!FPClass || typeof FPClass !== 'function'){ stats.engineMissing = true; return null; }
  var nameToIdx = {};
  rawSheets.forEach(function(s, i){ nameToIdx[s.name] = i; });
  var grids = rawSheets.map(function(s){
    var g = {};
    s.cells.forEach(function(row, ri){
      row.forEach(function(cell, ci){
        if (cell) g[(ri + 1) + '|' + (ci + 1)] = cell;
      });
    });
    return g;
  });
  var nrows = rawSheets.map(function(s){ return s.cells.length; });
  var ncols = rawSheets.map(function(s){
    return s.cells.reduce(function(a, r){ return Math.max(a, r.length); }, 0);
  });
  var memo = {}, busy = {};
  var parser;
  function k3(si, r, c){ return si + '|' + r + '|' + c; }
  /* Parser values -> display values. Error objects become Excel error text
   * (faithful: Excel shows #DIV/0! in the cell). Arrays collapse to their
   * first scalar. Anything else unrecognized becomes blank, never a crash. */
  function scalarOf(v){
    if (v === undefined || v === null) return v;
    if (v instanceof Date) return v;
    if (typeof v === 'object'){
      if (typeof v._error === 'string') return v._error;
      if (Array.isArray(v)){
        var f = v.length ? v[0] : null;
        if (Array.isArray(f)) f = f.length ? f[0] : null;
        return (f === undefined || f === null) ? '' : scalarOf(f);
      }
      return '';
    }
    return v;
  }
  function evalCell(si, r, c){
    var k = k3(si, r, c), m = memo[k];
    if (m) return m;
    var cell = grids[si][r + '|' + c];
    var out;
    if (!cell) out = { v: undefined };              /* empty cell */
    else if (cell.k === 'v') out = { v: cell.v };
    else if (busy[k]) out = { v: null, cycle: true }; /* circular reference */
    else {
      busy[k] = 1;
      try {
        out = { v: scalarOf(parser.parse(cell.f, { sheet: rawSheets[si].name, row: r, col: c })) };
      } catch(e){
        out = { v: null, failed: true };
      }
      delete busy[k];
    }
    memo[k] = out;
    return out;
  }
  function onCell(ref){
    var si = nameToIdx[ref.sheet];
    if (si === undefined) return undefined;
    return evalCell(si, ref.row, ref.col).v;
  }
  function onRange(ref){
    var si = nameToIdx[ref.sheet];
    var arr = [];
    if (si === undefined) return arr;
    /* clamp to the used area: anything outside is empty anyway */
    var r1 = Math.max(1, ref.from.row), r2 = Math.min(nrows[si], ref.to.row);
    var c1 = Math.max(1, ref.from.col), c2 = Math.min(ncols[si], ref.to.col);
    for (var r = r1; r <= r2; r++){
      var inner = [];
      for (var c = c1; c <= c2; c++) inner.push(evalCell(si, r, c).v);
      arr.push(inner);
    }
    return arr;
  }
  parser = new FPClass({ onCell: onCell, onRange: onRange, functions: FORMULA_POLYFILLS });
  /* Drive evaluation for every uncached formula cell so chains resolve
   * once, up front, instead of lazily mid-render. */
  var evals = rawSheets.map(function(s, si){
    return s.cells.map(function(row, ri){
      return row.map(function(cell, ci){
        if (cell && cell.k === 'f'){
          var e = evalCell(si, ri + 1, ci + 1);
          if (e.failed || e.cycle){
            stats.evalFailed = (stats.evalFailed || 0) + 1;
            return { ok: false };
          }
          stats.evaluated = (stats.evaluated || 0) + 1;
          return { ok: true, v: e.v };
        }
        return undefined;
      });
    });
  });
  return evals;
}

/* One ExcelJS cell -> raw formula/value record (or null when empty).
 * Formulas come back without the leading '=' from ExcelJS. */
function rawOfExcelCell(cell){
  var v = cell.value;
  if (v === undefined || v === null) return null;
  if (v instanceof Date) return { k:'v', v:v };
  if (typeof v === 'object'){
    var formula = (v.formula !== undefined && v.formula !== null) ? v.formula : v.sharedFormula;
    if (formula !== undefined && formula !== null){
      if (v.result === undefined || v.result === null) return { k:'f', f: String(formula) };
      var r = v.result;
      if (r && typeof r === 'object' && r.error) return { k:'v', v: String(r.error) };
      return { k:'v', v: r };
    }
    if (v.error !== undefined) return { k:'v', v: String(v.error) };
    if (v.richText) return { k:'v', v: v.richText.map(function(rt){ return rt.text; }).join('') };
    return null;
  }
  return { k:'v', v:v };
}

/* deps.XLSX optional: used for number-format rendering */
function ssfFormat(XLSX, fmt, v){
  try {
    if (fmt && XLSX && XLSX.SSF && XLSX.SSF.format){
      if (v instanceof Date && (!fmt || /general/i.test(String(fmt)))) return v.toLocaleDateString();
      return XLSX.SSF.format(String(fmt), v);
    }
  } catch(e){}
  if (v instanceof Date) return v.toLocaleDateString();
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function mapHAlign(h){
  h = String(h || '').toLowerCase();
  if (h === 'center' || h === 'centercontinuous') return 'center';
  if (h === 'right') return 'right';
  /* Excel default is General: text left, numbers right (resolved at draw time). */
  return 'general';
}
function mapVAlign(v){
  v = String(v || '').toLowerCase();
  if (v === 'middle' || v === 'center') return 'middle';
  if (v === 'top') return 'top';
  /* Excel's default vertical alignment is bottom. */
  return 'bottom';
}

function styleFromCell(cell, theme){
  var rc = { bold:false, italic:false, fill:null, color:null,
    font:null, border:null, align:null };
  var f = cell.font || {};
  var fcolor = resolveColor(f.color || f.colour, theme);
  rc.bold = !!f.bold; rc.italic = !!f.italic;
  /* Keep explicit font colors, even white: white text on a colored fill is
   * meaningful (and faithful). */
  rc.color = fcolor;
  rc.font = {
    name: f.name || null,
    size: (typeof f.size === 'number' && f.size > 0) ? f.size : 11,
    bold: !!f.bold, italic: !!f.italic,
    underline: !!(f.underline === true || f.underline === 'single' || f.underline === 'double'),
    strike: !!f.strike,
    color: fcolor
  };
  var fl = cell.fill;
  if (fl && (fl.pattern === 'solid' || fl.type === 'pattern') && fl.fgColor){
    var fr = resolveColor(fl.fgColor, theme);
    if (fr && !isWhite(fr)) rc.fill = fr;
  }
  /* borders: {top:{style,color},...}; color may sit under `color` or `colour` */
  var b = cell.border || {}, out = { top:null, right:null, bottom:null, left:null };
  ['top','right','bottom','left'].forEach(function(side){
    var bs = b[side];
    if (bs && bs.style && bs.style !== 'none'){
      var bc = resolveColor(bs.color || bs.colour, theme);
      out[side] = { style: String(bs.style), color: bc };
    }
  });
  rc.border = out;
  var a = cell.alignment || {};
  var rot = a.textRotation;
  if (rot === 255) rot = 0; /* 255 = stacked text; not supported, keep unrotated */
  rc.align = {
    h: mapHAlign(a.horizontal),
    v: mapVAlign(a.vertical),
    wrap: !!a.wrapText,
    rotation: (typeof rot === 'number') ? rot : 0,
    indent: Math.max(0, Math.min(8, a.indent | 0))
  };
  return rc;
}

/* One ExcelJS cell -> rich cell. stats.noCache counts formulas that had no
 * cached value AND could not be evaluated: they render blank (never as raw
 * "=FORMULA" text in final mode; raw mode is the user's explicit choice).
 * `computed` is the evaluateFormulas() entry for this cell, when present. */
function richFromExcelCell(cell, mode, XLSX, stats, theme, computed){
  var rc = emptyRich();
  var st = styleFromCell(cell, theme);
  rc.bold = st.bold; rc.italic = st.italic; rc.fill = st.fill; rc.color = st.color;
  rc.font = st.font; rc.border = st.border; rc.align = st.align;
  var v = cell.value;
  if (v == null){ rc.t = ''; }
  else if (v instanceof Date){ rc.t = ssfFormat(XLSX, cell.numFmt, v); rc.num = true; }
  else if (typeof v === 'object'){
    if (v.formula !== undefined){
      var formula = (v.formula !== undefined && v.formula !== null) ? v.formula : v.sharedFormula;
      if (mode === 'raw'){ rc.t = '=' + formula; }
      else if (computed && computed.ok){
        var cv = computed.v;
        if (typeof cv === 'string' && cv.charAt(0) === '#'){ rc.t = cv; rc.num = false; }
        else if (cv === undefined || cv === null || cv === ''){ rc.t = ''; }
        else {
          rc.t = ssfFormat(XLSX, cell.numFmt, cv);
          rc.num = (typeof cv === 'number') || (cv instanceof Date);
        }
      }
      else if (v.result !== undefined && v.result !== null){
        var fr = v.result;
        if (fr && typeof fr === 'object' && fr.error){ rc.t = String(fr.error); }
        else {
          rc.t = ssfFormat(XLSX, cell.numFmt, fr);
          rc.num = (typeof fr === 'number') || (fr instanceof Date);
        }
      }
      else {
        /* Final mode, no cached value, not evaluated: blank. Raw formula
         * text in a client-facing PDF is never acceptable. */
        rc.t = '';
        stats.noCache++;
      }
    }
    else if (v.error !== undefined){ rc.t = String(v.error); }
    else if (v.richText){ rc.t = v.richText.map(function(rt){ return rt.text; }).join(''); }
    else { rc.t = ''; }
  }
  else {
    rc.t = ssfFormat(XLSX, cell.numFmt, v);
    rc.num = (typeof v === 'number');
  }
  return rc;
}

function applyMerges(rows, merges, stats){
  merges.forEach(function(ref){
    var m = parseMergeRef(ref);
    if (!m) return;
    if (m.r2 > m.r1) stats.vMerge = (stats.vMerge || 0) + 1;
    for (var r = m.r1; r <= m.r2; r++){
      for (var c = m.c1; c <= m.c2; c++){
        if (!rows[r] || !rows[r][c]) continue;
        if (r === m.r1 && c === m.c1){
          rows[r][c].span = m.c2 - m.c1 + 1;
          rows[r][c].spanR = m.r2 - m.r1 + 1;
        }
        else { rows[r][c].skip = true; }
      }
    }
  });
}

/* Print headers/footers: Excel stores &[Picture] images for header/footer
 * sections in per-sheet VML drawings (xl/drawings/vmlDrawingN.vml), not in
 * the cell drawing. This walks workbook.xml -> sheet rels -> vmlDrawing ->
 * imagedata r:ids -> media parts, and returns {sheetName: [{data, ext}]}
 * with images in VML document order. Never throws; a miss means []. */
function hfXmlAttr(tag, name){
  var m = tag.match(new RegExp(name.replace(/:/g, ':') + '="([^"]*)"'));
  return m ? m[1] : null;
}
function hfUnesc(s){
  return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}
function hfResolve(base, target){
  var parts = base.split('/'); parts.pop();
  String(target).split('/').forEach(function(seg){
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  });
  return parts.join('/');
}
function hfRelsMap(xml){
  var map = {};
  var re = /<Relationship\b[^>]*>/g, m;
  while ((m = re.exec(xml))) map[hfXmlAttr(m[0], 'Id')] = hfXmlAttr(m[0], 'Target');
  return map;
}
async function parseHFPictures(JSZip, buf){
  var out = {};
  try {
    var zip = await JSZip.loadAsync(buf);
    var wbF = zip.file('xl/workbook.xml');
    if (!wbF) return out;
    var wbXml = await wbF.async('string');
    var sheets = [];
    var re = /<sheet\b[^>]*>/g, m;
    while ((m = re.exec(wbXml))){
      var nm = hfXmlAttr(m[0], 'name');
      var rid = hfXmlAttr(m[0], 'r:id') || hfXmlAttr(m[0], 'id');
      if (nm && rid) sheets.push({ name: hfUnesc(nm), rid: rid });
    }
    var relsF = zip.file('xl/_rels/workbook.xml.rels');
    if (!relsF) return out;
    var wbRels = hfRelsMap(await relsF.async('string'));
    for (var si = 0; si < sheets.length; si++){
      var sTarget = wbRels[sheets[si].rid];
      if (!sTarget) continue;
      var sheetPath = hfResolve('xl/workbook.xml', sTarget);
      var sheetRelsPath = sheetPath.replace(/\/([^\/]+)$/, '/_rels/$1.rels');
      var srF = zip.file(sheetRelsPath);
      if (!srF) continue;
      var srXml = await srF.async('string');
      var vmlTarget = null;
      var re2 = /<Relationship\b[^>]*>/g, m2;
      while ((m2 = re2.exec(srXml))){
        if (/vmlDrawing/i.test(hfXmlAttr(m2[0], 'Type') || '')){
          vmlTarget = hfXmlAttr(m2[0], 'Target'); break;
        }
      }
      if (!vmlTarget) continue;
      var vmlPath = hfResolve(sheetPath, vmlTarget);
      var vmlF = zip.file(vmlPath);
      if (!vmlF) continue;
      var vmlXml = await vmlF.async('string');
      var ids = [];
      var re3 = /<\w+:imagedata\b[^>]*>/g, m3;
      while ((m3 = re3.exec(vmlXml))){
        var iid = hfXmlAttr(m3[0], 'r:id') || hfXmlAttr(m3[0], 'id');
        if (iid) ids.push(iid);
      }
      if (!ids.length) continue;
      var vmlRelsPath = vmlPath.replace(/\/([^\/]+)$/, '/_rels/$1.rels');
      var vrF = zip.file(vmlRelsPath);
      if (!vrF) continue;
      var vRels = hfRelsMap(await vrF.async('string'));
      var imgs = [];
      for (var k = 0; k < ids.length; k++){
        var mt = vRels[ids[k]];
        if (!mt) continue;
        var mediaPath = hfResolve(vmlPath, mt);
        var mf = zip.file(mediaPath);
        if (!mf) continue;
        var data = await mf.async('uint8array');
        var ext = (mediaPath.split('.').pop() || 'png').toLowerCase();
        imgs.push({ data: data, ext: ext });
      }
      if (imgs.length) out[sheets[si].name] = imgs;
    }
  } catch(e){}
  return out;
}

/* Raw print header/footer strings from the workbook (odd/even/first
 * variants). Code parsing and &[Picture] resolution happen in the print
 * layer; this stays the faithful record of what the file defines. */
function printHFRaw(ws){
  try {
    var hf = ws.headerFooter || {};
    if (!(hf.oddHeader || hf.oddFooter || hf.evenHeader || hf.evenFooter ||
          hf.firstHeader || hf.firstFooter)) return null;
    return { oddHeader: hf.oddHeader || null, oddFooter: hf.oddFooter || null,
      evenHeader: hf.evenHeader || null, evenFooter: hf.evenFooter || null,
      firstHeader: hf.firstHeader || null, firstFooter: hf.firstFooter || null,
      differentFirst: !!hf.differentFirst, differentOddEven: !!hf.differentOddEven,
      images: [] };
  } catch(e){ return null; }
}

async function readViaExcelJS(buf, mode, deps){
  var ExcelJS = deps.ExcelJS, XLSX = deps.XLSX;
  var wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  var theme = themeFromWorkbook(wb);
  var stats = { noCache: 0, vMerge: 0, sheets: 0, themeColors: !!theme };
  /* First pass: collect raw cells for every sheet, so formulas can be
   * evaluated workbook-wide (cross-sheet references resolve). */
  var pending = [];
  wb.eachSheet(function(ws){
    stats.sheets++;
    var nrows = ws.rowCount || 0;
    var cellGrid = [];
    var maxC = 0;
    for (var r = 1; r <= nrows; r++){
      var row = [], wsRow = ws.getRow(r);
      var nc = wsRow.cellCount || 0;
      for (var c = 1; c <= nc; c++){
        var cell;
        try { cell = wsRow.getCell(c); } catch(e){ cell = null; }
        row.push(cell);
      }
      if (nc > maxC) maxC = nc;
      cellGrid.push(row);
    }
    /* normalize row lengths */
    cellGrid.forEach(function(crow){ while (crow.length < maxC) crow.push(null); });
    pending.push({ ws: ws, cellGrid: cellGrid, maxC: maxC });
  });
  /* Raw value grids for ALL modes (chart series resolve against these).
   * Formula evaluation happens only in final mode with uncached formulas. */
  var hasUncached = false;
  var rawSheets = pending.map(function(p){
    var rawCells = p.cellGrid.map(function(crow){
      return crow.map(function(ccell){
        var raw = ccell ? rawOfExcelCell(ccell) : null;
        if (raw && raw.k === 'f') hasUncached = true;
        return raw;
      });
    });
    return { name: p.ws.name, cells: rawCells };
  });
  /* Formula evaluation (final mode only): every uncached formula in the
   * workbook gets a computed value BEFORE rendering, so the PDF shows
   * numbers, never raw "=SUM(...)" text. */
  var allEvals = null;
  if (mode === 'final' && hasUncached){
    if (deps.FormulaParser) allEvals = evaluateFormulas(rawSheets, deps.FormulaParser, stats);
    else stats.engineMissing = true;
  }
  var sheets = [];
  pending.forEach(function(p, pi){
    var ws = p.ws, cellGrid = p.cellGrid, maxC = p.maxC;
    var evals = allEvals ? allEvals[pi] : null;
    var rows = cellGrid.map(function(crow, ri){
      return crow.map(function(ccell, ci){
        return ccell ? richFromExcelCell(ccell, mode, XLSX, stats, theme,
            evals ? evals[ri][ci] : undefined)
          : emptyRich();
      });
    });
    while (rows.length && rows[rows.length-1].every(function(rc){ return rc.t === ''; })) rows.pop();
    /* normalize row lengths */
    rows.forEach(function(row){ while (row.length < maxC) row.push(emptyRich()); });
    /* column widths (Excel width units) + row heights (points) */
    var colWidths = [];
    for (var c = 1; c <= maxC; c++){
      var col = null;
      try { col = ws.getColumn(c); } catch(e){}
      colWidths.push(col && typeof col.width === 'number' ? col.width : null);
    }
    var rowHeights = [];
    for (var r2 = 1; r2 <= rows.length; r2++){
      var rh = null;
      try { var wr = ws.getRow(r2); rh = (typeof wr.height === 'number') ? wr.height : null; } catch(e){}
      rowHeights.push(rh);
    }
    var merges = (ws.model && ws.model.merges) || [];
    applyMerges(rows, merges, stats);
    /* Conditional formatting: evaluate rules against computed values and
     * merge the winning dxf into the rich cells. Never breaks the sheet. */
    try {
      applyConditionalFormatting(ws, rows, maxC, ws.name, evals, deps.FormulaParser, theme, stats);
    } catch(e){}
    /* Embedded images: anchor rects in 0-based grid units + raw bytes.
     * Charts are not exposed by ExcelJS; they are reported in stats. */
    var images = [];
    try {
      var wbImages = ws.getImages ? ws.getImages() : [];
      wbImages.forEach(function(im){
        var meta = null;
        try { meta = wb.getImage(im.imageId); } catch(e){}
        if (!meta || !meta.buffer) return;
        var r = im.range || {}, tl = r.tl || {}, br = r.br || {};
        var buf = meta.buffer;
        /* Normalize to Uint8Array (browser ExcelJS may give base64 or ArrayBuffer). */
        if (typeof buf === 'string'){
          var bin = null;
          if (typeof Buffer !== 'undefined') buf = Buffer.from(buf, 'base64');
          else if (typeof atob !== 'undefined'){
            bin = atob(buf); var u8 = new Uint8Array(bin.length);
            for (var bi = 0; bi < bin.length; bi++) u8[bi] = bin.charCodeAt(bi);
            buf = u8;
          }
        } else if (buf instanceof ArrayBuffer){
          buf = new Uint8Array(buf);
        }
        images.push({ c0: tl.col || 0, r0: tl.row || 0,
          c1: (br.col == null ? (tl.col || 0) + 1 : br.col),
          r1: (br.row == null ? (tl.row || 0) + 1 : br.row),
          ext: (meta.extension || '').toLowerCase(), buf: buf });
      });
    } catch(e){}
    sheets.push({ name: ws.name, rows: rows, header: rows.length > 0,
      merges: merges.length, colWidths: colWidths, rowHeights: rowHeights,
      images: images, printHF: printHFRaw(ws) });
    /* Pad rows so image anchors beyond the data still have a row to draw on. */
    if (images.length){
      var needRows = 0;
      images.forEach(function(im){ needRows = Math.max(needRows, Math.ceil(im.r1)); });
      var sh = sheets[sheets.length - 1];
      var ncolsPad = sh.rows.reduce(function(a, r){ return Math.max(a, r.length); }, maxC);
      while (sh.rows.length < needRows){
        var er = [];
        for (var ec = 0; ec < ncolsPad; ec++) er.push(emptyRich());
        sh.rows.push(er);
      }
    }
  });
  /* Charts: DrawingML is invisible to ExcelJS, so parse the zip directly
   * via injected JSZip. Series data resolves against the raw/evaluated
   * value grids (computed values, not cached). Never breaks the sheet. */
  try {
    if (deps.JSZip && deps.XlsxCharts){
      var getChartVal = function(sheetName, r, c){ /* 1-based */
        for (var gi = 0; gi < rawSheets.length; gi++){
          if (rawSheets[gi].name !== sheetName) continue;
          var rawRow = rawSheets[gi].cells[r - 1];
          var raw = rawRow ? rawRow[c - 1] : null;
          if (!raw) return null;
          if (raw.k === 'v') return (raw.v instanceof Date) ? null : raw.v;
          if (raw.k === 'f' && allEvals && allEvals[gi] && allEvals[gi][r - 1]){
            var ev = allEvals[gi][r - 1][c - 1];
            return ev ? ev.v : null;
          }
          return null;
        }
        return null;
      };
      var chartMap = await deps.XlsxCharts.parseCharts(deps.JSZip, buf, getChartVal);
      sheets.forEach(function(sh){
        var list = chartMap[sh.name] || [];
        sh.charts = list;
        var needRows = sh.rows.length, needCols = 0;
        list.forEach(function(entry){
          if (entry.chart.kind === 'unsupported')
            stats.chartsUnsupported = (stats.chartsUnsupported || 0) + 1;
          else {
            stats.chartsDrawn = (stats.chartsDrawn || 0) + 1;
            if (entry.anchor){
              needRows = Math.max(needRows, entry.anchor.r1);
              needCols = Math.max(needCols, entry.anchor.c1);
            }
          }
        });
        /* Pad rows AND columns so chart anchors outside the data grid
         * still have cells to sit on (Excel's print area includes charts). */
        var ncolsPad = sh.rows.reduce(function(a, r){ return Math.max(a, r.length); }, 0);
        ncolsPad = Math.max(ncolsPad, needCols);
        sh.rows.forEach(function(row){
          while (row.length < ncolsPad) row.push(emptyRich());
        });
        while (sh.rows.length < needRows){
          var er2 = [];
          for (var ec2 = 0; ec2 < ncolsPad; ec2++) er2.push(emptyRich());
          sh.rows.push(er2);
        }
      });
    }
  } catch(e){}
  /* Print headers/footers: resolve &[Picture] images from the header/footer
   * VML drawings via the raw zip. Never breaks the sheet. */
  try {
    if (deps.JSZip){
      var hfMap = await parseHFPictures(deps.JSZip, buf);
      sheets.forEach(function(sh){
        if (sh.printHF) sh.printHF.images = hfMap[sh.name] || [];
      });
    }
  } catch(e){}
  return { sheets: sheets, stats: stats };
}

async function readViaSheetJS(buf, mode, deps){
  var XLSX = deps.XLSX;
  var wb = XLSX.read(buf, { type: 'array' });
  var stats = { noCache: 0, vMerge: 0, sheets: 0, themeColors: false };
  var sheets = wb.SheetNames.map(function(name){
    stats.sheets++;
    var ws = wb.Sheets[name];
    var range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    var rows = [];
    for (var r = range.s.r; r <= range.e.r; r++){
      var row = [];
      for (var c = range.s.c; c <= range.e.c; c++){
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        var rc = emptyRich();
        if (cell){
          if (cell.f !== undefined){
            /* SheetJS quirk (verified 0.18.5): a formula cell with no cached
             * value and no t="str" is dropped from the sheet entirely, so a
             * formula cell can only reach here when SheetJS kept it. */
            if (mode === 'raw'){ rc.t = '=' + cell.f; }
            else if (cell.v === undefined || cell.v === null || cell.v === ''){ rc.t = '=' + cell.f; stats.noCache++; }
            else { try { rc.t = XLSX.utils.format_cell(cell); } catch(e){ rc.t = String(cell.v); } }
          } else { try { rc.t = XLSX.utils.format_cell(cell); } catch(e){ rc.t = String(cell.v); } }
        }
        row.push(rc);
      }
      rows.push(row);
    }
    while (rows.length && rows[rows.length-1].every(function(rc){ return rc.t === ''; })) rows.pop();
    var maxC = rows.reduce(function(a, row){ return Math.max(a, row.length); }, 0);
    rows.forEach(function(row){ while (row.length < maxC) row.push(emptyRich()); });
    return { name: name, rows: rows, header: rows.length > 0, merges: 0,
      colWidths: [], rowHeights: [], printHF: null };
  });
  /* Charts on the SheetJS fallback path (used when ExcelJS cannot parse the
   * file): same DrawingML parse, values from cached cell values. */
  try {
    if (deps.JSZip && deps.XlsxCharts){
      var getVal2 = function(sheetName, r, c){
        var ws2 = wb.Sheets[sheetName];
        if (!ws2) return null;
        var cell2 = ws2[XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })];
        return cell2 ? cell2.v : null;
      };
      var chartMap2 = await deps.XlsxCharts.parseCharts(deps.JSZip, buf, getVal2);
      sheets.forEach(function(sh){
        var list2 = chartMap2[sh.name] || [];
        sh.charts = list2;
        var needRows2 = sh.rows.length, needCols2 = 0;
        list2.forEach(function(entry){
          if (entry.chart.kind === 'unsupported')
            stats.chartsUnsupported = (stats.chartsUnsupported || 0) + 1;
          else {
            stats.chartsDrawn = (stats.chartsDrawn || 0) + 1;
            if (entry.anchor){
              needRows2 = Math.max(needRows2, entry.anchor.r1);
              needCols2 = Math.max(needCols2, entry.anchor.c1);
            }
          }
        });
        var ncolsPad2 = sh.rows.reduce(function(a, r){ return Math.max(a, r.length); }, 0);
        ncolsPad2 = Math.max(ncolsPad2, needCols2);
        sh.rows.forEach(function(row){
          while (row.length < ncolsPad2) row.push(emptyRich());
        });
        while (sh.rows.length < needRows2){
          var er3 = [];
          for (var ec3 = 0; ec3 < ncolsPad2; ec3++) er3.push(emptyRich());
          sh.rows.push(er3);
        }
      });
    }
  } catch(e){}
  return { sheets: sheets, stats: stats };
}

/* Entry point. deps: {ExcelJS?, XLSX, log?}. Never throws for parseable files. */
async function readBuffer(buf, mode, deps){
  deps = deps || {};
  mode = mode === 'raw' ? 'raw' : 'final';
  var log = deps.log || function(){};
  if (deps.ExcelJS){
    try { return await readViaExcelJS(buf, mode, deps); }
    catch(e){ log('ExcelJS could not parse this file (' + (e && e.message) + '), used the basic reader.', 'warn'); }
  } else {
    log('ExcelJS did not load, used the basic reader (no cell styling).', 'warn');
  }
  return await readViaSheetJS(buf, mode, deps);
}

/* ---- conditional formatting ----
 * ExcelJS parses CF rules on read (ws.conditionalFormattings): no raw XML
 * needed for cellIs / expression / colorScale / dataBar / iconSet.
 * Rules evaluate against COMPUTED values (our formula engine), so CF on
 * formula cells works even when the workbook has no cached values.
 * Priority semantics: first hit in priority order wins (Excel's default
 * stopIfTrue). v1 does not honor stopIfTrue=false. Ranges are clamped to
 * the used area so full-column CF can't hang the tab. Never throws. */
function parseCFRange(ref){
  var m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(String(ref || '').split(' ')[0]);
  if (!m) return null;
  var c1 = colLettersToIndex(m[1]), r1 = +m[2] - 1, c2 = c1, r2 = r1;
  if (m[3]){ c2 = colLettersToIndex(m[3]); r2 = +m[4] - 1; }
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2),
           r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

/* Shift relative A1 refs in a formula by (dr, dc). Honors $ absolutes,
 * skips string literals and sheet-qualified refs. */
function shiftFormula(f, dr, dc){
  f = String(f || '');
  var parts = [], cur = '', inStr = false, i, ch;
  for (i = 0; i < f.length; i++){
    ch = f[i];
    if (ch === '"'){ inStr = !inStr; cur += ch; }
    else if (!inStr && /[A-Za-z0-9_.$!]/.test(ch)) cur += ch;
    else { parts.push(cur); cur = ''; parts.push(ch); }
  }
  parts.push(cur);
  function colToLetters(c){
    var s = ''; c += 1;
    while (c > 0){ var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  return parts.map(function(p){
    if (!/[A-Za-z]/.test(p.charAt(0))) return p;
    return p.replace(/(?<![A-Za-z0-9_.$!])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g,
      function(m, ac, col, ar, row){
        var c = colLettersToIndex(col), r = +row - 1;
        if (!ac) c += dc;
        if (!ar) r += dr;
        if (c < 0 || r < 0) return '#REF!';
        return (ac ? '$' : '') + colToLetters(c) + (ar ? '$' : '') + (r + 1);
      });
  }).join('');
}

var CF_OPS = {
  greaterThan: function(a, b){ return a > b; },
  lessThan: function(a, b){ return a < b; },
  equal: function(a, b){ return a == b; }, /* eslint eqeqeq:0 */
  notEqual: function(a, b){ return a != b; }, /* eslint eqeqeq:0 */
  greaterThanOrEqual: function(a, b){ return a >= b; },
  lessThanOrEqual: function(a, b){ return a <= b; },
  between: function(a, b, c){ return a >= b && a <= c; },
  notBetween: function(a, b, c){ return !(a >= b && a <= c); }
};

/* Resolve a cfvo ({type, value}) to a number given the range's numeric values. */
function cfvoValue(cfvo, vals, parser){
  var t = String(cfvo.type || '').toLowerCase();
  if (t === 'min' || t === 'automin') return vals.length ? Math.min.apply(null, vals) : 0;
  if (t === 'max' || t === 'automax') return vals.length ? Math.max.apply(null, vals) : 0;
  if (t === 'num'){ var n = Number(cfvo.value); return isNaN(n) ? 0 : n; }
  if (t === 'percent' || t === 'percentile'){
    var p = Number(cfvo.value);
    if (isNaN(p)) p = 0;
    if (!vals.length) return 0;
    var lo = Math.min.apply(null, vals), hi = Math.max.apply(null, vals);
    if (t === 'percent') return lo + (hi - lo) * p / 100;
    var s = vals.slice().sort(function(a, b){ return a - b; });
    return s[Math.min(s.length - 1, Math.floor(p / 100 * s.length))];
  }
  if (t === 'formula' && parser){
    try { var v = parser.parse(String(cfvo.value)); return typeof v === 'number' ? v : 0; }
    catch(e){ return 0; }
  }
  return 0;
}
function rangeNumbers(getVal, R){
  var out = [];
  for (var r = R.r1; r <= R.r2; r++)
    for (var c = R.c1; c <= R.c2; c++){
      var v = getVal(r, c);
      if (typeof v === 'number' && !isNaN(v)) out.push(v);
    }
  return out;
}
function lerpColor(A, B, t){
  return [A[0] + (B[0] - A[0]) * t, A[1] + (B[1] - A[1]) * t, A[2] + (B[2] - A[2]) * t];
}

/* Evaluate one CF rule against cell (r,c) 0-based. Returns null (no hit) or
 * a result: { fill, fontColor, bold, border } and/or
 * { dataBar:{frac,color}, iconSet:{index,total,set}, colorScaleFill }. */
function evalCFRule(rule, r, c, R, getVal, parser, sheetName, theme){
  var type = String(rule.type || '');
  var cellV = getVal(r, c);
  function parseWithCtx(f){
    return parser.parse(String(f), { sheet: sheetName, row: r + 1, col: c + 1 });
  }
  if (type === 'cellIs'){
    var op = CF_OPS[rule.operator];
    if (!op) return null;
    var args;
    try { args = (rule.formulae || []).map(parseWithCtx); }
    catch(e){ return null; }
    var hit = false;
    try { hit = !!op.apply(null, [cellV].concat(args)); } catch(e){ hit = false; }
    return hit ? { dxf: rule.style } : null;
  }
  if (type === 'expression'){
    var f0 = (rule.formulae && rule.formulae[0]) || '';
    var shifted = shiftFormula(f0, r - R.r1, c - R.c1);
    try { return parser.parse(shifted, { sheet: sheetName, row: r + 1, col: c + 1 }) ? { dxf: rule.style } : null; }
    catch(e){ return null; }
  }
  if (type === 'colorScale'){
    if (typeof cellV !== 'number' || isNaN(cellV)) return null;
    var cfvos = rule.cfvo || [], colors = (rule.color || []).map(function(cc){ return argbToRgb(cc.argb); });
    if (colors.length < 2 || colors.some(function(x){ return !x; })) return null;
    var vals = rangeNumbers(getVal, R);
    var stops = cfvos.map(function(cv){ return cfvoValue(cv, vals, parser); });
    var lo = stops[0], hi = stops[stops.length - 1], t;
    if (hi <= lo) t = 0.5; else t = Math.max(0, Math.min(1, (cellV - lo) / (hi - lo)));
    var fill;
    if (colors.length === 2) fill = lerpColor(colors[0], colors[1], t);
    else {
      var mid = stops.length === 3 ? (stops[1] - lo) / (hi - lo || 1) : 0.5;
      fill = t < mid
        ? lerpColor(colors[0], colors[1], mid ? t / mid : 0)
        : lerpColor(colors[1], colors[2], mid < 1 ? (t - mid) / (1 - mid) : 0);
    }
    return { colorScaleFill: fill };
  }
  if (type === 'dataBar'){
    if (typeof cellV !== 'number' || isNaN(cellV)) return null;
    var dcfvos = rule.cfvo || [];
    var dvals = rangeNumbers(getVal, R);
    var dlo = dcfvos[0] ? cfvoValue(dcfvos[0], dvals, parser) : (dvals.length ? Math.min.apply(null, dvals) : 0);
    var dhi = dcfvos[1] ? cfvoValue(dcfvos[1], dvals, parser) : (dvals.length ? Math.max.apply(null, dvals) : 100);
    var frac = dhi > dlo ? (cellV - dlo) / (dhi - dlo) : 0;
    frac = Math.max(0, Math.min(1, frac));
    var dcolor = (rule.color && rule.color.argb) ? argbToRgb(rule.color.argb) : [0.39, 0.56, 0.78];
    return { dataBar: { frac: frac, color: dcolor || [0.39, 0.56, 0.78] } };
  }
  if (type === 'iconSet'){
    if (typeof cellV !== 'number' || isNaN(cellV)) return null;
    var iset = rule.iconSet || {};
    var setName = String(iset.iconSet || '3Arrows');
    var total = parseInt(setName, 10) || 3;
    var icfvos = iset.cfvo || [];
    var ivals = rangeNumbers(getVal, R);
    var ilo = icfvos[0] ? cfvoValue(icfvos[0], ivals, parser) : (ivals.length ? Math.min.apply(null, ivals) : 0);
    var ihi = icfvos[icfvos.length - 1] ? cfvoValue(icfvos[icfvos.length - 1], ivals, parser)
      : (ivals.length ? Math.max.apply(null, ivals) : 100);
    var it = ihi > ilo ? (cellV - ilo) / (ihi - ilo) : 0;
    it = Math.max(0, Math.min(0.9999, it));
    return { iconSet: { index: Math.floor(it * total), total: total, set: setName } };
  }
  return null;
}

/* Merge a dxf style (ExcelJS {fill,font,border}) into a rich cell. */
function applyCFdxf(rc, style, theme){
  if (!style) return;
  try {
    if (style.fill && style.fill.fgColor){
      var fr = resolveColor(style.fill.fgColor, theme);
      if (fr) rc.fill = fr;
    }
    if (style.font){
      var fc = resolveColor(style.font.color || style.font.colour, theme);
      if (fc) rc.font.color = fc;
      if (style.font.bold){ rc.font.bold = true; rc.bold = true; }
      if (style.font.italic){ rc.font.italic = true; rc.italic = true; }
    }
    if (style.border){
      ['top', 'right', 'bottom', 'left'].forEach(function(side){
        var bs = style.border[side];
        if (bs && bs.style && bs.style !== 'none'){
          rc.border[side] = {
            style: String(bs.style),
            color: resolveColor(bs.color || bs.colour, theme)
          };
        }
      });
    }
  } catch(e){}
}

/* Main driver: evaluate every CF rule on the sheet and merge winners.
 * ws: ExcelJS worksheet; rows: rich cells; ncols: used columns;
 * evals: per-cell {ok,v} from evaluateFormulas (may be null in raw mode). */
function applyConditionalFormatting(ws, rows, ncols, sheetName, evals, FPClass, theme, stats){
  var cfs = ws.conditionalFormattings;
  if (!cfs || !FPClass) return;
  var nrows = rows.length;
  if (!nrows || !ncols) return;
  function getVal(ri, ci){
    var ev = evals && evals[ri] && evals[ri][ci];
    if (ev && ev.ok){
      var v = ev.v;
      return (v === '' || v === undefined || v === null) ? undefined : v;
    }
    var cell = null;
    try { cell = ws.getRow(ri + 1).getCell(ci + 1); } catch(e){}
    if (!cell) return undefined;
    var cv = cell.value;
    if (cv == null) return undefined;
    if (cv instanceof Date) return cv;
    if (typeof cv === 'object'){
      if (cv.result !== undefined && cv.result !== null){
        if (cv.result && typeof cv.result === 'object' && cv.result.error) return undefined;
        return cv.result;
      }
      return undefined;
    }
    return cv;
  }
  var parser;
  try {
    parser = new FPClass({
      onCell: function(ref){ return getVal(ref.row - 1, ref.col - 1); },
      onRange: function(ref){
        var arr = [];
        for (var r = ref.from.row; r <= ref.to.row; r++){
          var inner = [];
          for (var c = ref.from.col; c <= ref.to.col; c++) inner.push(getVal(r - 1, c - 1));
          arr.push(inner);
        }
        return arr;
      }
    });
  } catch(e){ return; }
  var rules = [];
  Object.keys(cfs).forEach(function(k){
    var b = cfs[k];
    var R = parseCFRange(b.ref);
    if (!R) return;
    /* clamp to the used area */
    R.r2 = Math.min(R.r2, nrows - 1); R.c2 = Math.min(R.c2, ncols - 1);
    if (R.r2 < R.r1 || R.c2 < R.c1) return;
    (b.rules || []).forEach(function(rule){ rules.push({ R: R, rule: rule }); });
  });
  if (!rules.length) return;
  stats.cfRules = (stats.cfRules || 0) + rules.length;
  rules.sort(function(a, b){ return (a.rule.priority || 0) - (b.rule.priority || 0); });
  var done = {};
  rules.forEach(function(entry){
    var R = entry.R, rule = entry.rule;
    for (var r = R.r1; r <= R.r2; r++){
      for (var c = R.c1; c <= R.c2; c++){
        var key = r + '|' + c;
        if (done[key]) continue; /* first hit in priority order wins */
        var res = null;
        try { res = evalCFRule(rule, r, c, R, getVal, parser, sheetName, theme); }
        catch(e){ res = null; }
        if (res){
          var rc = rows[r] && rows[r][c];
          if (rc && !rc.skip){
            if (res.dxf) applyCFdxf(rc, res.dxf, theme);
            if (res.colorScaleFill) rc.fill = res.colorScaleFill;
            if (res.dataBar) rc.dataBar = res.dataBar;
            if (res.iconSet){ rc.iconSet = res.iconSet; rc.iconPadU = 14; }
            stats.cfCells = (stats.cfCells || 0) + 1;
          }
          done[key] = true;
        }
      }
    }
  });
}

var api = {
  readBuffer: readBuffer,
  readViaExcelJS: readViaExcelJS,
  readViaSheetJS: readViaSheetJS,
  richFromExcelCell: richFromExcelCell,
  argbToRgb: argbToRgb,
  parseMergeRef: parseMergeRef,
  colLettersToIndex: colLettersToIndex,
  ssfFormat: ssfFormat,
  emptyRich: emptyRich,
  parseThemeColors: parseThemeColors,
  resolveColor: resolveColor,
  applyTint: applyTint,
  themeFromWorkbook: themeFromWorkbook,
  evaluateFormulas: evaluateFormulas,
  rawOfExcelCell: rawOfExcelCell,
  FORMULA_POLYFILLS: FORMULA_POLYFILLS,
  applyConditionalFormatting: applyConditionalFormatting,
  evalCFRule: evalCFRule,
  shiftFormula: shiftFormula,
  parseCFRange: parseCFRange,
  parseHFPictures: parseHFPictures
};

if (typeof module !== 'undefined' && module.exports){ module.exports = api; }
else if (typeof globalThis !== 'undefined'){ globalThis.RichXlsx = api; }
})();
