/* @29court/xlsxtopdf — XLSX to beautiful PDF in pure JavaScript.
 *
 * Bytes in, bytes out. Unzips the workbook, evaluates every formula
 * on-device (even workbooks with no cached values), typesets each sheet
 * like a printer would, and draws the PDF directly. No browser, no
 * headless Chrome, no LibreOffice, no WASM, no native code, no DOM.
 * Runs in Node, in browsers (via bundler), and in AWS Lambda.
 *
 * Dependencies are MIT/Apache-2.0 only (exceljs, xlsx, pdf-lib,
 * fast-formula-parser). The formula engine is fast-formula-parser, not
 * HyperFormula — HyperFormula is GPL-3.0, which would copyleft this
 * package. That decision is permanent.
 */
"use strict";

var ExcelJS = require("exceljs");
var XLSX = require("xlsx");
var PDFLib = require("pdf-lib");
var FormulaParser = require("fast-formula-parser");
var JSZip = require("jszip");
var fontkit = require("@pdf-lib/fontkit");

var RichXlsx = require("./xlsx-rich.js");
var XlsxPrint = require("./xlsx-print.js");
var PdfKit = require("./pdf-kit.js");
var XlsxCharts = require("./xlsx-charts.js");

/* pdf-kit.js is browser-script heritage: it reads PDFLib as a global at
 * call time. Provide it when the host hasn't. */
if (typeof globalThis !== "undefined" && !globalThis.PDFLib) {
  globalThis.PDFLib = PDFLib;
}

function toBytes(input, name) {
  if (input == null) throw new TypeError("xlsxtopdf: " + (name || "input") + " is required");
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (typeof input === "string") {
    throw new TypeError("xlsxtopdf: expected binary bytes (Uint8Array/Buffer/ArrayBuffer), got a string path. Read the file first, e.g. fs.readFileSync(path).");
  }
  throw new TypeError("xlsxtopdf: expected Uint8Array, Buffer, or ArrayBuffer");
}

function dateStrOf(d) {
  if (d == null) return new Date().toLocaleDateString();
  if (d instanceof Date) return d.toLocaleDateString();
  return String(d);
}

function normalizePage(p) {
  if (Array.isArray(p) && p.length === 2 && +p[0] > 0 && +p[1] > 0)
    return [+p[0], +p[1]];
  var s = String(p || "a4").toLowerCase();
  return s === "letter" ? "letter" : "a4";
}

function normalizeOrientation(o) {
  var s = String(o || "auto").toLowerCase();
  return (s === "portrait" || s === "landscape") ? s : "auto";
}

/* Parse a workbook into render-ready sheets (rich cells: text, fonts,
 * fills, borders, alignment, merges, column widths, number formats).
 * options: { evaluateFormulas=true, log }
 * Returns { sheets, stats }. stats.evaluated counts formulas the engine
 * calculated on-device; stats.noCache counts formulas left blank. */
async function parseWorkbook(xlsxBytes, options) {
  options = options || {};
  var buf = toBytes(xlsxBytes, "xlsxBytes");
  var mode = options.evaluateFormulas === false ? "raw" : "final";
  return RichXlsx.readBuffer(buf, mode, {
    ExcelJS: ExcelJS,
    XLSX: XLSX,
    FormulaParser: FormulaParser,
    JSZip: JSZip,
    XlsxCharts: XlsxCharts,
    log: typeof options.log === "function" ? options.log : undefined
  });
}

/* Render parsed sheets to PDF bytes.
 * options: { filename, date, page='a4'|'letter'|[w,h], orientation='auto'|'portrait'|'landscape',
 *            margins, cover=true, coverTitle, coverTagline, logoB64, logoCover, logoFooter,
 *            header='auto'|'none'|{left,center,right}, footer='auto'|'none'|'default'|{left,center,right},
 *            headerStyle, footerStyle, useWorkbookHeaders=true,
 *            includeSheets, sheetOrder, cjkFontBytes }
 * header/footer 'auto' (default): the workbook's own print headers/footers
 *   win when the file defines them; otherwise the built-in default footer
 *   (filename, date, Page N of M) and no header. 'none' disables.
 *   {left,center,right} gives custom text; Excel & codes (&P &N &D &T &F &A,
 *   &"Font,Bold" &size &B &I &U &KRRGGBB &[Picture] &&) are honored.
 * headerStyle/footerStyle: { font='reg'|'bold'|'italic'|'bi'|'mono', size,
 *   color ([r,g,b] or '#rrggbb'), rule, ruleColor, pictureHeight }.
 * margins: { top, right, bottom, left } in points (default 36 each).
 * sheetOrder: array of sheet names; listed sheets render first in that
 *   order, unlisted sheets keep their relative order after.
 * cjkFontBytes: async (script) => Uint8Array of the Noto Sans SC TTF, for
 *   CJK text. If omitted and CJK text is present, the browser build fetches
 *   it from a pinned CDN URL; Node/Lambda callers should supply it to avoid
 *   the network fetch. See README for the asset path. */
async function renderPdf(sheets, options) {
  options = options || {};
  var pdfBytes = await XlsxPrint.sheetsToPdf(sheets, {
    filename: options.filename || "workbook.xlsx",
    dateStr: dateStrOf(options.date),
    logoB64: options.logoB64 || null,
    logoCover: options.logoCover,
    logoFooter: options.logoFooter,
    pageSize: normalizePage(options.page),
    orientation: normalizeOrientation(options.orientation),
    margins: options.margins,
    cover: options.cover !== false,
    coverTitle: options.coverTitle,
    coverTagline: options.coverTagline,
    header: options.header,
    footer: options.footer,
    headerStyle: options.headerStyle,
    footerStyle: options.footerStyle,
    useWorkbookHeaders: options.useWorkbookHeaders,
    includeSheets: options.includeSheets,
    sheetOrder: options.sheetOrder
  }, { PDFLib: PDFLib, PDF: PdfKit, fontkit: fontkit,
       XlsxCharts: XlsxCharts, cjkFontBytes: options.cjkFontBytes });
  return pdfBytes instanceof Uint8Array ? pdfBytes : new Uint8Array(pdfBytes);
}

/* One call: xlsx bytes in, PDF bytes out.
 * options: { filename='workbook.xlsx', date, page='a4'|'letter'|[w,h],
 *   orientation='auto'|'portrait'|'landscape', margins, fit='width',
 *   evaluateFormulas=true, includeSheets, sheetOrder, cover=true,
 *   coverTitle, coverTagline, logoB64, logoCover, logoFooter,
 *   header='auto', footer='auto', headerStyle, footerStyle,
 *   useWorkbookHeaders=true, log }
 * includeSheets: array of sheet names; only those sheets render.
 * sheetOrder: array of sheet names; reorders after includeSheets filtering.
 * fit: 'width' is the behavior — fit all columns to the page width in the
 *   best orientation, tiling below the readability floor. (Reserved for
 *   future modes; anything else is accepted and ignored.)
 * cover: the 29Court-branded cover page. Set false for a neutral document
 *   (footers then number from page 1).
 * Returns Promise<Uint8Array>. */
async function xlsxToPdf(xlsxBytes, options) {
  options = options || {};
  var data = await parseWorkbook(xlsxBytes, options);
  return renderPdf(data.sheets, options);
}

function hostOf(url) {
  var m = String(url).match(/^https?:\/\/([^\/:?#]+)/i);
  return m ? m[1] : String(url);
}

/* Download an xlsx from a URL, then convert it. Never hangs (default 30s
 * timeout, configurable via options.timeout) and never silently produces
 * garbage: DNS failures, HTTP errors, timeouts, oversized responses, and
 * non-xlsx responses all throw honest Errors.
 * options: same as xlsxToPdf, plus { timeout=30000, maxBytes=104857600 }.
 * When options.filename is omitted, it is derived from the URL path. */
async function fetchXlsxBytes(url, options) {
  var timeoutMs = options.timeout != null ? +options.timeout : 30000;
  var maxBytes = options.maxBytes != null ? +options.maxBytes : 100 * 1024 * 1024;
  if (typeof url !== "string" || !/^https?:\/\//i.test(url))
    throw new Error("xlsxToPdfUrl: URL must start with http:// or https://, got " +
      JSON.stringify(url));
  var fetchFn = (typeof fetch !== "undefined") ? fetch : null;
  if (!fetchFn)
    throw new Error("xlsxToPdfUrl: no fetch() available in this environment " +
      "(Node 18+ or a browser is required); pass xlsx bytes to xlsxToPdf() instead");
  var ctrl = null, timer = null;
  try { ctrl = new AbortController(); } catch (e) { ctrl = null; }
  if (ctrl && timeoutMs > 0)
    timer = setTimeout(function () { try { ctrl.abort(); } catch (e) {} }, timeoutMs);
  var res;
  try {
    res = await fetchFn(url, ctrl ? { signal: ctrl.signal } : {});
  } catch (e) {
    if (timer) clearTimeout(timer);
    if (ctrl && ctrl.signal && ctrl.signal.aborted)
      throw new Error("xlsxToPdfUrl: timed out after " + timeoutMs + "ms waiting for " +
        hostOf(url) + " (options.timeout)");
    throw new Error("xlsxToPdfUrl: could not reach " + hostOf(url) +
      " (" + String((e && e.message) || e) + ")");
  }
  if (timer) clearTimeout(timer);
  if (!res.ok)
    throw new Error("xlsxToPdfUrl: download failed with HTTP " + res.status +
      " " + (res.statusText || "") + " for " + url);
  var cl = null;
  try { cl = res.headers.get("content-length"); } catch (e) {}
  if (cl != null && +cl > maxBytes)
    throw new Error("xlsxToPdfUrl: file is " + cl + " bytes, over the " + maxBytes +
      "-byte limit (options.maxBytes)");
  var ab = await res.arrayBuffer();
  var bytes = new Uint8Array(ab);
  if (bytes.length > maxBytes)
    throw new Error("xlsxToPdfUrl: downloaded " + bytes.length + " bytes, over the " +
      maxBytes + "-byte limit (options.maxBytes)");
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4B)
    throw new Error("xlsxToPdfUrl: response is not an xlsx file (bad zip signature) for " + url);
  return bytes;
}

async function xlsxToPdfUrl(url, options) {
  options = options || {};
  var buf = await fetchXlsxBytes(url, options);
  if (!options.filename) {
    try {
      var base = new URL(url).pathname.split("/").pop();
      if (base) options = Object.assign({}, options,
        { filename: decodeURIComponent(base) });
    } catch (e) {}
  }
  return xlsxToPdf(buf, options);
}

module.exports = {
  xlsxToPdf: xlsxToPdf,
  xlsxToPdfUrl: xlsxToPdfUrl,
  parseWorkbook: parseWorkbook,
  renderPdf: renderPdf,
  /* lower-level core, for advanced use */
  layoutTiles: XlsxPrint.layoutTiles,
  version: "0.2.0"
};
