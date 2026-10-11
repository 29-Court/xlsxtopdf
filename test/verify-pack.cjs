/* Packaged-build verification for @29court/xlsxtopdf.
 * Run from a scratch dir with the packed tarball installed:
 *   npm install /path/to/29court-xlsxtopdf-0.1.0.tgz
 *   node /path/to/test/verify-pack.cjs
 * Env: CORPUS=/path/to/23-file/corpus METHODA=/path/to/methodA_single.xlsx
 */
"use strict";
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");
const { xlsxToPdf, parseWorkbook } = require("@29court/xlsxtopdf");
const PDFLib = require("pdf-lib");

const CORPUS = process.env.CORPUS;
const METHODA = process.env.METHODA;
if (!CORPUS || !METHODA) {
  console.error("need CORPUS and METHODA env vars");
  process.exit(2);
}

let pass = 0, fail = 0;
function ok(cond, name, extra) {
  if (cond) { pass++; /* console.log("  PASS " + name); */ }
  else { fail++; console.log("  FAIL " + name + (extra ? " :: " + extra : "")); }
}

/* crude PDF text extraction: uncompress streams, collect (...) and <...> runs */
async function pdfInfo(bytes) {
  const pdf = await PDFLib.PDFDocument.load(bytes);
  let text = "";
  const re = /stream\r?\n([\s\S]*?)endstream/g;
  let m;
  const latin = Buffer.from(bytes).toString("latin1");
  while ((m = re.exec(latin))) {
    let chunk = Buffer.from(m[1], "latin1");
    try { chunk = zlib.inflateSync(chunk); } catch (e) {}
    const s = chunk.toString("latin1");
    const t1 = [...s.matchAll(/\((?:\\.|[^\\()])*\)/g)].map((x) => x[0].slice(1, -1).replace(/\\(.)/g, "$1"));
    const t2 = [...s.matchAll(/<([0-9A-Fa-f]+)>/g)].map((x) => Buffer.from(x[1], "hex").toString("latin1"));
    text += t1.join(" ") + " " + t2.join(" ") + "\n";
  }
  const sizes = pdf.getPages().map((p) => p.getSize());
  return { pages: pdf.getPageCount(), text, sizes };
}

(async () => {
  console.log("== 1. corpus: 23/23 convert, 0 crashes ==");
  const files = fs.readdirSync(CORPUS).filter((f) => f.endsWith(".xlsx")).sort();
  ok(files.length === 23, "corpus has 23 files", "found " + files.length);
  let converted = 0;
  for (const f of files) {
    try {
      const pdf = await xlsxToPdf(fs.readFileSync(path.join(CORPUS, f)), { cover: false });
      const isPdf = pdf instanceof Uint8Array && Buffer.from(pdf.slice(0, 5)).toString() === "%PDF-";
      const info = await pdfInfo(pdf);
      if (isPdf && info.pages >= 1) converted++;
      else console.log("    bad output: " + f);
    } catch (e) {
      console.log("    threw: " + f + " (" + e.message + ")");
    }
  }
  ok(converted === 23, "23/23 converted to valid PDFs", converted + "/23");

  console.log("== 2. Method A: all 41 formulas evaluated ==");
  const mbuf = fs.readFileSync(METHODA);
  const { stats } = await parseWorkbook(mbuf);
  ok(stats.sheets === 3, "3 sheets parsed", stats.sheets);
  ok(stats.evaluated === 41, "41 formulas evaluated on-device", stats.evaluated);
  ok((stats.noCache | 0) === 0, "0 uncalculated formulas", stats.noCache);
  ok((stats.evalFailed | 0) === 0, "0 evaluation failures", stats.evalFailed);

  console.log("== 3. Method A: computed values in PDF, zero raw formulas ==");
  const mpdf = await xlsxToPdf(mbuf, { filename: "methodA_single.xlsx" });
  const mi = await pdfInfo(mpdf);
  ok(mi.pages === 4, "4 pages (cover + 3 sheets)", mi.pages);
  const t = mi.text;
  const expected = ["32,890.15", "27,265.55", "119.75", "27,385.30", "989.75",
    "31,900.39", "28,879.55", "245.01", "29,124.56", "65,780.30", "64,790.54",
    "56,145.10", "364.76", "56,509.86", "3,411.97", "2,092.87", "2,775.83",
    "1.07642", "1.09531", "29,478.18", "1.05559", "31,116.84", "33,673.70",
    "1.13626", "1.15620", "1.14654", "4,868.70", "61,378.57"];
  const missing = expected.filter((v) => !t.includes(v));
  ok(missing.length === 0, "all " + expected.length + " computed values present",
    missing.length ? "missing: " + missing.join(",") : "");
  ok(!/=[A-Z]+\(/.test(t), "zero raw =FORMULA( text in PDF");
  ok(!t.includes("undefined"), "no undefined leaks");

  console.log("== 4. options ==");
  const letter = await xlsxToPdf(mbuf, { cover: false, page: "letter" });
  const li = await pdfInfo(letter);
  const lw = li.sizes[0].width, lh = li.sizes[0].height;
  ok((Math.abs(lw - 612) < 1 && Math.abs(lh - 792) < 1) ||
     (Math.abs(lw - 792) < 1 && Math.abs(lh - 612) < 1),
    "page:'letter' -> letter dimensions (landscape for this wide sheet)",
    JSON.stringify(li.sizes[0]));
  const letterPor = await xlsxToPdf(mbuf, { cover: false, page: "letter", orientation: "portrait" });
  const lpi = await pdfInfo(letterPor);
  ok(Math.abs(lpi.sizes[0].width - 612) < 1 && Math.abs(lpi.sizes[0].height - 792) < 1,
    "page:'letter' + portrait -> 612x792", JSON.stringify(lpi.sizes[0]));
  const land = await xlsxToPdf(mbuf, { cover: false, orientation: "landscape" });
  const ldi = await pdfInfo(land);
  ok(ldi.sizes[0].width > ldi.sizes[0].height, "orientation:'landscape' honored");
  const nocover = await xlsxToPdf(mbuf, { cover: false });
  ok((await pdfInfo(nocover)).pages === 3, "cover:false -> 3 content pages");
  const one = await xlsxToPdf(mbuf, { cover: false, includeSheets: ["Floor 1"] });
  const oi = await pdfInfo(one);
  ok(oi.text.includes("Floor 1") && !oi.text.includes("56,509.86"),
    "includeSheets filters to Floor 1");
  const raw = await xlsxToPdf(mbuf, { cover: false, evaluateFormulas: false });
  ok(/=[A-Z]+\(/.test((await pdfInfo(raw)).text), "evaluateFormulas:false -> raw =FORMULA text");

  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error("HARNESS FAIL", e); process.exit(1); });
