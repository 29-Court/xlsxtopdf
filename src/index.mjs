/* @29court/xlsxtopdf — ESM entry.
 * Single source of truth is ./index.cjs; this re-exports it so bundlers,
 * Lambda ESM handlers, and `import` users get the same API. */
import pkg from "./index.cjs";

const { xlsxToPdf, xlsxToPdfUrl, parseWorkbook, renderPdf, layoutTiles, version } = pkg;

export { xlsxToPdf, xlsxToPdfUrl, parseWorkbook, renderPdf, layoutTiles, version };
export default pkg;
