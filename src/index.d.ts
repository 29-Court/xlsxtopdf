export type PageSize = "a4" | "letter" | [number, number];
export type Orientation = "auto" | "portrait" | "landscape";

/** Page margins in points. */
export interface Margins {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}

/** Custom header/footer content. Each section is a string; Excel & codes
 *  are honored: &L &C &R, &"Font,Bold", &12, &B &I &U, &KRRGGBB, &P (page),
 *  &N (total), &D (date), &T (time), &F (filename), &A (sheet), &[Picture],
 *  && (literal &). */
export interface CustomHeaderFooter {
  left?: string;
  center?: string;
  right?: string;
}

/** Header/footer visual style. Codes inside the content (&B, &12, &KRRGGBB)
 *  add to these base values. */
export interface HeaderFooterStyle {
  /** Base typeface: 'reg' | 'bold' | 'italic' | 'bi' | 'mono'. Default
   *  header 'reg', footer 'reg'. */
  font?: "reg" | "bold" | "italic" | "bi" | "mono" | string;
  /** Base size in points. Default header 10, footer 8. */
  size?: number;
  /** Base color: [r,g,b] in 0..1 or '#rrggbb'. */
  color?: [number, number, number] | string;
  /** Draw a horizontal rule with the block. Default header false, footer true. */
  rule?: boolean;
  /** Rule color: [r,g,b] or '#rrggbb'. Default gold. */
  ruleColor?: [number, number, number] | string;
  /** Height of &[Picture] images in points. Default header 22, footer 16. */
  pictureHeight?: number;
}

export interface XlsxToPdfOptions {
  /** Shown on the cover page and footer. Default 'workbook.xlsx'. */
  filename?: string;
  /** Conversion date. Default now. */
  date?: Date | string;
  /** 'a4' (default), 'letter', or [widthPt, heightPt]. */
  page?: PageSize;
  /** 'auto' (default, picks the orientation with the largest type),
   *  'portrait', or 'landscape'. */
  orientation?: Orientation;
  /** Page margins in points. Default { top: 36, right: 36, bottom: 36, left: 36 }. */
  margins?: Margins;
  /** 'width' (default): fit all columns to the page width, tiling across
   *  pages below the readability floor instead of shrinking to ants. */
  fit?: "width";
  /** Calculate formulas on-device (default true). Set false for raw =FORMULA text. */
  evaluateFormulas?: boolean;
  /** Only render these sheet names. Default: all sheets. */
  includeSheets?: string[];
  /** Render order for sheets. Listed sheets come first in this order;
   *  unlisted sheets keep their relative order after. Applies after
   *  includeSheets filtering. */
  sheetOrder?: string[];
  /** 29Court-branded cover page. Default true; set false for a neutral
   *  document (footers then number from page 1). */
  cover?: boolean;
  /** Cover title. Default '29COURT'; set '' to hide. */
  coverTitle?: string;
  /** Cover tagline. Default 'Made with 29Court Convert. Your file never left
   *  this device.'; set '' to hide. */
  coverTagline?: string;
  /** Base64 PNG logo for the cover/footer. Default none. */
  logoB64?: string | null;
  /** Show the logo on the cover. Default true. */
  logoCover?: boolean;
  /** Show the logo in the built-in default footer. Default true.
   *  (Workbook-defined or custom footers carry their own &[Picture].) */
  logoFooter?: boolean;
  /** Page header: 'auto' (default, honors the workbook's own print header
   *  when defined, else none), 'none', or { left, center, right } custom. */
  header?: "auto" | "none" | CustomHeaderFooter;
  /** Page footer: 'auto' (default, honors the workbook's own print footer
   *  when defined, else the built-in filename/date/Page N of M footer),
   *  'none', 'default' (always the built-in footer), or
   *  { left, center, right } custom. */
  footer?: "auto" | "none" | "default" | CustomHeaderFooter;
  /** Visual style for the header block. */
  headerStyle?: HeaderFooterStyle;
  /** Visual style for the footer block. */
  footerStyle?: HeaderFooterStyle;
  /** Honor the workbook's own print headers/footers (default true).
   *  Set false to ignore them (header/footer 'auto' then behaves like
   *  'none'/'default'). */
  useWorkbookHeaders?: boolean;
  /** Optional log sink: (message, level) => void. */
  log?: (message: string, level?: string) => void;
}

export interface XlsxToPdfUrlOptions extends XlsxToPdfOptions {
  /** Download timeout in ms. Default 30000. */
  timeout?: number;
  /** Maximum download size in bytes. Default 104857600 (100MB). */
  maxBytes?: number;
}

export interface WorkbookStats {
  sheets: number;
  evaluated: number;
  evalFailed: number;
  noCache: number;
  themeColors: boolean;
  [key: string]: unknown;
}

export interface ParsedSheet {
  name: string;
  rows: unknown[][];
  /** Raw print header/footer strings from the workbook (null when the file
   *  defines none), plus VML-ordered &[Picture] images. */
  printHF: {
    oddHeader: string | null;
    oddFooter: string | null;
    evenHeader: string | null;
    evenFooter: string | null;
    firstHeader: string | null;
    firstFooter: string | null;
    differentFirst: boolean;
    differentOddEven: boolean;
    images: Array<{ data: Uint8Array; ext: string }>;
  } | null;
  [key: string]: unknown;
}

export function xlsxToPdf(
  xlsxBytes: Uint8Array | ArrayBuffer | Buffer,
  options?: XlsxToPdfOptions
): Promise<Uint8Array>;

/** Download an xlsx from a URL and convert it. Honest errors on DNS
 *  failure, HTTP errors, timeouts, oversize, and non-xlsx responses. */
export function xlsxToPdfUrl(
  url: string,
  options?: XlsxToPdfUrlOptions
): Promise<Uint8Array>;

export function parseWorkbook(
  xlsxBytes: Uint8Array | ArrayBuffer | Buffer,
  options?: Pick<XlsxToPdfOptions, "evaluateFormulas" | "log">
): Promise<{ sheets: ParsedSheet[]; stats: WorkbookStats }>;

export function renderPdf(
  sheets: ParsedSheet[],
  options?: XlsxToPdfOptions
): Promise<Uint8Array>;

export function layoutTiles(
  colWpt: number[],
  usableW: number,
  mergeIvs?: Array<[number, number]>
): Array<{ c0: number; c1: number; k: number }>;

export const version: string;

declare const _default: {
  xlsxToPdf: typeof xlsxToPdf;
  xlsxToPdfUrl: typeof xlsxToPdfUrl;
  parseWorkbook: typeof parseWorkbook;
  renderPdf: typeof renderPdf;
  layoutTiles: typeof layoutTiles;
  version: typeof version;
};
export default _default;
