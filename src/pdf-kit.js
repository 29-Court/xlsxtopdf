/* 29Court Convert: shared pdf-lib primitives (pdf-lib injected as global PDFLib).
 * Used by app.js and xlsx-print.js so both render through identical code.
 * DOM-free; loadable in Node for tests (PDFLib must be set as a global). */
(function(){
"use strict";
var Kit = {
  A4: [595.28, 841.89],
  LETTER: [612, 792],
  doc: function(){ return PDFLib.PDFDocument.create(); },
  fonts: null,
  loadFonts: function(d){
    var self = this;
    if (self.fonts) return Promise.resolve(self.fonts);
    return Promise.all([
      d.embedFont(PDFLib.StandardFonts.Helvetica),
      d.embedFont(PDFLib.StandardFonts.HelveticaBold),
      d.embedFont(PDFLib.StandardFonts.HelveticaOblique),
      d.embedFont(PDFLib.StandardFonts.HelveticaBoldOblique),
      d.embedFont(PDFLib.StandardFonts.Courier)
    ]).then(function(f){
      self.fonts = { reg: f[0], bold: f[1], italic: f[2], bi: f[3], mono: f[4] };
      return self.fonts;
    });
  },
  wrap: function(text, font, size, maxW, meas){
    /* meas(text)->width override lets callers measure with mixed fonts
     * (e.g. Latin + CJK runs). Defaults to the given font. */
    meas = meas || function(t){
      try { return font.widthOfTextAtSize(t, size); }
      catch(e){ return t.length * size * 0.55; }
    };
    var words = String(text == null ? '' : text).split(/\s+/), lines = [], line = '';
    words.forEach(function(w){
      var t = line ? line + ' ' + w : w;
      if (meas(t) <= maxW){ line = t; return; }
      /* CJK has no spaces: break an over-wide word between characters. */
      if (/[\u3040-\u30FF\u4E00-\u9FFF]/.test(w)){
        if (line){ lines.push(line); line = ''; }
        var cur = '';
        for (var i = 0; i < w.length; i++){
          var t2 = cur + w[i];
          if (cur && meas(t2) > maxW){ lines.push(cur); cur = w[i]; }
          else cur = t2;
        }
        line = cur;
      } else {
        if (line) lines.push(line);
        line = w;
      }
    });
    if (line) lines.push(line);
    return lines.length ? lines : [''];
  },
  /* ---- CJK font support ----
   * Helvetica is WinAnsi-only. When a sheet contains CJK text (hiragana,
   * katakana, CJK ideographs), we lazy-embed Noto Sans SC with subsetting:
   * only the glyphs used end up in the PDF (~0.04% of the 10.5MB source).
   * SC (not JP) is the CJK font because it covers simplified + traditional
   * Chinese AND Japanese kana/kanji; Noto Sans JP lacks simplified-only
   * characters (计, 语, ...). Korean Hangul is not covered (honest out);
   * some Han characters render in their Chinese glyph variant.
   * Font bytes come from env.cjkFontBytes(script) when provided
   * (Node/Lambda/tests), otherwise from the pinned jsDelivr URL in browsers.
   * Bold/italic CJK renders in the regular weight (documented honest limit).
   * Never throws: on any failure the caller keeps the □-degradation path. */
  CJK_URL: 'https://cdn.jsdelivr.net/npm/@expo-google-fonts/noto-sans-sc@0.4.4/400Regular/NotoSansSC_400Regular.ttf',
  CJK_RE: /[\u3040-\u30FF\u4E00-\u9FFF]/,
  _cjkBytes: null,
  _cjkBytesFetch: function(env){
    var self = this;
    if (self._cjkBytes) return Promise.resolve(self._cjkBytes);
    function done(b){
      if (b && b.length > 1000){ self._cjkBytes = b; return b; }
      return null;
    }
    var p = null;
    try {
      if (env && typeof env.cjkFontBytes === 'function') p = Promise.resolve(env.cjkFontBytes('jp'));
      else if (typeof fetch !== 'undefined') p = fetch(self.CJK_URL).then(function(r){
        if (!r.ok) return null;
        return r.arrayBuffer().then(function(ab){ return new Uint8Array(ab); });
      });
    } catch(e){ p = null; }
    if (!p) return Promise.resolve(null);
    return p.then(done, function(){ return null; });
  },
  hasCJKText: function(sheets){
    for (var si = 0; si < sheets.length; si++){
      var rows = sheets[si].rows;
      for (var r = 0; r < rows.length; r++){
        var row = rows[r];
        for (var c = 0; c < row.length; c++){
          var t = row[c] && row[c].t;
          if (t && this.CJK_RE.test(t)) return true;
        }
      }
    }
    return false;
  },
  /* Embed the CJK font into doc d. env.fontkit required (pdf-lib needs it
   * for non-WinAnsi fonts). Returns the embedded font or null. The font is
   * per-document; the raw bytes are cached across documents. */
  loadCJK: function(d, env){
    var self = this;
    return self._cjkBytesFetch(env).then(function(bytes){
      if (!bytes || !env || !env.fontkit) return null;
      try { d.registerFontkit(env.fontkit); } catch(e){ return null; }
      return d.embedFont(bytes, { subset: true }).then(function(f){ return f; },
        function(){ return null; });
    });
  }
};
if (typeof module !== 'undefined' && module.exports){ module.exports = Kit; }
else if (typeof globalThis !== 'undefined'){ globalThis.PdfKit = Kit; }
})();
