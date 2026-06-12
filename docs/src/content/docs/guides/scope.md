---
title: Scope & limitations
description: What Ream implements today, and where the edges are.
---

The conversion core is broad and spec-driven, but ECMA-376 is vast and Ream does
not implement all of it — this page is an honest map of what works and what
doesn't yet.

## Implemented

**Output** — `convert('pdf')`, `convert('svg')` (a page-stack preview),
`convert('html')` (flowed, needs no fonts) and `convert('docx')` (write
WordprocessingML back out — for normalization, sanitization or in-browser
editing, and as a docx → docx round-trip). The docx round-trip is semantic, not
byte-exact; footnotes, charts and OfficeMath are not yet written.

**WordprocessingML (§17)**
- Text, runs and the full style cascade (`docDefaults` → styles → direct formatting).
- Tables — auto / fixed layout, §17.4 border-conflict resolution, cell shading,
  vertical merge and grid span, nested tables, **table styles** (`w:tblStyle` with
  conditional formats: banding, first/last row/column).
- Lists and numbering (`abstractNum`, level overrides), multi-level.
- Sections — per-section page size and orientation, headers and footers,
  **multi-column layout** (`w:cols`).
- **Hyperlinks** — external (clickable PDF annotations + HTML `<a>`, scheme-allowlisted)
  and internal: bookmarks become named destinations / `#`-anchors.
- **Fields** — `PAGE` / `NUMPAGES` render real page numbers in headers and footers.
- **Footnotes and endnotes** — notes at the bottom of the referencing page behind
  Word's separator rule; endnotes after the body.
- Inline and floating images (PNG / JPEG / JPEG2000), including **legacy VML
  pictures** (`<w:pict>` / `<w:object>` — ActiveX and OLE-object previews,
  images from older Word); floating drawings (`wp:anchor`) render outside the
  text flow — wrap-none (incl. `behindDoc`) for watermarks/stamps/text boxes,
  and side wrapping (`square`/`tight`/`through`) where the body text flows
  around the exclusion area.
- Tracked changes (`w:ins` / `w:del`).
- Reads both **Transitional and Strict** (ISO 29500) packages; block-level
  content controls (`w:sdt`) flow through.

**SpreadsheetML (§18)**
- Grids, shared strings, number formats and dates (incl. the 1904 date system).
- The print model — gridline suppression, print area, fit-to-page scaling, repeated
  print titles, manual page breaks, horizontal/vertical centering.
- Charts anchored to the sheet (the worksheet drawing part) render after the grid.

**Graphics & math**
- DrawingML shapes (preset and custom geometry, gradients, group shapes, theme colors).
- Charts — bar/column, line, pie/doughnut, area, scatter, stacked.
- OfficeMath — fractions, scripts, radicals, n-ary operators, functions, limits,
  delimiters, matrices, accents; inline and display.

**Typography**
- Type0 + CIDFontType2 embedding with subsetting.
- Knuth–Plass line breaking, Liang hyphenation (en / ru).
- OpenType ligatures and kerning (GSUB/GPOS), mark positioning.
- BiDi (UAX #9), Arabic cursive joining.

**PDF / compliance**
- PDF/A-1, -2, -3 at levels a / b / u — all formally **veraPDF-validated**.
- **PDF/UA-1** (ISO 14289-1) — veraPDF-validated, alone or combined with PDF/A-2a
  in a single file.
- Tagged PDF — logical structure tree, headings, tables, lists (with `Lbl`
  markers), figures with alt text, links with alternate descriptions, footnote
  `Note` elements, `/Lang`, pagination artifacts.
- **Encryption** — AES-256 (ISO 32000-2 R6) via WebCrypto, with permission flags.
- Digital signatures — PKCS#7 detached, ECDSA; object streams; JPEG2000 images.

## Not yet

- **Legacy `.doc` / `.xls`** (the binary OLE/CFB formats) — Ream parses only the OOXML
  ZIP+XML formats. Re-save as `.docx` / `.xlsx`.
- **Font measurement parity with a specific renderer.** Ream lays out correctly for the
  font you supply; pixel-exact agreement with another tool requires that tool's exact
  font metrics.
- Assorted deep-tail OOXML features are tracked as the project grows.

## Validation

Development is corpus-driven: documents are converted, compared against a LibreOffice
"golden" render (structural text diff + rasterized visual diff), and PDF/A output is
gated through veraPDF. Untrusted corpus files run inside a locked-down Docker sandbox.
