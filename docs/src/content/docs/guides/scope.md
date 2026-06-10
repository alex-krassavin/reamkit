---
title: Scope & limitations
description: What Ream implements today, and where the edges are.
---

Ream is an **alpha**. The conversion core is broad and spec-driven, but it is still
maturing — this page is an honest map of what works and what doesn't yet.

## Implemented

**WordprocessingML (§17)**
- Text, runs and the full style cascade (`docDefaults` → styles → direct formatting).
- Tables — auto / fixed layout, §17.4 border-conflict resolution, cell shading,
  vertical merge and grid span, nested tables.
- Lists and numbering (`abstractNum`, level overrides), multi-level.
- Sections — per-section page size and orientation, headers and footers.
- Inline and floating images (PNG / JPEG / JPEG2000).
- Tracked changes (`w:ins` / `w:del`).

**SpreadsheetML (§18)**
- Grids, shared strings, number formats and dates (incl. the 1904 date system).
- The print model — gridline suppression, print area, fit-to-page scaling, repeated
  print titles, manual page breaks, horizontal/vertical centering.

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
- Tagged PDF — logical structure tree, headings, tables, lists, figures with alt text,
  `/Lang`, pagination artifacts.
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
