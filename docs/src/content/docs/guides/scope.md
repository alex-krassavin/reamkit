---
title: Scope & limitations
description: What Ream implements today, and where the edges are.
---

The conversion core is broad and spec-driven, but ECMA-376 is vast and Ream does
not implement all of it — this page is an honest map of what works and what
doesn't yet.

## Implemented

**Input** — Ream parses **Word (`.docx`)**, **Excel (`.xlsx`)** and **PDF**,
sniffed from the bytes. PDF input handles classic and modern compressed files
(cross-reference streams, object streams) and encrypted files (RC4 / AES; the
user password is passed to `Ream.parse`, defaulting to the empty permissions-only
case). A **tagged** PDF (including the ones Ream writes) is rebuilt from
its structure tree — headings, paragraphs, tables, list items, reading order; an
**untagged** PDF is reconstructed heuristically from glyph positions (lines by
baseline, paragraphs by spacing, headings by relative font size, and a clean
two-column page split at its central gutter), which is approximate. Text comes back via each font's `/ToUnicode` map; **raster images,
hyperlinks and vector shapes** are lifted back out too (JPEG verbatim, other
images re-encoded as PNG with soft-mask alpha, `/Link` URIs re-attached to the
text, filled paths, stroked lines and shading-pattern gradients turned into
shapes). Clipping paths and clip-bounded (`sh`) shadings are not read.

**Output** — `convert('pdf')`, `convert('svg')` (a page-stack preview),
`convert('html')` (flowed, needs no fonts), `convert('docx')` (write
WordprocessingML back out) and `convert('xlsx')` (write SpreadsheetML back out —
spreadsheet input only). The writers are for normalization, sanitization,
in-browser editing, and round-tripping. The docx round-trip is semantic, not
byte-exact, but complete — text, tables, images, lists, links, headers/footers,
multi-section geometry, footnotes/endnotes, charts and OfficeMath all write
back. The xlsx round-trip preserves the whole grid surface — cells, styles,
merges, the print model, conditional formatting, sparklines, tables and embedded
charts — and is byte-stable across a read↔write loop.

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
  print titles, manual page breaks, horizontal/vertical centering, and **column-band
  pagination**: a sheet wider than the page (and not fit-to-width) splits across
  pages, all rows of the left columns first, then the next band ("down, then over"),
  honouring manual column breaks — instead of being squeezed onto one page width.
- **Frozen panes** round-trip through the writer and become sticky header rows /
  columns in HTML output. They do not affect PDF — in Excel freezing is a view
  setting that does not print (the printed repeat is the print titles above).
- **Conditional formatting** — `cellIs` (compare-to-constant highlights),
  `colorScale` (2/3-stop gradients), `dataBar` (in-cell bars, with a zero axis
  so negative values run the other way) and `iconSet` — traffic lights, arrows,
  signs, symbols (check / exclamation / cross), flags, ratings (a bar meter) and
  quarters (a clock pie). The cross-cell rules resolve against the range's value
  extent.
- **Sparklines** — per-cell line / column / win-loss mini charts, including
  cross-sheet data ranges and blank-cell gaps.
- **Excel tables** (`xl/tables`) — banded rows and a styled header row, the
  colours resolved from the named table style against the workbook theme.
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
