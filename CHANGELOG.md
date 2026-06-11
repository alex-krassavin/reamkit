# Changelog

All notable changes to **Ream** (`reamkit`) are documented here. The format is
based on [Keep a Changelog](https://keepachangelog.com/), and the project
follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- **Charts on xlsx sheets** — the worksheet's drawing part now loads:
  chart frames anchored to cell ranges render after their sheet's grid,
  sized from the anchor's column/row tracks. Custom chart colour themes
  (`colorsN.xml`) apply to series in both Word and Excel documents.
- **Tagged lists: Lbl elements** — list-item markers ("1.", "•") get their
  own `Lbl` structure element, so assistive technology announces the label
  separately from the item body.

## [1.2.0] - 2026-06-11

### Added

- **Bookmarks and internal links** — `w:bookmarkStart` + `w:hyperlink
  @anchor` become real GoTo links: PDF annotations with named destinations
  (`/Names /Dests`, only referenced names), tagged `Link` structure, HTML
  `id` anchors with `#`-fragment links.
- **PDF/UA-1** — `pdfUA: true` produces ISO 14289-1-conformant output
  (veraPDF-validated, alone and combined with PDF/A-2a in one file): tagged
  structure, `pdfuaid` XMP identification, alternate descriptions on link
  annotations, unique IDs on footnote Note elements, an always-present
  document title.
- **Multi-column sections** — `w:cols` lays content out column by column
  (equal-width with a shared gutter, or explicit per-column widths);
  headers, footers and footnotes keep the full page width.
- **Floating drawings** — `wp:anchor` placement: wrap-none drawings
  (watermarks, stamps, text boxes; including `behindDoc`) render at their
  anchored page/margin/paragraph-relative position without disturbing the
  text flow. Side-wrapping modes stay in flow (v1).
- **HTML writer: charts and shapes render as inline SVG** — bar/line/pie/
  area/scatter charts emit the same geometry scene as the PDF path (labels
  as native `<text>` with anchors, so the browser's fonts do the rendering);
  shape geometry (preset + custom, fills, strokes, dash patterns, rotation/
  flips) emits as `<path>` with the exact transform matrix the PDF layout
  computes, and text boxes overlay their content inside the body insets with
  the source vertical anchor. Charts and shapes are no longer reported as
  dropped losses for HTML output.
- **Footnotes and endnotes** — `w:footnoteReference`/`w:endnoteReference`
  render superscript numbers; footnote content lands at the bottom of the
  referencing page behind Word's short separator rule (the line and its note
  always travel together), endnotes flow after the body. Tagged PDFs wrap
  each note in a `Note` structure element (veraPDF-validated); the HTML
  writer renders anchored references with a notes section and backlinks.
- **PAGE / NUMPAGES fields** — page-number fields in headers and footers now
  render the real page number and total (both the `fldSimple` and the
  `fldChar` complex-field syntax). Bands containing fields re-lay out per
  page after pagination; other field instructions keep their cached result
  exactly as before, and documents without fields are byte-identical.
- **Table styles** — `w:tblStyle` referenced styles now render: the base
  layer (grid borders, default cell margins) plus `w:tblStylePr` conditional
  regions (first/last row and column, row/column banding, corner cells) gated
  by `w:tblLook` (modern attributes and the legacy bitmask). Resolved in the
  reader, so PDF, SVG and HTML all pick it up; tables without a style are
  byte-identical to before.
- **Hyperlinks** — `w:hyperlink` external targets now become clickable: PDF
  output gets `/Link` annotations (one rect per rendered line, merged into
  each page's `/Annots`; in tagged/PDF-A mode the annotation is enclosed in a
  `Link` structure element with `OBJR` + `/StructParent`, veraPDF-validated),
  and HTML output wraps the text in `<a href>`. Targets pass a scheme
  allowlist (`http`/`https`/`mailto`) — anything else renders as plain text
  with a degraded-`hyperlinks` loss; documents without links are
  byte-identical to before.

## [1.1.0]

### Added

- **HTML writer** — `doc.convert('html')` renders the parsed document as a
  single self-contained flowed HTML file (headings, run styling, tables with
  spans/borders/shading, images as `data:` URIs, list markers, RTL). A flow
  medium needs no pagination and no fonts, so the conversion performs zero
  I/O; chart/shape geometry, inline math and headers/footers are reported in
  the loss report. Also exposed as `htmlWriter`/`writeHtml` and
  `createConverter` `to: 'html'`.

### Changed

- **Page model frozen** (`@experimental` API): `PageItem` page-frame
  coordinates are now **top-left / y-down** (CSS/SVG convention) and branded
  as `Pt`; the PDF emitter converts into PDF's y-up frame at emission. PDF
  output is byte-identical; SVG output changes coordinates only, not
  geometry.
- `LaidOutDocument` narrowed to the page model proper (`pages`, `resources`,
  `fontResources`, `imageResources`); the PDF-only state rides on
  `layoutStyledDocument(...).pdf`. The internal `DrawCommand` alias is gone —
  the schema name is `PageItem`.

## [1.0.0]

The interlayer release — and the first stable major. Documents parse once
into a format-neutral tree (**FlowDoc**) and convert to any target from
there; the public face of the library is the `Ream` class.

### Added

- **`Ream`** — the object API: `Ream.parse(bytes)` (format sniffed) →
  `doc.convert('pdf' | 'svg', options)`, with `doc.flow` (the parsed tree),
  `doc.format`, `doc.losses`, and `doc.convertWithReport()` returning
  `{ bytes, losses }`. Conversion output is byte-identical to the per-format
  converters.
- **Intermediate representation (`@experimental`)** — `DocumentReader` /
  `DocumentWriter` interfaces, `docxReader`/`xlsxReader`, the
  `createConverter` facade, branded `Pt` units, a content-addressed
  `ResourceStore`, a `Feature` registry, and a loss protocol
  (`Loss`, `ConversionLossError`, `strict` mode).
- **SVG writer** — `doc.convert('svg')` renders the laid-out pages as a
  stacked-page SVG preview (no PDF involved); the third adapter, written
  purely against the page model.
- **Font provider chain** — `callerFontProvider` / `embeddedDocFontProvider` /
  `localFontProvider` (Chromium Local Font Access, with the OS/2 `fsType`
  licensing gate: embedding-restricted fonts are never used) /
  `remoteFontProvider`, composable via `fontProviders: [...]`; a remote or
  local winner is reported as a `substituted` loss.
- A cross-revision **byte gate** test suite: PDF output of fixed fixtures is
  snapshot-hashed, so pure refactorings are provably byte-identical.

### Changed

- `src/` reorganized into format modules: `core/` (format-agnostic), `word/`,
  `excel/`, `pdf/`, `svg/`. The public `reamkit/document-model` subpath is
  unchanged.
- The PDF renderer is split into layout and emit phases
  (`styled-page-emitter`); image embedding is split into pure
  `prepareImage` + `addImage` (PNG is no longer decoded twice per
  conversion).
- README and the docs site are rewritten around the `Ream` API, with an
  Examples page (PDF/A, signatures, providers, strict mode, recipes).

### Removed (breaking)

- The one-shot `convertDocxToPdf` / `convertXlsxToPdf` (+`Sync`) functions
  and their option types are no longer exported — use
  `Ream.parse(bytes).convert('pdf', options)`. Custom pipelines and
  bundle-size-sensitive consumers build on the `@experimental`
  reader/writer interfaces instead.

### Fixed

- `remoteFontProvider` lost the boldItalic→bold/italic degradation: a
  bold-italic run against a partial CDN set fell back to regular even when
  bold was available.
- Images in headers/footers resolved through the MAIN document's
  relationships (OPC ids are per-part): a colliding rId rendered the wrong
  picture, a non-colliding one dropped. Each header/footer part now resolves
  through its own `.rels`.
- `signPdf` located its placeholder by scanning for the first bare
  `/ByteRange` — an embedded attachment could hijack the signature bytes. The
  scan now matches the full fixed-width placeholder.

## [0.1.0-alpha.0]

First tagged alpha. DOCX/XLSX → PDF, implemented directly from the ECMA-376 and
ISO 32000 / 19005 specifications — no third-party converter, PDF writer, or
layout engine. Browser-first; the caller supplies fonts.

- **WordprocessingML** — text, styles, tables (§17.4 border-conflict resolution,
  shading, vMerge/gridSpan, nested tables), lists, multi-section layout,
  headers/footers, inline/floating images, tracked changes.
- **SpreadsheetML** — grids, number formats, dates, and the print model
  (gridlines, print area, fit-to-page, repeated titles, page breaks).
- **Fonts/typography** — Type0 + CIDFontType2 embedding with subsetting,
  Knuth–Plass line breaking, Liang hyphenation, OpenType ligatures/kerning,
  BiDi (UAX #9), Arabic cursive joining.
- **Graphics** — DrawingML shapes, charts, and OMML math.
- **PDF** — PDF/A-1/2/3 (a/b/u), veraPDF-validated; digital signatures
  (PKCS#7 detached, ECDSA); object streams; JPEG2000.
- **Tooling** — aligned to `@tanstack/config` (Vite build, ESLint, publint +
  are-the-types-wrong); MIT-licensed; tag-triggered npm release workflow.

[Unreleased]: https://github.com/alex-krassavin/reamkit/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/alex-krassavin/reamkit/compare/v0.1.0-alpha.0...v1.0.0
[0.1.0-alpha.0]: https://github.com/alex-krassavin/reamkit/releases/tag/v0.1.0-alpha.0
