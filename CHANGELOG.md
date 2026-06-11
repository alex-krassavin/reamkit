# Changelog

All notable changes to **Ream** (`reamkit`) are documented here. The format is
based on [Keep a Changelog](https://keepachangelog.com/), and the project
follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

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
