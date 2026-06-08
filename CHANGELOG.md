# Changelog

All notable changes to **Ream** (`reamkit`) are documented here. The format is
based on [Keep a Changelog](https://keepachangelog.com/), and the project
follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

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

[Unreleased]: https://github.com/alex-krassavin/reamkit/compare/v0.1.0-alpha.0...HEAD
[0.1.0-alpha.0]: https://github.com/alex-krassavin/reamkit/releases/tag/v0.1.0-alpha.0
