---
title: Concepts
description: How Ream converts a document — the from-scratch, spec-driven pipeline.
---

Ream is written **from the specifications**, not as a wrapper. There is no
LibreOffice, headless Office, ZetaJS/WASM, or commercial SDK underneath — the
conversion is its own code. Existing implementations (the LibreOffice source, Apache
POI) are read only as a reference for the edge cases, never used as a dependency.

## The pipeline

A conversion walks the same stages a word processor would, each mapped to a section of
the standard:

1. **OPC unpacking** (ECMA-376 Part 2) — the `.docx`/`.xlsx` is a ZIP of XML parts.
   `fflate` unzips; `fast-xml-parser` parses. (ZIP and XML are the only "incidental"
   pieces we don't write ourselves — they are separate standards, not the task.)
2. **OOXML parsing** (Part 1) — WordprocessingML (§17), SpreadsheetML (§18), DrawingML
   (§20), OfficeMath (§22) become a typed, in-memory **document model**.
3. **Style cascade** (§17.7) — `docDefaults` → named styles → direct formatting are
   resolved into the effective properties of every run, paragraph and cell.
4. **Layout** — the box model is laid out into lines and pages: **Knuth–Plass**
   paragraph breaking, **Liang** hyphenation, table auto-layout, the spreadsheet print
   model (gridlines, print area, fit-to-page, repeated titles, page breaks).
5. **Text shaping** — Unicode text becomes positioned glyphs: OpenType ligatures and
   kerning (GSUB/GPOS), BiDi reordering (UAX #9), Arabic cursive joining.
6. **PDF writing** (ISO 32000) — content streams, the cross-reference table, Type0 +
   CIDFontType2 font embedding **with subsetting**, image XObjects. Optionally tagged
   PDF, a PDF/A profile, or a digital signature.

## The interlayer (FlowDoc)

The stages above are decoupled by an intermediate representation: readers
parse a format into a **FlowDoc** (a semantic, format-neutral document tree —
no pages or coordinates yet), the layout turns it into positioned pages (the
frozen page model: `PageItem`s in a top-left, point-unit frame), and writers
emit a target format from there:

```
bytes → reader → FlowDoc → layout → pages → writer → bytes
```

`Ream.parse` runs a reader once and hands you the FlowDoc (`doc.flow`); every
`doc.convert` renders from it without re-reading the source. New formats plug
in as `DocumentReader`/`DocumentWriter` implementations (the `@experimental`
interfaces) instead of new end-to-end converters — the SVG preview writer
(consuming positioned pages) and the HTML writer (consuming the FlowDoc
directly, no layout or fonts involved) are exactly such adapters.

## Bytes in, bytes out

The public API is deliberately small and I/O-free: `Uint8Array` in, `Uint8Array` out.
Ream never touches the filesystem or the network on the conversion path (zero `node:*`,
`Buffer`, `process` or `__dirname` imports), which is what lets the identical code run
in a browser tab, a Node server, a Lambda, or an edge worker. How the bytes reach you —
a `File`, a `fetch`, `fs` — is your call.

## You supply the fonts

PDF embeds the fonts it draws with, so a converter needs font bytes. Ream's model is
**caller-supplies-fonts**: the synchronous API takes the TTF bytes you pass; the async
API, as a convenience, fetches an open substitute (Roboto / Tinos / Cousine — the same
families LibreOffice substitutes) based on the document's referenced fonts. There are
no bundled fonts on the main path — the library renders faithfully with whatever font
you give it.

## Document model

Beyond the converters, the typed document model is exported from the
`reamkit/document-model` subpath — paragraphs, runs, tables, sections, numbering,
shapes, math and more. `renderStyledPdf` drives the layout engine directly if you want
to build or transform a document programmatically rather than parse one.
