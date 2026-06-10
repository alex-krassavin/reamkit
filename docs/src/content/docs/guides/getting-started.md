---
title: Getting started
description: Install Ream and convert your first document to PDF.
---

Ream converts Word (`.docx`) and Excel (`.xlsx`) documents to PDF, implemented from
the ECMA-376 and ISO 32000 specifications. It works on `Uint8Array` in and
`Uint8Array` out, so the same code runs in the browser, Node.js, serverless and edge
runtimes.

## Install

The package is published as `reamkit`. It is currently an **alpha**, so install the
`alpha` dist-tag:

```sh
npm install reamkit@alpha
```

Runtime dependencies are minimal: `fflate` (ZIP/Deflate) and `fast-xml-parser`.

## Convert a document

Give it the document bytes, get the PDF bytes back. No fonts to wire up — an open
substitute font (Roboto for sans, Tinos for serif, Cousine for monospace — the same
families LibreOffice substitutes) is fetched automatically based on the document's
referenced fonts:

```ts
import { convertDocxToPdf } from 'reamkit';

// e.g. from an <input type="file"> or a fetch() — anything that yields bytes.
const docx = new Uint8Array(await file.arrayBuffer());

const pdf = await convertDocxToPdf(docx); // async — fetches a font if needed

// Hand the bytes to the browser: preview, download, upload, …
const url = URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' }));
window.open(url);
```

Excel works the same way via `convertXlsxToPdf`. The input and output are plain
`Uint8Array`s, so wiring this to files, the network, or disk is up to you.

## Bring your own fonts (synchronous, no network)

To embed specific fonts — or to avoid the network entirely — pass the font bytes in
and use the synchronous variant:

```ts
import { convertDocxToPdfSync } from 'reamkit';

const fonts = {
  regular: new Uint8Array(await fetch('/fonts/MyFont-Regular.ttf').then((r) => r.arrayBuffer())),
  bold: new Uint8Array(await fetch('/fonts/MyFont-Bold.ttf').then((r) => r.arrayBuffer())),
};
const pdf = convertDocxToPdfSync(docx, { fonts });
```

The async `convertDocxToPdf` also accepts `fonts` (then it does no network I/O), plus
`fontFamily` to force a substitute and `fontFetch` to inject a custom fetch.

## Options

Both converters accept (beyond `fonts`):

- `info` — PDF `/Info` metadata (`title`, `author`, `subject`, …). Also read
  automatically from the document's `docProps/core.xml`.
- `pdfA: 'PDF/A-2b'` — emit an archival PDF/A file (embedded sRGB OutputIntent, XMP,
  deterministic `/ID`, subset-tagged fonts). Levels `1b/1a/2b/2u/2a/3b/3u/3a`.
- `hyphenator` — a Liang hyphenator for nicer justified text:

```ts
import { getHyphenator } from 'reamkit';
const hyphenator = await getHyphenator('en-us'); // or 'ru'
const pdf = await convertDocxToPdf(docx, { hyphenator });
```

See the **API Reference** for the full surface, including `renderStyledPdf`,
`signPdf`, and the typed document model under `reamkit/document-model`.
