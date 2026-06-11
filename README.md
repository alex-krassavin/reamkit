# Ream

> DOCX & XLSX → PDF, from scratch — no LibreOffice, no headless Office, no commercial SDK.

Convert Word (`.docx`) and Excel (`.xlsx`) documents to PDF, **in the browser** —
implemented from the **ECMA-376** (OOXML) and **ISO 32000** (PDF) specifications,
with no wrapper around LibreOffice, headless Office, or any commercial SDK. Pure
TypeScript/JavaScript working on `Uint8Array` in and `Uint8Array` out, so it also
runs unchanged in Node.js, serverless, and edge runtimes.

## Install

```sh
npm install reamkit
```

Runtime dependencies are minimal: `fflate` (ZIP/Deflate) and `fast-xml-parser`.

## Usage

Parse once into the format-neutral interlayer, convert to any target. The
format (docx/xlsx) is sniffed from the bytes; no fonts to wire up — an open
substitute font (Roboto for sans, Tinos for serif, Cousine for monospace — the
same families LibreOffice substitutes) is fetched automatically based on the
document's referenced fonts:

```ts
import { Ream } from 'reamkit';

// e.g. from an <input type="file"> or a fetch() — anything that yields bytes.
const bytes = new Uint8Array(await file.arrayBuffer());

const doc = Ream.parse(bytes);            // docx or xlsx — sniffed
const pdf = await doc.convert('pdf');     // async — fetches a font if needed
const svg = await doc.convert('svg');     // same parse, different target

// Hand the bytes to the browser: preview, download, upload, …
const url = URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' }));
window.open(url);
```

`doc.flow` exposes the parsed document tree, `doc.format` the detected format,
and `doc.convertWithReport(...)` returns `{ bytes, losses }` (pass
`strict: true` to throw on the first conversion loss instead). The one-shot
functions `convertDocxToPdf` / `convertXlsxToPdf` remain for single
conversions. Input/output are plain `Uint8Array`s, so wiring this to files,
the network, or disk is up to you.

### Bring your own fonts (no network)

To embed specific fonts — or to avoid the network entirely — pass the font
bytes in. `convert` then does zero I/O:

```ts
const fonts = {
  regular: new Uint8Array(await fetch('/fonts/MyFont-Regular.ttf').then((r) => r.arrayBuffer())),
  bold: new Uint8Array(await fetch('/fonts/MyFont-Bold.ttf').then((r) => r.arrayBuffer())),
  // italic, boldItalic — optional; missing faces degrade gracefully
};

const pdf = await Ream.parse(bytes).convert('pdf', { fonts });
```

### Font resolution chain

For finer control, chain font providers — first byte answer wins. A remote or
local winner is recorded as a `substituted` loss in the report:

```ts
import { Ream, callerFontProvider, localFontProvider, remoteFontProvider } from 'reamkit';

const pdf = await doc.convert('pdf', {
  fontProviders: [
    callerFontProvider(myFonts), // your bytes — highest priority
    localFontProvider(),         // system fonts (Chromium Local Font Access,
                                 //   embedding-restricted fonts are never used)
    remoteFontProvider(),        // open substitute set from CDN, last resort
  ],
});
```

Fonts the document itself embeds (`w:embed`, including obfuscated `.odttf`)
are always used first — glyph-exact, no substitution.

### Archival PDF/A + embedded source

The whole PDF/A family is supported (1a/1b, 2a/2b/2u, 3a/3b/3u —
veraPDF-validated). PDF/A-3 can carry the source document inside the PDF:

```ts
const { bytes: pdfa, losses } = await doc.convertWithReport('pdf', {
  fonts,
  pdfA: 'PDF/A-3b',
  embedSource: true, // the parsed .docx/.xlsx rides along as /AF Source
});
```

### Digital signatures

PKCS#7 detached signatures (ISO 32000 §12.8) via WebCrypto — RSA or ECDSA,
optional PAdES and RFC 3161 timestamping:

```ts
const signed = await doc.convert('pdf', {
  fonts,
  signature: { certificate: certDer, privateKey: cryptoKey },
});
```

### Strict mode and the loss report

Every conversion can report what was dropped, degraded or substituted. For
compliance-critical flows, make any loss fatal:

```ts
const { bytes, losses } = await doc.convertWithReport('pdf', { fonts });
// losses: [{ severity: 'substituted', feature: 'fonts.substitution', … }]

await doc.convert('pdf', { fonts, strict: true }); // throws ConversionLossError on the first loss
```

### Inspect the interlayer

`parse` produces a format-neutral document tree (the interlayer) before any
rendering — inspect or analyze it without converting:

```ts
const doc = Ream.parse(bytes);
doc.format;     // 'docx' | 'xlsx'
doc.flow.body;  // paragraphs / tables / images / charts …
doc.losses;     // read-time losses
```

### Hyphenation (optional)

```ts
import { getHyphenator } from 'reamkit';
const hyphenator = await getHyphenator('en-us'); // or 'ru'
const pdf = await doc.convert('pdf', { fonts, hyphenator });
```

### More options

`convert` accepts (beyond the above): `info` (PDF `/Info` metadata — also read
automatically from the document's `docProps/core.xml`), `attachments`
(PDF/A-3 associated files), `tagged` (logical structure without full PDF/A),
`pageWidth`/`pageHeight`/margins overrides.

### Lower-level APIs

- `convertDocxToPdf` / `convertXlsxToPdf` (+`Sync`) — one-shot, per-format
  functions; importing only one keeps the other format out of your bundle.
- `createConverter({ readers })` — the functional facade behind `Ream`.
- `docxReader` / `xlsxReader`, `layoutStyledDocument`, `svgWriter` — the
  `@experimental` reader/writer interfaces of the interlayer, for building
  custom pipelines.
- `renderStyledPdf` drives the layout engine directly; the typed document
  model is on the `reamkit/document-model` subpath.

## Scope

Implemented: WordprocessingML text/styles/tables/lists/multi-section
layout/headers-footers/images/tracked changes, SpreadsheetML grids,
number formats and the print model (gridlines, print area, fit-to-page,
repeated titles, page breaks), DrawingML shapes and charts, OMML math,
Type0+CIDFontType2 embedding with subsetting, Knuth-Plass line breaking,
Liang hyphenation, OpenType ligatures/kerning + Arabic cursive joining,
BiDi (UAX #9), tagged PDF, PDF/A-1/2/3 (a/b/u), digital signatures
(PKCS#7/ECDSA/PAdES/RFC 3161), SVG page preview.

See `handoff.md` for the full feature matrix and known limitations.

## License

[MIT](./LICENSE) © Alex Krassavin
