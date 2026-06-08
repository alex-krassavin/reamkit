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

Give it the document bytes, get the PDF bytes back. No fonts to wire up — an open
substitute font (Roboto for sans, Tinos for serif, Cousine for monospace — the
same families LibreOffice substitutes) is fetched automatically based on the
document's referenced fonts:

```ts
import { convertDocxToPdf } from 'reamkit';

// e.g. from an <input type="file"> or a fetch() — anything that yields bytes.
const docx = new Uint8Array(await file.arrayBuffer());

const pdf = await convertDocxToPdf(docx); // async — fetches a font if needed

// Hand the bytes to the browser: preview, download, upload, …
const url = URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' }));
window.open(url);
```

Excel works the same way via `convertXlsxToPdf`. The input/output are plain
`Uint8Array`s, so wiring this to files, the network, or disk is up to you.

### Bring your own fonts (synchronous, no network)

To embed specific fonts — or to avoid the network entirely — pass the font bytes
in and use the synchronous variant:

```ts
import { convertDocxToPdfSync } from 'reamkit';

const fonts = {
  regular: new Uint8Array(await fetch('/fonts/MyFont-Regular.ttf').then((r) => r.arrayBuffer())),
  bold: new Uint8Array(await fetch('/fonts/MyFont-Bold.ttf').then((r) => r.arrayBuffer())),
};
const pdf = convertDocxToPdfSync(docx, { fonts });
```

The async `convertDocxToPdf` also accepts `fonts` (then it does no network I/O),
plus `fontFamily` to force a substitute and `fontFetch` to inject a custom
fetch implementation.

### Options

Both converters accept (beyond `fonts` / `fontBytes`):

- `info` — PDF `/Info` metadata (`title`, `author`, `subject`, …). Also read
  automatically from the document's `docProps/core.xml`.
- `pdfA: 'PDF/A-1b'` — emit an archival PDF/A-1b file (embedded sRGB
  OutputIntent, XMP, deterministic `/ID`, subset-tagged fonts, no transparency).
- `hyphenator` — a Liang hyphenator (see below) for nicer justified text.

### Hyphenation (optional)

```ts
import { getHyphenator } from 'reamkit';
const hyphenator = await getHyphenator('en-us'); // or 'ru'
const pdf = await convertDocxToPdf(docx, { hyphenator });
```

### Advanced

`renderStyledPdf` drives the layout engine directly. The typed document model is
available from the `reamkit/document-model` subpath.

## Scope

Implemented: WordprocessingML text/styles/tables/lists/sections/headers-footers/
images, SpreadsheetML grids/number-formats/dates, Type0+CIDFontType2 font
embedding with subsetting, Knuth-Plass line breaking, Liang hyphenation,
OpenType ligatures + kerning, BiDi (UAX #9), and PDF/A-1b.

See `handoff.md` for the full feature matrix and known limitations.

## License

[MIT](./LICENSE) © Alex Krassavin
