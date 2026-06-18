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
metric-compatible substitute font (Arimo for sans, Tinos for serif, Cousine for
monospace, plus Carlito/Caladea for Calibri/Cambria — the same families LibreOffice
substitutes) is fetched automatically based on the document's referenced fonts:

```ts
import { Ream } from 'reamkit';

// e.g. from an <input type="file"> or a fetch() — anything that yields bytes.
const bytes = new Uint8Array(await file.arrayBuffer());

const doc = Ream.parse(bytes);            // docx, xlsx, pptx or pdf — sniffed
const pdf = await doc.convert('pdf');     // async — fetches a font if needed
const svg = await doc.convert('svg');     // same parse, different target
const html = await doc.convert('html');   // flowed HTML — needs no fonts at all
const docx = await doc.convert('docx');   // write WordprocessingML back out
const xlsx = await doc.convert('xlsx');   // write SpreadsheetML back (xlsx source)

// Hand the bytes to the browser: preview, download, upload, …
const url = URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' }));
window.open(url);
```

`doc.flow` exposes the parsed document tree, `doc.format` the detected format,
and `doc.convertWithReport(...)` returns `{ bytes, losses }` (pass
`strict: true` to throw on the first conversion loss instead). Input/output
are plain `Uint8Array`s, so wiring this to files, the network, or disk is up
to you.

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
veraPDF-validated), plus accessible **PDF/UA-1** (`pdfUA: true`, also
veraPDF-validated and combinable with PDF/A in one file). PDF/A-3 can carry
the source document inside the PDF:

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
doc.format;     // 'docx' | 'xlsx' | 'pptx' | 'pdf'
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

- `docxReader` / `xlsxReader`, `svgWriter`, `htmlWriter`, `docxWriter` — the `@experimental`
  reader/writer interfaces of the interlayer, for building custom pipelines (and
  keeping unused formats out of your bundle); `layoutStyledDocument` produces the
  frozen page model (`PageItem` pages in a top-left `Pt` frame) the page-based
  writers consume (`docxWriter` works from the flow model, before layout).
- `renderStyledPdf` drives the layout engine directly; the typed document
  model is on the `reamkit/document-model` subpath.

## Scope

Implemented: WordprocessingML text/styles/tables (incl. table styles)/lists/
multi-section and multi-column layout/headers-footers (incl. PAGE/NUMPAGES
fields)/footnotes and endnotes/hyperlinks and bookmarks/floating drawings/
images/tracked changes, SpreadsheetML grids,
number formats and the print model (gridlines, print area, fit-to-page,
repeated titles, page breaks), **conditional formatting** (color scales, data
bars, icon sets), **sparklines** and **Excel tables**, DrawingML shapes and
charts, OMML math, Type0+CIDFontType2 embedding with subsetting, Knuth-Plass
line breaking, Liang hyphenation, OpenType ligatures/kerning + Arabic cursive
joining, BiDi (UAX #9), hyperlinks (PDF link annotations + HTML anchors,
scheme-allowlisted), tagged PDF, PDF/A-1/2/3 (a/b/u), PDF/UA-1, AES-256
encryption, digital signatures (PKCS#7/ECDSA/PAdES/RFC 3161), SVG page
preview, flowed HTML export, and **docx + xlsx output** (write WordprocessingML
/ SpreadsheetML back out, incl. round-trips). Reads OOXML Transitional and Strict.

**Reads PDF, too.** `Ream.parse` accepts a PDF and reconstructs a `FlowDoc` — a
tagged PDF from its structure tree (headings, tables, lists, reading order), an
untagged one heuristically from glyph positions (lines, paragraphs, headings,
and a clean two-column split). It lifts back the text (via each font's
`/ToUnicode`), raster images (JPEG verbatim; PNG/Flate/LZW/CCITT-fax decoded and
re-encoded), `/Link` hyperlinks, form-XObject content, and filled / stroked /
gradient vector shapes. It reads modern compressed files (cross-reference + object
streams) and encrypted ones (RC4 / AES — the user password is passed to
`Ream.parse(bytes, { password })`, defaulting to the permissions-only case).

**Reads PowerPoint, too.** `Ream.parse` accepts a `.pptx` and turns each slide
into a page at the deck size — text boxes (with run formatting, alignment,
bullets and indents), layout/master placeholders, pictures, shapes, DrawingML
tables, embedded charts, theme colours, slide backgrounds, grouped shapes and
hyperlinks — then converts onward to PDF, SVG, HTML or DOCX like any source.

**Reads legacy `.doc`, `.xls` and `.ppt`, too.** The binary Word / Excel /
PowerPoint 97–2003 formats (OLE2/CFB) parse through a shared container reader: a
`.doc` yields its text with run and paragraph formatting, tables, inline images,
fields, headers/footers and lists; an `.xls` yields the grid with styling, embedded
images, charts, drawing shapes, cell hyperlinks, the page-setup print model and
defined names (named ranges, print area, repeated titles) and cell comments; a
`.ppt` yields each slide's
text (with run and paragraph formatting), embedded images, per-shape placement
(anchored text boxes and pictures at their slide rectangles) and decorative
autoshapes (with fill / line colours resolved through the slide's colour scheme),
one page per slide — all convert onward to PDF, SVG, HTML, or back to `.docx` /
`.xlsx` like any source.

See [`CHANGELOG.md`](./CHANGELOG.md) for the release history; the docs
[**Scope**](https://reamkit.dev/guides/scope/) guide has the full feature matrix
and known limitations.

## License

[MIT](./LICENSE) © Alex Krassavin
