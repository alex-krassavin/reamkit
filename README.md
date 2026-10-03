# Ream

> Read Word, Excel, PowerPoint and PDF — and convert any of them to PDF, SVG, HTML, Markdown, DOCX or XLSX. From scratch, in the browser. No LibreOffice, no headless Office, no commercial SDK.

Ream parses **seven document formats** — the modern Office Open XML trio (`.docx`,
`.xlsx`, `.pptx`), `.pdf`, and the legacy binary `.doc` / `.xls` / `.ppt` — into one
format-neutral interlayer, then renders that to **PDF, SVG, HTML, Markdown, DOCX or XLSX**.
It is implemented directly from the **ECMA-376** (OOXML), **ISO 32000 / 19005**
(PDF / PDF/A), and Microsoft's binary-format specifications — no wrapper around
LibreOffice, headless Office, or any commercial SDK. Pure TypeScript/JavaScript on
`Uint8Array` in and `Uint8Array` out, so the same code runs unchanged in the browser,
Node.js, serverless and edge runtimes.

|            |                                                                                             |
| ---------- | ------------------------------------------------------------------------------------------- |
| **Reads**  | `.docx` · `.xlsx` · `.pptx` · `.pdf` · legacy `.doc` · `.xls` · `.ppt`                       |
| **Writes** | `.pdf` (incl. PDF/A-1/2/3, PDF/UA-1, signed, encrypted) · `.svg` · `.html` · `.md` · `.docx` · `.xlsx` |

## Install

```sh
npm install reamkit
```

Runtime dependencies are minimal: `fflate` (ZIP/Deflate) and `fast-xml-parser`.

## Usage

Parse once into the format-neutral interlayer, convert to any target. The
format is sniffed from the bytes; no fonts to wire up — an open
metric-compatible substitute font (Arimo for sans, Tinos for serif, Cousine for
monospace, plus Carlito/Caladea for Calibri/Cambria — the same families LibreOffice
substitutes) is fetched automatically based on the document's referenced fonts,
with a Noto face for the Japanese, Korean, Chinese, Arabic, Hebrew or Thai text
the document holds:

```ts
import { Ream } from 'reamkit';

// e.g. from an <input type="file"> or a fetch() — anything that yields bytes.
const bytes = new Uint8Array(await file.arrayBuffer());

const doc = Ream.parse(bytes);            // docx, xlsx, pptx, pdf, doc, xls or ppt — sniffed
const pdf = await doc.convert('pdf');     // async — fetches a font if needed
const svg = await doc.convert('svg');     // same parse, different target
const html = await doc.convert('html');   // flowed HTML — needs no fonts at all
const md   = await doc.convert('md');     // GitHub-Flavored Markdown — same, narrower
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

An encrypted source opens with its password — a PDF's user password, or an
Office package's (MS-OFFCRYPTO, Agile and Standard):

```ts
const doc = Ream.parse(bytes, { password: 'secret' });
```

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
doc.format;     // 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'doc' | 'xls' | 'ppt'
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
`encrypt` (AES-256 with a user and an owner password and the permissions a
reader keeps: printing, copying, modifying, …), `pageWidth`/`pageHeight`/margins
overrides.

### Lower-level APIs

- `docxReader` / `xlsxReader`, `svgWriter`, `htmlWriter`, `markdownWriter`, `docxWriter` — the `@experimental`
  reader/writer interfaces of the interlayer, for building custom pipelines (and
  keeping unused formats out of your bundle); `layoutStyledDocument` produces the
  frozen page model (`PageItem` pages in a top-left `Pt` frame) the page-based
  writers consume (`docxWriter` works from the flow model, before layout).
- `createConverter({ readers })` — a converter over a reader registry of your
  own: `detect` names the reader that recognises the bytes, and `convert` reads
  and converts them in one call, returning `{ bytes, losses }`.
- `renderStyledPdf` drives the layout engine directly; the typed document
  model is on the `reamkit/document-model` subpath.

## Scope

Implemented: WordprocessingML text/styles/tables (incl. table styles)/lists/
multi-section and multi-column layout/headers-footers (incl. PAGE/NUMPAGES
fields)/footnotes and endnotes/hyperlinks and bookmarks/floating drawings/
images/tracked changes, SpreadsheetML grids,
number formats and the print model (gridlines, print area, fit-to-page,
repeated titles, page breaks), **conditional formatting** (color scales, data
bars, icon sets, and `expression` rules evaluated by a ~140-function formula
engine), **sparklines** and **Excel tables**, DrawingML shapes and
charts, OMML math, Type0+CIDFontType2 embedding with subsetting, Knuth-Plass
line breaking, Liang hyphenation, OpenType ligatures/kerning + Arabic cursive
joining, BiDi (UAX #9), hyperlinks (PDF link annotations + HTML anchors,
scheme-allowlisted), tagged PDF, PDF/A-1/2/3 (a/b/u), PDF/UA-1, AES-256
encryption, digital signatures (PKCS#7/ECDSA/PAdES/RFC 3161), SVG page
preview, flowed HTML export, and **docx + xlsx output** (write WordprocessingML
/ SpreadsheetML back out, incl. round-trips). Reads OOXML Transitional and Strict.

A Word document is laid out by Word's own rules, each measured in Word: a line
as tall as Word sets its faces, two paragraphs standing the larger of their
spacings apart rather than the sum, widow and orphan control, a heading kept with
what it heads, table borders that take the room they are wide, a Word 2010
table placed by its first cell's text, and sections whose lines run down the
sheet.

A workbook is drawn the way Excel draws it, each rule checked against Excel's
own PDF: columns in the Normal font's digit as Excel counts it, a table in the
style it names, a theme colour's tint to the digit, a sheet that reads from the
right turned round, the notes it shows beside their cells, its drawings over
the cells and cut where its pages are, a chart's axes scaled and crossed where
Excel puts them, its text in the face Excel sets it in, its 3-D bars and pies
standing in three dimensions. As HTML a workbook comes out as Excel's window
shows it — every tab, its gridlines, a noted cell flagged with the note on
hover — and as SVG as pictures of its sheets, each whole, its text drawn from
the faces' own outlines.

**Reads PDF, too.** `Ream.parse` accepts a PDF and reconstructs a `FlowDoc` — a
tagged PDF from its structure tree (headings, tables, lists, reading order), an
untagged one from where its glyphs stand: lines and paragraphs at the page's own
pitch, with their indents and alignment; headings and list items; columns of
any number, breaking and balancing where the page does; tables ruled or set out
on stops, within a column or across them; code listings; figures, a drawing and
the labels set on it kept as one group; the hanging entries of an index; and a
heading's bar as its shading. Running heads and feet become headers and footers,
their numbers PAGE fields counted in the page's own numerals — a book's roman
front matter, and a foot of its own for each side of its spreads.

It lifts back the text (via each font's `/ToUnicode`, or the embedded program's
own `cmap` where there is none, or the glyph names its `/Encoding` states — which
is all a PDF from TeX gives, its Greek and mathematics included) and the font
programs themselves: written to `.docx`, each face the page drew with travels
with the document, rebuilt as TrueType with the page's own advances, kerning and
ligatures, and embedded where its licence (OS/2 `fsType`) allows — so a line
breaks where the page broke it, on a machine that has none of its fonts. Raster
images come back (JPEG verbatim, a CMYK one re-coloured; PNG/Flate/LZW/CCITT-fax
and **JBIG2** decoded and re-encoded), with `/Link` hyperlinks, form-XObject
content, annotation and form appearances (drawing one itself where the file
supplies none), colour set through a named space — device, CIE (`CalGray`,
`CalRGB`, `Lab`), ICC-based, or a `Separation`/`DeviceN` run through its own
tint transform — and the page's artwork: filled / stroked / gradient shapes,
dashes and caps, clipping paths, tiling patterns, stencil image masks, opacity,
and the Type 3 glyphs that are drawings rather than letters.

It honours the layers a file turns off (§8.11 optional content), the box it says
to show (`/CropBox`) and the way it turns a page (`/Rotate` — a page whose words
run down its sheet comes back as a section set that way), and it decides for
itself whether a file is a document to re-flow or a page to keep — a paper is
mostly lines, a form is mostly marks — and there is nothing to configure: a
caller cannot know which of the two it was handed. It reads modern compressed
files (cross-reference + object streams) and encrypted ones (RC4 / AES — the
user password is passed to `Ream.parse(bytes, { password })`, defaulting to the
permissions-only case); a filter it does not carry can be handed to it through
`Ream.parse(bytes, { filters })`.

A form or a drawing is not a reflowable document, so the reader keeps it as a
page instead: every line stands where its glyphs stand, beside the artwork, the
way up the page is shown. Which of the two readings a file gets it decides
itself and records in the losses.

**Reads PowerPoint, too.** `Ream.parse` accepts a `.pptx` and turns each slide
into a page at the deck size — text boxes (with run formatting, alignment,
bullets and indents), layout/master placeholders, pictures, shapes, DrawingML
tables, embedded charts, theme colours, slide backgrounds, grouped shapes and
hyperlinks — then converts onward to PDF, SVG, HTML, Markdown or DOCX like any source.

**Reads legacy `.doc`, `.xls` and `.ppt`, too.** The binary Word / Excel /
PowerPoint 97–2003 formats (OLE2/CFB) parse through a shared container reader: a
`.doc` yields its text with run and paragraph formatting, tables (with cell
borders, vertical merges and background shading), inline images, fields,
headers/footers and lists (numbered or bulleted, in their number format);
an `.xls` yields the grid with styling, embedded
images, charts, drawing shapes, cell hyperlinks, the page-setup print model,
defined names (named ranges, print area, repeated titles), cell comments, data
validation, frozen panes, custom row heights and conditional formatting (the
classic cellIs / expression rules **and** the 2007 colour-scale / data-bar /
icon-set extensions); a `.ppt` yields each slide's
text (with run and paragraph formatting), embedded images, per-shape placement
(anchored text boxes and pictures at their slide rectangles) and decorative
autoshapes (preset or exact freeform geometry, with fill / line colours resolved
through the slide's colour scheme), one page per slide — all convert onward to PDF,
SVG, HTML, Markdown, or back to `.docx` / `.xlsx` like any source.

See [`CHANGELOG.md`](./CHANGELOG.md) for the release history; the docs
[**Scope**](https://reamkit.dev/guides/scope/) guide has the full feature matrix
and known limitations.

## License

[MIT](./LICENSE) — Copyright (c) 2026 Alexandr Krassavin
