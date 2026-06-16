---
title: Examples
description: Working recipes — PDF/A, digital signatures, font providers, SVG and HTML output, strict mode, the interlayer.
---

Every snippet below is runnable as-is; they all start from document bytes
(`Uint8Array`) however you obtained them — a `File`, a `fetch`, `fs`.

## One parse, many targets

`Ream.parse` reads the document once into the interlayer; every `convert`
renders from it without re-parsing:

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(bytes);

const pdf = await doc.convert('pdf', { fonts });
const svg = await doc.convert('svg', { fonts }); // page-stack preview, no PDF involved
const html = await doc.convert('html');          // flowed HTML — no fonts, zero I/O
const docx = await doc.convert('docx');          // WordprocessingML back out — no fonts, no layout
const xlsx = await doc.convert('xlsx');          // SpreadsheetML back out — from an .xlsx source
```

## docx → docx: normalize, sanitize, edit

`convert('docx')` writes the parsed document back to a valid `.docx`. The
round-trip is **semantic, not byte-exact** — the writer emits the resolved
formatting as direct properties rather than named styles — so use it to
normalize, sanitize or programmatically edit a document in the browser, not to
preserve the original markup verbatim. Images, tables, lists, links, bookmarks,
shapes, headers/footers, multi-section geometry, footnotes/endnotes, charts and
OfficeMath all round-trip; floating (anchored) placement collapses to inline.

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(bytes); // a .docx (xlsx has no docx writer)
const out = await doc.convert('docx');
// `out` is a fresh, valid .docx — hand it to a download, an upload, or re-parse it.
```

## xlsx → xlsx: re-emit a workbook

`convert('xlsx')` writes a spreadsheet's grid back to a valid `.xlsx`. Unlike the
docx writer it consumes the native grid tree, so the round-trip is **lossless on
the grid surface** — cells, styles, merges, the print model, conditional
formatting, sparklines, tables and embedded charts all survive a read → write →
read loop byte-stably. It requires a spreadsheet source; a `.docx` has no grid.

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(xlsxBytes); // a .xlsx
const out = await doc.convert('xlsx');
// `out` is a fresh, valid .xlsx — normalize, sanitize, or re-parse it.
```

## pdf → html / docx: read a PDF back

`Ream.parse` also accepts a **PDF** — including a modern compressed one
(cross-reference streams, object streams) or an encrypted one. A tagged PDF (the
ones Ream writes) is rebuilt from its structure tree — headings, paragraphs,
tables, lists in reading order; an untagged PDF is reconstructed heuristically
from glyph positions. **Raster images, hyperlinks and vector shapes come
back too** — images lifted out and sized from their placement, link annotations
re-attached to the text, filled paths, stroked lines and shading-pattern
gradients turned into shapes. The result is an ordinary `FlowDoc`, so it
converts onward like any other source. Clipping paths and clip-bounded (`sh`)
shadings are not read (reported as a loss).

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(pdfBytes); // doc.format === 'pdf'
const html = await doc.convert('html'); // the PDF's text as flowed HTML
const docx = await doc.convert('docx'); // …or an editable Word document

const { bytes, losses } = await doc.convertWithReport('html');
// losses note the untagged-heuristic degradation and any unread vector art.
```

### Encrypted PDFs

A PDF locked with a **user password** is opened by passing it to `Ream.parse`.
The empty-string default unlocks the common permissions-only encryption (where
the owner set restrictions but no open password), so most encrypted PDFs need no
password at all:

```ts
const doc = Ream.parse(pdfBytes, { password: 'letmein' });
```

A **wrong or missing** password is not thrown — `Ream.parse` still succeeds, but
the encrypted content can't be decrypted, so it's recorded as a read-time **loss**
and the text simply doesn't come back. Inspect `doc.losses`:

```ts
const doc = Ream.parse(lockedPdf); // no/incorrect password
doc.losses;
// [
//   {
//     severity: 'dropped',
//     feature: 'text',
//     detail: 'encrypted PDF — the user password was missing or incorrect, or the handler is unsupported',
//   },
// ]
```

To make that loss fatal instead, convert in **strict** mode: the first loss
throws a `ConversionLossError`, with the offending `Loss` on its `.loss`
property.

```ts
import { Ream, ConversionLossError } from 'reamkit';

try {
  await Ream.parse(lockedPdf).convert('html', { strict: true });
} catch (err) {
  if (err instanceof ConversionLossError) {
    err.loss.feature; // 'text'
    err.loss.detail; // 'encrypted PDF — the user password was missing or incorrect, …'
  }
}
```

## pptx → pdf: render a slide deck

`Ream.parse` also accepts a **PowerPoint** `.pptx`. Each slide becomes a page at
the deck size, its shapes read as positioned content — text boxes, placeholders,
pictures, shapes, tables, charts, theme colours, backgrounds, groups and
hyperlinks. The result is an ordinary `FlowDoc`, so it converts onward like any
other source:

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(pptxBytes); // doc.format === 'pptx'
const pdf = await doc.convert('pdf', { fonts }); // a page per slide
const html = await doc.convert('html'); // …or the slides as flowed HTML
```

## Browser: file input → PDF preview

```ts
import { Ream } from 'reamkit';

input.addEventListener('change', async () => {
  const bytes = new Uint8Array(await input.files![0].arrayBuffer());
  const pdf = await Ream.parse(bytes).convert('pdf');
  window.open(URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' })));
});
```

## Node: file in, file out

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import { Ream } from 'reamkit';

const doc = Ream.parse(new Uint8Array(readFileSync('report.docx')));
writeFileSync('report.pdf', await doc.convert('pdf'));
```

## Archival PDF/A with the source embedded

The whole PDF/A family is supported (1a/1b, 2a/2b/2u, 3a/3b/3u). PDF/A-3 can
carry the source document inside the PDF as an associated file:

```ts
const { bytes: pdfa, losses } = await doc.convertWithReport('pdf', {
  fonts,
  pdfA: 'PDF/A-3b',
  embedSource: true, // .docx/.xlsx rides along as /AF with /AFRelationship /Source
});
```

`'PDF/A-1a'` / `'PDF/A-2a'` / `'PDF/A-3a'` additionally emit the tagged
logical structure (headings, tables, lists, figure alt text).

## Accessible PDF/UA-1

`pdfUA: true` produces ISO 14289-1-conformant output — tagged structure,
alternate descriptions on links, an always-announced document title. It
combines with PDF/A in a single file (both veraPDF-validated):

```ts
const accessible = await doc.convert('pdf', { fonts, pdfUA: true });
const archival = await doc.convert('pdf', { fonts, pdfA: 'PDF/A-2a', pdfUA: true });
```

## Digital signature

PKCS#7 detached (ISO 32000 §12.8) via WebCrypto — RSA or ECDSA:

```ts
const signed = await doc.convert('pdf', {
  fonts,
  signature: {
    certificate: certificateDer,  // Uint8Array, DER
    privateKey: cryptoKey,        // WebCrypto CryptoKey
    // optional: signingTime, reason, location, pades: true, timestampUrl
  },
});
```

## Font resolution chain

Chain providers; the first byte answer wins. A remote or local winner is
recorded as a `substituted` loss:

```ts
import { Ream, callerFontProvider, localFontProvider, remoteFontProvider } from 'reamkit';

const { bytes, losses } = await doc.convertWithReport('pdf', {
  fontProviders: [
    callerFontProvider(myFonts), // your bytes first
    localFontProvider(),         // system fonts (Chromium Local Font Access;
                                 //   embedding-restricted fonts are never used)
    remoteFontProvider(),        // open substitute set, last resort
  ],
});
// losses[0] → { severity: 'substituted', feature: 'fonts.substitution', … }
```

Fonts embedded in the document itself (`w:embed`, including obfuscated
`.odttf`) always win — glyph-exact, no substitution.

## Renderer parity

Ream is a correct typesetter: it lays out faithfully for the font you give it. For
closer _visual parity_ with a specific renderer, pass a `layoutProfile` — it switches
the line-height model, line breaking and default kerning to match that tool. Paired
with the metric-compatible substitutes above (so the same glyph advances are in play),
the page geometry tracks the target closely:

```ts
const pdf = await doc.convert('pdf', { fonts, layoutProfile: 'libreoffice' });
// 'word' targets Microsoft Word; 'ream' (the default) is Ream's own typesetter.
```

`'libreoffice'` derives line height from the font's hhea metrics and breaks lines
greedily (first-fit); `'word'` uses the OS/2 win metrics and turns kerning off (Word's
default). The profile applies to flowing text (DOCX/PPTX); spreadsheet geometry follows
Excel's row model regardless.

## Strict mode (compliance flows)

Make any loss fatal instead of reported:

```ts
import { ConversionLossError } from 'reamkit';

try {
  await doc.convert('pdf', { fonts, strict: true });
} catch (e) {
  if (e instanceof ConversionLossError) {
    // e.loss — what exactly would have been dropped/degraded/substituted
  }
}
```

## Inspect the interlayer

`parse` produces a format-neutral tree before any rendering:

```ts
const doc = Ream.parse(bytes);

doc.format;          // 'docx' | 'xlsx'
doc.losses;          // read-time losses
for (const el of doc.flow.body) {
  // el.kind: 'paragraph' | 'table' | 'image' | 'chart' | 'shape'
}
```

## PDF metadata

`/Info` is read automatically from the document's `docProps/core.xml`;
caller values override it:

```ts
const pdf = await doc.convert('pdf', {
  fonts,
  info: { title: 'Q4 Report', author: 'Finance' },
});
```

## Hyphenation

```ts
import { getHyphenator } from 'reamkit';

const hyphenator = await getHyphenator('en-us'); // or 'ru'
const pdf = await doc.convert('pdf', { fonts, hyphenator });
```
