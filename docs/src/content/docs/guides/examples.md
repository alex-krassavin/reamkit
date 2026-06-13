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
shapes, headers/footers and multi-section geometry round-trip; footnotes, charts
and OfficeMath are reported as losses (see [strict mode](#strict-mode-compliance-flows)).

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
formatting, sparklines and tables all survive a read → write → read loop
byte-stably (embedded charts are the one piece not yet written). It requires a
spreadsheet source; a `.docx` has no grid.

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(xlsxBytes); // a .xlsx
const out = await doc.convert('xlsx');
// `out` is a fresh, valid .xlsx — normalize, sanitize, or re-parse it.
```

## pdf → html / docx: read a PDF back

`Ream.parse` also accepts a **PDF**. A tagged PDF (including the ones Ream
writes) is rebuilt from its structure tree — headings, paragraphs, tables, lists
in reading order; an untagged PDF is reconstructed heuristically from glyph
positions. **Raster images come back too** — lifted out of the page, sized from
their placement, and carried into the HTML `<img>` / docx media part. The result
is an ordinary `FlowDoc`, so it converts onward like any other source. Vector
graphics and encrypted PDFs are not read (reported as losses).

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(pdfBytes); // doc.format === 'pdf'
const html = await doc.convert('html'); // the PDF's text as flowed HTML
const docx = await doc.convert('docx'); // …or an editable Word document

const { bytes, losses } = await doc.convertWithReport('html');
// losses note the untagged-heuristic degradation and any unread vector art.
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
