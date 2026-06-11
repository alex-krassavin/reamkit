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
