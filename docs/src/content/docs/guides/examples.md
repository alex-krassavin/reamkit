---
title: Examples
description: Working recipes — PDF/A, digital signatures, font providers, SVG, HTML and Markdown output, strict mode, the interlayer.
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
const md   = await doc.convert('md');            // GitHub-Flavored Markdown — same, narrower
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
OfficeMath all round-trip, and a drawing that states where on the page it
belongs is written back as the `wp:anchor` that puts it there rather than
flattened into the text flow. No image part the format cannot show goes into the
package: JPEG 2000 is dropped with a loss that names it.

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(bytes); // a .docx (xlsx has no docx writer)
const out = await doc.convert('docx');
// `out` is a fresh, valid .docx — hand it to a download, an upload, or re-parse it.
```

## docx → markdown: structure without geometry

`convert('md')` writes GitHub-Flavored Markdown. Like HTML it is a flow medium —
no pagination, no layout engine, no fonts — but a much narrower one, so it keeps
what the document *says* and reports everything it cannot say back.

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(bytes);
const { bytes: md, losses } = await doc.convertWithReport('md');

// Headings, lists (with their real numbers and nesting), GFM tables,
// links, footnotes and the text inside shapes all survive.
console.log(new TextDecoder().decode(md));

// Everything markdown has no syntax for is reported — once each, however
// many paragraphs it happened in:
//   [dropped] text: paragraph alignment has no markdown expression
//   [degraded] tables: merged cells flattened — markdown has no spans
for (const loss of losses) console.log(loss.severity, loss.feature, loss.detail);
```

Pictures are inlined as `data:` URIs so the output is a single self-contained
file. When you would rather write the image bytes yourself — a docs site, a
repo — ask for links instead and the writer names them predictably:

```ts
const md = await doc.convert('md', { images: 'link' }); // ![](./media/image1.png)
const bare = await doc.convert('md', { images: 'drop' }); // no pictures at all
```

A deck's only structure is where one slide ends and the next begins, and the
`.pptx` / `.ppt` readers carry that as a page break. Markdown has no pages, so
by default the break is dropped and reported — ask for the rule and each slide
gets its own section:

```ts
const deck = Ream.parse(pptxBytes);
const md = await deck.convert('md', { pageBreaks: 'rule' }); // --- between slides
```

A workbook has the same problem and a different answer: markdown cannot tell
one sheet from the next, so each opens with a heading carrying its tab name.
A printed page never shows one — Excel and Calc emit it nowhere — so this is
markdown's alone, and `{ sheetNames: false }` turns it off:

```ts
const book = Ream.parse(xlsxBytes);
const md = await book.convert('md'); // # Revenue, then its grid; # Costs, then its
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

## Password-protected Word and Excel documents

A password-protected `.docx` / `.xlsx` is not a zip at all — ECMA-376 §2.3 puts
the whole package inside an OLE container as ciphertext — so it has to be
decrypted before any reader sees it. The same `password` option does that, for
both Office schemes (the 2007 standard one and the agile one 2010 and later
write):

```ts
const doc = Ream.parse(docxBytes, { password: 'letmein' });
const pdf = await doc.convert('pdf');
```

Unlike a PDF, here a wrong password is an **error**, not a loss: every scheme
stores a verifier, so the file can say the password is wrong instead of handing
back rubbish. It comes as `WrongPasswordError` — a class, so a corrupt file is
not mistaken for a bad password:

```ts
import { Ream, WrongPasswordError } from 'reamkit';

try {
  Ream.parse(bytes, { password });
} catch (e) {
  if (e instanceof WrongPasswordError) promptAgain();
  else throw e; // not a password problem
}
```

A document that needs a password and gets none throws too, naming the option to
pass. An interface usually wants to know before that — to show a password field
rather than an error — and `isEncryptedPackage` answers from the bytes:

```ts
import { Ream, isEncryptedPackage } from 'reamkit';

const password = isEncryptedPackage(bytes) ? await askTheUser() : undefined;
const doc = Ream.parse(bytes, { password });
```

Decryption only: Ream never writes an encrypted package, so `convert` output is
always in the clear.

## pdf → html / docx: read a PDF back

`Ream.parse` also accepts a **PDF** — including a modern compressed one
(cross-reference streams, object streams) or an encrypted one. A tagged PDF (the
ones Ream writes) is rebuilt from its structure tree — headings, paragraphs,
tables, lists in reading order; an untagged PDF is reconstructed heuristically
from glyph positions. **Raster images, hyperlinks and the page's artwork come
back too** — images lifted out and sized from their placement (including the
ones written into the content stream, and stencil masks painted in the page's
own colour), link annotations re-attached to the text, filled paths, stroked
lines and shading-pattern gradients turned into shapes, the clipping paths that
limit them, tiling patterns, constant alpha, the appearance an annotation
carries — or one drawn from its properties where the file supplies none — and
the Type 3 glyphs that are drawings rather than letters. Colour comes back
through whatever space states it: device, CIE (`CalGray`, `Lab`), or a
`Separation`/`DeviceN` run through its own tint transform. The layers a file
turns off stay off, and the box it says to show is the box you get. The result
is an ordinary `FlowDoc`, so it converts onward like any other source.
Clip-bounded (`sh`) shadings are not read (reported as a loss).

A form or a drawing is not a reflowable document, though — its rules and boxes
are placed absolutely, and a label an inch from the box it labels says nothing.
**The file decides for itself**: a paper is mostly lines and is re-flowed, a
form or a drawing is mostly marks and is kept as a page, with every line where
its glyphs stand. `pdfLayout` overrides the choice when you know better:

```ts
const page = Ream.parse(pdfBytes, { pdfLayout: 'positional' }); // keep the page
const flow = Ream.parse(pdfBytes, { pdfLayout: 'flow' }); // …or re-flow it
const docx = await page.convert('docx'); // the form, not the form's words
```

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

## A PDF whose streams use a filter Ream does not carry

§7.4 leaves the filter set open, and PDF 2.0 added Brotli. A stream Ream cannot
decode is a stream it has not read — and when that stream is the
cross-reference, the file has no pages at all. Supply the decoder and it does:

```ts
import { brotliDecompressSync } from 'node:zlib';
import { Ream } from 'reamkit';

const doc = Ream.parse(pdfBytes, {
  filters: { BrotliDecode: (bytes) => brotliDecompressSync(bytes) },
});
```

In a browser, pass any Brotli implementation you already ship. Without one the
loss report names the filter, so a document that comes back empty says why.

## pdf → pdf: a form keeps its form

A PDF reads back as a re-flowable document — paragraphs and tables in reading
order. A form is not that: its grid is painted at fixed coordinates, and text
that flows beside the grid sits in none of its boxes. Ask for the page instead
and every line stands where its glyphs do:

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(pdfBytes, { pdfLayout: 'positional' });
const pdf = await doc.convert('pdf', { fonts });
```

Keep the default for anything going onward to DOCX, Markdown or HTML — the
placed reading has no reading order, no paragraphs and no tables to give them.

## pptx → pdf: render a slide deck

`Ream.parse` also accepts a **PowerPoint** `.pptx`. Each slide becomes a page at
the deck size, its shapes read as positioned content — text boxes, placeholders,
pictures, shapes, tables, charts, embedded objects, groups and hyperlinks. What
a deck states once reaches every slide that relies on it: the master's and the
layouts' own decoration (drawn under the slide's content), the background —
solid, gradient, picture or a theme reference — the deck's colour map, the text
style its `p:defaultTextStyle` and `p:txStyles` set, and the table styles in
`tableStyles.xml`. The result is an ordinary `FlowDoc`, so it converts onward
like any other source:

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(pptxBytes); // doc.format === 'pptx'
const pdf = await doc.convert('pdf', { fonts }); // a page per slide
const html = await doc.convert('html'); // …or the slides as flowed HTML
```

## Word review comments

Review comments (`w:commentReference`) render as a bracketed superscript marker in
the text and a "Comments" section after the body — author, date and content, with
reply threads nested and resolved threads flagged, the commented range highlighted,
and author identities resolved from `people.xml`. In PDF the marker is a clickable
jump to the entry; pass `commentAnnotations: true` to _also_ attach each comment as a
native sticky-note annotation (interactive output only — suppressed under PDF/A and
tagged):

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(docxBytes);

const pdf = await doc.convert('pdf', { fonts }); // markers + Comments section + highlights
const sticky = await doc.convert('pdf', { fonts, commentAnnotations: true }); // …plus /Text pop-ups
const html = await doc.convert('html'); // replies nested, resolved threads flagged

doc.flow.comments?.get('0'); // { content, author?, date?, authorId?, parentId?, done? }
```

Comments — threads and resolved flags included — also write back through
`convert('docx')`.

## Legacy .doc / .xls / .ppt

The binary Office formats parse through the same entry point and the same
interlayer — `Ream.parse(bytes)` sniffs OLE2/CFB and picks the reader, so every
target and every option below works on them unchanged. A legacy deck brings its
master's decoration and background, the text its slide list holds with the
sizes, colours, typefaces and bullets its master states, grouped shapes and
tables, picture fills and the metafiles it draws with; a legacy workbook brings
its styling, its drawings placed on the sheet's own grid, its charts, its print
model and its conditional formatting.

```ts
const doc = Ream.parse(bytes);     // .doc / .xls / .ppt — sniffed like the rest
const pdf = await doc.convert('pdf');
const xlsx = await doc.convert('xlsx'); // a .xls re-writes as a modern workbook
```

## Excel pivot tables

A pivot table's cached output renders as an ordinary grid; on top of that Ream applies
the table's named pivot style (`pivotTableStyleInfo`) — banded rows and a styled
header — and emphasises the grand-total / subtotal rows and columns. Nothing is
recomputed from the cache:

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(xlsxBytes);
const pdf = await doc.convert('pdf', { fonts });
const html = await doc.convert('html');
```

## Excel data validation

A `list` data validation (`<dataValidations>`) paints an in-cell dropdown affordance
— a small button with a ▾ — on every cell of its range, in both PDF and HTML. The
validation, its formulas and the input/error prompts also round-trip through
`convert('xlsx')`:

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(xlsxBytes);
const pdf = await doc.convert('pdf', { fonts }); // list cells show a ▾ dropdown button
const html = await doc.convert('html');
const xlsx = await doc.convert('xlsx'); // the <dataValidations> survive the round-trip
```

## Excel slicers

A slicer (`xl/slicers` + `xl/slicerCaches`) renders as a captioned button box after the
grid — the way a chart frame does. A native-table slicer fills its buttons from the
referenced table column's distinct values and highlights the ones the column's
autofilter keeps; an OLAP/pivot slicer whose items live in a pivot cache degrades to a
caption-only box.

```ts
import { Ream } from 'reamkit';

const doc = Ream.parse(xlsxBytes);
const pdf = await doc.convert('pdf', { fonts }); // a button box per slicer, after the grid
const html = await doc.convert('html');
```

## SmartArt diagrams

SmartArt (DOCX and PPTX) renders either way round. Where the file caches a
pre-rendered DrawingML drawing (`diagrams/drawing#.xml`) that is drawn directly;
where it does not, the diagram is laid out from its own `data`, `layout`,
`colors` and `quickStyle` parts — the algorithms (`lin`, `composite`, `cycle`,
the hierarchy pair, `snake` and the rest), the constraint solver that sizes and
spaces the boxes, and the colour and style parts that paint them. Scheme colours
resolve through the document or deck theme, and box text is measured so a box
grows to what it holds.

Nothing special is needed to get it — it is part of the ordinary conversion:

```ts
import { Ream } from 'reamkit';

const { bytes, losses } = await Ream.parse(docxBytes).convertWithReport('pdf', { fonts });
// `losses` is where anything the file asks for and Ream cannot draw is reported.
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
`.odttf`, and a deck's `p:embeddedFontLst`) always win — glyph-exact, no
substitution. A `w:rFonts` that names a **theme** slot rather than a family
resolves through the theme's major / minor fonts before any of this runs.

The substitute set is Latin, so a document holding another writing system is
served separately: a Noto face is fetched per SCRIPT — Japanese, Korean,
Chinese, Arabic, Hebrew, Thai, geometric symbols — and only for the scripts the
document actually contains. It is chosen per character, so a paragraph that
mixes Latin and Arabic is drawn in both faces rather than in one with holes:

```ts
// A sheet of Korean: NotoSansKR is fetched, in the regular weight only.
// The same call on a Latin-only document downloads nothing extra.
const pdf = await Ream.parse(bytes).convert('pdf');
```

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
