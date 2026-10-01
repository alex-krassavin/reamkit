---
title: Scope & limitations
description: What Ream implements today, and where the edges are.
---

The conversion core is broad and spec-driven, but ECMA-376 is vast and Ream does
not implement all of it — this page is an honest map of what works and what
doesn't yet.

## Implemented

**Input** — Ream parses **Word (`.docx` and legacy `.doc`)**, **Excel (`.xlsx`
and legacy `.xls`)**, **PowerPoint (`.pptx` and legacy `.ppt`)** and **PDF**, sniffed
from the bytes. The legacy binary `.doc` / `.xls` / `.ppt` (the OLE2/CFB formats) are
read through a shared container reader — see WordprocessingML / SpreadsheetML /
PresentationML below. A **PowerPoint** deck becomes
one page per slide at the deck size, its shapes read as positioned content: text
boxes (run formatting, alignment, vertical anchor, bullets, indents), the text
style the deck states once, layout/master placeholders and the decoration they
carry, backgrounds (solid, gradient, picture or theme reference), pictures with
their recolouring, shapes (geometry/fill/stroke/gradient), DrawingML tables with
the style they name, embedded charts and objects, theme colours through the
deck's colour map, grouped shapes and run hyperlinks; text set to shrink is
measured and shrunk to fit its box, and a SmartArt is laid out from its own
parts whether or not the file cached a drawing. Not read (a graceful loss):
alpha/roman list numbering.
PDF input handles classic and modern compressed files
(cross-reference streams, object streams) and encrypted files (RC4 / AES; the
user password is passed to `Ream.parse`, defaulting to the empty permissions-only
case). A **tagged** PDF (including the ones Ream writes) is rebuilt from
its structure tree — headings, paragraphs, tables, list items, reading order,
the elements a document names for itself followed through its `/RoleMap` to the
standard ones they stand for, and an inline element (a span, a link, a formula
set in a line, MathML) read as a stretch of its parent's line rather than a
block of its own (§14.8); an
**untagged** PDF is reconstructed heuristically from glyph positions (lines by
baseline, paragraphs by spacing and by the indent the page set them with,
headings by relative font size, a word the line broke in half put back together,
and the page read for its own gutters: as many columns of PROSE as it was set
in — a page of labels with figures against the right margin breaks a dozen lines
at the same x and is not two columns, so it is read across and each label keeps
its figure — blocks set SIDE BY SIDE, each at a leading of its own, set as the
borderless table a writer sets them in and each read down on its own (an
invoice's stack of dates and numbers beside the address it bills; blocks whose
lines share their baselines stay one text, read across on its stops), a page
RULED into a grid rebuilt as a table with its measured columns and rows, a gap
written as a TAB with the stop the page set the piece after it at, where no
word space could make it or where it lands on a stop the lines around it share,
consecutive lines standing on the SAME stops rebuilt as a table of their own,
with a column whose figures end together set against its right edge, a line
drawn between two blocks given to the block it separates as that paragraph's
border, a rule typed out of hyphens left on the line of its own it was written
on, a band repeated at the top or bottom of every sheet as a running head or
foot with its regions and its page number as a field, and a bracketed block of
rows and columns as an OfficeMath matrix), which is approximate. A page in
columns is read down them: a gutter is the white between two columns, not where
most of a ragged column's lines end; a JUSTIFIED column, flush on both sides, is
prose as surely as a ragged one; and a column a few ems wide beside a gap the
page's lines run across is a table's. A page that turns to its next column short
of the foot breaks there, and a last page set in balanced columns ends their
section on that page. A line set wholly in a typewriter face is CODE, a listing
kept line for line with its comments on the stops the page set them at; a row of
four cells and more is a TABLE's, read whole where it reaches across the gutter,
the line over it its heading, a rule across it the border of the row under it and
a rule drawn down it the division between two of its columns; a drawing and the
words set on it — boxes and arrows, a bar chart — are one FIGURE, set as an
inline group in its column; a line that opens with a bullet or a hanging label
opens a list item, and an INDEX's entries hang, each a paragraph with the lines
that carry it on. A box filled behind a single line — a heading on a dark bar —
is that paragraph's shading. A running band's numeral that changes from page to
page is the page's NUMBER, counted in the numerals the page prints: front matter
in roman starts a section of its own, a book that leaves its blank pages out is
counted all the same, and a book that signs its even and odd pages at opposite
edges gets a foot to each side (§17.10.1). A page whose words run down its sheet
comes back as a section set that way (§17.6.20 `tbRl`). The margins are
measured to the page's INK — a blank is
not ink, and a picture is: a sheet whose one line of text sits over a picture
wider than that line keeps a measure the picture still fits. Text comes
back via each font's `/ToUnicode` map (a range of it judged code by code,
§9.10.3), or — where a composite font ships none — from the reverse `cmap` of
the program it embeds (§9.10.2), or — where a SIMPLE
font ships none, which is every PDF from TeX — from the glyph NAMES its
`/Encoding /Differences` states (§9.6.6.1) over the base encoding the font
declares (Annex D.2 Standard, WinAnsi and MacRoman, or a symbolic font's own
built-in one — ZapfDingbats is a font of PICTURES and is read as pictures,
Annex D.6), and a code the differences leave `.notdef` is no character at all.
The names TeX gives its Greek and its mathematics — `alpha`, `epsilon1`,
`asteriskmath` — are read like any other, and an accent TeX strikes over a letter
it has no glyph for is composed with that letter: "naïve", not "na¨ıve". A word
space is read where the page STEPS one as wide as most of its face's space as
well as where it draws one, a dot stepped on after a word is the rest of an
ellipsis, and a dash closed up to its word takes no space where a line breaks.
A ligature comes back as the letters it joins — "ﬁrst" is `first` to a search
and to an edit — and a composite font the file names but does not embed, with
`Identity-H` and no `/ToUnicode`, is read in the glyph order of the face it
names where that is one of the Microsoft core faces (§9.7.4 — Arial, Times New
Roman, Verdana and their like keep the Macintosh standard order for their first
258 glyphs). A face that cannot be addressed by character at all — a TrueType
program with no usable `cmap` (§9.6.6.4) — is not guessed at: its codes are
glyph indices, the text is marked unrecoverable and the glyphs are DRAWN from
their outlines, a word at a time, as are the glyphs of a font that names them
after nothing a reader knows (CFF charstrings run as the programs they are,
`glyf` outlines and their composites, `post` names over the Macintosh standard
order). Punctuation a font names no character for — a producer that maps every
mark to U+0000 leaves an invoice with no hyphen, colon or bracket in its text —
is read off the SHAPE the page draws, where that shape says one thing: a flat
bar at the middle of the lower case is a dash, two stacked dots a colon, one
small mark on the baseline a full stop or a comma. The
predefined CMaps a file names instead of embedding are known (§9.7.5.2),
including the vertical ones that advance the pen down the page. A run names its
face's FAMILY — `Inter` for `Inter-SemiBold`, without the producer's debris (a
`,Bold`, a `-Identity-H`, a serial number) — with the kind of face it is, so a
program that lacks it substitutes one of the same kind. The weight and slant
come from the `/FontDescriptor`, and where it states no weight from the header
of the program itself; a face's own name outweighs the flags that contradict it
(a `-Light` face is not bold for a ForceBold hint), and a family the
substitution tables know decides whether it is a serif, a sans or a typewriter
face. The embedded font programs themselves are carried into the output, so a
rebuilt page is set in the type it was set in. Written to `.docx`, each face the
page drew with is rebuilt as TrueType — the outline of every character it drew,
at the advance the page set it at, with its kerning pairs (GPOS `kern`) and the
ligatures the page drew (GSUB `liga`) — and embedded where its licence (OS/2
`fsType`) lets it, and a word space TeX shrank to fit a line is written as narrow
as TeX set it (§17.3.2.43 `w:w`): a re-set line breaks where the page broke it,
on a machine that has none of the document's fonts.
**Raster images, hyperlinks and vector artwork** are lifted back out too (JPEG
verbatim but for a baseline CMYK one, which every viewer reads in Adobe's
inverted convention and which is decoded instead, Adobe's YCCK included, and
given back in the colours its inks paint; other images re-encoded as PNG with
soft-mask
alpha — including JBIG2, the bilevel coding a scanner stores a page of text in
(ISO/IEC 14492:
generic, refinement, symbol-dictionary, text and halftone regions, arithmetic
and Huffman, and a dictionary whose symbols are REFINEMENTS of the ones it
already holds), and CCITT Group 3 two-dimensional, where every line begins with
the bit that says how it is coded — `/Link` URIs re-attached to the text,
filled paths, stroked lines — dashed in the page's own pattern and capped as it
caps them (§8.4.3.6 `d`, §8.4.3.3 `J`) — and shading-pattern gradients turned
into shapes), colour set through a named space (§8.6.8 `cs` / `sc`, which is how
every PDF a browser prints states it) — including the CIE
spaces `CalGray` (§8.6.5.6), `CalRGB` (§8.6.5.7, gamma and matrix into XYZ,
adapted onto D65 by Bradford) and `Lab` (§8.6.5.8), an `/ICCBased` stream
decoded through the profile it carries (§8.6.5.5), and a `Separation` or
`DeviceN` run through its own tint transform (§8.6.6.4/§8.6.6.5) — for an
IMAGE as much as for a fill — which needs
all four kinds of PDF function (§7.10: sampled, exponential, stitching, and the
type-4 PostScript calculator) — along with clipping
paths (§8.5.4 `W` / `W*`, applied to paths and pictures alike, and carried into
a turned picture's OWN axes), stencil image masks (§8.9.6.2, painted in the
page's non-stroking colour), colour-key masks (§8.9.6.4 — the samples a `/Mask`
array names left unpainted, held against the samples before any `/Decode`),
images written into the content stream (§8.9.7 `BI` … `EI`), tiling patterns
(§8.7.3 — read as the pattern's colour at the strength its tile covers where
they fill type, and where they fill a shape with a tile drawn in lines and dots,
a hatch or a grid, which comes back as the tone it gives the page rather than as
the hatch itself; a tile that holds a picture is lifted as that picture; an
uncoloured pattern, §8.7.3.3, takes the colour `scn` gives it), constant alpha
(§11.6.4.4 — a fill's `/ca` and a stroke's `/CA` through the `gs` operator),
Type 3 glyphs (§9.6.5 — content streams, drawn as the artwork they are in every
reading: such a font has no program and no face, so its codes are never re-set
in a substitute, which would draw the letter a second time),
annotation appearance streams (§12.5.5, so a form field draws itself — inside
its `/BBox`, which is its clip, and what it writes, a button's caption or a
field's value, stands where the annotation stands, over the box it is written
on, rather than being read into the page's lines; a form that asks for its
appearances rebuilt, §12.7.2
`/NeedAppearances`, shows its fields' values rather than the stale text of the
old appearance), the text
render modes (§9.3.6 `Tr` — stroked type keeps its outline, and the invisible
modes an OCR layer uses are marked rather than painted), the box a viewer shows
(§14.11.2 `/CropBox` — which sizes the page AND clips the marks outside it; a
media box with no area is the Letter sheet a page that states none gets) and
the page's own `/Rotate` (§14.11.1). A file's optional content is honoured
(§8.11 — the layers its default configuration turns off stay off, on `/OC … BDC`
marked content and on an XObject's own `/OC`, through an `/OCMD`'s policies).
An annotation the file gives no `/AP` is drawn from its own properties (§12.5.5
— Ink, Line, Square, Circle, Polygon, PolyLine, a text field's `/V` set in its
`/DA`, a free-text note's `/Contents` set in its `/DA` (§12.5.6.6), and the text
markups' `QuadPoints` applied to the words they cover). The
built-in metrics of the 14 standard fonts (§9.6.2.2) stand in where a file
embeds neither the face nor a `/Widths` array. A gradient the page paints with
a bare `sh` is read as the shape it covers (§8.7.4.5, function-based type 1
included), in the colour space the shading dictionary states; an axial one runs
the way its matrix carries it onto the page, and a shape it fills shows the part
of its axis the shape covers (§8.7.4.5.3) rather than the whole axis squeezed
across it; text filled with
one comes back in its colours rather than in black, and the loss report says
where the gradient's shape was flattened to a single colour.

**Output** — `convert('pdf')`, `convert('svg')` (a page-stack preview),
`convert('html')` (flowed, needs no fonts) — a workbook's HTML is Excel's window and
its SVG a picture of each sheet, whole (SpreadsheetML below) — `convert('md')` (GitHub-Flavored
Markdown), `convert('docx')` (write WordprocessingML back out) and
`convert('xlsx')` (write SpreadsheetML back out — spreadsheet input only). The writers are for normalization, sanitization,
in-browser editing, and round-tripping. The docx round-trip is semantic, not
byte-exact, but complete — text, tables, images, lists, links, headers/footers,
multi-section geometry, footnotes/endnotes, charts and OfficeMath all write
back, and so do a paragraph's tab stops (§17.3.1.38) and its own rules
(§17.3.1.24 `w:pBdr`), and the page-number fields a running foot is written with
(§17.16.19 `w:fldSimple` — `PAGE` and `NUMPAGES`, so the second sheet does not
say it is the first). So do a paragraph's shading (§17.3.1.31 `w:shd`) and how
far its rules stand off it, its widow control and whether it keeps with the next
paragraph or keeps its lines together (§17.3.1.44, §17.3.1.14–15); a run's width
and character spacing (§17.3.2.43 `w:w`, §17.3.2.35 `w:spacing`), its kerning and
its ligatures (`w:kern`, `w14:ligatures`); a line break, as `w:br` rather than a
newline inside the text; a line spaced in lines, in the 240ths `w:lineRule="auto"`
counts; a table's indent (§17.4.65 `w:tblInd`); a drawing group (`wpg:wgp`); and a
section's page-number numerals (§17.6.12 `w:pgNumType`) and the way its lines run
(§17.6.20 `w:textDirection`). The document's settings write back too — headers of
its own for the even pages, a gutter at the head, `doNotExpandShiftReturn` and the
compatibility mode (§17.15.1) — and the styles part states the defaults the
writer leaves out, which Word would otherwise take from its own Normal template.
A picture keeps how it is drawn into its frame — its crop
(§20.1.8.55 `a:srcRect`), its turn and mirror (§20.1.7.6 `a:xfrm`) and its
opacity (`a:alphaModFix`) — so a cropped picture read from Word and saved again
no longer comes back whole and squeezed into the frame its crop was sized for; a
fill and a line keep the opacity they are painted at (§20.1.2.3.1 `a:alpha`),
and a line its own dash pattern (§20.1.8.21 `a:custDash`). A run names its
font's family, and the font table says what kind of face that is (§17.8.3), so
a reader that lacks the family substitutes one of the same kind. A section
states where it must OPEN as well as where it begins: a
`w:type` of `oddPage` or `evenPage` (§17.6.22) survives the round trip and is
laid out as Word lays it out, printing the blank sheet it implies — charged to
the section that ends, not the one that starts. A drawing that states where on
the page it belongs is written back as
the `wp:anchor` that puts it there (§20.4.2.3) rather than flattened into the
text flow. No image part the format cannot show goes into the package: JPEG
2000 is not among the parts §15.2.14 admits, and is dropped with a loss that
names it rather than written as a hole. The xlsx round-trip preserves the whole grid surface — cells, styles,
merges, the print model, conditional formatting, sparklines, tables and embedded
charts — and is byte-stable across a read↔write loop.

**Stream filters the reader does not implement (§7.4)** — a filter it cannot
undo leaves that stream unread, and when the unread one is the cross-reference
the whole document is missing. Rather than carry a decoder for every filter
anyone might write, `Ream.parse(bytes, { filters })` takes one from the caller:

```ts
import brotliPromise from 'brotli-dec-wasm'; // ~200 KB wasm, browsers and workers
const brotli = await brotliPromise;
Ream.parse(pdf, { filters: { BrotliDecode: (b) => brotli.decompress(b) } });
```

On a server the runtime already carries one, so `brotliDecompressSync` from
`node:zlib` goes in the same slot. Brotli is RFC 7932's context-modelled
Huffman scheme over a 122 KB static dictionary, and a filter almost nothing
produces does not belong in every browser bundle — which is why this is a hook
rather than a dependency. Absent, or throwing, the filter is reported
unreadable by name rather than producing an empty document silently. FlateDecode, LZW, RunLength, ASCII85,
ASCIIHex, CCITT, DCT, JPX and JBIG2 need nothing supplied.

**The two PDF readings** — a
conversion cannot have both readings at once: words that move cannot agree with
rules that do not, so anchored artwork over reflowed text lines up with none of
it. The FILE decides which it gets. A paper is mostly LINES with a rule or two
between them and is read as a re-flowable document — paragraphs and tables in
reading order, from the structure tree where the file has one, and where it has
none, in the columns of PROSE the page was set in, with its running head and
foot kept out of the body, its blocks set side by side read down each on its
own, its ruled grids, its rows-on-the-same-stops and its tables across the
columns rebuilt as tables, its listings kept line for line, its figures kept
with their labels, and the lines it drew between blocks given to the paragraphs
they separate; what a form
field or a button writes still stands where the page placed it, over the box it
is written on.
A form or a
drawing is mostly MARKS (one form sets 28 numbered rows in 355 ruled boxes) and
is read as a page: every line stands where its glyphs stand, beside the artwork,
with no reading order, no paragraphs and no tables. The marks are counted
against the baselines on the MEDIAN page, so one plan folded into a report does
not make the report a plan; the reader records which reading it took and why.
There is no option to override it. A caller holding a `Uint8Array` cannot know
whether it is a paper or a form, so asking it to choose only moved the guess
outward — and left the choice the library actually makes untested, because
everything that measured it pinned the other one.

**Markdown (`convert('md')`)** — GitHub-Flavored Markdown, a flow medium like
HTML: no pagination, no layout, no fonts, zero I/O. It keeps what a document
*says* and drops how it *looks*. Kept: headings (from `w:outlineLvl`, falling
back to a `Heading N` / `Title` style id), bold, italic, strikethrough, ordered
and bullet lists with their real numbers and nesting, tables as GFM pipe tables
with per-column alignment, hyperlinks (through the same scheme allowlist the
HTML writer uses), pictures, footnotes and endnotes as GFM footnotes, review
comments as footnotes attributed to their author, bookmarks as inline anchors,
and the text inside shapes. A spreadsheet opens each sheet with a heading
carrying its tab name — markdown has no pages to tell one sheet from the next
by, so `{ sheetNames: false }` is there for the bare tables. Underline and super/subscript survive as `<u>` /
`<sup>` / `<sub>`, which GFM parses as inline html.

Dropped, each reported once in the loss report rather than silently: alignment,
indents, colour, font family and size, tab stops, page and column breaks,
headers and footers, charts, shape geometry, and math. A page break can be kept
instead — `{ pageBreaks: 'rule' }` writes the `---` thematic break a slide deck
wants, since the `.pptx` and `.ppt` readers mark each slide boundary with one
and it is the only structure a deck has. Merged cells flatten —
markdown's table is a plain grid — and a nested table flattens into the cell
that holds it. Pictures are inlined as `data:` URIs by default; pass
`{ images: 'link' }` to reference them under `./media/` and write the bytes
yourself, or `{ images: 'drop' }` to omit them.

**Password-protected packages (ECMA-376 §2.3, MS-OFFCRYPTO)**
- An encrypted `.docx` / `.xlsx` / `.pptx` is not a zip at all: the whole OPC
  package is ciphertext inside an OLE container. Given the password
  (`Ream.parse(bytes, { password })`) it is decrypted and read like any other
  document. Both schemes in the wild are supported — the **standard** one
  (Office 2007 and LibreOffice: AES-ECB under a key spun from 50 000 SHA-1
  rounds) and the **agile** one (Office 2010 and later: an XML descriptor
  naming its own hash and cipher, the package cut into 4096-byte segments with
  an IV each). `isEncryptedPackage(bytes)` says a file wants a password before
  the parse does, and each scheme carries a verifier, so a **wrong password is
  refused** (`WrongPasswordError`) rather than yielding rubbish. Decryption
  only — Ream never writes an encrypted package.

**WordprocessingML (§17)**
- Text, runs and the full style cascade (`docDefaults` → styles → direct formatting).
- Tables — auto / fixed layout, §17.4 border-conflict resolution, cell shading,
  vertical merge and grid span, nested tables, **table styles** (`w:tblStyle` with
  conditional formats: banding, first/last row/column).
- Lists and numbering (`abstractNum`, level overrides), multi-level.
- Sections — per-section page size and orientation, headers and footers (a
  pair for the even and the odd pages, `w:evenAndOddHeaders`, and a first page
  of its own), **multi-column layout** (`w:cols`), and **lines that run down the
  sheet** (§17.6.20 `w:textDirection` `tbRl` / `btLr` — the text turned a
  quarter, the header and footer left across the sheet, as Word sets them).
- **Hyperlinks** — external (clickable PDF annotations + HTML `<a>`, scheme-allowlisted)
  and internal: bookmarks become named destinations / `#`-anchors.
- **Fields** — `PAGE` / `NUMPAGES` render real page numbers in headers and
  footers, in the numerals the section names (§17.6.12 `w:pgNumType w:fmt` —
  front matter in roman, a body that starts again at 1).
- **Laid out the way Word sets it** — each rule measured in Word for Mac. A line
  stands as tall as Word sets its faces: the face's hhea line, or its typo line
  where it asks for that, the tallest ascent and the deepest descent on the line
  taken apart, a list marker adding its ascent, a picture standing on the
  baseline. Two paragraphs stand the larger of their spacings apart, not the
  sum, and a space before is dropped at the head of a page or column where Word
  drops it; automatic spacing is 14pt in every compatibility mode. A paragraph's
  first and last lines never stand alone (§17.3.1.44 widow control), a heading
  goes to the next page with what it heads (`w:keepNext`), and a paragraph asked
  to stay whole does (`w:keepLines`). A table's horizontal borders take the room
  they are wide, and a Word 2010 table (compatibility mode 14 and earlier) stands
  out by its first cell's margin. Each page sets its body below the header and
  above the footer it shows, and a page break typed on a line of its own stays
  on the page it ends.
- **Footnotes and endnotes** — notes at the bottom of the referencing page behind
  Word's separator rule; endnotes after the body.
- **Review comments** (`w:commentReference`) — a bracketed superscript marker in the
  text and a "Comments" section after the body with each comment's author, date and
  content (PDF and HTML). Reply threads and resolved state come from
  `commentsExtended.xml` (replies nest under their parent, resolved threads are
  flagged); the commented range (`w:commentRangeStart/End`) is highlighted; author
  identities resolve from `people.xml`. As an opt-in, comments can also be emitted as
  native PDF sticky-note annotations (`commentAnnotations`, interactive output only).
- **SmartArt** — the diagram's own parts (`data`, `layout`, `colors`,
  `quickStyle`) are laid out by the same engine the deck reader uses, so a
  document that cached no drawing renders rather than degrading; where a
  pre-rendered drawing (`diagrams/drawing#.xml`) is present it is used directly.
- Inline and floating images (PNG / JPEG / JPEG2000 / GIF / **TIFF** / **BMP** —
  baseline TIFF 6.0 and Windows bitmaps are decoded to samples, since PDF has no
  filter that carries either; a BMP is read at every depth from one bit to
  thirty-two, run-length and bit-field forms included),
  including **legacy VML pictures** (`<w:pict>` / `<w:object>` — ActiveX and
  OLE-object previews, images from older Word); floating drawings
  (`wp:anchor`) render outside the text flow — wrap-none (incl. `behindDoc`,
  which Word reads beside no other wrap) for watermarks/stamps/text boxes, and
  side wrapping (`square`/`tight`/`through`) where the body text flows around
  the exclusion area, on **both sides** of it where the drawing leaves room,
  each line holding what its indents and the drawing's stand-off leave. A
  drawing anchored to a paragraph that breaks to a new page goes with it. A drawing
  **anchored inside a table cell** is placed against the cell (or against the
  page, where `layoutInCell` is off), and a picture's **washout**
  (`@gain`/`@blacklevel`) prints as the contrast and brightness it asks for.
- **Metafiles** — EMF (MS-EMF) and WMF (MS-WMF), the format Word writes for
  anything it draws itself: an embedded workbook's preview, a legacy clipart,
  an ActiveX control's face, an equation. The records are played through their
  own device context — pen, brush, font, current point, window/viewport and
  world transform — and drawn as paths and text: lines, polygons, polylines,
  Béziers, rectangles, ellipses, path brackets, brush blits and both text
  records. In the flow, floating, in a header, inside a cell, or inline in a
  line of text.
- **Text boxes chained** with `wps:linkedTxbx`: what overruns one box continues
  in the next of its chain.
- Tracked changes (`w:ins` / `w:del`).
- Reads both **Transitional and Strict** (ISO 29500) packages; block-level
  content controls (`w:sdt`) flow through.
- **Legacy `.doc`** (Word 97–2003) — the binary `WordDocument` stream inside the
  OLE2/CFB container is read for its **text and formatting**: the FIB locates the
  piece table (the CLX), whose pieces — 16-bit Unicode or 8-bit Windows-1252
  ("compressed") — are stitched back into the document text and split into
  paragraphs, while the CHPX and PAPX runs (located through the `PlcfBteChpx` /
  `PlcfBtePapx` bin tables and decoded from their sprms) carry **bold / italic /
  underline / font size** onto each run and **alignment / indentation / spacing**
  onto each paragraph. **Tables** are reconstructed too — the in-table paragraphs
  (marked by the `fInTable` / `fTtp` PAPX flags, cells delimited by the `0x07`
  cell mark) become a row-and-cell grid, with per-column widths, **per-cell borders
  and vertical merges** (the table definition's `TC80` array) and **per-cell
  background shading** (`sprmTDefTableShd`'s `cvBack` fill) — and **inline images**
  are extracted (the
  picture character's CHPX points at a PICF in the `Data` stream; the raster blip
  is pulled out and sized from the PICF). **Fields** resolve to their cached
  result — the field code (`PAGE`, `NUMPAGES`, `REF`, …) is dropped and the stored
  result text kept. The section's **headers and footers** are lifted from the
  `PlcfHdd` stories (best-effort: the binary story ordering can't be ground-truthed
  here, so only well-formed stories are surfaced). **List items** (`sprmPIlfo` /
  `sprmPIlvl`) render with their resolved **number format** — a real "1." / "a)" /
  "iii." or the bullet glyph, from the `LST` / `LVL` / `LFO` tables. So an old `.doc`
  renders to PDF/SVG/HTML and re-writes to `.docx`. Legacy drawing shapes / text
  boxes and comments are not read (re-save as `.docx` for full fidelity); an
  encrypted `.doc` yields no text — the binary formats lock themselves with the
  older RC4 scheme, which is not the OOXML one above.
  The shared CFB reader
  (`src/core/ole`) is the same keystone `.xls` uses.

**SpreadsheetML (§18)**
- Grids, shared strings, number formats and dates (incl. the 1904 date system).
- **Measured against Excel itself** — its own PDF of a probe workbook built for each
  question, and the absolute positions it writes beside every anchored drawing.
  Columns are counted in the Normal font's digit as Excel counts it: a Chinese,
  Japanese or Korean face's is a pixel more than its half-width figure (宋体 12 is
  9px). A column no `<col>` covers and no `defaultColWidth` sets is `baseColWidth`
  digits and padding, rounded up to Excel's step of eight pixels (Calibri's 64px). A
  colour's `tint` is worked as Excel works it, in the Windows HLS model, on a theme,
  `rgb` or indexed colour alike; a workbook with no theme part takes the Office 2023
  colours Excel starts from.
- **On a screen — HTML and SVG — a workbook is what Excel's window shows**, not its
  printed pages: every visible tab, an empty one too; one grid a sheet, all of its
  columns at full size; no print area, repeated titles, page breaks, printed headings
  or centring; text running on across the row; the drawings over the cells where they
  are anchored; the window's gridlines where `showGridLines` asks for them, on every
  edge no border claims and no fill covers; every noted cell flagged with Excel's
  red corner, purple for a threaded comment. HTML makes each tab a
  `<section data-sheet>` under its name and shows a note on hover; SVG draws each
  sheet whole on a page as large as what is on it, every glyph from its face's own
  outline with the words laid over it unseen for search and selection. PDF is the
  printed pages, by the print model below.
- **Sheets that read from the right** (`rightToLeft`) — the grid, each cell's left
  and right borders and the drawings turn round as Excel turns them (a right arrow
  still points right), and on paper the sheet meets the right margin. A cell's text
  keeps the side its alignment names, and General aligns text by its first strong
  character on any sheet.
- **Legacy `.xls`** (BIFF8, Excel 97–2003) — the binary `Workbook` stream inside the
  OLE2/CFB container is read into the same grid model, so an old `.xls` renders to
  PDF/SVG/HTML and even re-writes to `.xlsx`. Cell values, structure (sheets, shared
  strings, merges, column widths, custom row heights, frozen panes, the 1904 flag),
  **styling** — fonts, fills,
  borders, number formats and alignment from the FONT/FORMAT/XF records, with colours
  resolved through the BIFF colour palette — **embedded pictures** (from the
  Office-Drawing/Escher BLIP store, metafiles included), **embedded charts** (the
  BIFF chart substream, plotted from the worksheet cells its AI records reference)
  and **drawing shapes** (autoshapes + text boxes, from the Escher shape records and
  their TXO text) are read — each placed and sized on the sheet's own columns and
  rows, since a `.xls` drawing anchors to cells, and a grouped shape through the
  group that anchors it — plus **cell hyperlinks** (the HLINK record's URL moniker), the **page-setup
  print model** (orientation, scale, fit-to-page, margins, gridlines, centering,
  header/footer and manual page breaks), **defined names** (named ranges plus the
  print area and repeated titles, from the NAME records), **cell comments** (the
  Note record's author + the text-box text), **data validation** (the rule type,
  ranges and a `list` rule's in-cell dropdown) and **conditional formatting** — both
  the classic `cellIs` / `expression` rules (with their differential fill / font
  colour) and the 2007 **colour-scale / data-bar / icon-set** extensions (the CF12
  records, present only in a `.xls` re-saved by Excel 2007+); only a graphical rule
  whose colour is theme-relative rather than a literal value degrades gracefully.
- The print model — gridline suppression, print area, fit-to-page scaling, repeated
  print titles, manual page breaks, horizontal/vertical centering, and **column-band
  pagination**: a sheet wider than the page (and not fit-to-width) splits across
  pages, all rows of the left columns first, then the next band ("down, then over"),
  honouring manual column breaks — instead of being squeezed onto one page width.
  Printed row and column headings shrink with a sheet the print scale fits; a chart
  sheet prints landscape unless it says otherwise, its chart filling the page.
- **Frozen panes** round-trip through the writer and become sticky header rows /
  columns in HTML output. They do not affect PDF — in Excel freezing is a view
  setting that does not print (the printed repeat is the print titles above).
- **Conditional formatting** — the highlight rules: `cellIs` (compare-to-constant),
  `top10` (top/bottom N or N %), `aboveAverage` (mean, optionally shifted by N
  standard deviations), `duplicateValues` / `uniqueValues` (value frequency across
  the range, numbers by value and text case-insensitively) and the text tests
  (`containsText` / `notContainsText` / `beginsWith` / `endsWith`); plus the
  visual encodings `colorScale` (2/3-stop gradients), `dataBar` (in-cell bars,
  with a zero axis so negative values run the other way) and `iconSet` — traffic
  lights, arrows, signs, symbols (check / exclamation / cross), flags, ratings (a
  bar meter) and quarters (a clock pie). The cross-cell rules resolve against the
  range's value extent. Also `expression` — an arbitrary formula evaluated per cell
  by a built-in formula engine against the workbook's cached values (no
  recalculation): ~140 functions (logic / info incl. `IFS`/`SWITCH`/`XOR` and the
  `IS*` family; the math, trig and exponential set; the `SUM`/`COUNT`/`MEDIAN`/
  `SUMPRODUCT`/`STDEV`/`VAR`/`PERCENTILE` aggregates and the `COUNTIF(S)`/`SUMIF(S)`/
  `AVERAGEIF(S)` predicates; text; date / time; the `MATCH`/`INDEX`/`VLOOKUP`/
  `HLOOKUP` lookups and `ROW`/`COLUMN`), sheet-qualified references (`Sheet2!A1`),
  defined names, inline array constants (`OR(A1={1,3,5})`) and the per-cell
  relative-reference shift. A construct genuinely beyond a deterministic per-cell
  predicate — a 3-D reference, a dynamic-array / `LAMBDA` idiom, or a volatile /
  dynamic-reference function (`RAND`/`INDIRECT`/`OFFSET`) — evaluates to an error, so
  the rule simply does not paint rather than misrender. And `timePeriod` (today /
  this-week / last-month … windows). Both stay deterministic: `timePeriod` and
  `TODAY()`/`NOW()` read an explicit reference date you pass as `now` (never the
  system clock), so without one those clock-relative rules simply don't paint. The
  highest-priority matching rule claims the cell's fill / font; a data bar or icon
  applies on top. Rules the 2009 extension carries in `<extLst>` (`x14:cfRule` —
  a child `<xm:sqref>`, `<xm:f>` formulas and an inline `<x14:dxf>`) are read
  beside the ones the base schema declares; a `cellIs` operand may be text as
  well as a number.
- **Sparklines** — per-cell line / column / win-loss mini charts, including
  cross-sheet data ranges and blank-cell gaps.
- **Excel tables** (`xl/tables`) — drawn in the style they name: the sixty styles Excel
  builds in (TableStyleLight1…21, Medium1…28, Dark1…11, from the standard's own
  `presetTableStyles.xml`) and the ones a workbook defines in `<tableStyles>`. Each
  region is laid over the table in §18.8.41's order — the whole table, column and
  row stripes, last and first columns, header and total rows, their corner cells —
  with its fill, weight, colour and rules; stripes run through the data rows only.
  The cell's own formatting stays on top, and its text takes the table's colour only
  where its font names its colour as the Normal style does. **Structured references**
  (`Table[Column]`, `[#All]`, `[#Data]`, `[#Headers]`, `[#Totals]`, `[@Column]`,
  column ranges) evaluate in conditional formatting — a rule's formula, a colour
  scale's, data bar's or icon set's stops — and so does a defined name that is
  itself a formula.
- **Pivot tables** (`xl/pivotTables`) — Excel caches the pivot's output cells in the
  sheet, so the grid renders as data; on top of that Ream applies the named pivot
  style (`pivotTableStyleInfo`) — banded rows and a styled header — and emphasises
  grand-total / subtotal rows. The pivot is not recomputed from its cache.
- **Data validation** (`<dataValidations>`) — a `list` validation paints an in-cell
  dropdown affordance (a small button + ▾ at the cell's right edge) on every cell of
  its range, in PDF and HTML; the constraint, its formulas and the input/error
  prompts round-trip through `convert('xlsx')`.
- **Slicers** (`xl/slicers` + `xl/slicerCaches`) — a slicer panel renders as a
  captioned button box after the grid (the way chart frames do). A native-table
  slicer fills its buttons from the referenced table column's distinct values and
  highlights the items the column's autofilter keeps; an OLAP/pivot slicer whose
  items live in a pivot cache degrades to a caption-only box.
- Charts, **pictures and shapes** anchored to the sheet (the worksheet drawing part)
  stand where they are anchored, over the cells — a picture keeps its bytes; a shape
  its preset/custom geometry, fill, outline, text body and `a:xfrm` rotation/flips
  (reusing the DrawingML shape readers). All three anchor kinds place a drawing:
  `twoCellAnchor`, `oneCellAnchor` and `absoluteAnchor`, counted in the grid's own
  column unit, a hidden row or column taking no room; a group (`xdr:grpSp`) maps its
  children through the group's own coordinate space. On paper each page shows the
  part of every drawing over its own rows and columns, cut off where they end — a
  chart across a page break prints in two pieces, a drawing in the print titles
  repeats, nothing outside the print area prints, and the rows a drawing reaches
  print. A drawing marked `fPrintsWithSheet="0"` shows on a screen only.
- **SmartArt** — the pre-rendered drawing (`diagrams/drawing#.xml`, the `dsp:`
  namespace) where the file caches one, and otherwise the diagram laid out from
  its own `data`/`layout`/`colors`/`quickStyle` parts; either way each label
  lands in the rectangle the diagram sets aside for it.
- **Cell hyperlinks** (`<hyperlinks>`) — an external `r:id` resolves to a URL and the
  covered cell becomes a clickable link (PDF `/Link` annotation, HTML `<a>`).
- **Header/footer text** (`<headerFooter>`) — Excel's `&`-code mini-language (`&L`/`&C`/
  `&R` regions, `&P`/`&N` page-number fields resolved per page, `&A` sheet name,
  `&B`/`&I` bold/italic) renders in the page margins.
- **Cell formatting details** — **in-cell rich text** (a shared string built from
  several `<r>` runs renders one document-model run per run, each with its own
  bold / italic / underline / colour / size / super- or sub-script); **wrapped
  text** (`wrapText` cells keep their full text and wrap to the cell, growing the
  row); **indent** (`indent` — three spaces of the Normal font a level, from the side
  the cell is aligned to); **non-solid pattern fills** (gray / hatch
  patterns blend foreground over background to a representative solid) and
  **gradient fills** (summarised to the mean of their stops); **diagonal cell
  borders** (up / down strokes across the cell); **text rotation** (`textRotation`
  — rotated / vertical cells render their text stacked top-to-bottom); and
  **shrink-to-fit** (`shrinkToFit` scales the cell's font down to its column width).
- **Cell comments / notes** — legacy notes (`xl/comments`) and modern threaded
  comments (`xl/threadedComments`, authors resolved through `xl/persons`) are read.
  A note the sheet shows (its VML shape visible) is drawn where Excel draws it: in its
  `x:Anchor` box, the comment's runs in their own faces, VML's hard shadow and a line
  to the cell's corner. A window always shows them; paper follows the page setup's
  `cellComments` — nothing by default (reported as a loss), the shown notes in place
  under `asDisplayed`, and under `atEnd` a "Comments" section after the grid, one
  line per comment: `<cell> — <author>: <text>`.
- **Form controls** — checkboxes, option buttons, spinners, scroll bars, list /
  drop-downs and buttons (the worksheet's `<controls>`, each resolved to its
  `ctrlProp` part for type + state) are drawn in place, where their anchor or their
  VML shape puts them — a shape inside a `<v:group>` through the group's own
  coordinate space, a box given in points alone on the rows Excel counted it by.
  One the file gives no geometry is listed in a "Form controls" section after the
  grid, with a type-appropriate affordance and its state (`[x]` / `[ ]` for a
  checked box, `(o)` for an option button, the value for a spinner).
- **ActiveX controls** — the embedded OLE controls (`<oleObjects>` → `xl/activeX`)
  are drawn in place from the VML shape that shares their id, and otherwise listed in
  an "ActiveX controls" section the same way: the `progId` gives the
  control type and the `<ax:ocxPr>` property bag its visible state (caption,
  checked/value, group). A control persisted only to its binary `.bin`
  (MS-OFORMS) renders as its type without the caption — reading that property bag
  from the OLE/CFB stream is the remaining piece.

**PresentationML (§19)**
- Each slide is a page at the deck size (`p:sldSz`); shapes are floating content
  positioned from their `a:xfrm`, turned and mirrored by its `@rot`/`@flip`.
- Text boxes (`p:sp`) — runs (size, bold/italic/underline stated either way,
  colour, latin font, `cap`), paragraph alignment, the body vertical anchor,
  bullets and per-level indents.
- **The style a deck states once** — the presentation's `p:defaultTextStyle`, the
  master's three text families (`p:txStyles`) and each prototype's own
  `a:lstStyle`, in that order, under the slide's own runs. Placeholders inherit
  their geometry, fill and outline from the layout → master prototype as well.
- **Bullets** — a literal `a:buChar` (read out of the symbol face `a:buFont`
  names, so a Wingdings `l` is the circle it draws), an auto-numbered
  `a:buAutoNum`, or none; with the colour and size the level gives them
  (`a:buClr`, `a:buSzPct`/`a:buSzPts`) and a tab out to the hanging indent.
- **The deck's decoration** — the shapes a master and its layouts carry are drawn
  under the slide's own content, unless the slide or the layout turns them off
  (`@showMasterSp`).
- **Backgrounds** (`p:bg`) — a solid fill, a gradient, a picture, or a theme
  reference (`p:bgRef` into `a:bgFillStyleLst`); a shape marked `useBgFill` is
  painted with the piece of the background that lies under it.
- Pictures (`p:pic`) with the recolouring the file asks for — `a:duotone`,
  `a:clrChange`, `a:alphaModFix` — shapes with geometry/fill/stroke/gradient,
  and embedded charts (`c:chart`), including a chart papered with a picture.
- **WordArt** (`a:prstTxWarp`) — the text bent through the preset curve its body
  names and stretched onto the shape's box: thirty envelope presets and the two
  rings, each glyph placed on its own along the curve.
- **Tables** (`a:tbl`) — the style the table names in `tableStyles.xml` with the
  parts it switches on (header row, banding, first/last column), composed under
  the cell's own `a:tcPr` fill, `a:noFill` and four rules; row heights from
  `a:tr@h`.
- **Embedded objects** (`p:oleObj`) — the snapshot the producer wrote beside the
  object, in either the modern (`p:pic`) or the legacy (VML drawing) spelling.
- **SmartArt** — the pre-rendered drawing (`dsp:spTree`) where the deck caches
  one, and otherwise the diagram laid out from its own parts: eleven algorithms
  (`lin`, `composite`, `cycle`, `hierChild`/`hierRoot`, `snake`, `tx`, `sp`,
  `conn` among them), the `forEach`/`choose` axis walks the layouts are written
  in, the constraint solver that sizes and spaces the boxes, and the colour and
  style parts that paint them.
- **Theme** colours (`a:clrScheme`) through the deck's own colour map
  (`p:clrMap`, and a layout's or slide's override), and groups (`p:grpSp`)
  mapped through their child transform.
- Run hyperlinks (`a:hlinkClick`) → clickable PDF annotations / HTML `<a>`.
- **Legacy `.ppt`** (PowerPoint 97–2003) — the binary `PowerPoint Document` stream
  inside the OLE2/CFB container, reached through the Current User → UserEditAtom →
  PersistDirectoryAtom indirection. Each slide becomes one page at the deck size
  (the DocumentAtom slide size, in master units).

  A slide's **title and body come from the document's own slide list**, not from the
  shape that draws them: the placeholder holds an `OutlineTextRefAtom` naming which
  of the slide's texts is its, so each lands in the rectangle the file states for it.
  Their **size, colour, typeface and bullet come from the master** — the
  `TextMasterStyleAtom` for that text type and indent level — because a run usually
  states none of them; a run that does states only where it differs. A colour may be
  a literal sRGB or a slot in the slide's colour scheme, and the typeface is an index
  into the deck's font collection, so a deck set in Times is measured in Times. A
  bullet is a character, translated out of the symbol face that states it, and it
  **hangs** on the margins the paragraph or its master states — the body indent, the
  first-line indent and the space set around each paragraph. The three furniture
  placeholders — **date, footer and slide number** — take their text from the deck's
  own header/footer record rather than from the shape, which holds none of it.

  The slide is drawn on **the background it states** — a solid colour, a shade, or a
  picture — or on its master's, when its flags say to follow it, and under **the
  decoration the master carries**: the rules, the logo, the footer band, with the
  master's own placeholders left behind as the prototypes they are.

  Shapes are placed by their slide anchor; a **grouped** shape by the ChildAnchor its
  group resolves, which is what makes a `.ppt` table — a group of cell rectangles —
  read as one. Each carries its fill: a literal or scheme colour, a shade, or a
  **picture, stretched or tiled**, including a picture stored inline on the shape
  rather than in the deck's picture stream. A freeform carries its **exact custom
  geometry** (the `pVertices` / `pSegmentInfo` arrays), and a piece of **WordArt** is
  drawn as the word it is rather than as the rectangle its shape type would give.
  A picture shows the part of itself the shape **crops** to, and drops the colour
  the shape names as **transparent** — how clip art older than the alpha channel
  says a rectangle of ground is not part of the drawing. A **metafile** picture
  (EMF / WMF, deflated inside its Escher blip) is replayed like any other, and a
  headerless **DIB** blip is read as the bitmap it is. Only a palette-relative
  colour degrades gracefully.

**Graphics & math**
- DrawingML shapes (preset and custom geometry, gradients, group shapes, theme colors).
- Charts — bar/column, line, pie/doughnut, area, scatter, stacked. A value axis Excel
  is left to scale takes Excel's ends and step; the category axis lies where it
  crosses (`crosses`, `crossesAt`) with its labels where `tickLblPos` puts them, a
  reversed one keeps its values at its start, and a crowded one is thinned by the
  labels it draws; categories on several levels (`c:multiLvlStrCache`) are labelled
  in rows; a long title wraps; system colours (`sysClr`) are read and transformed.
- OfficeMath — fractions, scripts, radicals, n-ary operators, functions, limits,
  delimiters, matrices, accents; inline and display.

**Typography**
- Type0 + CIDFontType2 embedding with subsetting.
- Knuth–Plass line breaking, Liang hyphenation (en / ru).
- OpenType ligatures and kerning (GSUB/GPOS), mark positioning.
- BiDi (UAX #9), Arabic cursive joining.
- **A face for the writing system the document is in.** The curated substitutes are
  Latin, so text in Han, Kana, Hangul, Arabic, Hebrew, Thai or the geometric symbol
  block reaches a Noto face for that script instead of a row of notdef boxes —
  chosen per CHARACTER, fetched only for the scripts a document actually holds, and
  in one weight rather than four. Unified Han is the document's call: the same
  character is drawn differently in Japanese, Korean and Chinese, so its neighbours
  decide which face draws it.
- **The font a document brings with it** is used ahead of any substitute — a
  `.docx` font-table part (de-obfuscated against the GUID in its name) or a `.pptx`
  `p:embeddedFontLst` face. A `w:rFonts` that names a **theme** slot rather than a
  family (`asciiTheme="minorHAnsi"`) resolves through the theme's major / minor
  fonts, and a **narrow** family is measured at the narrow advance widths rather
  than the regular ones.
- **Renderer-compatibility `layoutProfile`** (`'word'` / `'libreoffice'`) — matches how a
  target renderer measures text, breaks it into lines and stacks the lines; with the
  metric-compatible open substitutes (Carlito / Caladea / Arimo / Tinos / Cousine) this
  tracks the target closely **without its private font metrics**. Under the default
  profile a document Word sets — a `.docx` or a `.doc` — stands its lines as Word
  does, from a table of the line Word gives each family it is known for; `'word'`
  stands every document's lines that way, and `'libreoffice'` takes each line from
  the face in hand.

**PDF / compliance**
- PDF/A-1, -2, -3 at levels a / b / u — all formally **veraPDF-validated**.
- **PDF/UA-1** (ISO 14289-1) — veraPDF-validated, alone or combined with PDF/A-2a
  in a single file.
- Tagged PDF — logical structure tree, headings, tables, lists (with `Lbl`
  markers), figures with alt text, links with alternate descriptions, footnote
  `Note` elements, `/Lang`, pagination artifacts.
- **Encryption** — AES-256 (ISO 32000-2 R6) via WebCrypto, with permission flags.
- Digital signatures — PKCS#7 detached, ECDSA; object streams; JPEG2000 images.

## Not yet

- **Byte-for-byte visual reproduction of another renderer.** `layoutProfile` plus the
  metric-compatible substitutes get a target tool's page geometry close — without its
  private font metrics — but _pixel-identical_ output is a non-goal: that would need the
  exact same font file and the renderer's internal glyph rounding.
- **A few format edge cases degrade gracefully** — re-save the original in its modern
  format for byte-exact fidelity, but each is a _missing detail, never a wrong one_:
  a legacy `.ppt` shape's
  **palette-relative colour** (literal, scheme- and system-relative colours resolve),
  a rare **arc / ellipse freeform segment** (it falls back to the path's preset
  bounds) or a **texture drawn through WordArt's glyphs** (the word, its size, its
  face and the preset curve it is bent through are all drawn); in a legacy `.xls`, a 2007 **Excel table's** banded style / autofilter
  (its shared-feature record is one even Apache POI leaves unparsed, so the cell
  values still render — only the table's banding is absent) and a colour-scale /
  data-bar / icon-set rule (a 2007 `CF12` record) whose colour is **theme-relative**
  rather than a literal value (the rules themselves, with literal or palette colours,
  are read); and an ActiveX **CommandButton / Label** whose caption is persisted
  only to a binary `.bin` (the MorphData control family — check box / option / toggle
  / text / combo / list — and every property-bag control _are_ read, with their
  caption and value).
- **Inside a metafile**, a gradient fill and a flood fill are not drawn;
  everything else the file draws is — the arcs, pies and chords, the clipping
  region it limits its records to, and the rasters it blits. Each omission is
  reported as a named loss rather than guessed at.
- **A `.pptx` embedded font packed with MicroType Express.** PowerPoint wraps an
  embedded face in an EOT container and, unlike Word's, it is compressed with a
  codec of its own; the deck renders in a substitute and says so rather than
  silently dropping the family.
- **Where a section's lines run down the sheet, a table and an inline picture
  turn with them**, where Word sets a table flat in the corner of the sheet and
  keeps an inline picture upright in its line. A PDF page whose words run UP its
  sheet, or stand on their heads, has no section to be set in that way — Word and
  LibreOffice lay a section's text across the sheet or down it — and is set
  across it, with a loss that says so.
- **A book's running foot names no chapter.** Read from a PDF, each side's foot
  says what most of that side's pages say and the page's number; the chapter a
  page is in, named differently in every chapter, is left out rather than set
  once and printed on every page. LibreOffice also seats a book's left and right
  pages by the number its first page carries, and may sign a page with the other
  side's foot where Word, which takes each page's own number, does not.
## Validation

Development is corpus-driven: documents are converted, compared against a LibreOffice
"golden" render (structural text diff + rasterized visual diff), and PDF/A output is
gated through veraPDF. A PDF read back as `.docx` is set beside its source page on a
bench of some four hundred files, and the documents a change is for are opened in
Word and in LibreOffice and counted page for page against the source; a rule of
Word's layout is measured in Word before it is written. Untrusted corpus files run
inside a locked-down Docker sandbox.
