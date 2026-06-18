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
boxes (run formatting, alignment, vertical anchor, bullets, indents),
layout/master placeholders, pictures, shapes (geometry/fill/stroke/gradient),
DrawingML tables, embedded charts, theme colours, slide/master backgrounds,
grouped shapes and run hyperlinks; not read (each a graceful loss): text autofit
shrink, picture backgrounds, picture placeholders, alpha/roman list numbering.
PDF input handles classic and modern compressed files
(cross-reference streams, object streams) and encrypted files (RC4 / AES; the
user password is passed to `Ream.parse`, defaulting to the empty permissions-only
case). A **tagged** PDF (including the ones Ream writes) is rebuilt from
its structure tree — headings, paragraphs, tables, list items, reading order; an
**untagged** PDF is reconstructed heuristically from glyph positions (lines by
baseline, paragraphs by spacing, headings by relative font size, and a clean
two-column page split at its central gutter), which is approximate. Text comes back via each font's `/ToUnicode` map; **raster images,
hyperlinks and vector shapes** are lifted back out too (JPEG verbatim, other
images re-encoded as PNG with soft-mask alpha, `/Link` URIs re-attached to the
text, filled paths, stroked lines and shading-pattern gradients turned into
shapes). Clipping paths and clip-bounded (`sh`) shadings are not read.

**Output** — `convert('pdf')`, `convert('svg')` (a page-stack preview),
`convert('html')` (flowed, needs no fonts), `convert('docx')` (write
WordprocessingML back out) and `convert('xlsx')` (write SpreadsheetML back out —
spreadsheet input only). The writers are for normalization, sanitization,
in-browser editing, and round-tripping. The docx round-trip is semantic, not
byte-exact, but complete — text, tables, images, lists, links, headers/footers,
multi-section geometry, footnotes/endnotes, charts and OfficeMath all write
back. The xlsx round-trip preserves the whole grid surface — cells, styles,
merges, the print model, conditional formatting, sparklines, tables and embedded
charts — and is byte-stable across a read↔write loop.

**WordprocessingML (§17)**
- Text, runs and the full style cascade (`docDefaults` → styles → direct formatting).
- Tables — auto / fixed layout, §17.4 border-conflict resolution, cell shading,
  vertical merge and grid span, nested tables, **table styles** (`w:tblStyle` with
  conditional formats: banding, first/last row/column).
- Lists and numbering (`abstractNum`, level overrides), multi-level.
- Sections — per-section page size and orientation, headers and footers,
  **multi-column layout** (`w:cols`).
- **Hyperlinks** — external (clickable PDF annotations + HTML `<a>`, scheme-allowlisted)
  and internal: bookmarks become named destinations / `#`-anchors.
- **Fields** — `PAGE` / `NUMPAGES` render real page numbers in headers and footers.
- **Footnotes and endnotes** — notes at the bottom of the referencing page behind
  Word's separator rule; endnotes after the body.
- **Review comments** (`w:commentReference`) — a bracketed superscript marker in the
  text and a "Comments" section after the body with each comment's author, date and
  content (PDF and HTML). Reply threads and resolved state come from
  `commentsExtended.xml` (replies nest under their parent, resolved threads are
  flagged); the commented range (`w:commentRangeStart/End`) is highlighted; author
  identities resolve from `people.xml`. As an opt-in, comments can also be emitted as
  native PDF sticky-note annotations (`commentAnnotations`, interactive output only).
- **SmartArt** — rendered from the diagram's pre-rendered DrawingML drawing
  (`diagrams/drawing#.xml`) as positioned shapes; a file with no drawing fallback
  degrades to a graceful loss rather than an empty space.
- Inline and floating images (PNG / JPEG / JPEG2000), including **legacy VML
  pictures** (`<w:pict>` / `<w:object>` — ActiveX and OLE-object previews,
  images from older Word); floating drawings (`wp:anchor`) render outside the
  text flow — wrap-none (incl. `behindDoc`) for watermarks/stamps/text boxes,
  and side wrapping (`square`/`tight`/`through`) where the body text flows
  around the exclusion area.
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
  cell mark) become a row-and-cell grid, with per-column widths **and per-cell
  borders and vertical merges** from the table definition (`sprmTDefTable`'s
  `TC80` array) — and **inline images** are extracted (the
  picture character's CHPX points at a PICF in the `Data` stream; the raster blip
  is pulled out and sized from the PICF). **Fields** resolve to their cached
  result — the field code (`PAGE`, `NUMPAGES`, `REF`, …) is dropped and the stored
  result text kept. The section's **headers and footers** are lifted from the
  `PlcfHdd` stories (best-effort: the binary story ordering can't be ground-truthed
  here, so only well-formed stories are surfaced). **List items** (`sprmPIlfo` /
  `sprmPIlvl`) render with their resolved **number format** — a real "1." / "a)" /
  "iii." or the bullet glyph, from the `LST` / `LVL` / `LFO` tables. So an old `.doc`
  renders to PDF/SVG/HTML and re-writes to `.docx`. Cell background shading is not
  read yet (re-save as `.docx` for full fidelity); an encrypted file yields no text.
  The shared CFB reader
  (`src/core/ole`) is the same keystone `.xls` uses.

**SpreadsheetML (§18)**
- Grids, shared strings, number formats and dates (incl. the 1904 date system).
- **Legacy `.xls`** (BIFF8, Excel 97–2003) — the binary `Workbook` stream inside the
  OLE2/CFB container is read into the same grid model, so an old `.xls` renders to
  PDF/SVG/HTML and even re-writes to `.xlsx`. Cell values, structure (sheets, shared
  strings, merges, column widths, the 1904 flag), **styling** — fonts, fills,
  borders, number formats and alignment from the FONT/FORMAT/XF records, with colours
  resolved through the BIFF colour palette — **embedded pictures** (from the
  Office-Drawing/Escher BLIP store), **embedded charts** (the BIFF chart substream,
  plotted from the worksheet cells its AI records reference) and **drawing shapes**
  (autoshapes + text boxes, from the Escher shape records and their TXO text) are
  read, plus **cell hyperlinks** (the HLINK record's URL moniker), the **page-setup
  print model** (orientation, scale, fit-to-page, margins, gridlines, centering,
  header/footer and manual page breaks), **defined names** (named ranges plus the
  print area and repeated titles, from the NAME records), **cell comments** (the
  Note record's author + the text-box text), **data validation** (the rule type,
  ranges and a `list` rule's in-cell dropdown) and **conditional formatting** (the
  classic `cellIs` / `expression` rules with their differential fill / font colour).
  Only the 2007 colour-scale / data-bar / icon-set extensions (the CF12 records) are
  not read yet — see Not yet.
- The print model — gridline suppression, print area, fit-to-page scaling, repeated
  print titles, manual page breaks, horizontal/vertical centering, and **column-band
  pagination**: a sheet wider than the page (and not fit-to-width) splits across
  pages, all rows of the left columns first, then the next band ("down, then over"),
  honouring manual column breaks — instead of being squeezed onto one page width.
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
  range's value extent. Also `expression` (an arbitrary formula, evaluated per
  cell by a small built-in formula engine against the workbook's cached values —
  no recalculation) and `timePeriod` (today / this-week / last-month … windows).
  Both stay deterministic: `timePeriod` and `TODAY()`/`NOW()` read an explicit
  reference date you pass as `now` (never the system clock), so without one those
  clock-relative rules simply don't paint. The highest-priority matching rule
  claims the cell's fill / font; a data bar or icon applies on top.
- **Sparklines** — per-cell line / column / win-loss mini charts, including
  cross-sheet data ranges and blank-cell gaps.
- **Excel tables** (`xl/tables`) — banded rows and a styled header row, the
  colours resolved from the named table style against the workbook theme.
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
  render after the grid — a picture keeps its bytes; a shape its preset/custom
  geometry, fill, outline and text body (reusing the DrawingML shape readers).
- **Cell hyperlinks** (`<hyperlinks>`) — an external `r:id` resolves to a URL and the
  covered cell becomes a clickable link (PDF `/Link` annotation, HTML `<a>`).
- **Header/footer text** (`<headerFooter>`) — Excel's `&`-code mini-language (`&L`/`&C`/
  `&R` regions, `&P`/`&N` page-number fields resolved per page, `&A` sheet name,
  `&B`/`&I` bold/italic) renders in the page margins.
- **Cell formatting details** — **in-cell rich text** (a shared string built from
  several `<r>` runs renders one document-model run per run, each with its own
  bold / italic / underline / colour / size / super- or sub-script); **wrapped
  text** (`wrapText` cells keep their full text and wrap to the cell, growing the
  row); **left indent** (`indent`); **non-solid pattern fills** (gray / hatch
  patterns blend foreground over background to a representative solid) and
  **gradient fills** (summarised to the mean of their stops); **diagonal cell
  borders** (up / down strokes across the cell); **text rotation** (`textRotation`
  — rotated / vertical cells render their text stacked top-to-bottom); and
  **shrink-to-fit** (`shrinkToFit` scales the cell's font down to its column width).
- **Cell comments / notes** — legacy notes (`xl/comments`) and modern threaded
  comments (`xl/threadedComments`, authors resolved through `xl/persons`) are read
  and listed in a "Comments" section after the grid — a heading then one line per
  comment, `<cell> — <author>: <text>` — mirroring Excel's "print comments at end of
  sheet". The legacy VML note box is ignored; only the text + author are surfaced.
- **Form controls** — checkboxes, option buttons, spinners, scroll bars, list /
  drop-downs and buttons (the worksheet's `<controls>`, each resolved to its
  `ctrlProp` part for type + state) are listed in a "Form controls" section after
  the grid, each with a type-appropriate affordance and its state (`[x]` / `[ ]`
  for a checked box, `(o)` for an option button, the value for a spinner). The
  control's anchored VML shape isn't drawn in place.
- **ActiveX controls** — the embedded OLE controls (`<oleObjects>` → `xl/activeX`)
  are listed in an "ActiveX controls" section the same way: the `progId` gives the
  control type and the `<ax:ocxPr>` property bag its visible state (caption,
  checked/value, group). A control persisted only to its binary `.bin`
  (MS-OFORMS) renders as its type without the caption — reading that property bag
  from the OLE/CFB stream is the remaining piece.

**PresentationML (§19)**
- Each slide is a page at the deck size (`p:sldSz`); shapes are floating content
  positioned from their `a:xfrm`.
- Text boxes (`p:sp`) — runs (size, bold/italic/underline, colour, latin font),
  paragraph alignment, the body vertical anchor, bullets (`a:buChar` and
  auto-numbered `a:buAutoNum`) and per-level indents.
- **Placeholders** — title/body/number shapes inherit geometry and per-level text
  styles from the slide layout → master (`p:txStyles`).
- Pictures (`p:pic`), shapes with geometry/fill/stroke/gradient, DrawingML tables
  (`a:tbl`) and embedded charts (`c:chart`).
- **SmartArt** — rendered from the diagram's pre-rendered DrawingML drawing
  (`dsp:spTree`) as positioned shapes; no drawing fallback ⇒ a graceful loss.
- **Theme** colours (`a:clrScheme`), slide/master backgrounds (`p:bg`) painted
  behind the content, and groups (`p:grpSp`) mapped through their child transform.
- Run hyperlinks (`a:hlinkClick`) → clickable PDF annotations / HTML `<a>`.
- **Legacy `.ppt`** (PowerPoint 97–2003) — the binary `PowerPoint Document` stream
  inside the OLE2/CFB container, reached through the Current User → UserEditAtom →
  PersistDirectoryAtom indirection. Each slide becomes one page at the deck size
  (the DocumentAtom slide size, in master units); the text is read from the
  TextChars / TextBytes atoms with run formatting (bold / italic / underline / size
  / colour from the StyleTextPropAtom) and paragraph alignment / indent level, and
  embedded images are pulled from the Pictures stream (OfficeArtBlip referenced by a
  shape's `pib`). A shape that carries a slide anchor (OfficeArtClientAnchor) is
  positioned at its rectangle — text boxes and pictures become floating content,
  like the `.pptx` reader; an un-anchored shape (e.g. a placeholder that inherits
  master geometry) flows in reading order. Decorative autoshapes are read as vector
  shapes — the preset type from the OfficeArtFSP (or, for a freeform, its **exact
  custom geometry** walked from the `pVertices` / `pSegmentInfo` arrays) plus their
  fill / line colour, whether a literal sRGB value or one resolved through the
  slide's colour scheme (the master's when the slide follows it); only palette /
  system colours are not — see Not yet.

**Graphics & math**
- DrawingML shapes (preset and custom geometry, gradients, group shapes, theme colors).
- Charts — bar/column, line, pie/doughnut, area, scatter, stacked.
- OfficeMath — fractions, scripts, radicals, n-ary operators, functions, limits,
  delimiters, matrices, accents; inline and display.

**Typography**
- Type0 + CIDFontType2 embedding with subsetting.
- Knuth–Plass line breaking, Liang hyphenation (en / ru).
- OpenType ligatures and kerning (GSUB/GPOS), mark positioning.
- BiDi (UAX #9), Arabic cursive joining.
- **Renderer-compatibility `layoutProfile`** (`'word'` / `'libreoffice'`) — matches a
  target renderer's line-height model, line breaking and default kerning; with the
  metric-compatible open substitutes (Carlito / Caladea / Arimo / Tinos / Cousine) this
  tracks the target closely **without its private font metrics**.

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

- **The legacy `.ppt` reader does not yet read** (re-save as `.pptx` for these): a
  shape's **palette / system colours** (a shape's literal sRGB and colour-scheme
  fill / line _are_ resolved, but a palette- or system-relative colour is dropped),
  and the rare arc / ellipse freeform segment (a path using one falls back to its
  preset bounds rather than risk a mis-aligned curve). The slide text with run and
  paragraph formatting, embedded images, per-shape placement and autoshapes — preset
  geometry _and_ exact freeform geometry, with literal or scheme-resolved fill /
  line — _are_ read (see PresentationML above). All three legacy binary formats
  (`.doc` / `.xls` / `.ppt`) are read through the shared CFB container reader
  (`src/core/ole`).
- **The legacy `.doc` reader does not yet read** (re-save as `.docx` for these):
  table cell **background shading**. Everything else — text, run and paragraph
  formatting, tables with column widths, cell borders and vertical merges, list
  items (with their number format / bullet), inline images, fields, and the
  section's headers/footers — is read (see WordprocessingML above).
- **The legacy `.xls` reader does not yet read** (re-save as `.xlsx` for these): the
  **2007 conditional-format extensions** — colour scales, data bars and icon sets,
  which live in the separate `CF12` records and only exist in a `.xls` re-saved by
  Excel 2007+ (the 97–2003 UI had no such rules). Everything else _is_ read: the cell
  data, styling, embedded images, charts, drawing shapes, cell hyperlinks, the
  page-setup print model, defined names, cell comments, data validation and the
  classic `cellIs` / `expression` conditional-format rules (see SpreadsheetML above).
- **Byte-for-byte visual reproduction of another renderer.** `layoutProfile` plus the
  metric-compatible substitutes get a target tool's page geometry close — without its
  private font metrics — but _pixel-identical_ output is a non-goal: that would need the
  exact same font file and the renderer's internal glyph rounding.
- **Some Excel constructs are not rendered yet:**
  - **ActiveX control binary state** — an ActiveX control persisted only to its
    `.bin` (MS-OFORMS, not the `<ax:ocxPr>` property bag) renders as its control
    type without the caption/value. (Property-bag-persisted controls *are* listed
    with their visible state, above.)
  - The `expression` formula engine covers the functions conditional formats
    commonly use (logic, math, text, the date functions, `COUNTIF`/`SUMIF`); a
    formula that reaches outside it — a sheet-qualified reference, a defined name,
    or a function not in the library — evaluates to an error and the rule simply
    doesn't apply (a graceful loss, never a misrender). Rules are evaluated only
    on cells that carry a stored value.

## Validation

Development is corpus-driven: documents are converted, compared against a LibreOffice
"golden" render (structural text diff + rasterized visual diff), and PDF/A output is
gated through veraPDF. Untrusted corpus files run inside a locked-down Docker sandbox.
