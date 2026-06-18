# Changelog

All notable changes to **Ream** (`reamkit`) are documented here. The project
follows [Semantic Versioning](https://semver.org/).

## 1.15.1

Read-fidelity fixes across the binary and PDF readers, found by a cross-format
validation sweep against real third-party documents.

### Fixed

- **Legacy `.doc` / `.xls` / `.ppt` (CFB container).** Ignore the non-zero garbage
  some Office writers leave in the reserved high 4 bytes of a v3 stream-size field
  (MS-CFB §2.6.1) — affected files were wrongly rejected as exceeding the size
  limit. And resolve streams from the main document's storage, so an embedded OLE
  object's same-named `WordDocument` / `1Table` / `Workbook` no longer shadows the
  real one (a `.doc` that embeds an object could otherwise parse to nothing).
- **Read PDF.** Preserve the source page size: a reconstructed PDF re-renders at
  its real MediaBox size and orientation instead of a fixed A4 — an A3 page no
  longer splits across several A4 pages, and landscape / custom page sizes are
  kept.
- **Read `.pptx`.** Omit hidden slides (`p:sld@show="0"`) from the rendered deck,
  matching PowerPoint and LibreOffice; the omission is reported as a loss.
- **Read legacy `.doc`.** Read the section page size from the SEP (sprmSXaPage /
  sprmSYaPage), so an A4 or landscape document is no longer forced to US Letter.

## 1.15.0

Three new input formats — the legacy binary `.doc`, `.xls` and `.ppt` (Office
97–2003) — and the conditional-format expression formula engine. `Ream.parse` now
sniffs and reads **seven formats** in total.

### Added

- **Read legacy `.doc` (Word 97–2003).** The binary WordprocessingML format
  (OLE2/CFB) parses through a shared container reader into the same interlayer as
  `.docx`. Reads the document **text** (the piece table / CLX — 16-bit Unicode and
  8-bit Windows-1252 pieces), **run formatting** (bold / italic / underline / size
  from the CHPX), **paragraph formatting** (alignment, indents, spacing from the
  PAPX), **tables** (cells from the `0x07` mark, per-column widths, cell borders,
  vertical merges and background shading), **inline images** (the picture
  character's PICF in the `Data` stream), **fields** (resolved to their cached
  result), the section's **headers and footers** (the PlcfHdd stories) and **list
  items** (numbered or bulleted, in their resolved number format). Converts onward
  to PDF / SVG / HTML / DOCX / XLSX like any source.
- **Read legacy `.xls` (Excel 97–2003 / BIFF8).** Reads the grid (NUMBER / RK /
  MULRK / LABELSST / BOOLERR / FORMULA records) with full **styling** (fonts,
  fills, borders, number formats and the colour palette from the XF table),
  **embedded images** and **charts** (the Escher BLIP store; the BIFF chart
  substream), **drawing shapes and text boxes**, **cell hyperlinks** (HLINK), the
  **page-setup print model**, **defined names** (named ranges, print area, repeated
  titles), **cell comments**, **data validation**, **frozen panes**, **custom row
  heights** and **conditional formatting** — the classic `cellIs` / `expression`
  rules and the 2007 **colour-scale / data-bar / icon-set** extensions (CF12).
- **Read legacy `.ppt` (PowerPoint 97–2003).** Each slide becomes a page at the
  deck size. Reads the slide **text** with **run and paragraph formatting** (the
  StyleTextPropAtom), **embedded images** (the Pictures stream), **per-shape
  placement** (anchored text boxes and pictures at their slide rectangles) and
  **decorative autoshapes** — preset or exact freeform geometry, with fill / line
  colours resolved through the slide's colour scheme (literal, scheme- and
  system-relative).
- **Conditional-format expression formula engine (XLSX).** `<cfRule
  type="expression">` and `type="timePeriod">` now evaluate — closing the
  documented graceful loss from 1.14.0 — with a deterministic, no-recalculation
  engine over the workbook's cached values: ~140 functions (logic / info incl. the
  `IS*` family and `IFS` / `SWITCH` / `XOR`; math, trig and exponential; the `SUM` /
  `COUNT` / `MEDIAN` / `SUMPRODUCT` / `STDEV` / `VAR` / `PERCENTILE` aggregates and
  the `COUNTIF(S)` / `SUMIF(S)` / `AVERAGEIF(S)` predicates; text; date / time; the
  `MATCH` / `INDEX` / `VLOOKUP` / `HLOOKUP` lookups and `ROW` / `COLUMN`),
  sheet-qualified references (`Sheet2!A1`), defined names, inline array constants
  (`OR(A1={1,3,5})`) and the per-cell relative-reference shift. `timePeriod` and
  `TODAY()` / `NOW()` read an injected reference date (`options.now`), never the
  system clock. A construct beyond a deterministic per-cell predicate evaluates to
  an error and the rule simply does not paint — never a misrender.
- **ActiveX control visible state (XLSX).** An `<oleObject>` ActiveX form control
  renders its visible state — a check box / option button / toggle as checked or
  unchecked, a text / combo / list control with its value — resolved from the
  control's `ctrlProp` part and, for the MorphData control family, from the binary
  `.bin` stream (MS-OFORMS) when the caption / value is persisted only there.

### Changed

- The interlayer sniffs and reads **seven input formats** now (`.docx`, `.xlsx`,
  `.pptx`, `.pdf` and the legacy `.doc`, `.xls`, `.ppt`); the README and the docs
  site were rewritten to state the full read / write matrix.

## 1.14.0

### Added

- **Excel sheet pictures and shapes (XLSX).** Pictures (`xdr:pic`) and shapes /
  text boxes (`xdr:sp`) anchored to a worksheet's drawing render as blocks after
  the grid, anchor-ordered, the way charts already do. A picture keeps its bytes
  (re-encoded into the PDF image XObject / HTML data URI); a shape keeps its preset
  or custom geometry, fill, outline and text body, reusing the DrawingML shape
  readers (parsed a second time on the preserve-order tree, gated so chart- and
  picture-only drawings don't pay for it). Render-only.
- **Excel cell hyperlinks (XLSX).** A worksheet `<hyperlinks>` entry whose `r:id`
  resolves to an external URL turns every covered cell into a clickable link — a PDF
  `/Link` annotation and an HTML `<a>` (scheme-allowlisted). In-workbook
  (location-only) links carry no URL and are skipped. Render-only.
- **Excel header/footer text (XLSX).** `<headerFooter>` expands Excel's `&`-code
  mini-language into the page margins: `&L`/`&C`/`&R` regions, `&P`/`&N` page-number
  fields resolved per page, `&A` sheet name, `&B`/`&I` bold/italic; non-deterministic
  or unsupported codes (date/time/file/path, font/size/colour selections) are
  dropped. Each region becomes its own aligned paragraph.
- **Excel conditional-format rule types (XLSX).** Beyond `cellIs` / `colorScale` /
  `dataBar` / `iconSet`, the value- and text-driven families now resolve: `top10`
  (top/bottom N or N %), `aboveAverage` (mean, optionally shifted by N standard
  deviations), `duplicateValues` / `uniqueValues` (value frequency across the range —
  numbers by value, text case-insensitively) and the text tests (`containsText` /
  `notContainsText` / `beginsWith` / `endsWith`). They write back through
  `convert('xlsx')`. `expression` (needs a formula engine) and `timePeriod` (clock-
  relative — Ream's output is deterministic) stay a documented graceful loss.
- **Excel cell-format details (XLSX).** **In-cell rich text** — a shared string built
  from several `<r>` runs renders one run per `<r>`, each with its own bold / italic /
  underline / colour / size / super- or sub-script. **Wrapped text** (`wrapText`)
  keeps its full text and wraps to the cell, growing the row. **Non-solid fills**
  (gray / hatch patterns) blend foreground over background to a representative solid,
  and **gradient fills** are summarised to the mean of their stops. **Left indent**,
  **diagonal cell borders** (up / down strokes), **text rotation** (rotated / vertical
  cells render stacked top-to-bottom) and **shrink-to-fit** (the font scales down to
  the column width) all render; the alignment + border attributes round-trip.
- **Excel cell comments / notes (XLSX).** Legacy notes (`xl/comments`) and modern
  threaded comments (`xl/threadedComments`, authors resolved through `xl/persons`) are
  read and listed in a "Comments" section after the grid — a heading then one line per
  comment, `<cell> — <author>: <text>` — mirroring Excel's "print comments at end of
  sheet". The legacy VML note box is ignored; only the text + author surface.
  Render-only.
- **Excel form controls (XLSX).** Checkboxes, option buttons, spinners, scroll bars,
  list / drop-downs and buttons (the worksheet's `<controls>`, each resolved to its
  `ctrlProp` part for type and state) are listed in a "Form controls" section after
  the grid, each with a type-appropriate affordance and its state (`[x]` / `[ ]` for a
  checked box, `(o)` for an option button, the value for a spinner). ActiveX controls
  are OLE binaries and remain a graceful loss. Render-only.

## 1.13.0

### Added

- **Excel data validation (XLSX).** Worksheet `<dataValidations>` are read into the
  SpreadsheetML model. A `list` validation paints an in-cell dropdown affordance — a
  small button with a ▾ at the cell's right edge — on every cell of its range, in PDF
  (a gated shape pass reusing the conditional-format icon machinery) and HTML (a
  floated inline SVG). The constraint, its formulas and the input/error prompts write
  back through `convert('xlsx')`, so the SheetDoc stays a byte-stable round-trip
  fixpoint. `showDropDown` keeps ECMA's inverted sense ("1" hides the dropdown); x14
  cross-sheet list sources are a documented omission.
- **Excel slicers (XLSX).** Slicers (`xl/slicers` + `xl/slicerCaches`) are resolved in
  the reader and render as captioned button boxes after the grid, the way chart frames
  do. A native-table slicer fills its buttons from the referenced table column's
  distinct values and highlights the items the column's autofilter keeps; an OLAP/pivot
  slicer whose items live in a pivot cache degrades to a caption-only box. The panel
  reuses the existing table layout/emit path (a styled mini-table); style accents
  follow the table/pivot heuristic. Slicer parts are not written back (dropped on
  `convert('xlsx')`, like pivot tables).

## 1.12.0

### Added

- **SmartArt diagrams (DOCX + PPTX).** SmartArt renders from the diagram's
  pre-rendered DrawingML drawing (`diagrams/drawing#.xml`, `dsp:spTree`) as
  positioned shapes — reusing the existing DrawingML shape machinery rather than
  re-running Office's layout engine. Scheme colours resolve through the
  document/deck theme. A file that ships no drawing fallback degrades to a
  graceful loss (`shapes.smartArt`) instead of vanishing.
- **Word review comments (DOCX).** `w:commentReference` is read into the
  `FlowDoc.comments` map (author, date, initials and block content). PDF and
  HTML render a bracketed superscript marker in the text and a "Comments"
  section after the body; in PDF the marker is a clickable internal jump to its
  entry. Reply threads and resolved state are read from `commentsExtended.xml`
  (`w15:paraIdParent` / `w15:done`): HTML nests replies under their parent and
  flags resolved threads, and PDF indents replies and notes the parent. The
  commented range (`w:commentRangeStart/End`) is highlighted in HTML and PDF, and
  author identities resolve from `people.xml`. An opt-in `commentAnnotations`
  render option additionally emits each comment as a native PDF sticky-note
  annotation (interactive output only — suppressed under PDF/A and tagged output).
  Comments — threads and resolved flags included — write back through
  `convert('docx')`, surviving a read↔write round-trip.
- **Excel pivot tables (XLSX).** A pivot's cached output grid already rendered
  as data; on top of that Ream now applies the named pivot style
  (`pivotTableStyleInfo`) — banded rows and a styled header — and emphasises
  grand-total / subtotal rows and columns (parsed from `rowItems` / `colItems`).
  The pivot is not recomputed from its cache.

## 1.11.0

### Added

- **Renderer-compatibility `layoutProfile`.** `convert('pdf', { layoutProfile })`
  switches the line-height model, line breaking and default kerning to match a
  specific renderer, for closer visual parity:
  - `'libreoffice'` — line height from the font's hhea metrics; greedy
    (first-fit) line breaking.
  - `'word'` — line height from the OS/2 win metrics; greedy breaking; kerning
    off (Microsoft Word's default).
  - `'ream'` (the default) — Ream's own typesetter; output is unchanged.

  Validated against a LibreOffice golden render, `'libreoffice'` cuts the median
  baseline drift of flowing prose several-fold. The profile applies to DOCX/PPTX
  text; spreadsheet geometry follows the Excel row model regardless.

### Changed

- **Metric-compatible font substitutes.** The auto-substitution chain now maps
  each referenced family to an open font engineered to reproduce its advance
  widths, so text breaks into lines where the original would: Calibri → Carlito,
  Cambria → Caladea, Arial → Arimo (the sans default moves from Roboto to Arimo),
  alongside the existing Times New Roman → Tinos and Courier New → Cousine. These
  are the families LibreOffice substitutes, so a no-fonts conversion lands closer
  to the source layout.

## 1.10.0

### Added

- **PowerPoint (`.pptx`) input.** `Ream.parse` now reads PresentationML, so a
  deck converts onward to PDF, SVG, HTML or DOCX like any other source. Each
  slide becomes a page at the deck size; its shapes are read as positioned
  content:
  - **Text** — text boxes at their slide positions, with direct run formatting
    (size, bold/italic/underline, colour, font) and paragraph alignment, the
    vertical anchor, bullets (`•` and auto-numbered) and per-level indents.
  - **Placeholders** — title/body/number placeholders inherit their geometry and
    text styling from the slide layout and master (the PresentationML cascade).
  - **Pictures, shapes, tables and charts** — images, shapes with their geometry/
    fill/stroke/gradient, DrawingML tables and embedded charts all render.
  - **Theme, backgrounds and groups** — scheme colours resolve through the deck's
    theme; slide/master backgrounds paint behind the content; grouped shapes
    (`p:grpSp`) map through their child→slide transform.
  - **Hyperlinks** — a run's external link becomes a clickable PDF annotation /
    HTML `<a>`.

  Not yet read: text autofit shrink, picture/blip backgrounds, picture
  placeholders, alpha/roman list numbering, and SmartArt — each degrades
  gracefully rather than failing.

## 1.8.0

### Added

- **PDF form-XObject text.** Reading a PDF now recurses into the Form XObjects a
  page paints (a bare `/Name Do`), so text drawn inside a reusable form — which
  page-level interpretation missed — is recovered on both the tagged and
  heuristic paths.
- **Encrypted PDF with a user password.** A PDF locked with a real user password
  opens via `Ream.parse(bytes, { password })` (AES-256/R6 plus the legacy
  RC4/AES handlers); the empty-string default still opens the common
  permissions-only encryption.
- **PDF stroked vector graphics.** Lines, rules, dividers and shape outlines come
  back as line shapes carrying their stroke colour and width, alongside the
  filled paths already lifted.
- **LZW-encoded images.** Reading decodes `/LZWDecode` rasters — the TIFF/GIF-era
  codec, with `/EarlyChange` and a layered PNG/TIFF predictor — so legacy and
  scanned PDFs keep their pictures.
- **CCITT fax images.** Reading decodes `/CCITTFaxDecode` Group 4 (and Group 3
  one-dimensional) bilevel scans — the dominant encoding of fax-scanned PDFs —
  with a from-scratch ITU-T T.4 / T.6 codec.
- **Gradient fills are first-class.** A DrawingML `a:gradFill` parses into real
  colour stops and a direction (no longer averaged to a flat colour), renders
  faithfully to SVG, HTML and PDF (an axial/radial shading pattern), and
  round-trips through `convert('docx')`. Reading a PDF lifts a shading-pattern
  gradient back out into a gradient-filled shape.
- **Two-column PDF reconstruction.** An untagged two-column page is split at its
  central gutter and read column-by-column instead of interleaving the columns.
  The detection is conservative, so single-column and title pages are unaffected.

## 1.7.0

### Added

- **PDF raster images.** Reading a PDF now lifts its raster images back out —
  JPEG verbatim, everything else decoded and re-encoded as PNG (DeviceGray / RGB
  / CMYK, Indexed, ICCBased, with soft-mask transparency) — and places them in
  reading order, so `Ream.parse(pdf).convert('html' | 'docx')` carries the
  pictures instead of dropping them. A tagged `/Figure` keeps its alt text.
- **Compressed PDF input.** Reading handles modern PDFs whose cross-reference is
  a stream (`/Type /XRef`) and whose objects are packed into object streams
  (`/Type /ObjStm`) — previously those objects were unreachable, so
  heavily-compressed files lost most of their content.
- **Encrypted PDF input.** A PDF encrypted with the empty user password (the
  common permissions-only case) is read transparently — RC4, AES-128 and AES-256
  (R6) — with the cryptographic primitives implemented from scratch so the
  synchronous reader needs no asynchronous WebCrypto.
- **PDF hyperlinks** are recovered: a `/Link` annotation's URI is re-attached to
  the text beneath its rectangle, so a parsed PDF's links survive onward to the
  HTML `<a>` / docx hyperlink.
- **PDF filled vector graphics** are lifted out of untagged pages as shapes —
  filled rectangles and paths with their colour — interleaved with the text and
  images by position. Stroked / shaded art (lines, gradients, clips) is not read.
- **docx footnotes and endnotes** write back: a parsed note's reference and body
  are re-emitted, completing the note round-trip.
- **docx charts and OfficeMath** write back, so a `.docx` with an embedded chart
  or a mathematical equation round-trips through `convert('docx')` intact.
- **xlsx embedded charts** write back — the last piece of the spreadsheet grid
  surface that did not survive a read → write loop.
- **Excel fit-to-width pagination.** A sheet set to fit _N_ pages wide
  (`fitToWidth=N`) now scales its columns and paginates them across those pages,
  instead of being squeezed onto one page where it overflowed.

## 1.6.0

### Added

- **PDF input.** `Ream.parse` now also reads **PDF**, reconstructing a document
  tree from the page content. A tagged PDF — including the ones Ream writes — is
  rebuilt from its structure tree: headings, paragraphs, tables, list items and
  reading order. An untagged PDF is reconstructed heuristically from glyph
  positions (lines by baseline, paragraphs by spacing, headings by relative font
  size), which is approximate. PDF text is recovered through each font's
  `/ToUnicode` map. The result is an ordinary document tree, so a PDF converts
  onward like any other source — `Ream.parse(pdf).convert('html')` or
  `convert('docx')`. Images, vector graphics and encrypted PDFs are not read and
  are reported as losses. Ream is now a universal document engine: DOCX / XLSX /
  PDF in, PDF / SVG / HTML / DOCX / XLSX out.
- **Excel wide-sheet pagination.** A worksheet wider than the printable page now
  paginates across columns — all rows of the left columns first, then the next
  band ("down, then over") — honouring manual column breaks and repeating the
  print titles on every band, instead of being squeezed onto one page width.
- **Excel frozen panes** round-trip through the writer and become sticky header
  rows / columns in HTML output. They have no effect on PDF — in Excel, freezing
  is a view setting that does not print (the printed repeat is the print titles).
- **Conditional-format icon sets** — the symbols (check / exclamation / cross),
  ratings (a signal-bar meter) and quarters (a clock pie) families now draw
  faithfully instead of as plain circles.

## 1.5.0

### Added

- **Excel conditional formatting** — `cellIs` rules become per-cell highlights
  (dxf fill/font), `colorScale` a 2- or 3-stop gradient interpolated across the
  range's value extent, `dataBar` an in-cell bar (with a zero axis so a range
  spanning negatives draws them the other way, in red), and `iconSet` a per-cell
  glyph chosen by value bucket — traffic lights, arrows, flags, signs and the
  grey families. Rendered in PDF and HTML.
- **Sparklines** — the per-cell line / column / win-loss mini charts from the
  worksheet `extLst` render as vector graphics inside their host cell, including
  data ranges on another sheet and blank cells kept as gaps.
- **Excel tables** (`xl/tables`) — banded rows and a styled header row, with the
  header / band colours and white header text resolved from the named table
  style (`TableStyleMedium2`, …) against the workbook theme accents.
- **xlsx output (`convert('xlsx')`)** — Ream now writes SpreadsheetML as well as
  reading it. Unlike the docx writer it consumes the native grid tree, so the
  round-trip is lossless on the grid surface: cells, shared strings, the full
  style table, merges, the print model (margins, page setup, fit-to-page, print
  options, breaks), conditional formatting, sparklines and tables all survive a
  read → write → read loop byte-stably. Across the real-world xlsx corpus every
  readable workbook round-trips to a full grid-content identity; embedded charts
  are reported as losses, not yet written.

### Fixed

- Workbooks whose relationship parts (`.rels`) put the OPC namespace on a prefix
  (`<ns0:Relationship>`) instead of the default now read correctly — previously
  the relationships parsed to nothing, so such a file resolved to zero sheets.

## 1.4.0

### Added

- **docx output (`convert('docx')`)** — Ream now writes WordprocessingML as
  well as reading it. The parsed document re-serializes to a valid `.docx`:
  runs and paragraphs with their resolved formatting, page breaks, numbered
  lists, hyperlinks and bookmarks, tables (grid spans, borders, shading,
  nesting), images of every embedded format, DrawingML shapes (preset and
  custom geometry, fill, line, text), headers and footers, and multi-section
  page geometry. Use it to normalize, sanitize or edit a document in the
  browser and save it back, or for a docx → docx round-trip. The round-trip is
  semantic rather than byte-exact (direct formatting in place of named styles);
  across a 1100-document corpus every file re-writes without failure, 1099 of
  them to a full content identity. Footnotes, charts and OfficeMath are
  reported as losses, not yet written.
- **Legacy VML images now render** — the reader recovers pictures stored the
  old way (`<w:pict>` / `<w:object>` with `<v:imagedata>`): ActiveX and OLE
  object previews, and images from documents last saved by an older Word. They
  now appear in every output — PDF, SVG, HTML and docx.

## 1.3.0

### Added

- **Charts on xlsx sheets** — the worksheet's drawing part now loads:
  chart frames anchored to cell ranges render after their sheet's grid,
  sized from the anchor's column/row tracks. Custom chart colour themes
  (`colorsN.xml`) apply to series in both Word and Excel documents.
- **Float text wrapping** — side-wrapped anchored drawings (`wrapSquare`/
  `tight`/`through`) now claim an exclusion area: paragraph lines beside the
  float narrow to the wider side (Knuth-Plass re-breaks the paragraph with
  per-line widths) and resume full width below it.
- **Tagged lists: Lbl elements** — list-item markers ("1.", "•") get their
  own `Lbl` structure element, so assistive technology announces the label
  separately from the item body.
- **PDF encryption (AES-256)** — `encrypt: { userPassword, ownerPassword?,
  permissions? }` produces an ISO 32000-2 R6 encrypted PDF via WebCrypto
  (async conversion path only). PDF/A and encryption are mutually exclusive
  by standard; PDF/UA keeps the accessibility-extraction permission on.


### Changed

- **OOXML Strict (ISO 29500) packages** now load: relationship types are
  matched by name against both the Transitional (`schemas.openxmlformats.org`)
  and Strict (`purl.oclc.org`) namespaces.
- **Block-level content controls** (`w:sdt`) unwrap their content into the
  document flow instead of dropping it.
- **Password-protected (encrypted) OOXML** files now fail with a clear message
  identifying the file as an OLE compound file, instead of a cryptic ZIP error.

## 1.2.0

### Added

- **Bookmarks and internal links** — `w:bookmarkStart` + `w:hyperlink
  @anchor` become real GoTo links: PDF annotations with named destinations
  (`/Names /Dests`, only referenced names), tagged `Link` structure, HTML
  `id` anchors with `#`-fragment links.
- **PDF/UA-1** — `pdfUA: true` produces ISO 14289-1-conformant output
  (veraPDF-validated, alone and combined with PDF/A-2a in one file): tagged
  structure, `pdfuaid` XMP identification, alternate descriptions on link
  annotations, unique IDs on footnote Note elements, an always-present
  document title.
- **Multi-column sections** — `w:cols` lays content out column by column
  (equal-width with a shared gutter, or explicit per-column widths);
  headers, footers and footnotes keep the full page width.
- **Floating drawings** — `wp:anchor` placement: wrap-none drawings
  (watermarks, stamps, text boxes; including `behindDoc`) render at their
  anchored page/margin/paragraph-relative position without disturbing the
  text flow. Side-wrapping modes stay in flow (v1).
- **HTML writer: charts and shapes render as inline SVG** — bar/line/pie/
  area/scatter charts emit the same geometry scene as the PDF path (labels
  as native `<text>` with anchors, so the browser's fonts do the rendering);
  shape geometry (preset + custom, fills, strokes, dash patterns, rotation/
  flips) emits as `<path>` with the exact transform matrix the PDF layout
  computes, and text boxes overlay their content inside the body insets with
  the source vertical anchor. Charts and shapes are no longer reported as
  dropped losses for HTML output.
- **Footnotes and endnotes** — `w:footnoteReference`/`w:endnoteReference`
  render superscript numbers; footnote content lands at the bottom of the
  referencing page behind Word's short separator rule (the line and its note
  always travel together), endnotes flow after the body. Tagged PDFs wrap
  each note in a `Note` structure element (veraPDF-validated); the HTML
  writer renders anchored references with a notes section and backlinks.
- **PAGE / NUMPAGES fields** — page-number fields in headers and footers now
  render the real page number and total (both the `fldSimple` and the
  `fldChar` complex-field syntax). Bands containing fields re-lay out per
  page after pagination; other field instructions keep their cached result
  exactly as before, and documents without fields are byte-identical.
- **Table styles** — `w:tblStyle` referenced styles now render: the base
  layer (grid borders, default cell margins) plus `w:tblStylePr` conditional
  regions (first/last row and column, row/column banding, corner cells) gated
  by `w:tblLook` (modern attributes and the legacy bitmask). Resolved in the
  reader, so PDF, SVG and HTML all pick it up; tables without a style are
  byte-identical to before.
- **Hyperlinks** — `w:hyperlink` external targets now become clickable: PDF
  output gets `/Link` annotations (one rect per rendered line, merged into
  each page's `/Annots`; in tagged/PDF-A mode the annotation is enclosed in a
  `Link` structure element with `OBJR` + `/StructParent`, veraPDF-validated),
  and HTML output wraps the text in `<a href>`. Targets pass a scheme
  allowlist (`http`/`https`/`mailto`) — anything else renders as plain text
  with a degraded-`hyperlinks` loss; documents without links are
  byte-identical to before.

## 1.1.0

### Added

- **HTML writer** — `doc.convert('html')` renders the parsed document as a
  single self-contained flowed HTML file (headings, run styling, tables with
  spans/borders/shading, images as `data:` URIs, list markers, RTL). A flow
  medium needs no pagination and no fonts, so the conversion performs zero
  I/O; chart/shape geometry, inline math and headers/footers are reported in
  the loss report. Also exposed as `htmlWriter`/`writeHtml` and
  `createConverter` `to: 'html'`.

### Changed

- **Page model frozen** (`@experimental` API): `PageItem` page-frame
  coordinates are now **top-left / y-down** (CSS/SVG convention) and branded
  as `Pt`; the PDF emitter converts into PDF's y-up frame at emission. PDF
  output is byte-identical; SVG output changes coordinates only, not
  geometry.
- `LaidOutDocument` narrowed to the page model proper (`pages`, `resources`,
  `fontResources`, `imageResources`); the PDF-only state rides on
  `layoutStyledDocument(...).pdf`. The internal `DrawCommand` alias is gone —
  the schema name is `PageItem`.

## 1.0.0

The interlayer release — and the first stable major. Documents parse once
into a format-neutral tree (**FlowDoc**) and convert to any target from
there; the public face of the library is the `Ream` class.

### Added

- **`Ream`** — the object API: `Ream.parse(bytes)` (format sniffed) →
  `doc.convert('pdf' | 'svg', options)`, with `doc.flow` (the parsed tree),
  `doc.format`, `doc.losses`, and `doc.convertWithReport()` returning
  `{ bytes, losses }`. Conversion output is byte-identical to the per-format
  converters.
- **Intermediate representation (`@experimental`)** — `DocumentReader` /
  `DocumentWriter` interfaces, `docxReader`/`xlsxReader`, the
  `createConverter` facade, branded `Pt` units, a content-addressed
  `ResourceStore`, a `Feature` registry, and a loss protocol
  (`Loss`, `ConversionLossError`, `strict` mode).
- **SVG writer** — `doc.convert('svg')` renders the laid-out pages as a
  stacked-page SVG preview (no PDF involved); the third adapter, written
  purely against the page model.
- **Font provider chain** — `callerFontProvider` / `embeddedDocFontProvider` /
  `localFontProvider` (Chromium Local Font Access, with the OS/2 `fsType`
  licensing gate: embedding-restricted fonts are never used) /
  `remoteFontProvider`, composable via `fontProviders: [...]`; a remote or
  local winner is reported as a `substituted` loss.
- A cross-revision **byte gate** test suite: PDF output of fixed fixtures is
  snapshot-hashed, so pure refactorings are provably byte-identical.

### Changed

- `src/` reorganized into format modules: `core/` (format-agnostic), `word/`,
  `excel/`, `pdf/`, `svg/`. The public `reamkit/document-model` subpath is
  unchanged.
- The PDF renderer is split into layout and emit phases
  (`styled-page-emitter`); image embedding is split into pure
  `prepareImage` + `addImage` (PNG is no longer decoded twice per
  conversion).
- README and the docs site are rewritten around the `Ream` API, with an
  Examples page (PDF/A, signatures, providers, strict mode, recipes).

### Removed (breaking)

- The one-shot `convertDocxToPdf` / `convertXlsxToPdf` (+`Sync`) functions
  and their option types are no longer exported — use
  `Ream.parse(bytes).convert('pdf', options)`. Custom pipelines and
  bundle-size-sensitive consumers build on the `@experimental`
  reader/writer interfaces instead.

### Fixed

- `remoteFontProvider` lost the boldItalic→bold/italic degradation: a
  bold-italic run against a partial CDN set fell back to regular even when
  bold was available.
- Images in headers/footers resolved through the MAIN document's
  relationships (OPC ids are per-part): a colliding rId rendered the wrong
  picture, a non-colliding one dropped. Each header/footer part now resolves
  through its own `.rels`.
- `signPdf` located its placeholder by scanning for the first bare
  `/ByteRange` — an embedded attachment could hijack the signature bytes. The
  scan now matches the full fixed-width placeholder.

## 0.1.0-alpha.0

First tagged alpha. DOCX/XLSX → PDF, implemented directly from the ECMA-376 and
ISO 32000 / 19005 specifications — no third-party converter, PDF writer, or
layout engine. Browser-first; the caller supplies fonts.

- **WordprocessingML** — text, styles, tables (§17.4 border-conflict resolution,
  shading, vMerge/gridSpan, nested tables), lists, multi-section layout,
  headers/footers, inline/floating images, tracked changes.
- **SpreadsheetML** — grids, number formats, dates, and the print model
  (gridlines, print area, fit-to-page, repeated titles, page breaks).
- **Fonts/typography** — Type0 + CIDFontType2 embedding with subsetting,
  Knuth–Plass line breaking, Liang hyphenation, OpenType ligatures/kerning,
  BiDi (UAX #9), Arabic cursive joining.
- **Graphics** — DrawingML shapes, charts, and OMML math.
- **PDF** — PDF/A-1/2/3 (a/b/u), veraPDF-validated; digital signatures
  (PKCS#7 detached, ECDSA); object streams; JPEG2000.
- **Tooling** — aligned to `@tanstack/config` (Vite build, ESLint, publint +
  are-the-types-wrong); MIT-licensed; tag-triggered npm release workflow.
