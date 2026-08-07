# Changelog

All notable changes to **Ream** (`reamkit`) are documented here. The project
follows [Semantic Versioning](https://semver.org/).

## 1.24.0

A release about the diagrams. A SmartArt graphic keeps its own description —
the nodes, the layout that arranges them, the colours, the text style — beside
a drawing PowerPoint caches for readers that cannot run the layout. Ream read
the cache and dropped the rest, so a diagram saved without one was a hole in
the slide. The layout engine is now built: eleven algorithms, the colour and
style parts, and the text measurement that sizes a box to what it holds.

Measured against LibreOffice over the 2226 documents comparable both before and
after, the eleven corpora summed to 37.867 and now sum to 27.309. The pptx
regression corpus, which is where the diagrams are, went 14.688 to 6.324.

The rest is the gallery of preset shapes, WordArt's warps, and more of the
legacy work 1.23.0 began — including two metafile fixes every format shares.

### Added

- **A SmartArt diagram is laid out from its own parts.** `dgm:dataModel`,
  `dgm:layoutDef`, `dgm:colorsDef` and `dgm:styleDef` describe the graphic; a
  file that cached no drawing, or cached one a reader should not trust, used to
  render as nothing. Implemented: the `lin`, `composite`, `cycle`, `hierChild`
  and `hierRoot`, `snake`, `tx`, `sp` and `conn` algorithms, the `forEach` and
  `choose`/`if` axis walks the layouts are written in, the constraint solver
  that sizes and spaces the boxes, and the `presOf` paths that decide which of
  the model's points a shape shows. The colour list paints the boxes and their
  outlines, the style part their effects, and every run takes the point size
  its own level asks for.

- **The same diagrams in a `.docx`.** A Word document reaches its parts by the
  same relationships, and now lays them out the same way.

- **WordArt is bent through the curve its shape names.** §20.1.9.10
  `a:prstTxWarp` — thirty envelope presets and the two rings, each glyph placed
  on its own along the curve and stretched onto the shape's box.

- **Twelve more preset geometries**, the last of the gallery page that used to
  come out as plain rectangles, `halfFrame` and `bevel` among them.

- **Windows bitmaps.** BMP and the headerless DIB an Escher picture store keeps,
  in every depth from one bit to thirty-two, `BI_RLE4`/`BI_RLE8` and
  `BI_BITFIELDS` included — and the pattern fills that are made of them.

- **An EMF draws its pie, chord and arc records**, and honours the clip region
  a file states.

- **A `.ppt` draws the date, footer and slide number its deck states**
  (§2.4.15 `RT_HeadersFooters`), and keeps the space §2.9.32 sets around each
  paragraph.

- **A cell whose text is turned a quarter is drawn turned**, and a picture is
  clipped to the geometry it names for itself.

- **A slide's ActiveX controls are drawn** from the picture they cache, and
  text set to shrink is measured and shrunk to fit.

- **A drawing wider than its column band carries on across the next ones**,
  which is how a spreadsheet prints a picture that spans more than one page of
  columns.

### Fixed

- **A metafile keeps its lines thin, its figures whole, and its last blit
  blank.** Three faults in the EMF player, each of which every format shares. A
  blit with no source bitmap is the brush over a rectangle, but only for the
  raster operations that state a colour without reading the source or the
  destination — a picture ending in the plain `DSTCOPY` no-operation was
  painted white. A pen states its width in logical units, so the world
  transform scales it: a writer that shrinks the world by sixteen and asks for
  a sixteen-wide pen means one device unit, and at face value it drew every
  rule four points thick. And a `…To` record appended to an open figure was
  moving to the current point again, splitting one figure into three, which the
  even-odd rule then carved a wedge out of.

- **A WMF whose window runs upward is drawn the right way up.** §2.3.5.12 — a
  negative window extent turns that axis round, so the origin names the far
  corner. Read as the near one, a picture was laid out a full height above its
  box and hung off the top of the slide.

- **A `.ppt` paragraph keeps the style its own run states.** A
  StyleTextPropAtom's paragraph run counts CHARACTERS, not paragraphs, and one
  run routinely covers a whole placeholder. Read as one run per paragraph,
  every paragraph after the first fell through to the master's style — a body
  set flush left came out flush right because its master says so. The same
  exception's `leftMargin` and `indent` went unread, so no bullet ever hung.

- **A text box drawn on a slide is styled as one.** §2.13.33's text types skip
  3 — `other` is 4, not 3 — so a plain text box was taken for a body variant
  and inherited the outline's size, margins and spacing. The style it should
  take is stated once for the whole deck, in the document's own text info, and
  was never read at all.

- **A slide paints in the order its own shape tree gives**, which is what
  decides overlap; a pie labels its slices with what the chart asks to show;
  and a texture's grid of copies is centred on its shape.

- **A shadow falls under the picture its shape is filled with**, builds back up
  to the transparency a blur asks for, and keeps that alpha inside a picture
  too. A gradient stop away from the edge holds its colour there.

- **A run filled with a gradient takes its colour from one**, not from black; a
  gradient with no stops takes the gallery style's colours; a blip fill naming
  no mode papers the shape; and a shape-path gradient sweeps the way the shape
  runs.

- **A pattern fill that names no colour is the Automatic pair**, black on
  white, which §18.8.19 says an unstated pair means.

## 1.23.0

A release about the formats that came before the XML ones. `.ppt` and `.xls`
were read — every feature the specs list had been implemented — but no one had
ever compared one to a picture. Measured against LibreOffice, the twelve legacy
decks summed to 2.831 and the twelve legacy workbooks to 3.084; they now sum to
1.362 and 0.631. The deck that sat at a flat 1.000 sits at 0.076, and the one
whose slides overflowed onto eleven pages more than the reference prints now
paginates 30 for 30.

Two of the fixes are in the metafile reader, which every format shares.

### Added

- **A slide is drawn on the background it states.** A `.ppt` states its
  background as a shape like any other, marked `fBackground`: a solid colour, a
  shade, or a picture named by `fillBlip` — which is not the `pib` a picture
  SHAPE uses, and was the reason a photographic background read as nothing at
  all. A slide whose flags say `fMasterBackground` takes its master's, which is
  the only time the master's is consulted and the only version that is not a
  stale copy.

- **The decoration a master draws on the slides that follow it** — the rules,
  the logo, the footer band §2.4.24 `fMasterObjects` puts beneath what a slide
  draws itself. Its placeholders stay behind: they are prototypes, and drawn,
  every slide in a deck would carry "Click to edit Master title style".

- **A grouped shape lands where its group puts it.** Its rectangle is a
  ChildAnchor in the GROUP's coordinate space, resolved through the group's
  own FSPGR and client anchor. A `.ppt` table IS a group of cell rectangles, so
  this is what makes a table a table; the same correction lands in `.xls`,
  where forty of one sheet's forty-one shapes were grouped.

- **A picture fill, stretched or tiled**, including a blip carried INLINE as
  the property's own complex data — a deck with no picture stream at all, each
  table cell holding its own.

- **A metafile inside an Escher blip is unpacked** rather than walked past. A
  raster is its own bytes; a metafile sits behind an OfficeArtMetafileHeader
  and is deflated, so the signature scan missed every EMF and WMF in every
  `.ppt` and `.xls` there is. One slide whose whole content was one of them
  drew its shape's fill colour instead — a blue rectangle over a diagram.

- **A deck is set in the typefaces it names.** A run names its font by an index
  into the deck's font collection, which nothing read, so every `.ppt` rendered
  in whatever face the caller happened to hold. A substitute measures nothing
  like the Times a deck asks for, and one deck's body text ran off the bottom
  of every slide.

- **Bullets.** Stated as a character, almost always one from a symbol font, and
  translated out of it — untranslated, the circle every deck bullets with is
  the letter `l`. The paragraph decides whether it is bulleted and the master
  decides with what; where neither states the flag, a stated character is
  itself the statement, which is how a master names its outline levels.

- **WordArt is drawn as the word it is** — its text, size and face are shape
  properties, not a text box, and unread the shape was a coloured rectangle.
  The letters do not get their effects: no texture through the glyphs, no
  stretch to the frame.

- **A tiled picture repeats at the size the file asks for**: MS-ODRAW
  `fillWidth` / `fillHeight` in a `.ppt`, `a:tile @sx @sy` in DrawingML. Both
  default to the picture's own size, so a file that states neither renders
  exactly as before.

### Fixed

- **A slide's title and body come from where PowerPoint keeps them.** Not from
  the shape that draws them — from the document's own slide list, with the
  placeholder holding an OutlineTextRefAtom naming which text is its. Read the
  other way, a deck came out as one un-anchored heap of lines in the top-left
  corner, with the rectangles the file states going unused. The size was
  missing for the same reason: a run states none, because the master states it
  once per text type and indent level.

- **Text takes the colour the scheme names for it.** Half of what a themed deck
  states is a slot in its colour scheme, not an sRGB, and only the sRGB spelling
  was read — so a white-on-blue deck came out black on white. A master's slot
  resolves against the SLIDE's scheme, and a master lends its colour only to
  the text type that states it.

- **A `.xls` cell's fill is read from the bytes that hold it.** The pattern is
  the top six bits of the four bytes at offset 14 and the colours are the word
  at 18; both were read sixteen bits low, so every fill in every legacy
  workbook came back as black on black. The palette compounded it by returning
  eight hex digits where the style model holds six: silver reached the page as
  pink, and a red font as yellow.

- **A sheet's drawings are placed on the sheet's own grid.** They anchor to
  cells, and were sized against a made-up 48pt column and never positioned at
  all — on a sheet of 6.75pt columns each shape came out seven times too wide,
  and one engineering drawing flowed down eleven blank pages.

- **An object a WMF creates and the reader does not model still takes a handle
  slot.** A palette, a region, the two pattern brushes were passed over
  silently, so everything created after one landed a slot low and every
  selection picked its predecessor: a contour chart drew all 736 of its lines
  in the black pen its ten coloured ones had displaced.

- **A WMF that moves its window origin is read in one frame.** The origin names
  the frame the records AFTER it are drawn in, and a file may move it as often
  as it likes; the last one was applied to everything, so a drawing made under
  an earlier origin came out displaced by the difference. One slide's metafile
  moves its origin six times.

- **A CJK workbook measures its columns by a half-width digit.** The table of
  faces a column's width unit is taken from held only Latin ones, so every CJK
  file fell back to Excel's 7px where the digit is 8: a twelve-point PMingLiU
  workbook came out on forty-nine pages against the reference's seventy-eight.

- **A freeform's path, and a fill the shape says it does not use.** The vertex
  and segment arrays state their length without the six-byte header they carry,
  so a curve of twenty-one points read as a single line; and a shape states a
  fill colour whether or not it is filled, which §2.3.7.43's booleans decide.

## 1.22.0

A release about presentations. A deck was read but barely drawn: its slides
came out as the text on them, over white, with the design the deck states once
and every slide relies on left in the master. Measured against LibreOffice over
473 real decks, the pixel distance summed to 74.7 with 33 files rendering
almost nothing recognisable; it now sums to 22.1 with one, and 409 of the 472
comparable decks sit within a tenth of the reference.

### Added

- **A deck's own decoration reaches every slide.** The rules, bands, logos and
  panels a master and its layouts carry are drawn — under the slide's own
  content, where §19.3.1 puts them — and a slide or a layout that asks for
  none (`@showMasterSp`) gets none. The background comes with them: a solid
  fill, a theme reference (`p:bgRef` into the theme's background fill styles),
  a gradient or a picture, and a shape marked `useBgFill` is painted with the
  piece of it that lies under the shape.

- **A slide table stands where it is put and wears its style.** A slide table
  states almost nothing about how it looks: its `a:tblPr` names a style by GUID
  in `tableStyles.xml` and switches on the parts of it that apply — a header
  row, banded rows, a first column. Those are read and composed in the spec's
  order, with the cell's own `a:tcPr` over them: its fill, its `a:noFill`, its
  four rules and the row heights it asks for.

- **A slide is written in the style its deck states.** The presentation's
  `p:defaultTextStyle`, the master's three text families and each prototype's
  own list carry a slide's type: its sizes, its colours, its alignment and
  anchor, its bullets (character, colour and size, read out of the symbol face
  that states them) and `cap="all"`. A slide's own runs state only where they
  differ — including where they turn something OFF.

- **An embedded object shows the picture it keeps.** A `p:oleObj` is another
  document on a slide, and what the slide shows of it is the snapshot beside
  it: the modern spelling as a `p:pic`, the legacy one through the slide's VML
  drawing. A deck whose eighth slide is one embedded presentation is no longer
  a blank page.

- **A picture is painted the way the file recolours it.** `a:duotone` (through
  a luminosity soft mask, so a JPEG is never decoded), `a:clrChange` (one
  colour replaced or knocked out), `a:alphaModFix` (how opaque the picture is
  drawn) and `a:blipFill` on a chart's own frame, tiled at its natural size or
  stretched over it.

- **A gradient that fades out fades out.** Its stops' transparency is read as
  transparency, and in PDF the sweep is painted through a luminosity mask built
  from the same geometry — a shading carries no alpha of its own.

- **Twenty-five more preset shapes**, including the arrows, callouts, brackets
  and the flowchart family that slides are built from, plus the symbol fonts
  (Wingdings, Webdings, Symbol) whose letters are not letters.

### Fixed

- **Thirteen decks would not open at all.** A package was sniffed by scanning
  its bytes for a part name, and a deck that merely mentions `xl/workbook.xml`
  somewhere in a compressed stream was read as a workbook. The ZIP's own
  directory is asked instead.

- **A deck says for itself what `bg1` and `tx1` mean.** §19.3.1.6 `p:clrMap`
  maps the four background/text slots onto theme colours, and a layout or a
  slide may override it. Ignored, every deck whose master swaps light and dark
  came out inverted — dark text on dark paper.

- **The hue a theme multiplies.** §20.1.2.3.14 `a:hueMod` was not read at all,
  so the Ion theme's backdrop — one dark teal and a 108% hue — stayed teal
  where both references draw deep blue.

- **What stands behind the content paints before it.** The page paints by kind,
  every image and then every shape, so a slide's white backdrop landed on top
  of the photograph it stands on and a layout's card buried the picture the
  slide put on it. What a page puts behind now paints first, in its own order,
  and its text is set with the same emitter as the page's own — 112 documents
  in the Word corpus draw closer for that alone, the watermarks and header
  drawings most of all.

- **A text box applies the paragraph's indent**, which the page, the band and a
  table cell each already did. Slides' bullets stood in the margin with their
  text against them.

- **A SmartArt drawing is found where PowerPoint names it**, and each diagram
  on a slide gets its own: a deck with four of them drew the first one four
  times.

- **A picture fill's `a:fillRect` says where the picture goes**, not how far to
  zoom it, and a shape turned by `a:xfrm@rot` is turned.

## 1.21.0

### Added

- **A protected document is answerable for from outside.** 1.20.0 opens one
  with its password but kept the two things a caller needs to ask around it
  inside: `isEncryptedPackage(bytes)` says a file wants a password BEFORE the
  parse throws — so an interface can put up the prompt instead of an error —
  and `WrongPasswordError`, thrown all along, is now a class to catch by type
  rather than an `err.name` string to match. Both are on the package.

## 1.20.0

### Added

- **A password-protected document opens with its password.** An encrypted
  `.docx`/`.xlsx` is not a zip at all — ECMA-376 §2.3 puts the whole OPC
  package in the `EncryptedPackage` stream of an OLE container — so every
  reader declined it and the caller was told to re-save without a password.
  Pass the password to `Ream.parse` and it opens: both MS-OFFCRYPTO schemes
  are read, the standard one (AES-ECB under a key spun from 50 000 SHA-1
  rounds, Office 2007 and LibreOffice) and the agile one (an XML descriptor
  naming its own hash and cipher, the package cut into 4096-byte segments,
  Office 2010 and later). A wrong password is refused as a wrong password —
  each scheme carries a verifier — rather than yielding rubbish.

### Fixed

- **SHA-512 is sixteen times faster.** It was written in BigInt, which is the
  clear way to write it and forty times the cost of 32-bit pairs; an agile
  encrypted document spins the hash a hundred thousand times, so 3.7 seconds
  of key derivation became 0.23. The digests are checked against Node's for
  every message length around the block boundaries.

## 1.19.0

A second pass of the pixel ranking over all 1121 corpus documents, taken from
the top. Six changes, and what is left above 0.09 is either a recorded
difference of judgement or a font metric.

### Added

- **A tiled picture fill is laid, not stretched.** `@type="tile"` (VML) and
  `<a:tile>` (DrawingML) repeat the picture over the box at its own size —
  the pixels read at 96 dpi, the resolution Office measures a picture in.
  Read as nothing, a page papered with a texture fell back to its flat colour;
  read as a stretch, the texture became one blurred picture across the shape.
  A pattern fill is still not this: a two-colour tile at its own tiny scale,
  which stretched over a shape is a black slab.

### Fixed

- **A shade is a fraction of light, not of the byte.** `<a:shade>` and
  `<a:tint>` take their percentage on LINEAR values; taken on the gamma-encoded
  byte, every shade came out far too dark — the 51 % the Office theme builds
  its gradients from turned a medium red into a near-black maroon. Half of the
  Office accent blue is now 2F528F, the neighbourhood Word's own "darker"
  palette lives in.
- **A vertical merge is painted as the one cell it is.** Word writes the fill
  on the cell that OPENS a `<w:vMerge>` and nothing on the rows continuing it;
  painting only what is written left a white notch under every merged cell.
- **A floating table in a header stands at its anchor.** `<w:tblpPr>` was read
  everywhere but the band emitter, which walked the table's rows down the
  header's cursor — a banner pinned across the top of the page sat inside the
  margins and 78pt too low. It leaves the band's flow with that, so the band
  stops reserving room for it and the body starts where it should.
- **A measurement may name its unit** — `"85.05pt"`, `"1cm"` — wherever a count
  of twips is expected, not only in the page geometry: a document that writes
  its tab stops, indents and spacing that way lost every one of them, and with
  the tab stop went the dot leader hanging on it. A font size is read in ITS
  own unit (`<w:sz>` counts half-points), and an indent accepts the Strict
  `start`/`end` spelling of its sides.

## 1.18.0

What the pixel ranking said was left after 1.17.0, taken in order: the picture
formats a document may embed that nothing here could read, the drawings that
had nowhere to stand, and the places where a page breaks.

### Added

- **EMF and WMF are read and drawn.** The picture format Word writes for
  anything it draws itself — the preview of an embedded workbook, a legacy
  clipart, an ActiveX control's face, an equation — was carried through the
  package and never rendered: the box was reserved and left blank in forty
  corpus documents. Both formats are one device interface (a context with a
  pen, a brush, a font and a current point), so one reader vocabulary serves
  both: the state records, the window/viewport and world-transform mapping, and
  the drawing ones — lines, polygons, polylines, Béziers, rectangles, ellipses,
  path brackets, both text records, and the brush blits a WMF uses for every
  rule and panel. What needs a raster or a region — bitmap blits, gradients,
  flood fills, palettes — is named as a loss rather than drawn wrong. A picture
  draws in the flow, floating, in a header, inside a cell, and inline in a line
  of text, and turns with its drawing's own rotation.
- **TIFF.** A `.docx` embeds one as plainly as a PNG and PDF has no filter that
  carries one, so it is decoded to samples: both byte orders, WhiteIsZero /
  BlackIsZero / RGB / palette / CMYK, 1/2/4/8/16-bit, uncompressed, LZW,
  PackBits and Deflate, horizontal differencing, strips and tiles, and an
  ExtraSamples channel as the soft mask.
- **A drawing anchored inside a table cell** — a title over the first cell of a
  glossary, logos in the corner of a form, a page number in a footer table.
  Which box its position is measured in is `@layoutInCell` (VML
  `o:allowincell`), now read: by default the CELL is every box the anchor can
  name, and turned off the drawing reaches past the table to the paper.
- **Linked text boxes.** `wps:linkedTxbx`: the chain's first box holds the
  words and each box shows as much as its height allows, the rest carried on by
  `@seq` — not by document order, since a continuation may be written first.
- **Text on BOTH sides of a float.** Where a drawing leaves room either side,
  a line runs down one gap and continues in the other at the same baseline;
  `left`, `right` and `largest` keep to one side, as they say.
- **A picture's washout** (`@gain`/`@blacklevel`): contrast and brightness about
  mid grey, drawn as the flat veil that arithmetic is — exact for every wash a
  document can state, and it rewrites no pixels, which matters because a JPEG is
  embedded verbatim.
- **A transparent fill is transparent.** `<a:alpha>` and VML `@opacity` belong
  to the fill, not to its colour: composited over white, a translucent shape
  drawn over anything else was simply the wrong colour.

### Fixed

- **A table row breaks where the page ends** (§17.4.6). A row was split only
  when it was taller than a whole page, so one that merely outgrew the space
  left was drawn where it stood and ran off the bottom, over whatever followed.
  Under that, the chunk sizer measured lines and not the space between their
  paragraphs, so a cell reported 518pt where it drew 713 and everything past
  the first piece was silently dropped. A sheet row still moves whole: Excel
  and LibreOffice never break one.
- **A column of another width re-breaks what lands in it.** Every block is
  broken at the first column's measure, because nothing knows which column it
  will land in; a section of unequal columns set its wide column narrow and took
  two pages for one.
- **A frame that forbids text beside it** (`w:wrap="notBeside"`) stands where
  its `w:x` and `w:y` name and excludes the whole column for its height. Read as
  a top-and-bottom wrap it went back into the flow and lost its x; read as a
  square wrap it let text run up beside it, which is the one thing the mode
  forbids. A frame anchored to the TEXT stays in the flow — that is the only
  place it names.
- **A VML line stands where its ends put it.** `v:line` states `from` and `to`
  rather than a box, and its style carries no position for the anchor to read;
  a margin rule was drawn down the very edge of the paper.

### Internals

- The anchor arithmetic (§20.4.3.3/§20.4.3.4) moved out of the page assembler
  into two functions over an explicit frame, so a table cell resolves a
  drawing's position by exactly the same rules the body does, with its own
  boxes.
- LZW is one coder now (`core/lzw`), shared by the TIFF reader and the PDF one
  rather than kept in step as two copies.
- An image resource may carry a metafile picture instead of a prepared raster:
  a metafile has nothing to embed and takes no resource name.

## 1.17.0

The `.docx` half of the same sweep: every one of 1121 real Word documents opened
beside a reference conversion, the pages compared as pictures rather than as
text, and then ranked by how many pixels differ so the worst page was always the
next one to open. Nothing below was visible in a text dump.

### Added

- **The page background Word actually paints.** `<w:background w:color>` is only
  the flat fallback; the `<v:background>` beside it carries a gradient or a
  picture. Both render now — including the stops a VML gradient lists between
  its two named colours, the band an `@focus` puts the second colour in, the
  corner a radial sweep starts from, and the rectangular contours VML draws it
  with (a PDF calculator function, since no shading type describes them).
- **Shadows.** A shape's own `<a:outerShdw>`, the one a gallery style names
  through `<a:effectRef>`, and the one a picture carries on its `pic:spPr` —
  which rides the image so an inline picture gets it too. `blurRad` draws as a
  soft edge rather than a cut one.
- **A table inside a text box**, which is what a title page usually is: one box
  holding one table.
- **A gallery fill is the theme's whole fill** — a solid, a gradient, a picture —
  including the background fill styles an `<a:fillRef>` reaches past index 1000,
  where Word's cover pages live. A path gradient starts where its
  `<a:fillToRect>` says.
- **Drawings placed and sized as the document asks**: a keyword position
  (`<wp:align>`), an anchor or a relative size against a margin BAND rather than
  the text area, a share of the page, the room a rotation needs, and the
  distance text keeps from a float.
- **Numbering**: `chicago` reference marks, legal numbering in arabic, ideographs,
  a picture bullet at the size of the text beside it, an instance that redefines
  a level or starts it elsewhere, and a heading numbered by its style.
- **Text frames** (`w:framePr`) float where they say, at the height they state.
- **Legacy VML shapes and locked canvases** draw; so do pattern fills, arrowheads,
  underline styles, GIFs, and a shape written inside a table cell.
- **Sections**: a rule between columns, a column break, pages printed from the
  number the section names, and a table that reads right to left.
- **Charts** are drawn with the rules they state.

### Fixed

The style cascade

- **A table style ranks under the style a paragraph is written in** (§17.7.2), on
  the paragraph side and the run side both. Written on as if the document had
  stated them, a table style's spacing packed the headings of a styled cell
  together and its 8pt blue overwrote every colour and size the text named.
- **A paragraph that names no style is written in the default one**, and
  automatic spacing is a property of its own rather than a length that any
  literal beats.
- **A cell counts the space between its paragraphs once** — twice over, a 685pt
  cell measured 991pt and was split where no reader splits it.
- **A gallery shape lends its text the colour it names**, at the rank a theme
  colour holds: under every style the text itself names.

The page and its sections

- **A document that states no page size is on US Letter**, which is what both
  references use, not the library's own A4 fallback.
- **A measurement may name its unit** (`"28.35pt"`, `"1cm"`): a whole page
  geometry written that way had been read as nothing at all.
- **The body carries one section break, and it closes the body**; a section with
  no header of its own is not pushed down by a header it does not show.
- **A line stands as tall as what prints on it** — Word measures a line's spaces
  at the size of its text, not their own.

Drawings

- **A header's drawing keeps its size and its place in the stack**: an anchored
  picture may hang into the margins, and `@behindDoc` puts it behind everything
  in the band whatever order it was written in.
- **A paragraph that anchors a drawing keeps its own line** — the drawing leaves
  the flow, the paragraph mark does not.
- **A run that carries a picture is not an empty run**, so a break after one
  sends what follows to the next page instead of moving the picture.
- **A VML shape wraps text only where it says so**, keeps the z-order its style
  states, and a group nested in a group is scaled by both coordinate spaces.

Reading robustly

- **The main document part is whatever the package says it is** — a file naming
  it anything but `word/document.xml` had been refused outright.
- **A document of nested tables, several bodies, or Strict-namespace charts and
  diagrams is read** rather than partly dropped.

## 1.16.0

The end of a validation sweep that opened every one of 649 real `.xlsx`
documents beside a reference conversion and compared the two rendered pages.
Most of what follows was invisible in a text dump and obvious in a picture.

### Added

- **SmartArt in a spreadsheet.** A diagram is four parts of description — data,
  layout, quick style, colours — that a reader is meant to lay out itself. The
  producer also writes what it drew, as plain DrawingML under `dsp:`, and that
  is what renders: shapes at the frame's corner, each label in the rectangle
  the diagram sets aside for it. A file with no drawing fallback degrades to a
  graceful loss rather than an empty space.
- **Conditional formats carried in the 2009 extension.** ISO/IEC 29500 could not
  express everything Excel 2010 wanted, so the rest lives in `<extLst>` under
  `x14`: the range is a child `<xm:sqref>`, the formulas are `<xm:f>`, and the
  format is written inside the rule rather than pointing at the workbook's
  table. Nineteen corpus workbooks put rules there.
- **Printed row and column headings** (`<printOptions headings="1">`), repeated
  on every page and on every column band.
- **Charts**: a secondary value axis of its own, axis titles (including one read
  bottom-to-top), a legend key drawn per series type, data labels the author
  typed, and an axis drawn in the number format it declares.
- **Shape shadows** (`<a:effectRef>` and a direct `<a:outerShdw>`), and a shape's
  fill and outline width taken from the theme's style lists.
- **Superscript and subscript** draw off the baseline, and smaller.
- **Every PNG colour type**, interlaced or not.

### Fixed

Shapes and drawings

- **A turned shape draws turned.** `a:xfrm`'s rotation and flips were read
  nowhere, so an arrow pointing up and to the right lay flat across the page. A
  turned shape is also the one case where the anchor is not its box: Excel spans
  from/to across the ground the shape covers once rotated and keeps the shape's
  own size in `a:ext`.
- **A gallery shape finds its colour.** `<a:fillRef>` and `<a:lnRef>` each wrap
  one colour, and the reader took the reference's first *child* — which in a
  part saved with indentation is a newline. Such shapes drew with neither fill
  nor outline, which is to say not at all.
- **`absoluteAnchor`** names no cell: it gives the distance from the sheet's own
  corner in `<xdr:pos>`. Unread, every such shape piled into the corner.
- **A group draws the shapes inside it** — `xdr:grpSp` maps its children through
  `a:chOff`/`a:chExt` instead of being skipped.
- **Block-arrow geometry.** §20.1.9.18 measures both the shaft and the head
  against the *shorter* side of the box, not the direction the arrow points, so
  a tall thin arrow now has a compact head rather than a spike two thirds down
  the shaft.
- **`a:hueOff` / `a:satOff`** — hue is counted in sixtieth-thousandths of a
  degree, a different unit from every other transform in that family.

Reading robustly

- **An entity a DOCTYPE declares resolves** instead of printing its own spelling.
  Nested declarations are the billion-laughs bomb, so each expands once under a
  budget and a runaway one resolves to nothing.
- **A style behind `mc:AlternateContent` keeps its place.** Style collections are
  indexed by position, so a wrapper the reader skipped shifted every entry after
  it — a form lost its table borders, its title and its column alignment
  together.
- **A corrupt archive entry is refused by name**, not by the inflater's crash.
- The 2006 beta names a shared string `<sstItem>`; `_xHHHH_` escapes decode; an
  inline string written as a single formatted run is read.

The printed page

- **Fit-to-page overrides the breaks drawn into the sheet.** A manual break costs
  a page at every scale, so a sheet asking to fit on one page could never
  satisfy it and gave up scaling altogether.
- **A sheet that declares no row height takes the workbook's own.** The height is
  a line height, and a workbook naming Arial 10 gets 12.75pt where Calibri 11
  gets 15.
- **A line feed in a cell that does not wrap leaves nothing behind** — it is a
  control character with no width, not a space.
- The print scale reaches the header, the footer and the drawings; a column band
  with nothing in it is not printed; a band stops at its last row that draws
  something.

Conditional formats and number formats

- **A `cellIs` rule compares words**, not only numbers: `notEqual ""` is how a
  document asks to format every cell that holds something.
- A data bar fades away from the axis it grows out of; a colour scale keeps its
  gradient when a stop is a formula; a rule belongs to its cell, not to the
  columns its text borrowed.
- A format section may state a condition and then pick itself; a comma against
  the last placeholder scales the number; General spells an exponent Excel's way
  and gives up decimals in the document's unit.

Charts

- A line chart that asks for markers gets them; a scatter's points are the symbol
  the file names; a chart silent about its gap takes the schema's default rather
  than none; a ranked bar chart reads top-down; a series cycles the workbook's
  accents rather than a built-in palette.

## 1.15.4

Fixes a regression in 1.15.3 that could turn a spreadsheet into hundreds of
blank pages, and three more spreadsheet read gaps found by the same validation
sweep.

### Fixed

- **Blank pages from merged empty cells** (regression in 1.15.3). A sheet whose
  merges cover cells that hold nothing no longer paginates that emptiness: one
  corpus workbook of 50 000 such merges produced 1042 blank pages where it
  should produce one. The used range is built from cells that carry content,
  and a merge now extends it only when its origin does — a merge follows
  content, it does not create any. 1.15.3 made blank space paginate correctly,
  which was right, and exposed this.
- **Wide sheets were truncated.** A worksheet may use all 16 384 columns, and
  we capped the rendered grid at 1024 — a limit of ours, not the format's,
  which cut a two-row sheet spanning the full column space from 1639 pages to
  103. Memory is bounded by the total-cell budget instead, which such a sheet
  comes nowhere near.
- **Numeric cell references.** A producer that addresses cells as
  `"<column>_<row>"` instead of A1 no longer has its values packed
  consecutively — and so filed under the wrong column headings. The reading is
  only used where the file corroborates it (every reference inside a row must
  agree with that row); otherwise document order stands, as before.
- **ActiveX controls reached through `<control>`.** §18.3.1.19 `<control>`
  resolves to either a form control or an ActiveX one, and only the
  relationship target says which. Routing them all to the form-control path
  left a sheet of option buttons rendering as bare names, with their captions,
  states and group names unread beside them.

## 1.15.3

Spreadsheet rendering: the grid now lands where a spreadsheet puts it, and a
sheet that cannot be rendered in full says so instead of quietly shrinking.
Found by a validation sweep over 642 real `.xlsx` documents.

### Fixed

- **Column widths.** A `<col width>` is now honoured as written. Widths were
  treated as a hint and refitted to cell content — right for WordprocessingML,
  wrong for a spreadsheet, where it moved every column after the first. The
  width also regains the 5-pixel padding the measurement carries
  (ECMA-376 §18.3.1.13), and `<sheetFormatPr>`'s `defaultColWidth` /
  `baseColWidth` are read for columns no `<col>` covers.
- **Row heights.** `<sheetFormatPr defaultRowHeight>` is read; rows without an
  explicit `ht` take it, or Excel's 15pt when the sheet declares none. Row pitch
  was previously whatever leading the rendering font wanted, so a sheet asking
  for 30pt rows drew them at 13pt.
- **Page margins.** A sheet that declares no `<pageMargins>` now prints with
  Excel's defaults (0.7in sides, 0.75in top and bottom) rather than a word
  processor's inch, which shifted the whole grid.
- **Number alignment.** "General" alignment follows the value's type
  (§18.8.1): numbers and dates right, booleans and errors centred, text left.
  Columns of figures previously rendered flush left.
- **Cell padding.** Cell text is inset like a spreadsheet's — a couple of
  points at the sides, nothing above or below, where the row height is the box.
- **Print scale.** `<pageSetup scale>` now shrinks column widths along with
  fonts and row heights. A scaled sheet kept full-width columns and lost the
  overflow off the page edge.
- **Text overflow.** A cell without `wrapText` runs its text across the empty
  cells to its right on one line, as Excel and LibreOffice draw it, instead of
  wrapping it inside its own column. It stops at a neighbour that holds
  anything of its own — a value, a fill, a border, a sparkline.
- **Per-sheet page setup.** Each sheet prints on the paper and orientation it
  declares. A workbook mixing, say, A4 landscape and US Letter portrait was
  printed entirely on the first sheet's paper.
- **Blank rows paginate.** A sheet whose rows are mostly empty no longer piles
  them all onto page one: consumed space now breaks a page whether or not
  anything was drawn on it. Blank space is most of a spreadsheet.
- **Cell references.** A cell whose `r=` cannot be parsed is placed by document
  order — the fallback §18.3.1.4 already defines for an absent one — instead of
  being dropped. Files from at least one producer rendered blank.
- **Inline strings.** A cell typed `inlineStr` that writes its text into `<v>`
  instead of `<is>` renders that text rather than nothing.
- **Windows ZIP separators.** Packages whose entries use `\` instead of `/`
  now open; they were rejected outright as missing their root relationships.
  Affects `.docx`, `.xlsx` and `.pptx`, including format detection.
- **Zip64.** An entry declaring the 0xFFFFFFFF size sentinel is no longer read
  as a literal 4 GiB and refused by the archive-size guard.

### Changed

- **Spreadsheet conversions now report losses.** The xlsx reader always
  returned an empty loss report, while the print model quietly clipped
  pathological sheets — an over-wide grid, an exhausted per-sheet text budget,
  an oversized sparkline range. Those clips are now reported. Callers passing
  `strict: true` will see such a document throw where it previously converted
  silently; that is the point of the report, but it is a behaviour change.

### Security

- **Declared-extent amplification.** A 2.5 KB worksheet declaring
  `A1:XFE16777217` exhausted a 6 GB heap: the grid was bounded per dimension
  (1024 columns, 50 000 rows) but nothing bounded their product. The projected
  grid is now bounded by total cells, and the blank tail a bound leaves behind
  is trimmed rather than paginated.

## 1.15.2

A layout fix for CJK line wrapping.

### Fixed

- **CJK line wrapping.** Long Chinese / Japanese / kana text in a narrow container
  (a table cell, a narrow column) now wraps like Word instead of overflowing the
  edge. The line breaker only opened break opportunities at whitespace, so a
  space-less CJK run stayed one unbreakable box; it now breaks between adjacent
  ideographs (Unicode UAX #14) — keeping closing punctuation off the start of a
  line and openers off the end. Non-CJK layout is byte-identical.

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
