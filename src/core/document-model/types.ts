// Typed in-memory model produced by OOXML parsing and consumed by layout.
//
// All lengths are canonical points (Pt, 1/72"); readers convert format-native
// units (twips, half-points, EMU) at the parse boundary (ir-design.md §5).
//
// All "properties" objects use optional fields to allow inheritance / cascade:
// undefined = "inherit from parent in the style chain". The cascade resolver
// (style-cascade) merges these into ResolvedRunProperties / ResolvedParagraphProperties
// that the renderer consumes.

import type { NativeBag, Pt, ResourceId } from '@/core/ir';
import type { ShapeGradient } from '@/core/vector';

/** Paragraph horizontal alignment (`w:jc`); `both` = justified. */
export type Alignment = 'left' | 'right' | 'center' | 'both' | 'distribute';

/** Run underline style (ECMA-376 §17.18.99 ST_Underline, the subset Ream renders). */
export type UnderlineStyle =
  | 'none'
  | 'single'
  | 'double'
  | 'thick'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashHeavy'
  | 'wave';

/** Run vertical alignment (`w:vertAlign`): baseline or super/subscript. */
export type VerticalAlign = 'baseline' | 'superscript' | 'subscript';

/**
 * The four script slots of `w:rFonts` (§17.3.2.26). A character picks its font
 * from the slot its Unicode range maps to (ASCII, high-ANSI, complex-script,
 * East-Asian).
 */
export interface FontFamilyMap {
  readonly ascii?: string;
  readonly hAnsi?: string;
  readonly cs?: string;
  readonly eastAsia?: string;
}

/**
 * ECMA-376 Part 1 §17.3.2 — Run Properties (`rPr`). All fields optional so the
 * cascade can inherit them (undefined = "inherit from parent in the style
 * chain"). Boolean toggle properties follow §17.17.4: an absent `<w:b/>` means
 * "inherit"; a present `<w:b/>` with no val (or val=true/1/on) means true;
 * val=false/0/off means explicit false.
 */
export interface RunProperties {
  readonly styleId?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: UnderlineStyle;
  /** §17.3.2.40 `w:u @w:color` — the underline's own colour, when it has one. */
  readonly underlineColorHex?: string;
  readonly strike?: boolean;
  /** §17.3.2.5 `w:caps` — the run is DISPLAYED in capitals, whatever it stores. */
  readonly caps?: boolean;
  /**
   * §17.3.2.33 `w:smallCaps` — displayed in capitals, with the letters that
   * were lower case set smaller.
   */
  readonly smallCaps?: boolean;
  readonly fontSizePt?: Pt;
  readonly colorHex?: string;
  readonly fontFamily?: FontFamilyMap;
  readonly verticalAlign?: VerticalAlign;
  /**
   * ECMA-376 §17.3.2.30 — `w:rtl`. Marks the run as right-to-left, which seeds
   * the BiDi algorithm with an R/AL bias even for neutral characters.
   */
  readonly rtl?: boolean;
  /**
   * ECMA-376 §17.3.2.20 — `w:lang @w:val` (e.g. `"en-US"`, `"ru-RU"`). Used for
   * the tagged-PDF per-element `/Lang`; does not affect visual layout.
   */
  readonly lang?: string;
  /**
   * §17.3.2.32 `w:rPr/w:shd` — the background painted behind the run's own
   * glyphs, already blended from the pattern, its colour and the fill.
   */
  readonly shadingColorHex?: string;
  /**
   * §17.3.2.35 `w:rPr/w:spacing` — extra space between the run's characters
   * (negative tightens). Word states it in twentieths of a point.
   */
  readonly letterSpacingPt?: Pt;
}

/**
 * ECMA-376 Part 1 §17.3.1 — Paragraph Properties (`pPr`). All lengths in
 * canonical Pt (converted from twips by the parser); all fields optional for
 * cascade inheritance.
 *
 * `indentFirstLine` encodes both `<w:ind w:firstLine="…"/>` (positive value)
 * and `<w:ind w:hanging="…"/>` (stored as the negative twips, so a hanging
 * indent of 360 twips becomes `indentFirstLine = -18pt` — the first line ends up
 * 360 twips to the left of `indentLeft`).
 */
/**
 * §17.3.1.38 `w:tab` — one stop in a paragraph's tab list: where the text after
 * a tab character goes, how it sits against that position, and what fills the
 * gap on the way there.
 */
export interface TabStop {
  /**
   * Distance from the text margin, in points (`w:pos`, twips). Ignored when
   * {@link relativeTo} names an edge — a §17.3.3.15 `w:ptab` positions against
   * the column itself, whose width only the layout knows.
   */
  readonly positionPt: Pt;
  /**
   * §17.3.3.15 `w:ptab` — an ABSOLUTE position tab, which goes to the middle
   * or the far side of the text column rather than to a stated distance.
   */
  readonly relativeTo?: 'center' | 'right';
  /** §17.18.90 ST_TabJc — how the text after the tab sits against the stop. */
  readonly alignment: 'left' | 'center' | 'right' | 'decimal';
  /** §17.18.89 ST_TabTlc — the character drawn across the gap. */
  readonly leader?: 'dot' | 'hyphen' | 'underscore' | 'middleDot';
}

/**
 * §17.3.1.11 `w:framePr` — the paragraph is a floating TEXT FRAME: it leaves
 * the flow, takes the box this describes, and the body text wraps around it.
 * Consecutive paragraphs carrying the same frame are one frame together.
 */
export interface FrameProperties {
  /** `w:w` — the frame's width. Absent ⇒ as wide as the text needs. */
  readonly widthPt?: Pt;
  /** `w:h` with `w:hRule` — `exact` pins the height, anything else fits the text. */
  readonly heightPt?: Pt;
  readonly heightRule?: 'auto' | 'exact' | 'atLeast';
  /** `w:x`/`w:y` — the offset from whatever `hAnchor`/`vAnchor` names. */
  readonly xPt?: Pt;
  readonly yPt?: Pt;
  /** `w:xAlign`/`w:yAlign` — an alignment instead of an offset. */
  readonly xAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside';
  readonly yAlign?: 'top' | 'center' | 'bottom' | 'inside' | 'outside';
  readonly hAnchor?: 'text' | 'margin' | 'page';
  readonly vAnchor?: 'text' | 'margin' | 'page';
  /** §17.18.104 ST_Wrap — how the body text runs past the frame. */
  readonly wrap?: 'auto' | 'around' | 'none' | 'notBeside' | 'tight' | 'through';
  readonly hSpacePt?: Pt;
  readonly vSpacePt?: Pt;
}

export interface ParagraphProperties {
  readonly styleId?: string;
  /**
   * §17.7.2 — the `w:pPr` of the TABLE STYLE the paragraph's cell is formatted
   * by. It is not the paragraph's own formatting: the cascade puts it under the
   * paragraph's style (docDefaults < table style < numbering < paragraph style
   * < character style < direct), so it is carried apart from the direct
   * properties rather than merged into them.
   */
  readonly tableStyle?: ParagraphProperties;
  /**
   * §17.7.2 — a run layer that ranks BELOW the paragraph's own style and below
   * the character style a run names: the `w:rPr` of the table style the cell is
   * formatted by, or the `a:fontRef` colour a gallery-drawn shape lends the
   * text inside it (§20.1.4.2.14). Carried on the paragraph because every run
   * of it — and its MARK — reads the same layer.
   */
  readonly inheritedRun?: RunProperties;
  /**
   * §17.6.17 — the paragraph's mark carries a `w:sectPr`: it is the last
   * paragraph of its section, and the mark IS the section break. One with no
   * content of its own therefore prints nothing, not an empty line.
   */
  readonly sectionBreak?: boolean;
  /** §17.3.1.11 — the paragraph floats as a text frame. */
  readonly frame?: FrameProperties;
  /**
   * §17.3.1.41 `w:textDirection` — which way the paragraph's lines run.
   * `btLr` reads bottom-to-top, `tbRl` top-to-bottom; `lrTb` is the default
   * and is not recorded.
   */
  readonly textDirection?: 'btLr' | 'tbRl';
  /**
   * §17.3.1.32 `w:snapToGrid` — whether the paragraph's lines stand on the
   * section's document grid (§17.6.5). Absent ⇒ they do; Word's header and
   * footer styles are what usually says otherwise.
   */
  readonly snapToGrid?: boolean;
  readonly alignment?: Alignment;
  readonly spacingBefore?: Pt;
  readonly spacingAfter?: Pt;
  readonly spacingLine?: Pt;
  readonly spacingLineRule?: 'auto' | 'exact' | 'atLeast';
  /**
   * ECMA-376 Part 1 §17.3.1.9 — `w:contextualSpacing`. The space before and
   * after is dropped where the neighbour is a paragraph of the SAME style: a
   * list is written this way, so its items sit together while the list as a
   * whole keeps its space from the text around it.
   */
  readonly contextualSpacing?: boolean;
  /** ECMA-376 Part 1 §17.3.1.37 — the paragraph's own `w:tabs` stops. */
  readonly tabs?: ReadonlyArray<TabStop>;
  /**
   * §17.3.1.24 `w:pBdr` — rules drawn around the paragraph. `w:space` is the
   * gap each keeps from the text, carried on the {@link Border} as `spacePt`.
   */
  readonly borders?: CellBorders;
  /** §17.3.1.31 `w:pPr/w:shd` — the paragraph's own background fill. */
  readonly shading?: CellShading;
  readonly indentLeft?: Pt;
  readonly indentRight?: Pt;
  readonly indentFirstLine?: Pt;
  /** The implicit default run properties for runs in this paragraph (`w:pPr/w:rPr`). */
  readonly runProperties?: RunProperties;
  /**
   * §17.3.1.3/§17.3.1.1 — the gap `w:beforeAutospacing`/`w:afterAutospacing`
   * asks the consumer for, already resolved to a length by the reader (it
   * depends on the document's compatibility mode). Present only when the
   * autospacing flag is ON, and it beats the `w:before`/`w:after` beside it.
   * A separate property because style inheritance is per ATTRIBUTE: a style
   * based on one that says `beforeAutospacing` inherits the flag even when it
   * states a `w:before` of its own — which is how tdf104354_firstParaInSection
   * gets its 75pt before, and taking that literally spread one page over four.
   */
  readonly spacingBeforeAuto?: Pt;
  readonly spacingAfterAuto?: Pt;
  readonly numbering?: NumberingReference;
  /**
   * §17.9.4 — a `w:numPr` that names a LEVEL and no `w:numId`: the instance is
   * whatever the style this one is based on refers to. Word's own Heading 2
   * says exactly this, and dropped it numbered num-parent-style.docx's headings
   * 1, 2, 3, 4 where its own text says they should read 1, 1.1, 2, 2.1.
   */
  readonly numberingLevel?: number;
  /**
   * ECMA-376 Part 1 §17.3.1.21 — `w:pageBreakBefore`. When true, the paragraph
   * starts on a fresh page even if there is room on the current one.
   */
  readonly pageBreakBefore?: boolean;
  /**
   * ECMA-376 §17.3.1.6 — `w:bidi`. Sets the paragraph's base direction to RTL,
   * so the BiDi paragraph embedding level is 1 and default alignment is right.
   */
  readonly bidi?: boolean;
  /**
   * ECMA-376 §17.3.1.20 — `w:outlineLvl`. Outline level 0–8 maps to Heading 1–9;
   * 9 (or absent) is body text. Drives tagged-PDF heading structure (H1–H6).
   */
  readonly outlineLevel?: number;
}

/**
 * ECMA-376 §20.4.2.8 — `wp:inline` picture extent inside a `w:r`. Stored on
 * {@link Run} alongside (or instead of) `text` so layout can position the image
 * as if it were a glyph in the line box.
 */
/**
 * §20.1.2.2.24 `a:ln` on a `pic:spPr` (VML: `@stroked`/`v:stroke`) — the frame
 * a picture is drawn with. Word's "Picture Border": a rule around the picture
 * box, not part of the image itself.
 */
export interface PictureOutline {
  readonly colorHex: string;
  readonly widthPt: Pt;
}

export interface InlineImage {
  /**
   * Content-addressed bytes in the document's `ResourceStore`; absent when the
   * source relationship did not resolve (the layout box still reserves space).
   */
  readonly resource?: ResourceId;
  /** The picture's own frame, when it has one (see {@link PictureOutline}). */
  readonly outline?: PictureOutline;
  /**
   * §20.1.8.40 `a:outerShdw` — the drop shadow under the picture. Word writes
   * one on every screenshot pasted with a style; drawn nowhere, imgshadow.docx's
   * six stood flat on the page where both references lift them off it.
   */
  readonly shadow?: ShapeShadow;
  /**
   * §14.1.2.10 `@gain`/`@blacklevel` — the contrast and brightness the picture
   * is drawn through, about mid grey: `out = (in - 0.5) * gain + 0.5 + black`.
   * Word washes a watermark out this way.
   */
  readonly wash?: { readonly gain: number; readonly black: number };
  /**
   * §20.1.8.16 `a:clrChange` — one colour of the picture replaced by another,
   * or knocked out entirely when the destination states `a:alpha` at zero. A
   * logo on a white card goes onto a dark slide that way (corpus: tdf113163,
   * whose whole slide is a metafile with its white ground declared away).
   */
  readonly colorChange?: {
    readonly fromHex: string;
    readonly toHex: string;
    readonly transparent: boolean;
  };
  readonly width: Pt;
  readonly height: Pt;
  /** §20.1.8.55 `a:srcRect` — the part of the source the frame shows. */
  readonly crop?: ImageCrop;
  /** §20.1.7.6 `a:xfrm @rot` — the picture's own rotation (1/60000°, clockwise). */
  readonly rotation60k?: number;
  /** §20.1.7.6 `a:xfrm @flipH/@flipV` — the picture drawn mirrored. */
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /**
   * §20.4.2.6 `wp:effectExtent` — space the drawing needs BEYOND its extent,
   * for what a rotation or an effect throws outside the frame. The line box
   * reserves it; the picture itself is drawn inset by it.
   */
  readonly effectExtent?: {
    readonly leftPt: Pt;
    readonly topPt: Pt;
    readonly rightPt: Pt;
    readonly bottomPt: Pt;
  };
}

/**
 * ECMA-376 Part 1 §22 — OfficeMathML (OMML). A recursive math element tree.
 * The union grows by milestone; the layout engine renders structural elements
 * (fraction bars, radicals, big operators, stretchy delimiters) as vector
 * paths and ordinary symbols as font glyphs.
 */
export type MathNode =
  | MathRow
  | MathRun
  | MathFraction
  | MathScript
  | MathRadical
  | MathNary
  | MathFunc
  | MathLimit
  | MathDelimiter
  | MathMatrix
  | MathAccent
  | MathBar
  | MathGroupChr
  | MathEqArray;

/** A sequence of sibling nodes (`m:oMath`, `m:e`, `m:num`, `m:den`, `m:fName` …). */
export interface MathRow {
  readonly type: 'row';
  readonly children: ReadonlyArray<MathNode>;
}

/**
 * `m:r` / `m:t` — literal symbols. `italic`/`bold` come from `m:rPr`/`m:sty`;
 * `nor` (normal text) forces upright. With none set, letters auto-italicise.
 */
export interface MathRun {
  readonly type: 'run';
  readonly text: string;
  readonly italic?: boolean;
  readonly bold?: boolean;
  readonly nor?: boolean;
}

/** `m:f` — fraction. `barless` = `m:fPr/m:type val="noBar"`. */
export interface MathFraction {
  readonly type: 'fraction';
  readonly num: MathNode;
  readonly den: MathNode;
  readonly barless?: boolean;
}

/** `m:sSup` / `m:sSub` / `m:sSubSup` / `m:sPre` (pre-scripts). */
export interface MathScript {
  readonly type: 'script';
  readonly base: MathNode;
  readonly sub?: MathNode;
  readonly sup?: MathNode;
  readonly pre?: boolean; // scripts sit before the base (m:sPre)
}

/** `m:rad` — radical with optional degree (`m:deg`). */
export interface MathRadical {
  readonly type: 'radical';
  readonly radicand: MathNode;
  readonly degree?: MathNode;
}

/** `m:nary` — n-ary operator (∑ ∏ ∫ …) with optional sub/sup limits. */
export interface MathNary {
  readonly type: 'nary';
  readonly op: string;
  readonly body: MathNode;
  readonly sub?: MathNode;
  readonly sup?: MathNode;
  readonly limLoc?: 'undOvr' | 'subSup'; // limits below/above vs at sub/sup
}

/** `m:func` — function application (sin x, …): a name applied to an argument. */
export interface MathFunc {
  readonly type: 'func';
  readonly name: MathNode;
  readonly body: MathNode;
}

/** `m:limLow` / `m:limUpp` — a limit below/above the base. */
export interface MathLimit {
  readonly type: 'limit';
  readonly base: MathNode;
  readonly lim: MathNode;
  readonly pos: 'low' | 'upp';
}

/** `m:d` — delimiters around one or more elements. */
export interface MathDelimiter {
  readonly type: 'delimiter';
  readonly begChr: string;
  readonly endChr: string;
  readonly sepChr?: string;
  readonly children: ReadonlyArray<MathNode>;
}

/** `m:m` — matrix of cells (rows × columns). */
export interface MathMatrix {
  readonly type: 'matrix';
  readonly rows: ReadonlyArray<ReadonlyArray<MathNode>>;
}

/** `m:acc` — accent character over a base (hat, bar, vec, dot, tilde …). */
export interface MathAccent {
  readonly type: 'accent';
  readonly char: string;
  readonly base: MathNode;
}

/** `m:bar` — a bar above or below the base. */
export interface MathBar {
  readonly type: 'bar';
  readonly base: MathNode;
  readonly pos: 'top' | 'bot';
}

/** `m:groupChr` — a grouping character (brace …) above or below the base. */
export interface MathGroupChr {
  readonly type: 'groupChr';
  readonly char: string;
  readonly base: MathNode;
  readonly pos: 'top' | 'bot';
}

/**
 * `m:eqArr` — an equation array: a stack of equations, each on its own line,
 * left-aligned, the block vertically centred on the math axis.
 */
export interface MathEqArray {
  readonly type: 'eqArr';
  readonly rows: ReadonlyArray<MathNode>;
}

/**
 * ECMA-376 Part 1 §17.3.2 — a run: a span of `text` (or an inline image / math
 * object) sharing one set of {@link RunProperties}, plus the reference markers
 * (hyperlink, note, comment, page-number field) the layout acts on.
 */
export interface Run {
  readonly native?: NativeBag;
  readonly text: string;
  readonly properties: RunProperties;
  /**
   * ECMA-376 §17.16.22 — the run sits inside a `w:hyperlink` whose `r:id`
   * resolved to an external target. The URL is stored as written in the rels
   * part; writers MUST pass it through the scheme allowlist (`core/links`)
   * before emitting anything clickable.
   */
  readonly href?: string;
  /**
   * §17.16.5.35 PAGE / §17.16.5.33 NUMPAGES — the run is a page-number field;
   * `text` holds the source's cached result. Header/footer rendering
   * substitutes the real number per page; body rendering keeps the cache.
   */
  readonly field?: 'PAGE' | 'NUMPAGES';
  /**
   * §17.16.22 `w:hyperlink @w:anchor` — internal link target: a bookmark name
   * in this document (never a URL — bypasses the scheme allowlist).
   */
  // §17.11.14 w:footnoteReference / §17.11.6 w:endnoteReference — the run
  // marks a note reference; the layout assigns sequential numbers in reading
  // order and renders them superscript.
  readonly anchor?: string;
  readonly footnoteRef?: string;
  readonly endnoteRef?: string;
  /**
   * §17.13.4.1 `w:commentReference` — the run anchors a review comment by id;
   * the comment's content/author live in `FlowDoc.comments`.
   */
  readonly commentRef?: string;
  /**
   * §17.13.4.3/4 `w:commentRangeStart`/`End` — the ids of the comment ranges
   * this run falls inside (a run may be covered by several). Renderers highlight
   * the covered span; the marker run (`commentRef`) sits at the range's end.
   */
  readonly commentRangeRefs?: ReadonlyArray<string>;
  /**
   * §17.11.13 `w:footnoteRef` / §17.11.5 `w:endnoteRef` — inside note content:
   * render the OWNING note's number here.
   */
  readonly noteNumber?: true;
  /**
   * The run is a list-item marker materialized by `applyNumbering` (`"1."`,
   * `"•"`). Tagged PDF wraps its glyphs in a Lbl structure element (§14.8.4.3.3).
   */
  readonly listMarker?: true;
  /** When set, the run renders this image inline in the line; `text` is ignored. */
  readonly inlineImage?: InlineImage;
  /** When set, the run is an inline OfficeMath object; `text` is ignored. */
  readonly math?: MathNode;
  /**
   * ECMA-376 Part 1 §17.3.3.1 — `w:br w:type="page"`. A forced page break: the
   * paragraph's following content starts on a new page.
   */
  readonly pageBreak?: boolean;
  /**
   * §17.3.3.1 `w:br w:type="column"` — the text after this run continues in the
   * NEXT column of the section (or, past the last, on the next page).
   */
  readonly columnBreak?: boolean;
}

/** ECMA-376 Part 1 §17.3.1 — a paragraph: its {@link ParagraphProperties} and runs. */
export interface Paragraph {
  readonly native?: NativeBag;
  readonly properties: ParagraphProperties;
  readonly runs: ReadonlyArray<Run>;
  /**
   * §17.13.6.2 `w:bookmarkStart` — names of bookmarks opening in (or
   * immediately before) this paragraph. Paragraph-level v1: the destination
   * is the paragraph's first line.
   */
  readonly bookmarks?: ReadonlyArray<string>;
}

/**
 * ECMA-376 Part 1 §17.13.4.2 `w:comment` — a review comment: its block content
 * plus the author/date attribution. Anchored from a run's `commentRef`.
 */
export interface Comment {
  readonly content: ReadonlyArray<BodyElement>;
  readonly author?: string;
  readonly initials?: string;
  /** Raw w:date timestamp (ISO 8601), as authored — not reformatted. */
  readonly date?: string;
  /**
   * The author's resolved identity (usually an email) from word/people.xml
   * (`w15:person` → `w15:presenceInfo/@w15:userId`), matched on the author name.
   * Absent when the file ships no people part or the author is not listed.
   */
  readonly authorId?: string;
  /**
   * Microsoft commentsExtended (w15) — the id of the comment this one replies
   * to, forming a thread. Set only when word/commentsExtended.xml links this
   * comment's paragraph to a parent (w15:paraIdParent). Top-level comments omit it.
   */
  readonly parentId?: string;
  /** w15:done — the comment thread was marked resolved. */
  readonly done?: boolean;
}

/** ECMA-376 Part 1 §17.9 — `w:numFmt` list marker format. */
export type NumberingFormat =
  | 'decimal'
  | 'decimalZero'
  | 'decimalFullWidth'
  | 'ordinal'
  | 'lowerLetter'
  | 'upperLetter'
  | 'lowerRoman'
  | 'upperRoman'
  /** §17.18.59 — the ten heavenly stems 甲乙丙…, then plain digits. */
  | 'ideographTraditional'
  /** The twelve earthly branches 子丑寅…, then plain digits. */
  | 'ideographZodiac'
  /** The formal (anti-fraud) numerals 壹貳參…, composed with 拾佰仟. */
  | 'ideographLegalTraditional'
  /** Digit-by-digit ideographs: 10 is 一零, not 十. */
  | 'ideographDigital'
  | 'koreanDigital2'
  /** Counting ideographs: 10 is 十, 11 is 十一, 21 is 二十一. */
  | 'chineseCounting'
  | 'chineseCountingThousand'
  | 'japaneseCounting'
  | 'koreanCounting'
  | 'taiwaneseCounting'
  | 'taiwaneseCountingThousand'
  /** §17.18.59 — Hebrew numerals (gematria): א, ב, … ט״ו, ט״ז, י״ז … */
  | 'hebrew1'
  /** The Hebrew alphabet as a plain sequence, cycling past ת. */
  | 'hebrew2'
  /** ①②③… up to twenty, then plain digits. */
  | 'decimalEnclosedCircle'
  | 'chicago'
  | 'bullet'
  | 'none';

/** A paragraph's list reference (`w:numPr`): the numbering instance id + level. */
export interface NumberingReference {
  readonly numId: string;
  readonly ilvl: number;
}

/** One level of an abstract numbering definition (`w:lvl`): its format and chrome. */
export interface NumberingLevel {
  readonly ilvl: number;
  readonly start: number;
  readonly format: NumberingFormat;
  /** §17.9.11 `w:lvlText` — the marker template (e.g. `"%1."`). */
  readonly lvlText: string;
  /**
   * §17.9.9 `w:lvlPicBulletId` → §17.9.21 `w:numPicBullet` — the level's bullet
   * is a PICTURE, not a character, and the `w:lvlText` is only the fallback
   * glyph Word writes beside it.
   */
  readonly picBullet?: PictureBullet;
  /**
   * §17.9.10 `w:isLgl` — every level of this level's marker is printed in
   * DECIMAL, whatever format the level it names asks for. Word calls it legal
   * numbering: "Sect I.01" becomes "Sect 1.01".
   */
  readonly isLegal?: boolean;
  readonly paragraphProperties: ParagraphProperties;
  readonly runProperties: RunProperties;
}

/** §17.9.21 `w:numPicBullet` — the image a level uses in place of a bullet. */
export interface PictureBullet {
  /** Content-addressed bytes. A bullet whose picture does not resolve is not one. */
  readonly resource: ResourceId;
  readonly widthPt: Pt;
  readonly heightPt: Pt;
}

/** §17.9.1 `w:abstractNum` — a reusable list definition keyed by level. */
export interface AbstractNumbering {
  readonly id: string;
  readonly levels: ReadonlyMap<number, NumberingLevel>;
}

/** §17.9.18 `w:num` — a concrete list instance bound to an {@link AbstractNumbering}. */
export interface NumberingInstance {
  readonly numId: string;
  readonly abstractNumId: string;
  /**
   * §17.9.27/§17.9.28 `w:lvlOverride/w:startOverride` — where THIS instance
   * starts a level, by `w:ilvl`, whatever the abstract definition says. Absent
   * when the instance takes the abstract starts unchanged.
   */
  readonly startOverrides?: ReadonlyMap<number, number>;
  /**
   * §17.9.27 `w:lvlOverride/w:lvl` — a level this instance REDEFINES whole,
   * shadowing the abstract definition's. NumberingWOverrides.docx rewrites all
   * nine levels of one instance this way.
   */
  readonly levelOverrides?: ReadonlyMap<number, NumberingLevel>;
}

/** The parsed `word/numbering.xml`: abstract definitions + their instances. */
export interface Numbering {
  readonly abstractNums: ReadonlyMap<string, AbstractNumbering>;
  readonly numInstances: ReadonlyMap<string, NumberingInstance>;
}

/** ECMA-376 Part 1 §17.7 — `w:style @w:type` style category. */
export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering';

/**
 * §17.7.6 — one table-style formatting layer: the style's own base layer
 * (`w:style/tblPr` + `tcPr` + `rPr` + `pPr`) or one conditional override
 * (`w:tblStylePr`). Borders come from `tblBorders` (table layer) or `tcBorders`
 * (region layer) — whichever the layer carries.
 */
export interface TableStyleLayer {
  readonly borders?: CellBorders;
  readonly cellMargins?: CellMargins;
  /** §17.4.65 `w:tblInd` — the table indent a whole-table layer declares. */
  readonly indentPt?: Pt;
  readonly shading?: CellShading;
  readonly runProperties?: RunProperties;
  readonly paragraphProperties?: ParagraphProperties;
}

/** §17.7.6.3 `w:tblStylePr @w:type` — the table regions a conditional layer targets. */
export type TableStyleConditionType =
  | 'wholeTable'
  | 'band1Vert'
  | 'band2Vert'
  | 'band1Horz'
  | 'band2Horz'
  | 'firstCol'
  | 'lastCol'
  | 'firstRow'
  | 'lastRow'
  | 'nwCell'
  | 'neCell'
  | 'swCell'
  | 'seCell';

/** A conditional table-style override: the region `type` and the layer it applies. */
export interface TableStyleCondition {
  readonly type: TableStyleConditionType;
  readonly layer: TableStyleLayer;
}

/**
 * ECMA-376 §17.7 — one style definition: its id, category, `basedOn` parent,
 * and the run/paragraph properties it contributes (plus the table-style layers
 * for `type === 'table'`).
 */
export interface Style {
  readonly id: string;
  readonly type: StyleType;
  readonly basedOn?: string;
  readonly isDefault: boolean;
  readonly runProperties: RunProperties;
  readonly paragraphProperties: ParagraphProperties;
  /** Table styles only (§17.7.6): the base layer and conditional overrides. */
  readonly tableLayer?: TableStyleLayer;
  readonly tableConditions?: ReadonlyArray<TableStyleCondition>;
  /** `w:tblPr/w:tblStyleRowBandSize` / `ColBandSize` (default 1). */
  readonly rowBandSize?: number;
  readonly colBandSize?: number;
}

/**
 * The parsed `word/styles.xml`: document defaults (`docDefaults`) plus the
 * styles map keyed by id. The cascade resolver merges these into the resolved
 * property objects the renderer consumes.
 */
export interface StyleSheet {
  readonly defaultRunProperties: RunProperties;
  readonly defaultParagraphProperties: ParagraphProperties;
  readonly styles: ReadonlyMap<string, Style>;
}

// ECMA-376 Part 1 §17.4 — Tables.

/** §17.18.2 ST_Border — a cell-border line style (the subset Ream renders). */
export type BorderStyle =
  | 'none'
  | 'single'
  | 'double'
  | 'thick'
  | 'dotted'
  | 'dashed'
  // …and the same dash over a SHORTER gap, which both references draw as a
  // pattern of its own: `w:val="dashSmallGap"` in a document, a thin `dashed`
  // rule in a spreadsheet.
  | 'dashSmallGap'
  // §18.18.3 — the dash-DOT family. A dash and a dot alternate, which is what
  // tells these apart from `dashed` on the page: cell-borders.xlsx names five
  // of them and we drew a uniform dash for every one.
  | 'dashDot'
  | 'dashDotDot';

/** One border edge: its line style plus optional width and colour. */
export interface Border {
  readonly style: BorderStyle;
  readonly width?: Pt;
  readonly colorHex?: string;
  /** §17.3.1.24 `w:space` — how far a paragraph rule stands off its text. */
  readonly spacePt?: Pt;
}

/** §17.4.39 `w:tcBorders` — the per-edge borders of a cell (or table). */
export interface CellBorders {
  readonly top?: Border;
  readonly right?: Border;
  readonly bottom?: Border;
  readonly left?: Border;
  readonly insideH?: Border;
  readonly insideV?: Border;
  /**
   * Diagonal strokes across the cell box: `diagonalDown` runs top-left →
   * bottom-right, `diagonalUp` bottom-left → top-right (Excel diagonal borders).
   */
  readonly diagonalDown?: Border;
  readonly diagonalUp?: Border;
}

/** §17.4.42 `w:tcMar` — a cell's inner padding per side. */
export interface CellMargins {
  readonly top?: Pt;
  readonly right?: Pt;
  readonly bottom?: Pt;
  readonly left?: Pt;
}

/** §17.4.33 `w:shd` — a cell's solid background fill (6-hex). */
export interface CellShading {
  readonly colorHex: string;
}

/**
 * A conditional-format data bar: a horizontal bar of width `fraction` (`0..1` of
 * the cell), painted over the shading and under the text (E-SHEET SC1c). It
 * starts at `startFraction` from the cell's left (default 0); a mixed-sign range
 * puts the axis inside the cell so negative bars run left of it (tail TC4).
 */
export interface CellDataBar {
  readonly fraction: number;
  readonly colorHex: string;
  readonly startFraction?: number;
  /**
   * The bar runs LEFT from the axis (a negative value in a mixed-sign range),
   * so its solid end is its right one. Excel fades a data bar away from the
   * axis, and which way that is depends on the sign.
   */
  readonly negative?: boolean;
}

/**
 * A conditional-format icon: a small glyph at the cell's left, chosen by the
 * value's bucket (E-SHEET SC1c). Format-neutral — the xlsx layer maps Excel's
 * named icon families (3TrafficLights, 3Arrows, …) onto these shapes + colours.
 */
export type CellIconShape =
  | 'circle'
  | 'square'
  | 'diamond'
  | 'triangleUp'
  | 'triangleDown'
  | 'triangleRight'
  // Symbols families (3Symbols / 3Symbols2): a check / exclamation / cross mark.
  | 'check'
  | 'cross'
  | 'exclamation'
  // Meter families: ratings (a signal-strength bar histogram) and quarters (a
  // clock-fill pie). Both read `fill` for how many units are coloured in.
  | 'bars'
  | 'pie';
/** One resolved conditional-format icon: its {@link CellIconShape}, colour and fill level. */
export interface CellIcon {
  readonly shape: CellIconShape;
  readonly colorHex: string;
  /**
   * Meter glyphs (`bars` / `pie`): how many of `levels` units are filled with
   * `colorHex`; the rest are drawn in a neutral grey. Absent for single glyphs.
   */
  readonly fill?: { readonly filled: number; readonly levels: number };
}

/**
 * A sparkline: a mini chart filling the cell, plotting `values` (E-SHEET SC2).
 * Format-neutral — the xlsx layer resolves the data range to a value sequence;
 * the layout renders line / column / win-loss geometry sized to the cell.
 */
export interface CellSparkline {
  readonly kind: 'line' | 'column' | 'winLoss';
  /**
   * A blank/non-numeric cell in the range is a gap (`null`) so x-positions stay
   * aligned: a line breaks across it, a column/win-loss skips its slot.
   */
  readonly values: ReadonlyArray<number | null>;
  readonly colorHex?: string;
}

/**
 * Resolved position of a cell in a vertical merge group (ECMA-376 §17.4.85
 * `vMerge` markers are resolved by the reader): `start` opens a group that at
 * least one cell continues, `middle` / `end` are continuations; undefined =
 * not merged. Continuation cells stay in their rows (they hold the column
 * slot); writers that need an HTML-style rowSpan can derive it by counting a
 * start's middle/end run downwards.
 */
export type CellMerge = 'start' | 'middle' | 'end';

/** §17.4.30 `w:tcPr` — a cell's properties: span/merge, chrome, and CF overlays. */
export interface CellProperties {
  readonly width?: Pt;
  /**
   * §17.4.72 `w:tcW w:type="pct"` — the cell's preferred width as a share of
   * the table (`w`/5000). Held apart from {@link width} because a percentage
   * cannot be resolved until the table's own width is known.
   */
  readonly widthFraction?: number;
  /** §17.18.90 ST_TblWidth — which of the two above the cell actually declared. */
  readonly widthType?: 'auto' | 'dxa' | 'pct' | 'nil';
  readonly colSpan?: number;
  readonly merge?: CellMerge;
  readonly borders?: CellBorders;
  readonly margins?: CellMargins;
  /**
   * How many of the spanned columns the cell's own PAINT covers, when that is
   * fewer than `colSpan`.
   *
   * Text overflow is modelled as a span — a cell whose value does not fit runs
   * across the empty neighbours to its right — but Excel runs the TEXT over an
   * unpainted cell without painting it. With one width for both, a filled cell
   * carried its fill along: 54436.xlsx ran its pivot header's blue a whole
   * column past the pivot. Absent ⇒ the paint covers the whole span, which is
   * what a merge wants.
   */
  readonly paintColumns?: number;
  readonly shading?: CellShading;
  readonly dataBar?: CellDataBar;
  readonly icon?: CellIcon;
  readonly sparkline?: CellSparkline;
  /**
   * A data-validation `list` cell (E-SHEET SV1): the renderer paints an in-cell
   * dropdown affordance at the cell's right edge (a small button + ▾ glyph).
   */
  readonly dropdown?: boolean;
  /**
   * The cell's text is not allowed to wrap: it renders on one line and whatever
   * does not fit the cell box is cut, as a spreadsheet cell without `wrapText`
   * does. Only the paginated layout honours it — an HTML view has no page edge
   * to clip against and lets the browser decide.
   */
  readonly noWrap?: boolean;
  /**
   * §17.4.20 `w:hideMark` — the cell's end-of-cell mark is not counted when the
   * row's height is measured, so an EMPTY cell that says so adds nothing at all.
   */
  readonly hideMark?: boolean;
  /**
   * The cell holds a NUMBER under a format of its own, so it may not be shown
   * truncated: a date cut to "4/30/201" is not a shorter date, it is the wrong
   * one. Excel and LibreOffice fill such a cell with `#` instead, which says
   * "widen me" and cannot be misread. Text is exempt — a clipped word is still
   * recognisably that word.
   */
  readonly hashOnOverflow?: boolean;
  /**
   * Where the cell's content sits in a box taller than itself. A spreadsheet
   * cell defaults to `'bottom'` — §18.8.1, and both Excel and LibreOffice do it
   * — which is visible on any row taller than its text. Absent ⇒ the top, which
   * is a word-processor table's default.
   */
  readonly verticalAlign?: 'top' | 'center' | 'bottom';
}

/** §17.4.81 `w:trPr` — a table row's properties: height, split/header flags. */
export interface RowProperties {
  readonly height?: Pt;
  readonly heightRule?: 'auto' | 'atLeast' | 'exact';
  readonly cantSplit?: boolean;
  readonly isHeader?: boolean;
  /**
   * §17.4.14 `w:gridBefore` — how many grid columns the row leaves empty
   * before its first cell, which is how a row starts part-way across a table.
   */
  readonly gridBefore?: number;
  /**
   * Force this row to begin a new page (xlsx manual `<rowBreaks>`). The renderer
   * flushes the page before the row, then repeats any leading header rows.
   */
  readonly pageBreakBefore?: boolean;
  /**
   * §17.4.7 `w:cnfStyle` — the conditional formats of the table style this row
   * takes, whatever its position says. Word's calendar templates give a SECOND
   * header row `w:firstRow="1"` so it is painted like the first.
   */
  readonly conditional?: RowConditionalFormat;
}

/** §17.4.7 — the row-level conditional-format flags a `w:cnfStyle` declares. */
export interface RowConditionalFormat {
  readonly firstRow?: boolean;
  readonly lastRow?: boolean;
}

/**
 * §17.4.62 `w:tblLook` — which of the table style's conditional formats apply.
 * Modern files carry explicit attributes; legacy files a hex bitmask (both
 * parsed). Band flags are negative ("no band") per the spec.
 */
export interface TableLook {
  readonly firstRow?: boolean;
  readonly lastRow?: boolean;
  readonly firstColumn?: boolean;
  readonly lastColumn?: boolean;
  readonly noHBand?: boolean;
  readonly noVBand?: boolean;
}

/** §17.4.60 `w:tblPr` — a table's properties: style ref, width/layout, chrome. */
export interface TableProperties {
  /**
   * §17.7.6 — raw reference to a table style (resolved by the reader's
   * `resolveTableStyles` transform; round-trip material afterwards).
   */
  readonly styleId?: string;
  readonly look?: TableLook;
  readonly widthPt?: Pt;
  readonly widthFraction?: number; // tblW type=pct: w/5000 (1.0 = full content width)
  readonly widthType?: 'auto' | 'dxa' | 'pct' | 'nil';
  readonly layout?: 'auto' | 'fixed';
  readonly defaultCellMargins?: CellMargins;
  readonly borders?: CellBorders;
  /**
   * ECMA-376 §17.4.27 (`w:jc`) / xlsx `<printOptions horizontalCentered>`.
   * Centers or right-aligns a table narrower than the content width; absent ⇒ left.
   */
  readonly alignment?: 'left' | 'center' | 'right';
  /**
   * §17.4.65 `w:tblInd` — how far the table's leading edge stands in from the
   * text margin. Distinct from {@link alignment}, which shares out the slack a
   * narrow table leaves.
   */
  readonly indentPt?: Pt;
  /**
   * §17.4.58 `w:tblpPr` — the table FLOATS: it is placed at an anchor of its
   * own and the text runs past it, exactly as an anchored drawing does. Read
   * into the same {@link FloatAnchor} the drawings use.
   */
  readonly float?: FloatAnchor;
  /**
   * A sticky-pane hint from a frozen worksheet view (E-SHEET SE3): the first
   * `rows` rows / `cols` columns stay pinned while the rest scrolls. Consumed
   * only by the HTML writer (an interactive target); PDF/SVG ignore it.
   */
  readonly frozen?: { readonly rows: number; readonly cols: number };
}

/** §17.4.66 `w:tc` — a table cell: its {@link CellProperties} and block content. */
export interface TableCell {
  readonly properties: CellProperties;
  readonly content: ReadonlyArray<BodyElement>;
}

/** §17.4.79 `w:tr` — a table row: its {@link RowProperties} and cells. */
export interface TableRow {
  readonly properties: RowProperties;
  readonly cells: ReadonlyArray<TableCell>;
}

/**
 * §17.4.38 `w:tbl` — a table: its {@link TableProperties}, the `w:tblGrid`
 * column widths (`grid`), and its rows.
 */
export interface Table {
  readonly native?: NativeBag;
  readonly properties: TableProperties;
  readonly grid: ReadonlyArray<Pt>;
  readonly rows: ReadonlyArray<TableRow>;
}

/**
 * ECMA-376 Part 1 §20.4.2.8 — a block-level image (`wp:inline` picture extent).
 * EMU = English Metric Units: 914400 per inch (1 pt = 12700 EMU).
 */
/**
 * §20.1.8.55 `a:srcRect` — how much of each edge of the source picture is cut
 * away before it is fitted to its frame, as a fraction of the source (so 0.25
 * on `left` drops its left quarter). Absent edges are zero.
 */
export interface ImageCrop {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export interface ImageBlock {
  /** §20.4.2.3 — present when the drawing is anchored (floating). */
  readonly float?: FloatAnchor;
  readonly resource?: ResourceId;
  /** The picture's own frame, when it has one (see {@link PictureOutline}). */
  readonly outline?: PictureOutline;
  /**
   * §20.1.8.40 `a:outerShdw` — the drop shadow under the picture. Word writes
   * one on every screenshot pasted with a style; drawn nowhere, imgshadow.docx's
   * six stood flat on the page where both references lift them off it.
   */
  readonly shadow?: ShapeShadow;
  /**
   * §14.1.2.10 `@gain`/`@blacklevel` — the contrast and brightness the picture
   * is drawn through, about mid grey: `out = (in - 0.5) * gain + 0.5 + black`.
   * Word washes a watermark out this way.
   */
  readonly wash?: { readonly gain: number; readonly black: number };
  /**
   * §20.1.8.4 `a:alphaModFix` — how opaque the picture is DRAWN, `0..1`; absent
   * is opaque. A layout that lays a photograph behind its text sets it low
   * (ArtisticEffectSample's cover shows one at 52%), and drawn full-strength
   * the words on it are unreadable.
   */
  readonly alpha?: number;
  /**
   * §20.1.8.16 `a:clrChange` — one colour of the picture replaced by another,
   * or knocked out entirely when the destination states `a:alpha` at zero.
   */
  readonly colorChange?: {
    readonly fromHex: string;
    readonly toHex: string;
    readonly transparent: boolean;
  };
  readonly width: Pt;
  readonly height: Pt;
  /** §20.1.8.55 `a:srcRect` — the part of the source the frame shows. */
  readonly crop?: ImageCrop;
  /** §20.1.7.6 `a:xfrm @rot` — the picture's own rotation (1/60000°, clockwise). */
  readonly rotation60k?: number;
  /** §20.1.7.6 `a:xfrm @flipH/@flipV` — the picture drawn mirrored. */
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /** `wp14:sizeRelH/V` — a size stated as a share of the page or margins. */
  readonly relativeSize?: RelativeSize;
  /** §20.4.2.6 `wp:effectExtent` — space reserved around the picture (see {@link InlineImage}). */
  readonly effectExtent?: InlineImage['effectExtent'];
  readonly paragraphProperties: ParagraphProperties;
  /** `wp:docPr @descr/@title` — alternate text for the tagged-PDF Figure (`/Alt`). */
  readonly altText?: string;
}

// ECMA-376 Part 1 §20 — DrawingML shapes (wps:wsp inside a w:drawing).
// A standalone shape (a paragraph whose only content is a shape drawing)
// collapses to a ShapeBlock, mirroring ImageBlock; it carries the paragraph's
// properties for block spacing / alignment.

/**
 * §20.1.10.55 prstGeom path command (custGeom `<a:pathLst>`), in path-space
 * units (`a:path @w/@h`). `quad` is elevated to cubic at layout; `arc` uses the
 * DrawingML angle convention (1/60000°, clockwise, y-down) and is converted
 * to beziers by the geometry layer.
 */
export type CustomPathCmd =
  | { readonly cmd: 'move'; readonly x: number; readonly y: number }
  | { readonly cmd: 'line'; readonly x: number; readonly y: number }
  | {
      readonly cmd: 'cubic';
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly x: number;
      readonly y: number;
    }
  | {
      readonly cmd: 'quad';
      readonly x1: number;
      readonly y1: number;
      readonly x: number;
      readonly y: number;
    }
  | {
      readonly cmd: 'arc';
      readonly wR: number;
      readonly hR: number;
      readonly stAng: number;
      readonly swAng: number;
    }
  | { readonly cmd: 'close' };

/**
 * §20.1.9.11 `custGeom` — a custom geometry: its path-space extent
 * (`pathWidth`/`pathHeight`) and the {@link CustomPathCmd} list that draws it.
 */
export interface CustomGeometry {
  readonly pathWidth: number;
  readonly pathHeight: number;
  readonly commands: ReadonlyArray<CustomPathCmd>;
}

/** §20.1.10.55 `prstGeom` / §20.1.9.11 `custGeom` — a shape's geometry source. */
export interface ShapeGeometry {
  readonly kind: 'preset' | 'custom';
  readonly preset?: string; // kind==='preset', e.g. 'roundRect'
  readonly adjust?: ReadonlyMap<string, number>; // a:gd name → raw val (usually 0..100000)
  readonly custom?: CustomGeometry; // kind==='custom'
}

/** A shape's fill mode (`a:noFill`/`a:solidFill`/`a:gradFill`/`a:blipFill`). */
export type ShapeFillKind = 'none' | 'solid' | 'gradient' | 'picture';

/** A shape's fill: none, a solid colour, a {@link ShapeGradient}, or a picture. */
export interface ShapeFill {
  readonly kind: ShapeFillKind;
  readonly colorHex?: string; // resolved 6-hex (kind==='solid')
  readonly gradient?: ShapeGradient; // kind==='gradient' (a:gradFill, EP16)
  /**
   * §14.1.2.5 `@type="tile"` — the picture REPEATS at its own size rather than
   * stretching over the box. A parchment behind a page, a texture behind a
   * text box: stretched, such a fill is a blur where it should be a pattern.
   */
  readonly tiled?: boolean;
  /**
   * The size ONE copy of a tiled fill occupies, when the shape states it —
   * MS-ODRAW §2.3.7.11/.12 `fillWidth` / `fillHeight`. Absent, a tile is the
   * picture at its own resolution, which is the default both formats give.
   */
  readonly tileSizePt?: { readonly widthPt: number; readonly heightPt: number };
  /**
   * §20.1.8.58 `a:tile @sx @sy` — how far the picture is scaled BEFORE it is
   * repeated, as fractions of its own size (1 = 100 %). A texture halved tiles
   * four times as often, and read as a plain repeat it tiles once.
   */
  readonly tileScale?: { readonly sx: number; readonly sy: number };
  /**
   * §20.1.8.14 `a:blipFill` — the picture painted across the shape's box. A
   * DrawingML picture IS a shape with one of these, which is how a `pic:pic`
   * inside a group reaches the page.
   */
  readonly imageResource?: ResourceId;
  /**
   * §20.1.8.30 `a:stretch/a:fillRect` (or an `a:srcRect` beside it) — the part
   * of the picture the box shows, as the fractions cut from each side.
   */
  readonly imageCrop?: ImageCrop;
  /**
   * §20.1.8.30 `a:stretch/a:fillRect` with POSITIVE insets — the part of the
   * shape's box the picture is stretched into, as fractions of that box. A
   * slide backed by a picture inset 55 % from the left shows it in the corner,
   * not across the slide (corpus: tdf153466.pptx). Negative insets are the
   * other case entirely — the picture zooming past the box — and read as a
   * crop of the source.
   */
  /**
   * §20.1.8.23 `a:duotone` — the picture recoloured into two tones: its dark
   * end becomes the first colour, its light end the second. An Office theme
   * that ships a photograph tints it this way, so a deck whose background is a
   * brown ridged texture is stored as a grey one (corpus: themes.pptx).
   */
  readonly duotone?: { readonly shadowHex: string; readonly highlightHex: string };
  readonly imageFillRect?: {
    readonly left: number;
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
  };
  /**
   * §20.1.2.3.1 `a:alpha` / §14.1.2.5 `@opacity` — how opaque the fill is,
   * `0..1`. Absent is opaque. The colour above is the fill's own, NOT composited
   * over the paper: what is behind the shape shows through it.
   */
  readonly alpha?: number;
}

/** §20.1.10.49 ST_PresetLineDashVal — a shape outline's preset dash pattern. */
export type ShapeDash =
  | 'solid'
  | 'dot'
  | 'dash'
  | 'dashDot'
  | 'lgDash'
  | 'lgDashDot'
  | 'sysDash'
  | 'sysDot';

/** §20.1.2.2.24 `a:ln` — a shape's outline (stroke). */
export interface ShapeLine {
  readonly width?: Pt; // a:ln @w; default 0.75pt
  readonly colorHex?: string; // resolved 6-hex
  readonly dash?: ShapeDash; // a:prstDash @val
  /**
   * §20.1.8.21 `a:custDash` — the author's own pattern: dash/space lengths as
   * MULTIPLES of the line width, in the order they are drawn.
   */
  readonly customDash?: ReadonlyArray<number>;
  readonly cap?: 'flat' | 'round' | 'square'; // a:ln @cap (flat=butt)
  readonly fill?: 'solid' | 'none'; // a:ln/a:noFill ⇒ no visible stroke
  /** §20.1.8.24 `a:headEnd` — the decoration at the line's first point. */
  readonly headEnd?: LineEnd;
  /** §20.1.8.42 `a:tailEnd` — the decoration at the line's last point. */
  readonly tailEnd?: LineEnd;
}

/**
 * §20.1.8.24 / §20.1.8.42 — an arrowhead (or other decoration) at one end of a
 * line: its shape plus the width and length steps `ST_LineEndWidth` /
 * `ST_LineEndLength` name.
 */
export interface LineEnd {
  readonly type: 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
  readonly width?: 'sm' | 'med' | 'lg';
  readonly length?: 'sm' | 'med' | 'lg';
}

/**
 * `wp14:sizeRelH` / `wp14:sizeRelV` — the drawing's size as a PERCENTAGE of
 * the page or the margins rather than the extent beside it. Word 2010 writes
 * it in the `wp14` namespace, with the extent as the fallback for readers that
 * do not know it.
 */
export interface RelativeSize {
  readonly widthPct?: number; // 0..1
  /**
   * §20.4.3.6 ST_SizeRelFromH — what the percentage is OF. `margin` is the text
   * area; the rest are the bands around it, and a drawing sized against one is
   * a fraction of that band alone (tdf123324 asks for 150% of the top margin).
   */
  readonly widthFrom?: 'margin' | 'page' | 'leftMargin' | 'rightMargin';
  readonly heightPct?: number;
  /** §20.4.3.7 ST_SizeRelFromV — the vertical twin. */
  readonly heightFrom?: 'margin' | 'page' | 'topMargin' | 'bottomMargin';
}

/** §20.1.7.6 `a:xfrm` — a shape's rotation (1/60000°, clockwise) + flips. */
export interface ShapeTransform {
  readonly rotation60k?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
}

/** `wps:txbx` (`w:txbxContent`) + `wps:bodyPr` — text inside a shape. */
export interface ShapeTextBody {
  readonly content: ReadonlyArray<BodyElement>;
  readonly insetLeft?: Pt;
  readonly insetTop?: Pt;
  readonly insetRight?: Pt;
  readonly insetBottom?: Pt;
  readonly anchor?: 't' | 'ctr' | 'b'; // vertical anchor
  /**
   * §20.1.10.83 ST_TextVerticalType (`a:bodyPr @vert`) — text set along the
   * box's long axis rather than across it. `vert` reads top-to-bottom (turned a
   * quarter clockwise), `vert270` bottom-to-top.
   */
  readonly vertical?: 'vert' | 'vert270';
  /**
   * §20.1.10.28 `a:spAutoFit` — the SHAPE follows its text: its height is
   * whatever the text needs, whatever the stated box says.
   */
  readonly autoFit?: boolean;
  /**
   * §14.1.2.22 `v:textpath @fitshape` — the TEXT follows the shape: legacy
   * WordArt is set at whatever size fills the box it was drawn in.
   */
  readonly fitToBox?: boolean;
  /**
   * `wps:txbx @id` / `wps:linkedTxbx @id @seq` — the chain of boxes this one
   * belongs to. Text that overruns a box continues in the next of its chain;
   * `seq` 0 is the box that holds the words, and the rest carry none of their
   * own.
   */
  readonly chain?: { readonly id: string; readonly seq: number };
}

/**
 * A block-level DrawingML shape (§20.1): its size, {@link ShapeGeometry}, fill,
 * outline, transform and optional text body. A standalone shape carries the
 * paragraph's properties for block spacing / alignment, mirroring
 * {@link ImageBlock}.
 */
/**
 * §20.1.8.40 `a:outerShdw` — the drop shadow under a shape: how far it is
 * displaced, how soft its edge is, and in what colour. `dxPt`/`dyPt` are the
 * displacement in page coordinates (y grows DOWN), resolved from the spec's
 * polar `dist`/`dir`.
 */
export interface ShapeShadow {
  readonly dxPt: number;
  readonly dyPt: number;
  /** `blurRad` in points; 0 is a hard edge. */
  readonly blurPt: number;
  readonly colorHex: string;
  /** 0..1, from the shadow colour's `a:alpha` (absent ⇒ opaque). */
  readonly alpha: number;
}

/**
 * §20.5.2.17 `wpg:wgp` — one member of a drawing group, and where it sits: the
 * offsets are from the group's own top-left corner, already mapped out of the
 * group's child coordinate space.
 */
export interface ShapeGroupChild {
  readonly shape: ShapeBlock;
  readonly xPt: Pt;
  readonly yPt: Pt;
}

export interface ShapeBlock {
  /** §20.4.2.3 — present when the drawing is anchored (floating). */
  readonly float?: FloatAnchor;
  readonly width: Pt; // wp:extent cx (fallback a:ext cx)
  readonly height: Pt; // wp:extent cy
  /** §20.5.2.17 — the shapes a group holds, drawn inside its own box. */
  readonly children?: ReadonlyArray<ShapeGroupChild>;
  readonly geometry: ShapeGeometry;
  readonly fill: ShapeFill;
  readonly line?: ShapeLine;
  readonly transform?: ShapeTransform;
  /** `wp14:sizeRelH/V` — a size stated as a share of the page or margins. */
  readonly relativeSize?: RelativeSize;
  readonly text?: ShapeTextBody;
  /** §20.1.8.40 — the shape's drop shadow, direct or from its style reference. */
  readonly shadow?: ShapeShadow;
  readonly paragraphProperties: ParagraphProperties;
  /** `wp:docPr @descr/@title` — alternate text for the tagged-PDF Figure (`/Alt`). */
  readonly altText?: string;
}

/**
 * ECMA-376 Part 1 §21.2 — the kind of a DrawingML chart. A chart is referenced
 * from a `w:drawing` (`a:graphicData uri=…/chart` → `c:chart r:id`) and its data
 * lives in a separate chart part (e.g. `word/charts/chart1.xml`). The parsed
 * {@link Chart} is keyed by that relationship id and supplied to the renderer
 * alongside the body (mirroring image bytes), so {@link ChartBlock} only carries
 * the ref.
 */
export type ChartType = 'bar' | 'line' | 'pie' | 'area' | 'scatter' | 'unknown';

/** A per-point colour override (`c:dPt`) at series index `idx`. */
export interface ChartDataPoint {
  readonly idx: number;
  readonly colorHex: string; // c:dPt per-point colour override
}

/** One chart data series (`c:ser`): its name, values and colour overrides. */
export interface ChartSeries {
  readonly name?: string;
  /**
   * §21.2.2 — the chart group this series came from, when it is NOT the group
   * that gave {@link Chart.type}. A `c:plotArea` may hold several groups: a
   * combo chart writes `c:barChart` and `c:lineChart` side by side, and reading
   * only the first drops the other's series off the page entirely (57362.xlsx
   * loses its line). Absent ⇒ the series draws as the chart's own type.
   */
  readonly type?: ChartType;
  /**
   * §21.2.2.9 `c:axId` — the series plots against the SECONDARY value axis: its
   * group names a `c:valAx` other than the one the first group uses. 57362.xlsx
   * puts its line on an axis of its own at `axPos="r"`, and plotting it on the
   * primary scale puts every point at the wrong height whenever the two ranges
   * differ.
   */
  readonly secondaryAxis?: boolean;
  readonly values: ReadonlyArray<number>; // c:val/c:yVal numCache (idx-ordered, gaps → 0)
  readonly xValues?: ReadonlyArray<number>; // c:xVal numCache (scatter — paired with values)
  readonly colorHex?: string; // c:spPr solidFill
  readonly pointColors?: ReadonlyArray<ChartDataPoint>; // c:dPt overrides (pie slices)
  /**
   * §21.2.2.49 `c:dLbl/c:tx` — a data label the author typed rather than one
   * the chart computes, by point index. Excel and Calc print it verbatim: it
   * is the only place a label like "Промышленные потребители; 22,7млрд.кВтч;
   * 67,3%" exists, and generating one from the value instead loses the whole
   * sentence.
   */
  readonly pointLabels?: ReadonlyArray<{ readonly idx: number; readonly text: string }>;
  /**
   * §21.2.2.59 `c:val/c:numRef/c:f` and §21.2.2.215 `c:tx/c:strRef/c:f` — where
   * the series reads its numbers and its name FROM, when the chart part carries
   * no cache of them. A chart written without caches is not a chart without
   * data: the reader resolves these against the workbook.
   */
  readonly valuesRef?: string;
  readonly nameRef?: string;
  /**
   * §21.2.2.106 `c:marker` — the symbol this series stamps on each of its
   * points. Absent ⇒ the reader's own default.
   */
  readonly marker?: ChartMarker;
  /**
   * §21.2.2.198 `c:ser/c:spPr/a:ln` — the series' OWN rule. `<a:noFill/>` is
   * how Excel writes "scatter with markers only": the group still says
   * `lineMarker`, and it is this line that says the points are not joined.
   * SimpleScatterChart.xlsx is two loose dots in both references and we ran a
   * line between them.
   */
  readonly line?: ChartLineStyle;
}

/**
 * §21.2.2.107 `c:symbol` (ST_MarkerStyle) — the shape of a scatter/line
 * series' point marker. `auto` and `picture` are left to the reader, so they
 * are not carried here.
 */
export type ChartMarkerSymbol =
  | 'circle'
  | 'dash'
  | 'diamond'
  | 'dot'
  | 'none'
  | 'plus'
  | 'square'
  | 'star'
  | 'triangle'
  | 'x';

/**
 * §21.2.2.106 `c:marker` — a series' point symbol and its size in points
 * (§21.2.2.153 `c:size`, 2–72). chartTitle_noTitle.xlsx asks for a 5pt circle
 * at every point and both references draw one; we stamped a square.
 */
export interface ChartMarker {
  readonly symbol: ChartMarkerSymbol;
  readonly sizePt?: number;
}

/**
 * §21.2.2.196 `c:spPr/a:ln` on an axis or its gridlines — the rule the author
 * gave it. `none` is `<a:ln><a:noFill/>`, which means the axis draws no line at
 * all: 57362.xlsx hides its secondary value axis that way and keeps its labels.
 */
export interface ChartLineStyle {
  readonly none?: boolean;
  readonly colorHex?: string;
  readonly widthPt?: number;
  /** §20.1.10.49 `a:prstDash` — the rule's dash pattern, when it names one. */
  readonly dash?: ShapeDash;
}

/** A parsed chart (§21.2): its type, title, categories, series and rendering options. */
export interface Chart {
  readonly type: ChartType;
  readonly title?: string;
  readonly categories: ReadonlyArray<string>; // c:cat (shared across series)
  /** §21.2.2.24 `c:cat/…/c:f` — where the categories live, when uncached. */
  readonly categoriesRef?: string;
  readonly series: ReadonlyArray<ChartSeries>;
  readonly hasLegend: boolean;
  readonly legendPos?: 'r' | 'l' | 't' | 'b';
  readonly barDir?: 'col' | 'bar'; // c:barDir (bar charts)
  readonly grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  readonly doughnut?: boolean; // c:doughnutChart (a pie with a central hole)
  readonly showValues?: boolean; // c:dLbls/c:showVal — print each datum's value
  /**
   * §21.2.2.75 `c:gapWidth` — the gap between category slots as a percentage
   * of the bar width. Absent ⇒ the schema's 150.
   */
  readonly gapPercent?: number;
  readonly catAxisTitle?: string; // c:catAx/c:title
  /**
   * §21.2.2.134 `c:catAx/c:scaling/c:orientation` = `maxMin` — the category
   * axis runs the other way. A ranked bar chart is written this way so its
   * first row reads at the TOP (dataValidationTableRange.xlsx ranks 38 counties
   * and we printed the ranking upside down).
   */
  readonly catAxisReversed?: boolean;
  readonly valAxisTitle?: string; // c:valAx/c:title
  /** §21.2.2.168 — the title of the secondary value axis, when one is drawn. */
  readonly secondaryValAxisTitle?: string;
  /**
   * §21.2.2.161 `c:scatterStyle` — whether a scatter's points are joined by a
   * line, marked, or both. The schema's default is `marker`; Excel writes
   * `lineMarker` for the chart most people insert.
   */
  readonly scatterStyle?: 'none' | 'line' | 'lineMarker' | 'marker' | 'smooth' | 'smoothMarker';
  /**
   * §21.2.2.106 `c:marker` on a `c:lineChart` — whether the group's series mark
   * their data points. WithChart.xlsx asks for them and we drew bare lines.
   */
  readonly lineMarkers?: boolean;
  /** §21.2.2.196 — the category axis's own rule. */
  readonly catAxisLine?: ChartLineStyle;
  /** §21.2.2.196 — the value axis's own rule. */
  readonly valAxisLine?: ChartLineStyle;
  /** §21.2.2.196 — the secondary value axis's own rule. */
  readonly secondaryValAxisLine?: ChartLineStyle;
  /** §21.2.2.87 `c:majorGridlines/c:spPr/a:ln` — the gridlines' own rule. */
  readonly gridLine?: ChartLineStyle;
  /**
   * §21.2.2.157 `c:valAx/c:scaling/c:min|c:max` — the value axis the AUTHOR
   * fixed. Absent means "auto", and only then is the range read off the data:
   * a chart whose cells all read zero still has the axis its author pinned, and
   * scaling it to the data drew 0…1 where every reader draws 0…300.
   */
  readonly valAxisMin?: number;
  readonly valAxisMax?: number;
  /**
   * §21.2.2.198 `c:chartSpace/c:spPr` — the frame around the whole chart: its
   * background fill and its outline. Excel writes both on every chart it
   * creates, and both references draw them.
   */
  readonly frameFillHex?: string;
  /**
   * §20.1.8.14 `a:blipFill` — the PICTURE the frame is filled with, when it is
   * filled with one rather than a colour. chart-texture-bg.pptx papers its
   * whole chart with a woven cloth, and a chart on white is a different chart.
   */
  readonly frameFillImage?: { readonly resource: ResourceId; readonly tiled?: boolean };
  readonly frameLineHex?: string;
  /** The frame rule's width and dash, when it states them. */
  readonly frameLineWidthPt?: number;
  readonly frameLineDash?: ShapeDash;
  /**
   * §21.2.2.145 `c:plotArea/c:spPr` — the plot rectangle's own fill and rule,
   * which sit inside the frame. Chart_Plot_BorderLine_Style.docx rules its plot
   * in a heavy orange dash-dot.
   */
  readonly plotFillHex?: string;
  readonly plotLine?: ChartLineStyle;
  /**
   * §21.2.2.121 `c:valAx/c:numFmt@formatCode` — the number format the value
   * axis's tick labels and the data labels are drawn in, in the same code
   * grammar cells use (§18.8.31). Absent ⇒ a plain numeric render.
   */
  readonly numberFormat?: string;
  /**
   * MS-ODRAWXML `chartColorStyle` (`colorsN.xml`): the cycle of series colours;
   * overrides the built-in Office accent cycle when present.
   */
  readonly seriesColorCycle?: ReadonlyArray<string>;
}

/** A block-level chart reference (`c:chart`): its size and relationship id. */
export interface ChartBlock {
  /** §20.4.2.3 — present when the drawing is anchored (floating). */
  readonly float?: FloatAnchor;
  readonly chartRelId: string; // c:chart @r:id (resolve against the document's rels)
  readonly width: Pt;
  readonly height: Pt;
  readonly paragraphProperties: ParagraphProperties;
  /** `wp:docPr @descr/@title` — alternate text for the tagged-PDF Figure (`/Alt`). */
  readonly altText?: string;
}

/** One top-level body item: a discriminated union over the block kinds. */
export type BodyElement =
  | { readonly kind: 'paragraph'; readonly paragraph: Paragraph }
  | { readonly kind: 'table'; readonly table: Table }
  | { readonly kind: 'image'; readonly image: ImageBlock }
  | { readonly kind: 'shape'; readonly shape: ShapeBlock }
  | { readonly kind: 'chart'; readonly chart: ChartBlock };

/** ECMA-376 Part 1 §17.6.13 `w:pgSz` — the section's page dimensions + orientation. */
export interface PageSize {
  readonly width: Pt;
  readonly height: Pt;
  readonly orientation?: 'portrait' | 'landscape';
}

/** §17.6.11 `w:pgMar` — the section's page margins (and header/footer offsets). */
export interface PageMargins {
  readonly top: Pt;
  readonly right: Pt;
  readonly bottom: Pt;
  readonly left: Pt;
  readonly header?: Pt;
  readonly footer?: Pt;
  /**
   * §17.6.11 `w:gutter` — the binding space, added to the left margin (or to
   * the top, when `w:settings/w:gutterAtTop` says so).
   */
  readonly gutter?: Pt;
}

/** §17.10 — which page class a header/footer reference applies to. */
export type HeaderFooterType = 'default' | 'first' | 'even';

/** §17.10.5/§17.10.2 `w:headerReference`/`w:footerReference` — a typed rels pointer. */
export interface HeaderFooterReference {
  readonly type: HeaderFooterType;
  readonly relationshipId: string;
}

/** §17.6.17 `w:sectPr` — a section's page setup, header/footer refs and columns. */
export interface SectionProperties {
  readonly pageSize?: PageSize;
  readonly margins?: PageMargins;
  readonly headers: ReadonlyArray<HeaderFooterReference>;
  readonly footers: ReadonlyArray<HeaderFooterReference>;
  /**
   * ECMA-376 §17.10.6 — `w:titlePg` toggle in `sectPr`. When true the first page
   * of the section uses the `first` header/footer references.
   */
  readonly titlePg?: boolean;
  /**
   * §17.6.12 `w:pgNumType w:start` — the number the section's first page is
   * printed with. Absent ⇒ the count carries on from the section before.
   */
  readonly pageNumberStart?: number;
  /** §17.6.8 `w:lnNumType` — line numbers printed in the margin beside the text. */
  readonly lineNumbering?: {
    /** Print every `countBy`-th line (default 1). */
    readonly countBy: number;
    /** The number the count starts at (default 1). */
    readonly start: number;
    /** How far the number stands off the text (`w:distance`); absent ⇒ Word's quarter inch. */
    readonly distancePt?: Pt;
    /** §17.18.55 — where the count restarts. */
    readonly restart: 'newPage' | 'newSection' | 'continuous';
  };
  /**
   * ECMA-376 §17.15.1.36 — `w:evenAndOddHeaders` toggle in `word/settings.xml`
   * (document-wide, not per-section). When true even-numbered pages use the
   * `even` header/footer references.
   */
  readonly evenAndOddHeaders?: boolean;
  /** §17.6.4 `w:cols` — multi-column section layout. */
  readonly columns?: SectionColumns;
  /**
   * §17.6.10 `w:pgBorders` — the rules drawn around the page. `offsetFrom`
   * says what each edge's `spacePt` is measured from: the paper's edge, or the
   * text margin it stands outside of.
   */
  readonly pageBorders?: {
    readonly borders: CellBorders;
    readonly offsetFrom: 'page' | 'text';
  };
  /**
   * §17.6.22 `w:type` — where the section starts. `continuous` starts it on the
   * page already in hand rather than a fresh one; everything else (nextPage,
   * and the odd/even/column variants we do not distinguish) starts a page.
   */
  readonly sectionStart?: 'continuous' | 'nextPage';
  /**
   * §17.6.5 `w:docGrid` — the line grid a `lines`/`linesAndChars` section rules
   * its text onto, as the pitch in points. Every line of the section's text is
   * as tall as a whole number of these, however tall its own font makes it.
   * Absent ⇒ no grid (`w:type="default"`, or none stated).
   */
  readonly gridLinePitchPt?: Pt;
}

/**
 * §20.4.2.3 `wp:anchor` — a floating drawing's placement. v1 honours
 * out-of-flow placement for wrap `'none'` (incl. `behindDoc`); the side-wrapping
 * modes (square/tight/through) and `topAndBottom` stay in flow as blocks.
 */
export interface FloatAnchor {
  /**
   * §20.4.2.3 / §17.18.104 — how body text runs past the drawing. `notBeside`
   * is the FRAME's own mode: the drawing keeps the place its anchor names, but
   * no text may stand beside it, so its band spans the whole column.
   */
  readonly wrap: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom' | 'notBeside';
  readonly behind?: boolean; // wp:anchor @behindDoc
  /**
   * §20.4.2.3 `wp:anchor @relativeHeight` — the z-order among the floats on the
   * page: the higher number is drawn over the lower, whatever their order in
   * the document.
   */
  readonly zOrder?: number;
  readonly posH?: {
    /**
     * §20.4.3.3 ST_RelFromH. `leftMargin`/`rightMargin` are the margin BANDS
     * beside the text area, which is where a marginal note is placed:
     * tdf103573.docx centres one box in each, and read as plain `margin` they
     * both landed in the middle of the text and printed over each other.
     */
    readonly relativeFrom: 'margin' | 'page' | 'column' | 'leftMargin' | 'rightMargin';
    readonly offsetPt?: Pt; // wp:posOffset
    readonly align?: 'left' | 'center' | 'right'; // wp:align
  };
  readonly posV?: {
    /**
     * §20.4.3.4 ST_RelFromV. `topMargin`/`bottomMargin` measure from the top
     * edge of the margin BAND they name — page-content-bottom.docx hangs a
     * square 312pt above the bottom margin, and read as plain `margin` it went
     * off the top of the page.
     */
    readonly relativeFrom: 'margin' | 'page' | 'paragraph' | 'line' | 'topMargin' | 'bottomMargin';
    readonly offsetPt?: Pt;
    /**
     * §20.4.3.1 `wp:align` (VML: `mso-position-vertical`) — a KEYWORD rather
     * than an offset, which is how Word centres a watermark in its page.
     */
    readonly align?: 'top' | 'center' | 'bottom';
  };
  /**
   * §20.4.2.3 `@wrapText` on the wrap element — which side(s) of the drawing
   * text may stand on. `bothSides` is the default and the only value that
   * fills BOTH gaps; `left`/`right` name one, `largest` the wider.
   */
  readonly wrapSide?: 'bothSides' | 'left' | 'right' | 'largest';
  /**
   * §20.4.2.3 `wp:anchor @layoutInCell` (VML `o:allowincell`) — present and
   * false when a drawing anchored inside a TABLE CELL is placed against the
   * page rather than against the cell. Absent means the default: the cell is
   * the frame every `relativeFrom` is measured in.
   */
  readonly inCell?: boolean;
  /**
   * §20.4.2.3 `wp:anchor @distT/@distB/@distL/@distR` — how far the wrapped
   * text stands off each edge of the drawing. Absent sides are 0.
   */
  readonly wrapDist?: {
    readonly topPt: Pt;
    readonly bottomPt: Pt;
    readonly leftPt: Pt;
    readonly rightPt: Pt;
  };
}

/**
 * §17.6.4 — column definitions: an equal-width `count` + gutter, or explicit
 * per-column widths/gutters (`w:col` children).
 */
export interface SectionColumns {
  readonly count: number;
  readonly spacePt: number;
  readonly explicit?: ReadonlyArray<{ readonly widthPt: number; readonly spacePt: number }>;
  /** §17.6.4 `w:sep` — a vertical rule drawn down the middle of every gutter. */
  readonly separator?: boolean;
}

/**
 * One section descriptor for a multi-section document (ECMA-376 §17.6.17). Each
 * section's `sectPr` lives either inside a paragraph's `pPr` (mid-document
 * break) or as the final child of `w:body` (final section). Records the
 * properties and the upper-exclusive bound in `BodyElement[]` the section
 * covers — section i applies to
 * `body[sections[i-1].endIndex..sections[i].endIndex)`.
 */
export interface Section {
  readonly properties: SectionProperties;
  readonly endIndex: number;
}

/** The parsed WordprocessingML document: body, stylesheet, numbering and section setup. */
export interface DocumentModel {
  readonly body: ReadonlyArray<BodyElement>;
  readonly styleSheet: StyleSheet;
  readonly numbering?: Numbering;
  readonly section?: SectionProperties;
}

/**
 * Document metadata (PDF `/Info`-shaped, sourced from `docProps/core.xml` and/or
 * caller options). Lives in the model so FlowDoc can carry it format-neutrally.
 */
export interface DocumentInfo {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string;
  /** Producer; defaults to "Ream" if not provided. */
  readonly producer?: string;
  /** ISO 8601 date; converted to PDF date format (`D:YYYYMMDDHHmmSS`). */
  readonly creationDate?: Date;
  readonly modificationDate?: Date;
}
