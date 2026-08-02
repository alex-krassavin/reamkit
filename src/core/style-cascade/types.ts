// Fully-resolved properties consumed by the renderer.
// Every field is required — the cascade has already collapsed all undefined
// values against document defaults and style inheritance.

import type {
  Alignment,
  CellBorders,
  CellShading,
  FontFamilyMap,
  FrameProperties,
  NumberingReference,
  RunProperties,
  TabStop,
  UnderlineStyle,
  VerticalAlign,
} from '@/core/document-model';
import type { Pt } from '@/core/ir';
import { halfPtToPt, twipsToPt } from '@/core/ir';

/**
 * A run's fully-resolved properties: every field required because the cascade
 * has collapsed all inheritance against document defaults and the style chain.
 * Consumed directly by the renderer.
 */
export interface ResolvedRunProperties {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: UnderlineStyle;
  /** §17.3.2.40 `w:u @w:color` — the underline's own colour; absent means the text's. */
  readonly underlineColorHex?: string;
  readonly strike: boolean;
  /** §17.3.2.5 `w:caps` — displayed in capitals. */
  readonly caps: boolean;
  /** §17.3.2.33 `w:smallCaps` — displayed in capitals, the lower case smaller. */
  readonly smallCaps: boolean;
  readonly fontSizePt: Pt;
  readonly colorHex: string;
  readonly fontFamily: FontFamilyMap;
  readonly verticalAlign: VerticalAlign;
  readonly rtl: boolean;
  /**
   * Natural language (`w:lang @w:val`). Optional — only set when the source
   * specifies one; consumed by the tagged-PDF per-element `/Lang`, never layout.
   */
  readonly lang?: string;
  /** §17.3.2.32 — the background painted behind the run's glyphs. */
  readonly shadingColorHex?: string;
  /** §17.3.2.35 — extra space between the run's characters, in points. */
  readonly letterSpacingPt?: Pt;
}

/**
 * A paragraph's fully-resolved properties: every field required, the cascade
 * having collapsed all inheritance. Consumed directly by the renderer.
 */
export interface ResolvedParagraphProperties {
  readonly alignment: Alignment;
  readonly spacingBefore: Pt;
  readonly spacingAfter: Pt;
  readonly spacingLine: Pt;
  readonly spacingLineRule: 'auto' | 'exact' | 'atLeast';
  readonly indentLeft: Pt;
  readonly indentRight: Pt;
  readonly indentFirstLine: Pt;
  readonly pageBreakBefore: boolean;
  /** §17.3.1.9 — drop the space between this paragraph and a same-styled neighbour. */
  readonly contextualSpacing: boolean;
  /** §17.3.1.37 — the paragraph's tab stops, in ascending position order. */
  readonly tabs: ReadonlyArray<TabStop>;
  /** §17.3.1.24 `w:pBdr` — rules drawn around the paragraph. */
  readonly borders?: CellBorders;
  /** §17.3.1.31 `w:pPr/w:shd` — the paragraph's own background fill. */
  readonly shading?: CellShading;
  readonly bidi: boolean;
  /**
   * ECMA-376 §17.3.1.20 — resolved outline level (0–8). Undefined = body text.
   * Used only by tagged-PDF heading detection.
   */
  readonly outlineLevel?: number;
  /**
   * The paragraph's style id (e.g. `"Heading2"`), carried through for heading
   * detection when a style lacks an explicit outline level.
   */
  readonly styleId?: string;
  /**
   * §17.3.1.29 `w:pPr/w:rPr` — the formatting of the paragraph MARK. An empty
   * paragraph is as tall as its mark, so dropping this on resolution collapsed
   * a 36pt empty cell to a line of body text (conditionalstyles-tbllook.docx).
   */
  readonly runProperties?: RunProperties;
  /** §17.3.1.11 — the paragraph floats as a text frame. */
  readonly frame?: FrameProperties;
  /** §17.3.1.41 — which way the paragraph's lines run (`lrTb` is not recorded). */
  readonly textDirection?: 'btLr' | 'tbRl';
  /** §17.3.1.32 — whether these lines stand on the section's grid. Absent ⇒ they do. */
  readonly snapToGrid?: boolean;
  /**
   * §17.6.17 — the paragraph mark carries the section break, so a paragraph
   * with no content of its own prints nothing at all.
   */
  readonly sectionBreak?: boolean;
  /**
   * §17.9 list reference, carried through for tagged-PDF list structure
   * (L/LI nesting) — markers themselves are materialized by `applyNumbering`.
   */
  readonly numbering?: NumberingReference;
  /**
   * §17.6.5 — the line pitch of the section's document grid, in points. Not a
   * property of the paragraph at all: the section it falls in owns it, and the
   * layout stamps it here so a line knows the grid it stands on.
   */
  readonly gridLinePitchPt?: Pt;
}

/** Word's empty-document run defaults (used when `docDefaults` is absent). */
export const DEFAULT_RESOLVED_RUN: ResolvedRunProperties = {
  bold: false,
  italic: false,
  underline: 'none',
  strike: false,
  caps: false,
  smallCaps: false,
  fontSizePt: halfPtToPt(22),
  colorHex: '000000',
  fontFamily: {},
  verticalAlign: 'baseline',
  rtl: false,
};

/** Word's empty-document paragraph defaults (used when `docDefaults` is absent). */
export const DEFAULT_RESOLVED_PARAGRAPH: ResolvedParagraphProperties = {
  alignment: 'left',
  spacingBefore: twipsToPt(0),
  spacingAfter: twipsToPt(0),
  spacingLine: twipsToPt(240),
  spacingLineRule: 'auto',
  indentLeft: twipsToPt(0),
  indentRight: twipsToPt(0),
  indentFirstLine: twipsToPt(0),
  pageBreakBefore: false,
  contextualSpacing: false,
  tabs: [],
  bidi: false,
};
