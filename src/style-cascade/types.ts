// Fully-resolved properties consumed by the renderer.
// Every field is required — the cascade has already collapsed all undefined
// values against document defaults and style inheritance.

import type { Alignment, FontFamilyMap, UnderlineStyle, VerticalAlign } from '@/document-model';
import type { Pt } from '@/ir';
import { halfPtToPt, twipsToPt } from '@/ir';

export interface ResolvedRunProperties {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: UnderlineStyle;
  readonly strike: boolean;
  readonly fontSizePt: Pt;
  readonly colorHex: string;
  readonly fontFamily: FontFamilyMap;
  readonly verticalAlign: VerticalAlign;
  readonly rtl: boolean;
  // Natural language (w:lang @w:val). Optional — only set when the source
  // specifies one; consumed by the tagged-PDF per-element /Lang, never layout.
  readonly lang?: string;
}

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
  readonly bidi: boolean;
  // ECMA-376 §17.3.1.20 — resolved outline level (0–8). Undefined = body text.
  // Used only by tagged-PDF heading detection.
  readonly outlineLevel?: number;
  // The paragraph's style id (e.g. "Heading2"), carried through for heading
  // detection when a style lacks an explicit outline level.
  readonly styleId?: string;
}

// Word's empty document defaults (used when docDefaults is absent).
export const DEFAULT_RESOLVED_RUN: ResolvedRunProperties = {
  bold: false,
  italic: false,
  underline: 'none',
  strike: false,
  fontSizePt: halfPtToPt(22),
  colorHex: '000000',
  fontFamily: {},
  verticalAlign: 'baseline',
  rtl: false,
};

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
  bidi: false,
};
