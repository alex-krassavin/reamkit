// FlowDoc — the semantic IR tree (ir-design §5), v0.
//
// Everything a reader extracts from the document BYTES, format-neutrally:
// the flow content plus its document-scoped companions (styles, numbering,
// headers/footers, charts, binary resources, metadata). Caller-supplied
// conversion options (fonts, PDF/A profile, signature, …) are deliberately
// NOT part of the tree — they parameterize transforms, not the document.
//
// Stage 6 (closing the v0 deviation): `body` carries FINAL effective
// properties — readers materialize list markers (applyNumbering) and resolve
// the style cascade (resolveBodyStyles) while building the tree. `styles` and
// `numbering` remain as raw round-trip material; render projections must not
// re-apply them (resolving over the empty sheet is the identity).

import type {
  BodyElement,
  Chart,
  Comment,
  DocumentInfo,
  Numbering,
  Section,
  SectionProperties,
  ShapeFill,
  StyleSheet,
} from '@/core/document-model';
import type { FontRegistry, GlyphSeg } from '@/core/font';
import type { ResourceStore } from '@/core/ir/resources';

/**
 * A face's family as a word processor names it (ECMA-376 §17.8.3.9 `w:font`).
 */
export interface FaceFamily {
  /** The family's own name: `Inter` for `Inter-SemiBold`, `Arial` for `ArialMT`. */
  readonly family: string;
  /** §17.8.3.10 `w:family` — the kind of face, for a reader that has to substitute. */
  readonly generic: 'roman' | 'swiss' | 'modern';
}

/**
 * The outlines a face drew a document's characters with — what a writer needs
 * to EMBED the face, so a reader that lacks it sets the text in the face the
 * source was set in instead of a substitute (ECMA-376 §17.8.1).
 */
export interface FaceOutlines {
  /** Each character (one code point) the document shows in the face → its glyph. */
  readonly glyphs: ReadonlyMap<string, FaceGlyph>;
  /**
   * OS/2 `fsType` — the embedding the face's licence allows, as its program
   * states it; absent where the program states nothing.
   */
  readonly fsType?: number;
  /** The style the face IS, which is the slot a family embeds it in. */
  readonly bold: boolean;
  readonly italic: boolean;
  /** The program's own name for the face, without a subset tag. */
  readonly postScriptName: string;
  /**
   * The line to set the face in, above and below the baseline, in thousandths
   * of an em — `descent` negative. A reader that reconstructs a page gives the
   * line it measured the page against, so a paragraph set in the face's single
   * spacing lands where the page put it.
   */
  readonly ascent: number;
  readonly descent: number;
  readonly capHeight?: number;
  readonly xHeight?: number;
  /** Degrees counterclockwise from the vertical; a face slanted right is negative. */
  readonly italicAngle: number;
  readonly fixedPitch: boolean;
}

/** One glyph of a {@link FaceOutlines}: what it draws and how far it advances. */
export interface FaceGlyph {
  /** Its contours in a one-unit em, y up, filled by the nonzero rule; empty when blank. */
  readonly outline: ReadonlyArray<GlyphSeg>;
  /** How far the pen moves after it, in thousandths of an em. */
  readonly advance: number;
}

/**
 * The semantic IR tree (ir-design §5): everything a reader extracts from the
 * document bytes, format-neutrally — the flow `body` plus its document-scoped
 * companions (styles, numbering, header/footer bands, notes, charts, binary
 * resources, metadata). Caller-supplied conversion options (fonts, PDF/A
 * profile, signature, …) are deliberately NOT part of the tree; they
 * parameterize transforms, not the document.
 *
 * `body` carries FINAL effective properties — readers materialize list markers
 * and resolve the style cascade while building it — so render projections must
 * not re-apply `styles`/`numbering`, which remain only as round-trip material.
 */
export interface FlowDoc {
  /** Discriminant for {@link SourceDoc} (a FlowDoc passes through projection). */
  readonly kind: 'flow';
  /** The document flow content, carrying resolved, effective properties. */
  readonly body: ReadonlyArray<BodyElement>;
  /** Multi-section page geometry (docx). Empty for single-geometry sources. */
  readonly sections: ReadonlyArray<Section>;
  /** Single-section page geometry (xlsx print setup). */
  readonly section?: SectionProperties;
  /** Resolved style sheet, kept as round-trip material (already folded into `body`). */
  readonly styles: StyleSheet;
  /**
   * Raw numbering definitions (round-trip material). `body` already carries the
   * materialized list markers — readers apply numbering as a FlowDoc transform,
   * so render projections must not re-apply it.
   */
  readonly numbering?: Numbering;
  readonly headersFooters?: ReadonlyMap<string, ReadonlyArray<BodyElement>>;
  /** §17.11 footnotes/endnotes content by id (separator stubs excluded). */
  readonly footnotes?: ReadonlyMap<string, ReadonlyArray<BodyElement>>;
  readonly endnotes?: ReadonlyMap<string, ReadonlyArray<BodyElement>>;
  /** §17.13.4 review comments by id, anchored from a run's `commentRef`. */
  readonly comments?: ReadonlyMap<string, Comment>;
  /** Parsed charts keyed by relationship id (ChartBlock.chartRelId). */
  readonly charts?: ReadonlyMap<string, Chart>;
  /** Content-addressed binary resources (images). */
  readonly resources: ResourceStore;
  /** Fonts embedded in the source document itself (docx fontTable), by name. */
  readonly embeddedFonts?: ReadonlyMap<string, FontRegistry>;
  /**
   * The family each face a run names belongs to, where the two are not the
   * same name. A PDF names a FACE — `Inter-SemiBold` — where a word processor
   * names a family and states the weight beside it (`Inter`, bold). The layout
   * finds a document's own program by the face; a writer that hands the text
   * to another program names the family, which is the name that program knows.
   */
  readonly faceFamilies?: ReadonlyMap<string, FaceFamily>;
  /**
   * The outlines of the faces a run names, keyed as {@link faceFamilies} is —
   * for a writer that embeds them (see {@link FaceOutlines}).
   */
  readonly faceOutlines?: ReadonlyMap<string, FaceOutlines>;
  /** Document metadata from docProps/core.xml. */
  readonly info?: DocumentInfo;
  /** Document natural language hint (BCP-47), e.g. for tagged-PDF /Lang. */
  readonly language?: string;
  /**
   * ECMA-376 §17.15.1.35 `w:doNotExpandShiftReturn` — a justified line that
   * ends at a soft line break keeps its natural width.
   */
  readonly doNotExpandShiftReturn?: boolean;
  /**
   * ECMA-376 §17.2.1 `w:background` — the colour every page is painted, when
   * the document asks for one AND §17.15.1.28 `w:displayBackgroundShape` says
   * to draw it.
   */
  readonly pageBackgroundColorHex?: string;
  /**
   * ECMA-376 §17.2.1 — the same background as the FILL it is, when that is more
   * than a flat colour: the `v:background`'s gradient or picture. The colour
   * above stays the flat fallback, for writers that paint only colours.
   */
  readonly pageBackgroundFill?: ShapeFill;
  /**
   * ECMA-376 §17.15.1.38 `w:gutterAtTop` — the binding space `w:pgMar
   * @w:gutter` reserves belongs to the TOP margin rather than the left.
   */
  readonly gutterAtTop?: boolean;
}
