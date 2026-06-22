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
  StyleSheet,
} from '@/core/document-model';
import type { FontRegistry } from '@/core/font';
import type { ResourceStore } from '@/core/ir/resources';

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
  /** Document metadata from docProps/core.xml. */
  readonly info?: DocumentInfo;
  /** Document natural language hint (BCP-47), e.g. for tagged-PDF /Lang. */
  readonly language?: string;
}
