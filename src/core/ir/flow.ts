// FlowDoc — the semantic IR tree (ir-design §5), v0.
//
// Everything a reader extracts from the document BYTES, format-neutrally:
// the flow content plus its document-scoped companions (styles, numbering,
// headers/footers, charts, binary resources, metadata). Caller-supplied
// conversion options (fonts, PDF/A profile, signature, …) are deliberately
// NOT part of the tree — they parameterize transforms, not the document.
//
// v0 deviation from the target design, recorded on purpose: runs/paragraphs
// still carry cascade-able properties plus a StyleSheet, instead of fully
// resolved properties (the cascade currently runs inside layout). Folding the
// resolve into readers is part of stabilizing the schema (stage 6).

import type {
  BodyElement,
  Chart,
  DocumentInfo,
  Numbering,
  Section,
  SectionProperties,
  StyleSheet,
} from '@/core/document-model';
import type { FontRegistry } from '@/core/font';
import type { ResourceStore } from '@/core/ir/resources';

export interface FlowDoc {
  readonly kind: 'flow';
  readonly body: ReadonlyArray<BodyElement>;
  /** Multi-section page geometry (docx). Empty for single-geometry sources. */
  readonly sections: ReadonlyArray<Section>;
  /** Single-section page geometry (xlsx print setup). */
  readonly section?: SectionProperties;
  readonly styles: StyleSheet;
  // Raw definitions (round-trip material). `body` already carries the
  // materialized list markers — readers apply numbering as a FlowDoc
  // transform (stage 6); render projections must not re-apply it.
  readonly numbering?: Numbering;
  readonly headersFooters?: ReadonlyMap<string, ReadonlyArray<BodyElement>>;
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
