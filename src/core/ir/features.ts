// IR core — feature vocabulary for adapter capabilities (ir-design.md §9).
//
// Adapters declare what they understand (`supports: ReadonlySet<Feature>`) and
// conversions report what they lost (LossReport) in terms of these features.
// The vocabulary is hierarchical via dot-segments ('tables' ⊃ 'tables.nested')
// and OPEN: `Feature` is a plain string so third-party adapters can introduce
// their own; this registry only canonicalizes the names the core uses, so the
// capability matrix in the docs can be generated from code.

/**
 * A capability/loss feature name. A plain string, hierarchical via dot-segments
 * (`'tables.nested'` ⊂ `'tables'`); the vocabulary is OPEN so third-party
 * adapters can introduce their own. {@link FEATURES} canonicalizes the names the
 * core uses.
 */
export type Feature = string;

/** The feature names the core declares, so the docs' capability matrix is code-generated. */
export const FEATURES = {
  text: 'text',
  tables: 'tables',
  tablesNested: 'tables.nested',
  cellFormatting: 'cellFormatting',
  lists: 'lists',
  sections: 'sections',
  headersFooters: 'headersFooters',
  images: 'images',
  imagesJp2: 'images.jp2',
  hyperlinks: 'hyperlinks',
  shapes: 'shapes',
  smartArt: 'shapes.smartArt',
  charts: 'charts',
  math: 'math',
  rtl: 'rtl',
  trackedChanges: 'trackedChanges',
  hyphenation: 'hyphenation',
  fontsEmbedding: 'fonts.embedding',
  fontsSubstitution: 'fonts.substitution',
  pdfA: 'pdf.archival',
  pdfTagged: 'pdf.tagged',
  pdfSignatures: 'pdf.signatures',
} as const;

/** Union of the canonical {@link FEATURES} values. */
export type KnownFeature = (typeof FEATURES)[keyof typeof FEATURES];

/** True when `feature` equals `prefix` or sits under it ('tables.nested' ⊂ 'tables'). */
export function featureWithin(feature: Feature, prefix: Feature): boolean {
  return feature === prefix || feature.startsWith(`${prefix}.`);
}
