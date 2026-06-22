// FlowDoc → render options projection (oop-design §8, B2).
//
// One owner for the field-by-field mapping the converters and the facade all
// repeated by hand (and which had already drifted apart once). Key-presence
// matters downstream — conditional spreads are semantics, not style.

import type { FlowDoc } from '@/core/ir/flow';
import type { StyledRenderOptions } from '@/pdf';
import { EMPTY_STYLE_SHEET } from '@/core/style-cascade';

/** Styled-render options projected from a {@link FlowDoc} (everything but the font `registry`). */
export type FlowRenderOptions = Omit<StyledRenderOptions, 'registry'>;

/**
 * Project a {@link FlowDoc} to the {@link StyledRenderOptions} the layout/PDF path
 * consumes (minus the font `registry`, which the caller supplies). This is the
 * single owner of the field-by-field mapping the converters and the facade used
 * to repeat by hand. Key presence is significant — the conditional spreads are
 * semantics, not style.
 *
 * @param flow The parsed interlayer.
 * @returns The render options, threading the document's own sections,
 *          header/footer bands, notes, resources, charts, fonts and language.
 */
export function flowRenderOptions(flow: FlowDoc): FlowRenderOptions {
  return {
    // flow.body already carries materialized list markers AND resolved
    // properties (stage-6 reader transforms) — projecting styles/numbering
    // here would make the renderer apply them a second time. Resolving over
    // the empty sheet is the identity, so the renderer's own cascade pass
    // degrades to a memoized no-op.
    styles: EMPTY_STYLE_SHEET,
    ...(flow.sections.length > 0 ? { sections: flow.sections } : {}),
    ...(flow.section ? { section: flow.section } : {}),
    ...(flow.headersFooters ? { headersFooters: flow.headersFooters } : {}),
    ...(flow.footnotes ? { footnotes: flow.footnotes } : {}),
    ...(flow.endnotes ? { endnotes: flow.endnotes } : {}),
    ...(flow.comments ? { comments: flow.comments } : {}),
    resources: flow.resources,
    ...(flow.charts ? { charts: flow.charts } : {}),
    ...(flow.embeddedFonts ? { embeddedFonts: flow.embeddedFonts } : {}),
    ...(flow.language ? { language: flow.language } : {}),
  };
}
