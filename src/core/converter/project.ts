// FlowDoc → render options projection (oop-design §8, B2).
//
// One owner for the field-by-field mapping the converters and the facade all
// repeated by hand (and which had already drifted apart once). Key-presence
// matters downstream — conditional spreads are semantics, not style.

import type { FlowDoc } from '@/core/ir/flow';
import type { StyledRenderOptions } from '@/pdf';

export type FlowRenderOptions = Omit<StyledRenderOptions, 'registry'>;

export function flowRenderOptions(flow: FlowDoc): FlowRenderOptions {
  return {
    styles: flow.styles,
    ...(flow.numbering ? { numbering: flow.numbering } : {}),
    ...(flow.sections.length > 0 ? { sections: flow.sections } : {}),
    ...(flow.section ? { section: flow.section } : {}),
    ...(flow.headersFooters ? { headersFooters: flow.headersFooters } : {}),
    resources: flow.resources,
    ...(flow.charts ? { charts: flow.charts } : {}),
    ...(flow.embeddedFonts ? { embeddedFonts: flow.embeddedFonts } : {}),
    ...(flow.language ? { language: flow.language } : {}),
  };
}
