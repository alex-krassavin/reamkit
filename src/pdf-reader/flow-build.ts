// E-PDF — shared FlowDoc construction for the two reconstruction paths (the
// tagged fast-path EP3 and the heuristic layout path EP4). A reconstructed PDF
// carries only a body of paragraphs/tables over the empty style sheet.

import type { BodyElement, ParagraphProperties } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';

import { ResourceStore } from '@/core/ir';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';

export function paragraphBlock(text: string, outlineLevel?: number): BodyElement {
  const properties: ParagraphProperties = outlineLevel !== undefined ? { outlineLevel } : {};
  return {
    kind: 'paragraph',
    paragraph: { properties, runs: text.length > 0 ? [{ text, properties: {} }] : [] },
  };
}

export function buildFlowDoc(body: ReadonlyArray<BodyElement>): FlowDoc {
  return {
    kind: 'flow',
    body: resolveBodyStyles([...body], EMPTY_STYLE_SHEET),
    sections: [],
    styles: EMPTY_STYLE_SHEET,
    resources: new ResourceStore(),
  };
}
