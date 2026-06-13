// E-PDF — shared FlowDoc construction for the two reconstruction paths (the
// tagged fast-path EP3 and the heuristic layout path EP4). A reconstructed PDF
// carries a body of paragraphs/tables/images over the empty style sheet, with
// any lifted image bytes (EP6) in the resource store the writers embed from.

import type { BodyElement, ParagraphProperties } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';

import type { PdfImage } from './images';
import { ResourceStore, pt } from '@/core/ir';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';

// A reconstruction's document plus the losses incurred reading it (e.g. an
// undecodable image colour space) — surfaced through the reader's LossReport.
export interface Reconstruction {
  readonly doc: FlowDoc;
  readonly losses: ReadonlyArray<Loss>;
}

export function paragraphBlock(text: string, outlineLevel?: number): BodyElement {
  const properties: ParagraphProperties = outlineLevel !== undefined ? { outlineLevel } : {};
  return {
    kind: 'paragraph',
    paragraph: { properties, runs: text.length > 0 ? [{ text, properties: {} }] : [] },
  };
}

// One piece of reconstructed text, carrying any hyperlink (E-PDF EP8).
export interface TextSpan {
  readonly text: string;
  readonly href?: string;
}

// Build a paragraph from positioned spans, coalescing consecutive spans that
// share an href into one run (so a link survives as its own run) and squashing
// whitespace. With no hrefs this collapses to a single run — the same shape
// `paragraphBlock` produces.
export function paragraphFromRuns(
  spans: ReadonlyArray<TextSpan>,
  outlineLevel?: number,
): BodyElement {
  const merged: Array<{ text: string; href?: string }> = [];
  for (const s of spans) {
    const last = merged[merged.length - 1];
    if (last && last.href === s.href) last.text += s.text;
    else if (s.href !== undefined) merged.push({ text: s.text, href: s.href });
    else merged.push({ text: s.text });
  }
  const runs = merged
    .map((m) => ({ text: m.text.replace(/\s+/g, ' '), href: m.href }))
    .filter((m) => m.text.length > 0);
  // Trim the paragraph's outer whitespace.
  if (runs.length > 0) {
    runs[0]!.text = runs[0]!.text.replace(/^ /, '');
    runs[runs.length - 1]!.text = runs[runs.length - 1]!.text.replace(/ $/, '');
  }
  const properties: ParagraphProperties = outlineLevel !== undefined ? { outlineLevel } : {};
  return {
    kind: 'paragraph',
    paragraph: {
      properties,
      runs: runs
        .filter((r) => r.text.length > 0)
        .map((r) => ({ text: r.text, properties: {}, ...(r.href ? { href: r.href } : {}) })),
    },
  };
}

// Store the image bytes (content-addressed dedup) and build the block that
// references them, sized in points from the placement CTM.
export function imageBlock(image: PdfImage, resources: ResourceStore, alt?: string): BodyElement {
  const resource = resources.put(image.bytes);
  return {
    kind: 'image',
    image: {
      resource,
      width: pt(image.widthPt),
      height: pt(image.heightPt),
      paragraphProperties: {},
      ...(alt ? { altText: alt } : {}),
    },
  };
}

// Collapse losses sharing a message (the same colour space dropped on many pages).
export function dedupeLosses(losses: ReadonlyArray<Loss>): Array<Loss> {
  const byDetail = new Map<string, Loss>();
  for (const loss of losses) if (!byDetail.has(loss.detail)) byDetail.set(loss.detail, loss);
  return [...byDetail.values()];
}

export function buildFlowDoc(
  body: ReadonlyArray<BodyElement>,
  resources: ResourceStore = new ResourceStore(),
): FlowDoc {
  return {
    kind: 'flow',
    body: resolveBodyStyles([...body], EMPTY_STYLE_SHEET),
    sections: [],
    styles: EMPTY_STYLE_SHEET,
    resources,
  };
}
