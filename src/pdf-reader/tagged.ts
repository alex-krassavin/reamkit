// E-PDF EP3 — tagged fast-path reconstruction. Walks the structure tree (EP3
// struct-tree.ts), pulls each element's text from the per-page MCID → text map
// the content interpreter produced (EP2), and rebuilds a FlowDoc: headings
// (H1–H6 → outline level), paragraphs, and — flattened for now — the text of
// table cells and list items. The honest inverse of the tagged PDF Ream writes.

import { readStructTree } from './struct-tree';
import { extractPageText } from './text';
import type { BodyElement, ParagraphProperties } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';

import type { PdfFile } from './document';
import type { StructNode } from './struct-tree';
import { ResourceStore } from '@/core/ir';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';

export function reconstructTaggedPdf(file: PdfFile): FlowDoc | undefined {
  const root = readStructTree(file);
  if (!root) return undefined;

  // Per page: MCID → its text, concatenated in show order.
  const pageText = file.pages().map((page) => {
    const byMcid = new Map<number, string>();
    for (const run of extractPageText(file, page)) {
      if (run.mcid !== undefined) byMcid.set(run.mcid, (byMcid.get(run.mcid) ?? '') + run.text);
    }
    return byMcid;
  });

  const textOf = (node: StructNode): string =>
    node.mcids
      .map(({ page, mcid }) => pageText[page]?.get(mcid) ?? '')
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();

  const body: Array<BodyElement> = [];
  const walk = (node: StructNode): void => {
    if (node.children.length === 0) {
      const text = textOf(node);
      if (text.length > 0) body.push(paragraphBlock(text, headingLevel(node.type)));
      return;
    }
    for (const child of node.children) walk(child);
  };
  walk(root);
  if (body.length === 0) return undefined;

  return {
    kind: 'flow',
    body: resolveBodyStyles(body, EMPTY_STYLE_SHEET),
    sections: [],
    styles: EMPTY_STYLE_SHEET,
    resources: new ResourceStore(),
  };
}

function paragraphBlock(text: string, outlineLevel: number | undefined): BodyElement {
  const properties: ParagraphProperties = outlineLevel !== undefined ? { outlineLevel } : {};
  return {
    kind: 'paragraph',
    paragraph: { properties, runs: [{ text, properties: {} }] },
  };
}

// H1–H6 → outline level 0–5 (the FlowDoc heading representation, §17.3.1.20).
function headingLevel(type: string): number | undefined {
  const m = /^H([1-6])$/.exec(type);
  return m ? Number(m[1]) - 1 : undefined;
}
