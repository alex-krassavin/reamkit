// E-PDF EP3 — tagged fast-path reconstruction. Walks the structure tree (EP3
// struct-tree.ts), pulls each element's text from the per-page MCID → text map
// the content interpreter produced (EP2), and rebuilds a FlowDoc: headings
// (H1–H6 → outline level), paragraphs, tables (Table → TR → TH/TD), and list
// items (each LI → its label + body text). The honest inverse of the tagged PDF
// Ream writes.

import { buildFlowDoc, paragraphBlock } from './flow-build';
import { readStructTree } from './struct-tree';
import { extractPageText } from './text';
import type { BodyElement, Table, TableCell, TableRow } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { Pt } from '@/core/ir';

import type { PdfFile } from './document';
import type { StructNode } from './struct-tree';
import { pt } from '@/core/ir';

// Printable width assumed for a synthesized table grid (6.5" — the columns are
// equal because the structure tree carries no widths; layout auto-fits anyway).
const ASSUMED_CONTENT_WIDTH_PT = 468;

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
    squash(node.mcids.map(({ page, mcid }) => pageText[page]?.get(mcid) ?? '').join(' '));

  // All text under a node, in reading order (a list item's label + body).
  const collectText = (node: StructNode): string =>
    squash([textOf(node), ...node.children.map(collectText)].join(' '));

  function emit(node: StructNode, out: Array<BodyElement>): void {
    if (node.type === 'Table') {
      const table = buildTable(node);
      if (table) out.push(table);
      return;
    }
    if (node.type === 'LI') {
      const text = collectText(node);
      if (text.length > 0) out.push(paragraphBlock(text, undefined));
      return;
    }
    if (node.children.length === 0) {
      const text = textOf(node);
      if (text.length > 0) out.push(paragraphBlock(text, headingLevel(node.type)));
      return;
    }
    for (const child of node.children) emit(child, out);
  }

  function buildTable(tableNode: StructNode): BodyElement | undefined {
    const rows: Array<TableRow> = [];
    const collectRows = (n: StructNode): void => {
      for (const child of n.children) {
        if (child.type === 'TR') rows.push(buildRow(child));
        else if (child.type === 'THead' || child.type === 'TBody' || child.type === 'TFoot') {
          collectRows(child);
        }
      }
    };
    collectRows(tableNode);
    if (rows.length === 0) return undefined;
    const numCols = Math.max(
      1,
      ...rows.map((r) => r.cells.reduce((s, c) => s + (c.properties.colSpan ?? 1), 0)),
    );
    const colWidth = pt(Math.max(1, ASSUMED_CONTENT_WIDTH_PT / numCols));
    const grid: Array<Pt> = Array.from({ length: numCols }, () => colWidth);
    const table: Table = { properties: {}, grid, rows };
    return { kind: 'table', table };
  }

  function buildRow(trNode: StructNode): TableRow {
    const cells: Array<TableCell> = [];
    let allHeader = false;
    for (const cell of trNode.children) {
      if (cell.type !== 'TH' && cell.type !== 'TD') continue;
      if (cells.length === 0) allHeader = true;
      if (cell.type !== 'TH') allHeader = false;
      const content: Array<BodyElement> = [];
      for (const child of cell.children) emit(child, content);
      if (content.length === 0) content.push(paragraphBlock(textOf(cell), undefined));
      const colSpan = cell.colSpan ?? 1;
      cells.push({ properties: colSpan > 1 ? { colSpan } : {}, content });
    }
    return { properties: allHeader && cells.length > 0 ? { isHeader: true } : {}, cells };
  }

  const body: Array<BodyElement> = [];
  emit(root, body);
  if (body.length === 0) return undefined;
  return buildFlowDoc(body);
}

function squash(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

// H1–H6 → outline level 0–5 (the FlowDoc heading representation, §17.3.1.20).
function headingLevel(type: string): number | undefined {
  const m = /^H([1-6])$/.exec(type);
  return m ? Number(m[1]) - 1 : undefined;
}
