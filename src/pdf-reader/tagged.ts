// E-PDF EP3 — tagged fast-path reconstruction. Walks the structure tree (EP3
// struct-tree.ts), pulls each element's text from the per-page MCID → text map
// the content interpreter produced (EP2), and rebuilds a FlowDoc: headings
// (H1–H6 → outline level), paragraphs, tables (Table → TR → TH/TD), list items
// (each LI → its label + body text), and figures (EP6 — each /Figure's MCID
// resolves to a lifted image, carrying its /Alt). The honest inverse of the
// tagged PDF Ream writes.

import {
  buildFlowDoc,
  dedupeLosses,
  imageBlock,
  paragraphBlock,
  paragraphFromRuns,
} from './flow-build';
import { collectPageImages } from './images';
import { readStructTree } from './struct-tree';
import { extractPageText } from './text';
import type { BodyElement, Table, TableCell, TableRow } from '@/core/document-model';
import type { Pt } from '@/core/ir';

import type { TextRun } from './content';
import type { PdfFile } from './document';
import type { Reconstruction, TextSpan } from './flow-build';
import type { PdfImage } from './images';
import type { StructNode } from './struct-tree';
import { ResourceStore, pt } from '@/core/ir';

// Printable width assumed for a synthesized table grid (6.5" — the columns are
// equal because the structure tree carries no widths; layout auto-fits anyway).
const ASSUMED_CONTENT_WIDTH_PT = 468;

export function reconstructTaggedPdf(file: PdfFile): Reconstruction | undefined {
  const root = readStructTree(file);
  if (!root) return undefined;

  const pages = file.pages();
  // Per page: MCID → its runs, in show order (runs carry any hyperlink, EP8).
  const pageRuns = pages.map((page) => {
    const byMcid = new Map<number, Array<TextRun>>();
    for (const run of extractPageText(file, page)) {
      if (run.mcid === undefined) continue;
      const list = byMcid.get(run.mcid);
      if (list) list.push(run);
      else byMcid.set(run.mcid, [run]);
    }
    return byMcid;
  });
  const runsOfMcid = (page: number, mcid: number): Array<TextRun> =>
    pageRuns[page]?.get(mcid) ?? [];

  // Per page: the lifted images, indexed by their owning MCID (a /Figure's).
  const resources = new ResourceStore();
  const pageImages = pages.map((page) => collectPageImages(file, page));
  const imageLosses = dedupeLosses(pageImages.flatMap((p) => p.losses));
  const imagesByMcid = pageImages.map((p) => {
    const byMcid = new Map<number, Array<PdfImage>>();
    for (const img of p.images) {
      if (img.mcid === undefined) continue;
      const list = byMcid.get(img.mcid);
      if (list) list.push(img);
      else byMcid.set(img.mcid, [img]);
    }
    return byMcid;
  });
  const emitted = new Set<PdfImage>();
  const imagesForNode = (node: StructNode): Array<PdfImage> =>
    node.mcids.flatMap(({ page, mcid }) => imagesByMcid[page]?.get(mcid) ?? []);

  const textOf = (node: StructNode): string =>
    squash(
      node.mcids
        .map(({ page, mcid }) =>
          runsOfMcid(page, mcid)
            .map((r) => r.text)
            .join(''),
        )
        .join(' '),
    );

  // The node's own runs as link-carrying spans, with a space between MCIDs.
  const spansOf = (node: StructNode): Array<TextSpan> => {
    const spans: Array<TextSpan> = [];
    node.mcids.forEach(({ page, mcid }, i) => {
      if (i > 0) spans.push({ text: ' ' });
      for (const run of runsOfMcid(page, mcid)) {
        spans.push(
          run.href !== undefined ? { text: run.text, href: run.href } : { text: run.text },
        );
      }
    });
    return spans;
  };

  // All text under a node, in reading order (a list item's label + body).
  const collectText = (node: StructNode): string =>
    squash([textOf(node), ...node.children.map(collectText)].join(' '));

  function emit(node: StructNode, out: Array<BodyElement>): void {
    if (node.type === 'Table') {
      const table = buildTable(node);
      if (table) out.push(table);
      return;
    }
    if (node.type === 'Figure') {
      for (const img of imagesForNode(node)) {
        emitted.add(img);
        out.push(imageBlock(img, resources, node.alt));
      }
      return;
    }
    if (node.type === 'LI') {
      const text = collectText(node);
      if (text.length > 0) out.push(paragraphBlock(text, undefined));
      return;
    }
    if (node.children.length === 0) {
      if (textOf(node).length > 0) {
        out.push(paragraphFromRuns(spansOf(node), headingLevel(node.type)));
      }
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

  // Images not claimed by a /Figure (untagged figures, third-party PDFs) still
  // belong in the document — append them in page + top-down order so nothing is
  // silently lost.
  const orphans: Array<{ page: number; img: PdfImage }> = [];
  pageImages.forEach((p, page) => {
    for (const img of p.images) if (!emitted.has(img)) orphans.push({ page, img });
  });
  orphans.sort((a, b) => a.page - b.page || b.img.y - a.img.y);
  for (const { img } of orphans) body.push(imageBlock(img, resources));

  if (body.length === 0) return undefined;
  return { doc: buildFlowDoc(body, resources), losses: imageLosses };
}

function squash(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

// H1–H6 → outline level 0–5 (the FlowDoc heading representation, §17.3.1.20).
function headingLevel(type: string): number | undefined {
  const m = /^H([1-6])$/.exec(type);
  return m ? Number(m[1]) - 1 : undefined;
}
