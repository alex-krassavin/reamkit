// docx writer (E-DOCX): FlowDoc → WordprocessingML package — the inverse of
// the docx reader, and the fifth adapter overall. A flow medium with zero
// layout and zero I/O, like the HTML writer.
//
// v1 contract (epics.md, variant A): the writer emits a DENORMALIZED but
// valid document. FlowDoc's body carries RESOLVED properties (the stage-6
// cascade is already collapsed), so what we write is direct formatting — no
// named styles. The round-trip guarantee is therefore semantic, not textual:
// readDocx(writeDocx(flow)) yields an equivalent FlowDoc, never the original
// bytes. Anything the writer does not serialize yet is reported as a loss,
// exactly like the other writers.

import type { BodyElement, Paragraph, SectionProperties } from '@/core/document-model';
import type { DocumentWriter, WriteResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';

import { FEATURES } from '@/core/ir';
import { buildOpcPackage } from '@/core/opc';

const encoder = new TextEncoder();

const DOC_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const REL_OFFICE_DOCUMENT =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

export function writeDocx(flow: FlowDoc): WriteResult {
  const losses: Array<Loss> = [];
  const body: Array<string> = [];

  for (const el of flow.body) emitBlock(body, el, losses);

  const section = flow.sections[0]?.properties ?? flow.section;
  if (section) body.push(sectPrXml(section));

  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<w:body>${body.join('')}</w:body>` +
    '</w:document>';

  const bytes = buildOpcPackage({
    parts: [
      {
        path: 'word/document.xml',
        data: encoder.encode(documentXml),
        contentType: DOC_CONTENT_TYPE,
      },
    ],
    rootRelationships: [
      {
        id: 'rId1',
        type: REL_OFFICE_DOCUMENT,
        target: 'word/document.xml',
        targetMode: 'Internal',
      },
    ],
  });

  return { bytes, losses };
}

export const docxWriter: DocumentWriter<FlowDoc> = {
  id: 'docx',
  consumes: 'flow',
  supports: new Set([FEATURES.text]),
  write: (doc) => writeDocx(doc),
};

function emitBlock(out: Array<string>, el: BodyElement, losses: Array<Loss>): void {
  if (el.kind === 'paragraph') {
    out.push(paragraphXml(el.paragraph));
    return;
  }
  // Tables, images, charts, shapes: D4/D5 of the epic. Reported, not dropped
  // silently.
  const feature =
    el.kind === 'table'
      ? FEATURES.tables
      : el.kind === 'image'
        ? FEATURES.images
        : el.kind === 'chart'
          ? FEATURES.charts
          : FEATURES.shapes;
  losses.push({
    severity: 'dropped',
    feature,
    detail: `${el.kind} not written by the docx writer (v0)`,
  });
}

function paragraphXml(p: Paragraph): string {
  const runs: Array<string> = [];
  for (const run of p.runs) {
    // The reader materialized list markers into the body (stage 6); writing
    // them back as literal text would double the marker on re-read with
    // numbering re-applied — but v0 writes no numbering.xml, so the literal
    // marker IS the faithful rendition.
    if (run.text === '' || run.math !== undefined || run.inlineImage !== undefined) continue;
    runs.push(`<w:r><w:t xml:space="preserve">${escapeXml(run.text)}</w:t></w:r>`);
  }
  return `<w:p>${runs.join('')}</w:p>`;
}

// §17.6.17 — page size and margins in twentieths of a point.
function sectPrXml(s: SectionProperties): string {
  const parts: Array<string> = [];
  if (s.pageSize) {
    const orient = s.pageSize.orientation === 'landscape' ? ' w:orient="landscape"' : '';
    parts.push(
      `<w:pgSz w:w="${twips(s.pageSize.width)}" w:h="${twips(s.pageSize.height)}"${orient}/>`,
    );
  }
  if (s.margins) {
    const m = s.margins;
    const header = m.header !== undefined ? ` w:header="${twips(m.header)}"` : '';
    const footer = m.footer !== undefined ? ` w:footer="${twips(m.footer)}"` : '';
    parts.push(
      `<w:pgMar w:top="${twips(m.top)}" w:right="${twips(m.right)}"` +
        ` w:bottom="${twips(m.bottom)}" w:left="${twips(m.left)}"${header}${footer}/>`,
    );
  }
  if (parts.length === 0) return '';
  return `<w:sectPr>${parts.join('')}</w:sectPr>`;
}

const twips = (pt: number): number => Math.round(pt * 20);

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
