// XLSX reader (ir-design §7): bytes → FlowDoc. The grid is projected through
// the Excel print model into flow blocks (SheetDoc — the dedicated grid tree —
// is deliberately deferred tech debt; see handoff v1 §1). Document-derived
// state only; caller conversion options stay with the converter/facade.

import type { BodyElement, DocumentInfo, SectionProperties } from '@/document-model';
import type { DocumentReader, ReadResult } from '@/ir/adapters';
import type { FlowDoc } from '@/ir/flow';
import type { CoreProperties } from '@/opc';

import { EMPTY_STYLE_SHEET } from '@/ooxml/wordproc';
import { FEATURES, ResourceStore } from '@/ir';
import { OpcPackage, parseCoreProperties } from '@/opc';
import {
  EMPTY_XLSX_STYLES,
  parseSharedStrings,
  parseWorkbook,
  parseWorksheet,
  parseXlsxStyles,
} from '@/ooxml/spreadsheet';
import { bytesInclude } from '@/readers/docx-reader';
import {
  resolvePrintArea,
  resolvePrintTitleRows,
  sectionFromWorksheet,
  worksheetToBody,
} from '@/converter/xlsx-to-pdf';

const WORKBOOK_PART = 'xl/workbook.xml';
const SHARED_STRINGS_PART = 'xl/sharedStrings.xml';
const STYLES_PART = 'xl/styles.xml';
const CORE_PROPS_PART = 'docProps/core.xml';

export function readXlsx(xlsx: Uint8Array): ReadResult<FlowDoc> {
  const pkg = OpcPackage.open(xlsx);
  const workbookData = pkg.getPart(WORKBOOK_PART);
  if (!workbookData) throw new Error('Not a valid xlsx: missing xl/workbook.xml');
  const { sheets, date1904, definedNames } = parseWorkbook(workbookData);
  if (sheets.length === 0) throw new Error('xlsx has no sheets');

  const sharedStringsData = pkg.getPart(SHARED_STRINGS_PART);
  const sharedStrings = sharedStringsData ? parseSharedStrings(sharedStringsData) : [];

  const stylesData = pkg.getPart(STYLES_PART);
  const styles = stylesData ? parseXlsxStyles(stylesData) : EMPTY_XLSX_STYLES;

  const workbookRels = pkg.getPartRelationships(WORKBOOK_PART);
  const body: Array<BodyElement> = [];

  // Page geometry comes from the first sheet's <pageSetup>/<pageMargins>.
  // The renderer only supports one section, so subsequent sheets share the
  // first sheet's geometry; multi-section support is M2/M4-grade work.
  let firstSheetSection: SectionProperties | undefined;

  for (let sheetIdx = 0; sheetIdx < sheets.length; sheetIdx++) {
    const sheet = sheets[sheetIdx]!;
    const sheetRel = workbookRels.find((r) => r.id === sheet.relationshipId);
    if (!sheetRel) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, sheetRel);
    if (!resolved) continue;
    const worksheet = parseWorksheet(resolved.data);
    if (sheetIdx === 0) {
      firstSheetSection = sectionFromWorksheet(worksheet);
    }

    // Each sheet after the first starts on its own PDF page. We do NOT print the
    // sheet name: LibreOffice Calc / Excel `--convert-to pdf` emit it nowhere in
    // the body (nor the default header) — a synthetic title is pure extra text
    // that diverges from the print golden. So emit an empty page-break-only
    // paragraph (no runs ⇒ no glyphs ⇒ no text) for sheets > 0.
    if (sheetIdx > 0) {
      body.push({
        kind: 'paragraph',
        paragraph: { properties: { pageBreakBefore: true }, runs: [] },
      });
    }

    const printArea = resolvePrintArea(definedNames, sheetIdx);
    const titleRows = resolvePrintTitleRows(definedNames, sheetIdx);
    const gridLines = worksheet.printOptions?.gridLines === true;
    body.push(
      ...worksheetToBody(worksheet, sharedStrings, styles, date1904, {
        ...(printArea ? { printArea } : {}),
        ...(titleRows ? { titleRows } : {}),
        gridLines,
      }),
    );
  }

  const coreData = pkg.getPart(CORE_PROPS_PART);
  const coreProps = coreData ? parseCoreProperties(coreData) : undefined;
  const info = infoFromCore(coreProps);

  const doc: FlowDoc = {
    kind: 'flow',
    body,
    sections: [],
    ...(firstSheetSection ? { section: firstSheetSection } : {}),
    styles: EMPTY_STYLE_SHEET,
    resources: new ResourceStore(),
    ...(info ? { info } : {}),
  };
  return { doc, losses: [] };
}

export const xlsxReader: DocumentReader<FlowDoc> = {
  id: 'xlsx',
  produces: 'flow',
  supports: new Set([FEATURES.text, FEATURES.tables]),
  sniff: (bytes) =>
    bytes[0] === 0x50 && bytes[1] === 0x4b && bytesInclude(bytes, 'xl/workbook.xml'),
  read: (bytes) => readXlsx(bytes),
};

function infoFromCore(core: CoreProperties | undefined): DocumentInfo | undefined {
  if (!core) return undefined;
  return {
    ...(core.title ? { title: core.title } : {}),
    ...(core.creator ? { author: core.creator } : {}),
    ...(core.subject ? { subject: core.subject } : {}),
    ...(core.keywords ? { keywords: core.keywords } : {}),
    ...(core.created ? { creationDate: core.created } : {}),
    ...(core.modified ? { modificationDate: core.modified } : {}),
  };
}
