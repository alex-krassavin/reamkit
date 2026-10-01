// A workbook as images of its sheets — what the SVG target draws: each sheet
// whole on a page of its own, laid out as its window shows it and cut down to
// what is drawn on it, rather than the pages it prints on.

import type { SheetDoc } from '@/core/ir/sheet';
import type { ProjectSheetOptions } from '@/excel/sheet-to-flow';
import type { LaidOutDocument } from '@/layout/page-doc';
import type { StyledRenderOptions } from '@/pdf';
import { flowRenderOptions } from '@/core/converter/project';
import { projectSheetDoc } from '@/excel/sheet-to-flow';
import { fitPagesToContent } from '@/layout/fit-pages';
import { layoutStyledDocument } from '@/layout/styled-layout';

/** The room kept past a sheet's content, as the layout keeps it before it. */
const SHEET_IMAGE_PAD_PT = 2;

/**
 * Lay a workbook out as images of its sheets: the screen projection, one page a
 * sheet with room for all of it, each page then cut to its content.
 *
 * @param sheet   The workbook.
 * @param faces   The faces the sheets are measured and drawn in: the registry,
 *                and one per family where the document names several.
 * @param options The projection's other knobs (reference date, file name, the
 *                render face's digit width).
 * @param render  The caller's own layout options, over the document's.
 * @returns One page per visible sheet, each as large as what it holds.
 */
export function layoutSheetImages(
  sheet: SheetDoc,
  faces: Pick<StyledRenderOptions, 'registry' | 'registriesByFamily'>,
  options: Omit<ProjectSheetOptions, 'screen' | 'wholeSheetPages'> = {},
  render: Partial<StyledRenderOptions> = {},
): LaidOutDocument {
  const flow = projectSheetDoc(sheet, { ...options, screen: true, wholeSheetPages: true });
  return fitPagesToContent(
    layoutStyledDocument(flow.body, { ...faces, ...flowRenderOptions(flow), ...render }),
    SHEET_IMAGE_PAD_PT,
  );
}
