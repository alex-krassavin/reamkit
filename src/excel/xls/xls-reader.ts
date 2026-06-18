// Legacy `.xls` reader (XLS-3) — the DocumentReader that sniffs a BIFF8 binary
// workbook (an OLE2/CFB container holding a `Workbook` stream) and reads it into
// the same SheetDoc the OOXML xlsx reader produces, so the whole render pipeline
// (projection → PDF/SVG/HTML, and even re-write to .xlsx) works on a 1997–2003
// `.xls`. Cell data, styling, images, charts and drawing shapes are all read; the
// remaining gap is the secondary sheet features (recorded as a loss).

import type { DocumentReader } from '@/core/ir/adapters';
import type { Loss } from '@/core/ir';
import type { SheetDoc } from '@/core/ir/sheet';

import { FEATURES } from '@/core/ir';
import { isCfb, openCfb } from '@/core/ole/cfb';
import { readXlsToSheetDoc } from '@/excel/xls/biff-reader';

// A `.xls` is an OLE2 container with a `Workbook` (BIFF8) or `Book` (BIFF5)
// stream. The check also keeps a `.doc`/`.ppt` (or an encrypted OOXML, which is
// also a CFB but holds `EncryptedPackage`) from mis-routing here.
function looksLikeXls(bytes: Uint8Array): boolean {
  if (!isCfb(bytes)) return false;
  try {
    const cfb = openCfb(bytes);
    return cfb.hasStream('Workbook') || cfb.hasStream('Book');
  } catch {
    return false;
  }
}

// A `.xls` sheet's cell data, styling, images, charts, drawing shapes and its
// secondary features (conditional formatting — including the 2007 colour-scale /
// data-bar / icon-set extensions —, comments, hyperlinks, data validation, defined
// names, the page-setup print model, frozen panes and custom row heights) are read;
// what remains is the 2007 Excel table / autofilter feature, whose shared-feature
// payload is undocumented enough that Apache POI itself does not parse it — so it is
// left rather than guessed. Reported so a caller's loss report is honest about it.
const XLS_LOSSES: ReadonlyArray<Loss> = [
  {
    severity: 'degraded',
    feature: FEATURES.cellFormatting,
    detail:
      'legacy .xls Excel tables / autofilter are not read; the sheet’s cell data, styling, images, charts, drawing shapes, conditional formatting, comments, hyperlinks, data validation, defined names, the page-setup print model, frozen panes and custom row heights are',
  },
];

export const xlsReader: DocumentReader<SheetDoc> = {
  id: 'xls',
  produces: 'sheet',
  supports: new Set([FEATURES.text, FEATURES.tables]),
  sniff: looksLikeXls,
  read: (bytes) => ({ doc: readXlsToSheetDoc(bytes), losses: XLS_LOSSES }),
};
