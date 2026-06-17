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

// The visual content of a `.xls` — cell values, styling, images, charts and
// drawing shapes — is read; what remains are the secondary sheet features, which
// have BIFF records the reader does not yet parse. Reported so a caller's loss
// report is honest about the gap.
const XLS_LOSSES: ReadonlyArray<Loss> = [
  {
    severity: 'degraded',
    feature: FEATURES.cellFormatting,
    detail:
      "legacy .xls secondary features (conditional formatting, comments, hyperlinks, data validation, defined names, the page-setup print model) are not read; the sheet's cell data, styling, images, charts and drawing shapes are",
  },
];

export const xlsReader: DocumentReader<SheetDoc> = {
  id: 'xls',
  produces: 'sheet',
  supports: new Set([FEATURES.text, FEATURES.tables]),
  sniff: looksLikeXls,
  read: (bytes) => ({ doc: readXlsToSheetDoc(bytes), losses: XLS_LOSSES }),
};
