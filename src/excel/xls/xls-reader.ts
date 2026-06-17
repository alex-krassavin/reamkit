// Legacy `.xls` reader (XLS-3) — the DocumentReader that sniffs a BIFF8 binary
// workbook (an OLE2/CFB container holding a `Workbook` stream) and reads it into
// the same SheetDoc the OOXML xlsx reader produces, so the whole render pipeline
// (projection → PDF/SVG/HTML, and even re-write to .xlsx) works on a 1997–2003
// `.xls`. Styling and embedded drawings are not read yet — recorded as losses.

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

// What the BIFF8 reader does not surface yet — reported so a caller's loss report
// is honest about the gap. Cell content AND styling (fonts/fills/borders/number
// formats/alignment) are read; embedded drawings are not.
const XLS_LOSSES: ReadonlyArray<Loss> = [
  {
    severity: 'dropped',
    feature: FEATURES.charts,
    detail: 'legacy .xls embedded charts, images and drawing objects are not read',
  },
];

export const xlsReader: DocumentReader<SheetDoc> = {
  id: 'xls',
  produces: 'sheet',
  supports: new Set([FEATURES.text, FEATURES.tables]),
  sniff: looksLikeXls,
  read: (bytes) => ({ doc: readXlsToSheetDoc(bytes), losses: XLS_LOSSES }),
};
