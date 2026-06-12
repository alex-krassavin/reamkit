// E-PDF EP2 — page text extraction. Builds the page's font map from its
// /Resources /Font dictionary, then runs the content-stream interpreter to get
// the positioned text runs. The bridge from raw objects (EP1) to text (EP2).

import { interpretContent } from './content';
import { buildContentFont } from './font';
import type { ContentFont, TextRun } from './content';
import type { PdfFile, PdfPage } from './document';

export function extractPageText(file: PdfFile, page: PdfPage): Array<TextRun> {
  const fonts = new Map<string, ContentFont>();
  if (page.resources) {
    const fontContainer = file.get(page.resources, 'Font');
    if (fontContainer instanceof Map) {
      for (const [fontName, fontRef] of fontContainer) {
        const fontDict = file.resolve(fontRef);
        if (fontDict instanceof Map) {
          try {
            fonts.set(fontName, buildContentFont(file, fontDict));
          } catch {
            // A malformed font is skipped — its text falls back to Latin-1.
          }
        }
      }
    }
  }
  return interpretContent(file.pageContent(page), fonts);
}
