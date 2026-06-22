// ISO 32000-1:2008 §9.6.2.2 — The Standard Type 1 Fonts (14).
// These fonts are guaranteed to be available in every conforming PDF reader
// and do not need to be embedded.

/** One of the 14 Standard Type 1 font names (ISO 32000-1 §9.6.2.2). */
export type BuiltinFontName =
  | 'Helvetica'
  | 'Helvetica-Bold'
  | 'Helvetica-Oblique'
  | 'Helvetica-BoldOblique'
  | 'Times-Roman'
  | 'Times-Bold'
  | 'Times-Italic'
  | 'Times-BoldItalic'
  | 'Courier'
  | 'Courier-Bold'
  | 'Courier-Oblique'
  | 'Courier-BoldOblique'
  | 'Symbol'
  | 'ZapfDingbats';

/**
 * The 14 Standard Type 1 fonts, guaranteed available in every conforming PDF
 * reader and never embedded.
 */
export const BUILTIN_FONTS: ReadonlyArray<BuiltinFontName> = [
  'Helvetica',
  'Helvetica-Bold',
  'Helvetica-Oblique',
  'Helvetica-BoldOblique',
  'Times-Roman',
  'Times-Bold',
  'Times-Italic',
  'Times-BoldItalic',
  'Courier',
  'Courier-Bold',
  'Courier-Oblique',
  'Courier-BoldOblique',
  'Symbol',
  'ZapfDingbats',
];
