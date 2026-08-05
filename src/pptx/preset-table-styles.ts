// §20.1.4.2.24 — the table styles PowerPoint has BUILT IN.
//
// `ppt/tableStyles.xml` lists only what a deck customises. A table that wears
// one of the gallery's own styles names it by GUID and ships no definition at
// all, because every reader is expected to hold them: eight decks of the corpus
// do that, and `predefined-table-style` and `tdf149785` came out as bare text
// with no fill and no rule anywhere.
//
// The definitions are the ones PowerPoint itself writes when a deck DOES embed
// them — lifted from decks in the corpus that do, so they are the real thing
// and not a reconstruction. They are theme-relative by construction (`accent1`,
// `lt1`, `dk1`), which is what lets one definition serve every deck: the same
// style is blue under one theme and green under another.
//
// Two GUIDs the corpus references are NOT here, because no deck in it embeds
// them and inventing a definition would be worse than drawing none —
// `{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}` and
// `{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}`.

import type { PoNode } from '@/core/po-helpers';
import { poAttr, poIs } from '@/core/po-helpers';

/** The gallery styles this reader knows, by the GUID a table names. */
const PRESETS: ReadonlyMap<string, string> = new Map([
  // Medium Style 2 - Accent 1
  [
    '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
    '<a:tblStyle styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" styleName="Medium Style 2 - Acc' +
      'ent 1"><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><' +
      'a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:ln w="12700" cmpd="sng">' +
      '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:left><a:right><a:ln w="12700" ' +
      'cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:right><a:top><a:ln ' +
      'w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top><a:bo' +
      'ttom><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></' +
      'a:bottom><a:insideH><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:soli' +
      'dFill></a:ln></a:insideH><a:insideV><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val' +
      '="lt1"/></a:solidFill></a:ln></a:insideV></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="ac' +
      'cent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl><a' +
      ':band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40' +
      '000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H><a:band2H><a:tcStyle><a:tc' +
      'Bdr/></a:tcStyle></a:band2H><a:band1V><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr' +
      ' val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:ban' +
      'd1V><a:band2V><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2V><a:lastCol><a:tcTxStyle b="on"><a:' +
      'fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyl' +
      'e><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fi' +
      'll></a:tcStyle></a:lastCol><a:firstCol><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr' +
      ' val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill' +
      '><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstCol><a' +
      ':lastRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:s' +
      'chemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:top><a:ln w="38100" cmpd="sng"><a:s' +
      'olidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top></a:tcBdr><a:fill><a:solidFill' +
      '><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastRow><a:firstRow><a:t' +
      'cTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val=' +
      '"lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:bottom><a:ln w="38100" cmpd="sng"><a:solidFill><' +
      'a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:sch' +
      'emeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow></a:tblStyle>',
  ],
  // Medium Style 2
  [
    '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}',
    '<a:tblStyle styleId="{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}" styleName="Medium Style 2"><a:w' +
      'holeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr' +
      ' val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:ln w="12700" cmpd="sng"><a:solidFil' +
      'l><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:left><a:right><a:ln w="12700" cmpd="sng">' +
      '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:right><a:top><a:ln w="12700" c' +
      'mpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top><a:bottom><a:ln ' +
      'w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom><a' +
      ':insideH><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:l' +
      'n></a:insideH><a:insideV><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a' +
      ':solidFill></a:ln></a:insideV></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="dk1"><a:tint ' +
      'val="20000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl><a:band1H><a:tcSt' +
      'yle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="dk1"><a:tint val="40000"/></a:schemeClr' +
      '></a:solidFill></a:fill></a:tcStyle></a:band1H><a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></' +
      'a:band2H><a:band1V><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="dk1"><a:tint ' +
      'val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1V><a:band2V><a:tcStyl' +
      'e><a:tcBdr/></a:tcStyle></a:band2V><a:lastCol><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:' +
      'prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/>' +
      '<a:fill><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:fill></a:tcStyle></a:lastCol>' +
      '<a:firstCol><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><' +
      'a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr ' +
      'val="dk1"/></a:solidFill></a:fill></a:tcStyle></a:firstCol><a:lastRow><a:tcTxStyle b="on"><a' +
      ':fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxSty' +
      'le><a:tcStyle><a:tcBdr><a:top><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"' +
      '/></a:solidFill></a:ln></a:top></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="dk1"/></a:so' +
      'lidFill></a:fill></a:tcStyle></a:lastRow><a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="min' +
      'or"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:' +
      'tcBdr><a:bottom><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFil' +
      'l></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a' +
      ':fill></a:tcStyle></a:firstRow></a:tblStyle>',
  ],
  // No Style, Table Grid
  [
    '{5940675A-B579-460E-94D1-54222C63F5DA}',
    '<a:tblStyle styleId="{5940675A-B579-460E-94D1-54222C63F5DA}" styleName="No Style, Table Grid' +
      '"><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef' +
      '><a:schemeClr val="tx1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:ln w="12700" cmpd="sng' +
      '"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:left><a:right><a:ln w="12700' +
      '" cmpd="sng"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:right><a:top><a:l' +
      'n w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:top><a:' +
      'bottom><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>' +
      '</a:bottom><a:insideH><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="tx1"/></a:so' +
      'lidFill></a:ln></a:insideH><a:insideV><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr v' +
      'al="tx1"/></a:solidFill></a:ln></a:insideV></a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle' +
      '></a:wholeTbl></a:tblStyle>',
  ],
  // No Style, No Grid
  [
    '{2D5ABB26-0587-4C30-8999-92F81FD0307C}',
    '<a:tblStyle styleId="{2D5ABB26-0587-4C30-8999-92F81FD0307C}" styleName="No Style, No Grid"><' +
      'a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef><a' +
      ':schemeClr val="tx1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:ln><a:noFill/></a:ln></a:' +
      'left><a:right><a:ln><a:noFill/></a:ln></a:right><a:top><a:ln><a:noFill/></a:ln></a:top><a:bo' +
      'ttom><a:ln><a:noFill/></a:ln></a:bottom><a:insideH><a:ln><a:noFill/></a:ln></a:insideH><a:in' +
      'sideV><a:ln><a:noFill/></a:ln></a:insideV></a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle>' +
      '</a:wholeTbl></a:tblStyle>',
  ],
]);

// Parsed on first use and kept: a deck with many tables asks repeatedly, and
// the answer is the same every time.
const parsed = new Map<string, PoNode | undefined>();

/**
 * The built-in table style a GUID names, or `undefined` when it is not one this
 * reader holds.
 *
 * @param styleId The `a:tableStyleId` a table states.
 * @param parse   The module's XML reader, injected so this file needs none.
 */
export function presetTableStyle(
  styleId: string,
  parse: (xml: string) => ReadonlyArray<PoNode>,
): PoNode | undefined {
  const seen = parsed.get(styleId);
  if (seen !== undefined || parsed.has(styleId)) return seen;
  const xml = PRESETS.get(styleId);
  const style = xml
    ? parse(xml).find((n) => poIs(n, 'a:tblStyle') && poAttr(n, 'styleId') === styleId)
    : undefined;
  parsed.set(styleId, style);
  return style;
}
