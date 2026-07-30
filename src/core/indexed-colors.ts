// ECMA-376 §18.8.27 — the legacy indexed colour palette.
//
// It is shared: a cell's `<color indexed="N">` names it, and so does a number
// format's `[Color N]` (§18.8.31), which is why it lives here rather than in
// the SpreadsheetML style parser that used to own it.

/**
 * ECMA-376 §18.8.27 — the legacy indexed colour palette a `<color indexed="N">`
 * refers to. It predates `rgb` and theme colours and is still what Excel writes
 * for anything picked from the classic 40-colour dropdown, so a file using it
 * is not unusual: tdf58243.xlsx colours its header cells `indexed="10"`, and
 * ignoring the attribute rendered them black where every other reader shows red.
 *
 * Entries 0-7 are repeated at 8-15 (a quirk of the original palette); 64 and 65
 * are the system foreground and background.
 */
export const INDEXED_COLORS: ReadonlyArray<string> = [
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'C0C0C0',
  '808080',
  '9999FF',
  '993366',
  'FFFFCC',
  'CCFFFF',
  '660066',
  'FF8080',
  '0066CC',
  'CCCCFF',
  '000080',
  'FF00FF',
  'FFFF00',
  '00FFFF',
  '800080',
  '800000',
  '008080',
  '0000FF',
  '00CCFF',
  'CCFFFF',
  'CCFFCC',
  'FFFF99',
  '99CCFF',
  'FF99CC',
  'CC99FF',
  'FFCC99',
  '3366FF',
  '33CCCC',
  '99CC00',
  'FFCC00',
  'FF9900',
  'FF6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
  '000000',
  'FFFFFF',
];
