// Annex D.2 — the Latin-text encodings a simple font's codes are read through.
//
// A font says what its codes mean in three ways, and a reader that knows only
// the last two is wrong about the rest: `/ToUnicode`, `/Encoding /Differences`
// — and the BASE encoding underneath them, which is either named in the font
// dictionary or is the one built into the face. Read as Latin-1, which is all
// that is left without these tables, ZapfDingbats.pdf's title —
// "Character Sets — Zapf Dingbats", set in Times with no `/Encoding` at all —
// came back as "Character Sets Ð Zapf Dingbats": 0xD0 is Eth in Latin-1 and an
// em dash in StandardEncoding.
//
// Codes 0x20–0x7E are ASCII in all three (StandardEncoding's two quotes aside),
// and WinAnsiEncoding's upper half is Latin-1 by construction. So the tables
// hold what DIFFERS from Latin-1, and a code no table names keeps that reading.

/** One of the base encodings this reader has a table for (Annex D.2). */
export type BaseEncodingName = 'StandardEncoding' | 'WinAnsiEncoding' | 'MacRomanEncoding';

/**
 * Code → glyph name, from names read in code order. `.` marks a code the
 * encoding leaves unused, and the rows are the spec's own — sixteen to a line,
 * so a row can be checked against Annex D at a glance.
 */
function table(start: number, ...rows: ReadonlyArray<string>): Map<number, string> {
  const out = new Map<number, string>();
  let code = start;
  for (const row of rows) {
    for (const name of row.split(/\s+/u).filter(Boolean)) {
      if (name !== '.') out.set(code, name);
      code++;
    }
  }
  return out;
}

// Annex D.2 — the low half, which all three share. It says nothing Latin-1 does
// not, and it is here for the other thing a base encoding answers: which GLYPH
// a code selects in a program addressed by name (§9.6.6). A legacy eight-bit
// face is reached that way and no other.
const ASCII = table(
  0x20,
  'space exclam quotedbl numbersign dollar percent ampersand quotesingle parenleft parenright asterisk plus comma hyphen period slash',
  'zero one two three four five six seven eight nine colon semicolon less equal greater question',
  'at A B C D E F G H I J K L M N O',
  'P Q R S T U V W X Y Z bracketleft backslash bracketright asciicircum underscore',
  'grave a b c d e f g h i j k l m n o',
  'p q r s t u v w x y z braceleft bar braceright asciitilde .',
);

/** Annex D.2 STD — the built-in encoding of the standard Latin text faces. */
const STANDARD: ReadonlyMap<number, string> = new Map([
  ...ASCII,
  // The typewriter apostrophe and grave are the typographic quotes here, which
  // is how a TeX page writes `don't` and `‘this’`.
  ...table(0x27, 'quoteright'),
  ...table(0x60, 'quoteleft'),
  ...table(
    0xa0,
    '. exclamdown cent sterling fraction yen florin section currency quotesingle quotedblleft guillemotleft guilsinglleft guilsinglright fi fl',
    '. endash dagger daggerdbl periodcentered . paragraph bullet quotesinglbase quotedblbase quotedblright guillemotright ellipsis perthousand . questiondown',
    '. grave acute circumflex tilde macron breve dotaccent dieresis . ring cedilla . hungarumlaut ogonek caron',
    'emdash . . . . . . . . . . . . . . .',
    '. AE . ordfeminine . . . . Lslash Oslash OE ordmasculine . . . .',
    '. ae . . . dotlessi . . lslash oslash oe germandbls . . . .',
  ),
]);

/** Annex D.2 WIN — CP-1252. Only its punctuation block departs from Latin-1. */
const WIN_ANSI: ReadonlyMap<number, string> = new Map([
  ...ASCII,
  ...table(
    0x80,
    'Euro . quotesinglbase florin quotedblbase ellipsis dagger daggerdbl circumflex perthousand Scaron guilsinglleft OE . Zcaron .',
    '. quoteleft quoteright quotedblleft quotedblright bullet endash emdash tilde trademark scaron guilsinglright oe . zcaron Ydieresis',
  ),
  // Annex D.2 names 0xA0 `space` and 0xAD `hyphen` — the drawn glyphs, since a
  // PDF has no line to break. Left to Latin-1 they read as the no-break space
  // and the SOFT hyphen, which say the same about the ink and more about the
  // text: the soft hyphen is how a producer marks a word broken across a line,
  // and it is what the reconstruction joins that word back together on.
]);

/** Annex D.2 MAC — Mac OS Roman, as the PDF variant states it. */
const MAC_ROMAN: ReadonlyMap<number, string> = new Map([
  ...ASCII,
  ...table(
    0x80,
    'Adieresis Aring Ccedilla Eacute Ntilde Odieresis Udieresis aacute agrave acircumflex adieresis atilde aring ccedilla eacute egrave',
    'ecircumflex edieresis iacute igrave icircumflex idieresis ntilde oacute ograve ocircumflex odieresis otilde uacute ugrave ucircumflex udieresis',
    'dagger degree cent sterling section bullet paragraph germandbls registered copyright trademark acute dieresis notequal AE Oslash',
    'infinity plusminus lessequal greaterequal yen mu partialdiff summation product pi integral ordfeminine ordmasculine Omega ae oslash',
    'questiondown exclamdown logicalnot radical florin approxequal Delta guillemotleft guillemotright ellipsis space Agrave Atilde Otilde OE oe',
    'endash emdash quotedblleft quotedblright quoteleft quoteright divide lozenge ydieresis Ydieresis fraction currency guilsinglleft guilsinglright fi fl',
    'daggerdbl periodcentered quotesinglbase quotedblbase perthousand Acircumflex Ecircumflex Aacute Edieresis Egrave Iacute Icircumflex Idieresis Igrave Oacute Ocircumflex',
    '. Ograve Uacute Ucircumflex Ugrave dotlessi circumflex tilde macron breve dotaccent ring cedilla hungarumlaut ogonek caron',
  ),
]);

const TABLES: Readonly<Record<BaseEncodingName, ReadonlyMap<number, string>>> = {
  StandardEncoding: STANDARD,
  WinAnsiEncoding: WIN_ANSI,
  MacRomanEncoding: MAC_ROMAN,
};

/**
 * The glyph names of a base encoding, or `undefined` for a name this has no
 * table for — `/MacExpertEncoding`, whose repertory no reader here can set.
 *
 * @param name The `/Encoding` or `/BaseEncoding` name, without the slash.
 */
export function baseEncodingTable(name: string): ReadonlyMap<number, string> | undefined {
  return name in TABLES ? TABLES[name as BaseEncodingName] : undefined;
}

/** The encoding built into the standard Latin text faces (§9.6.2.2). */
export function standardEncodingTable(): ReadonlyMap<number, string> {
  return STANDARD;
}

// §9.6.2.2 Table 109 — the fourteen, by their PostScript names. Only an exact
// name counts: a file that says `Arial` and embeds nothing is asking for a
// substitute, and what a substitute's codes mean is whatever the producer
// assumed — nearly always Latin-1. The four the reader answers for with
// StandardEncoding are the Latin text faces; Symbol and ZapfDingbats have
// built-in encodings of their own (see `./dingbats`).
const STANDARD_LATIN_FACES = new Set([
  'courier',
  'courier-bold',
  'courier-oblique',
  'courier-boldoblique',
  'helvetica',
  'helvetica-bold',
  'helvetica-oblique',
  'helvetica-boldoblique',
  'times-roman',
  'times-bold',
  'times-italic',
  'times-bolditalic',
]);

// The Macintosh standard glyph ORDER — the names a `post` table means by an
// index below 258 (§post, format 2.0). Three names of its own, then the
// MacRoman repertory in code order, then a tail Apple appended later. A legacy
// eight-bit font is reached through this and nothing else: its program carries
// no `cmap`, and the shapes it draws are whatever the foundry put under those
// names — Masis, an Armenian face, draws ի under `i`.
const MAC_ORDER = `
.notdef .null nonmarkingreturn space exclam quotedbl numbersign dollar
percent ampersand quotesingle parenleft parenright asterisk plus comma
hyphen period slash zero one two three four
five six seven eight nine colon semicolon less
equal greater question at A B C D
E F G H I J K L
M N O P Q R S T
U V W X Y Z bracketleft backslash
bracketright asciicircum underscore grave a b c d
e f g h i j k l
m n o p q r s t
u v w x y z braceleft bar
braceright asciitilde Adieresis Aring Ccedilla Eacute Ntilde Odieresis
Udieresis aacute agrave acircumflex adieresis atilde aring ccedilla
eacute egrave ecircumflex edieresis iacute igrave icircumflex idieresis
ntilde oacute ograve ocircumflex odieresis otilde uacute ugrave
ucircumflex udieresis dagger degree cent sterling section bullet
paragraph germandbls registered copyright trademark acute dieresis notequal
AE Oslash infinity plusminus lessequal greaterequal yen mu
partialdiff summation product pi integral ordfeminine ordmasculine Omega
ae oslash questiondown exclamdown logicalnot radical florin approxequal
Delta guillemotleft guillemotright ellipsis nonbreakingspace Agrave Atilde Otilde
OE oe endash emdash quotedblleft quotedblright quoteleft quoteright
divide lozenge ydieresis Ydieresis fraction currency guilsinglleft guilsinglright
fi fl daggerdbl periodcentered quotesinglbase quotedblbase perthousand Acircumflex
Ecircumflex Aacute Edieresis Egrave Iacute Icircumflex Idieresis Igrave
Oacute Ocircumflex apple Ograve Uacute Ucircumflex Ugrave dotlessi
circumflex tilde macron breve dotaccent ring cedilla hungarumlaut
ogonek caron Lslash lslash Scaron scaron Zcaron zcaron
brokenbar Eth eth Yacute yacute Thorn thorn minus
multiply onesuperior twosuperior threesuperior onehalf onequarter threequarters franc
Gbreve gbreve Idotaccent Scedilla scedilla Cacute cacute Ccaron
ccaron dcroat
`
  .split(/\s+/u)
  .filter(Boolean);

/**
 * The name a `post` table's glyph-name index below 258 stands for.
 *
 * @param index The index as the table states it.
 * @returns The name, or `undefined` for an index outside the standard order.
 */
export function macGlyphName(index: number): string | undefined {
  return MAC_ORDER[index];
}

/**
 * Whether a `/BaseFont` names one of the standard Latin faces, whose built-in
 * encoding is StandardEncoding.
 *
 * @param baseFont The name as the font dictionary states it.
 */
export function isStandardLatinFace(baseFont: string): boolean {
  return STANDARD_LATIN_FACES.has(
    baseFont
      .trim()
      .replace(/^[A-Za-z]{6}\+/u, '')
      .toLowerCase(),
  );
}
