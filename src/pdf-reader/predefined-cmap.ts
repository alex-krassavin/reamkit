// §9.7.5.2 — the CMaps a file names instead of embedding.
//
// A composite font's `/Encoding` may be a NAME — `90ms-RKSJ-V`, `GBK-EUC-H`,
// `UniJIS-UCS2-H` — and the name stands for a mapping Adobe published, not for
// anything in the file. Read as `Identity-H` (two bytes, code = CID) a Japanese
// page comes apart: issue11555.pdf shows `<6162632082a082a282a4>`, which is
// "abc " in one-byte codes and あいう in two, and split down the middle it was
// six codes of nonsense.
//
// The way through is not the CID tables. A predefined CMap's codes ARE the
// bytes of a known character encoding — RKSJ is Shift-JIS, EUC is EUC-JP, B5 is
// Big5 — so the string decodes with a `TextDecoder` and the CIDs never come
// into it. What the CMap is still needed for is how many BYTES each code takes,
// which in every one of these is decided by the leading byte.
//
// Only the encodings a `TextDecoder` carries are here. The rest keep the
// Identity reading and the loss the reader already reports for a font that
// cannot say what its characters are.

/** How a named CMap breaks a string into codes, and what those codes mean. */
export interface PredefinedCMap {
  /** The `TextDecoder` label the code bytes are in. */
  readonly charset: string;
  /** Whether a byte begins a TWO-byte code; anything else stands alone. */
  readonly leadsPair: (byte: number) => boolean;
  /** §9.7.5.2 — a `…-V` CMap sets its text down the page. */
  readonly vertical: boolean;
}

/** Shift-JIS: 81–9F and E0–FC lead a pair; 00–80 and A0–DF stand alone. */
const shiftJis = (b: number): boolean => (b >= 0x81 && b <= 0x9f) || (b >= 0xe0 && b <= 0xfc);

/** The EUC family: A1–FE lead a pair, and 8E/8F introduce one as well. */
const euc = (b: number): boolean => b >= 0xa1 || b === 0x8e || b === 0x8f;

/** Big5, GBK and UHC: everything from 81 up leads a pair. */
const highLeads = (b: number): boolean => b >= 0x81;

/** Two bytes always, big-endian — the UCS-2 and UTF-16 CMaps. */
const always = (): boolean => true;

/** The families, longest name first so `90ms-RKSJ` is not read as `RKSJ`. */
const FAMILIES: ReadonlyArray<{
  match: RegExp;
  charset: string;
  leadsPair: (b: number) => boolean;
}> = [
  { match: /RKSJ/u, charset: 'shift_jis', leadsPair: shiftJis },
  { match: /-EUC(-|$)|^EUC-/u, charset: 'euc-jp', leadsPair: euc },
  { match: /^(ETen|ETenms|B5pc|HKscs-B5|CNS-EUC)/u, charset: 'big5', leadsPair: highLeads },
  { match: /^(GBK|GBpc|GBK2K|GB-EUC|GBKp)/u, charset: 'gbk', leadsPair: highLeads },
  { match: /^(KSC|KSCms|KSCpc)/u, charset: 'euc-kr', leadsPair: highLeads },
  { match: /UCS2|UTF16/u, charset: 'utf-16be', leadsPair: always },
];

/**
 * What a predefined CMap name comes to.
 *
 * @param name The `/Encoding` name, as the font states it.
 * @returns Its encoding and code structure, or `undefined` for `Identity` and
 *          for any name whose encoding this cannot decode.
 */
export function predefinedCMap(name: string): PredefinedCMap | undefined {
  if (name.startsWith('Identity')) return undefined;
  const vertical = name.endsWith('-V');
  // The EUC families name themselves with the ordering first — `KSCms-UHC-H`
  // is EUC-KR's UHC — so a plain suffix test is not enough.
  const family = FAMILIES.find((f) => f.match.test(name));
  if (!family) return undefined;
  if (!canDecode(family.charset)) return undefined;
  return { charset: family.charset, leadsPair: family.leadsPair, vertical };
}

/** Split a shown string the way this CMap's codespace says (§9.7.6.2). */
export function splitPredefined(cmap: PredefinedCMap, bytes: Uint8Array): Array<number> {
  const out: Array<number> = [];
  for (let i = 0; i < bytes.length; ) {
    const b = bytes[i]!;
    if (cmap.leadsPair(b) && i + 1 < bytes.length) {
      out.push((b << 8) | bytes[i + 1]!);
      i += 2;
    } else {
      out.push(b);
      i++;
    }
  }
  return out;
}

/** The character one code stands for: its own bytes, read in the CMap's encoding. */
export function decodePredefined(cmap: PredefinedCMap, code: number): string {
  const bytes = code > 0xff ? Uint8Array.from([code >> 8, code & 0xff]) : Uint8Array.from([code]);
  try {
    const text = new TextDecoder(cmap.charset, { fatal: false }).decode(bytes);
    // A byte pair the encoding does not know comes back as U+FFFD, which is
    // what the reconstruction already counts as unrecoverable.
    return text;
  } catch {
    return '';
  }
}

/** Whether this runtime's `TextDecoder` carries the legacy encoding. */
function canDecode(charset: string): boolean {
  const had = supported.get(charset);
  if (had !== undefined) return had;
  let ok = false;
  try {
    // A runtime built without the full encoding set throws on construction.
    new TextDecoder(charset);
    ok = true;
  } catch {
    ok = false;
  }
  supported.set(charset, ok);
  return ok;
}

const supported = new Map<string, boolean>();
