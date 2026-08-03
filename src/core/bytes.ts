// Byte-level helpers shared by format sniffers and writers.

/**
 * Base64-encode raw bytes for `data:` URIs (the svg + html writers). `btoa` is
 * available in browsers, workers and Node 16+; chunking keeps the intermediate
 * binary string small.
 */
export function toBase64(bytes: Uint8Array): string {
  let bin = '';
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin);
}

/**
 * Whether the ZIP in `bytes` holds a part named `partName` — the name exactly,
 * or anything under it when `partName` names a directory (ends with `/`).
 *
 * This is what a reader sniff should ask. A substring probe cannot tell a
 * package from one EMBEDDED in it: a deck with a chart carries the chart's
 * workbook as `ppt/embeddings/*.xlsx` STORED, uncompressed, so that workbook's
 * own `xl/workbook.xml` sits in the outer file's bytes verbatim — and the xlsx
 * sniff, which runs before the pptx one, claimed thirteen corpus presentations
 * and then failed on them. The ZIP's central directory lists what is actually
 * in this package, and walking that entry table is cheaper than scanning every
 * byte for a needle.
 *
 * Names are compared without case, as OPC compares part names (§9.1.1.1) and
 * as the package reader behind the sniff resolves them.
 *
 * @param bytes    The package bytes.
 * @param partName The part name, or a `/`-terminated directory prefix.
 * @returns Whether the package's directory names it.
 */
export function packageHasPart(bytes: Uint8Array, partName: string): boolean {
  const names = zipEntryNames(bytes);
  // No central directory to read: not an archive any reader here can open
  // either, so keep the old probe's answer and let the reader say why.
  if (names === undefined) return bytesIncludePartName(bytes, partName);
  const wanted = partName.toLowerCase();
  const isDir = wanted.endsWith('/');
  return names.some((raw) => {
    const name = raw.replace(/\\/g, '/').replace(/^\//u, '').toLowerCase();
    return isDir ? name.startsWith(wanted) : name === wanted;
  });
}

/**
 * The names in a ZIP's central directory (APPNOTE §4.3.12), or `undefined`
 * when it cannot be located or is malformed — including the ZIP64 spelling of
 * an archive with more than 65 534 entries.
 */
function zipEntryNames(bytes: Uint8Array): Array<string> | undefined {
  if (bytes.length < 22) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // §4.3.16 — the end record is last, behind at most 64 KB of archive comment.
  const floor = Math.max(0, bytes.length - 22 - 0xffff);
  let end = -1;
  for (let i = bytes.length - 22; i >= floor; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      end = i;
      break;
    }
  }
  if (end < 0) return undefined;
  let count = view.getUint16(end + 10, true);
  let start = view.getUint32(end + 16, true);
  if (count === 0xffff || start === 0xffffffff) {
    // §4.3.15 — the ZIP64 locator sits immediately before the end record, and
    // points at the ZIP64 end record holding the real count and offset.
    const locator = end - 20;
    if (locator < 0 || view.getUint32(locator, true) !== 0x07064b50) return undefined;
    const z64 = Number(view.getBigUint64(locator + 8, true));
    if (z64 + 56 > bytes.length || view.getUint32(z64, true) !== 0x06064b50) return undefined;
    count = Number(view.getBigUint64(z64 + 32, true));
    start = Number(view.getBigUint64(z64 + 48, true));
  }
  const names: Array<string> = [];
  let p = start;
  for (let i = 0; i < count; i++) {
    // §4.3.12 — each header is 46 bytes plus the name, the extra field and the
    // comment. A signature that is not the directory's means the offset lied.
    if (p + 46 > bytes.length || view.getUint32(p, true) !== 0x02014b50) return undefined;
    const nameLen = view.getUint16(p + 28, true);
    if (p + 46 + nameLen > bytes.length) return undefined;
    names.push(decodeUtf8(bytes.subarray(p + 46, p + 46 + nameLen)));
    p += 46 + nameLen + view.getUint16(p + 30, true) + view.getUint16(p + 32, true);
  }
  return names;
}

const utf8 = new TextDecoder();

/** An entry name as text. Names are UTF-8 or CP437; both decode ASCII alike. */
function decodeUtf8(bytes: Uint8Array): string {
  return utf8.decode(bytes);
}

/**
 * Scan raw package bytes for an OPC part name (e.g. `'xl/workbook.xml'`)
 * without unzipping. Prefer {@link packageHasPart}, which asks the archive
 * what it holds instead of hoping the name appears nowhere else. Accepts both
 * the spec's `/` separator and the `\` that Windows producers write.
 */
export function bytesIncludePartName(haystack: Uint8Array, partName: string): boolean {
  if (bytesInclude(haystack, partName)) return true;
  // APPNOTE §4.4.17.1 mandates '/', but Windows producers store backslashes
  // (corpus: tdf76115.xlsx). OpcPackage normalizes them, so a sniff that only
  // knows the spec spelling refuses documents the reader behind it can read.
  return partName.includes('/') && bytesInclude(haystack, partName.replace(/\//g, '\\'));
}

/**
 * Naive scan for an ASCII `needle` inside raw `haystack` bytes. Prefer
 * {@link bytesIncludePartName} for OPC part names — it also accepts the
 * backslash spelling real archives use.
 */
export function bytesInclude(haystack: Uint8Array, needle: string): boolean {
  const n = new TextEncoder().encode(needle);
  outer: for (let i = 0; i + n.length <= haystack.length; i++) {
    for (let j = 0; j < n.length; j++) {
      if (haystack[i + j] !== n[j]) continue outer;
    }
    return true;
  }
  return false;
}
