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
 * Naive scan for an ASCII `needle` inside raw `haystack` bytes — used by reader
 * sniffs to spot OPC part names (e.g. `'word/document.xml'`) without unzipping.
 */
/**
 * Scan raw package bytes for an OPC part name (e.g. `'xl/workbook.xml'`)
 * without unzipping — the reader sniffs' cheap format probe. Accepts both the
 * spec's `/` separator and the `\` that Windows producers write.
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
