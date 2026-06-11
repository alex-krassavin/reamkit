// Byte-level helpers shared by format sniffers.

// Naive scan for an ASCII needle inside raw bytes — used by reader sniffs to
// spot OPC part names (e.g. 'word/document.xml') without unzipping.
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
