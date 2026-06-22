// Big-endian binary reader. TrueType / OpenType (ISO/IEC 14496-22) files
// store all multi-byte integers big-endian.

/**
 * A cursor over a byte buffer that reads big-endian integers — the byte order
 * TrueType / OpenType (ISO/IEC 14496-22) files use for every multi-byte value.
 * Each read advances {@link BigEndianReader.offset}.
 */
export class BigEndianReader {
  /**
   * @param data   The buffer to read from.
   * @param offset The starting byte offset (default 0).
   */
  constructor(
    public readonly data: Uint8Array,
    public offset = 0,
  ) {}

  /** Move the cursor to an absolute byte offset. */
  seek(offset: number): void {
    this.offset = offset;
  }

  /** Advance the cursor by `bytes` without reading. */
  skip(bytes: number): void {
    this.offset += bytes;
  }

  /** Read an unsigned 8-bit integer; advances 1 byte. */
  u8(): number {
    return this.data[this.offset++]!;
  }

  /** Read an unsigned 16-bit big-endian integer; advances 2 bytes. */
  u16(): number {
    const d = this.data;
    const p = this.offset;
    this.offset += 2;
    return (d[p]! << 8) | d[p + 1]!;
  }

  /** Read a signed 16-bit big-endian integer; advances 2 bytes. */
  i16(): number {
    const v = this.u16();
    return v >= 0x8000 ? v - 0x10000 : v;
  }

  /** Read an unsigned 32-bit big-endian integer; advances 4 bytes. */
  u32(): number {
    const d = this.data;
    const p = this.offset;
    this.offset += 4;
    return d[p]! * 0x1000000 + ((d[p + 1]! << 16) | (d[p + 2]! << 8) | d[p + 3]!);
  }

  /** Read a signed 32-bit big-endian integer; advances 4 bytes. */
  i32(): number {
    const v = this.u32();
    return v >= 0x80000000 ? v - 0x100000000 : v;
  }

  /** Read a 4-byte ASCII tag (e.g. a TrueType table tag); advances 4 bytes. */
  tag(): string {
    const d = this.data;
    const p = this.offset;
    this.offset += 4;
    return String.fromCharCode(d[p]!, d[p + 1]!, d[p + 2]!, d[p + 3]!);
  }
}
