// Big-endian binary reader. TrueType / OpenType (ISO/IEC 14496-22) files
// store all multi-byte integers big-endian.

export class BigEndianReader {
  constructor(
    public readonly data: Uint8Array,
    public offset = 0,
  ) {}

  seek(offset: number): void {
    this.offset = offset;
  }

  skip(bytes: number): void {
    this.offset += bytes;
  }

  u8(): number {
    return this.data[this.offset++]!;
  }

  u16(): number {
    const d = this.data;
    const p = this.offset;
    this.offset += 2;
    return (d[p]! << 8) | d[p + 1]!;
  }

  i16(): number {
    const v = this.u16();
    return v >= 0x8000 ? v - 0x10000 : v;
  }

  u32(): number {
    const d = this.data;
    const p = this.offset;
    this.offset += 4;
    return d[p]! * 0x1000000 + ((d[p + 1]! << 16) | (d[p + 2]! << 8) | d[p + 3]!);
  }

  i32(): number {
    const v = this.u32();
    return v >= 0x80000000 ? v - 0x100000000 : v;
  }

  tag(): string {
    const d = this.data;
    const p = this.offset;
    this.offset += 4;
    return String.fromCharCode(d[p]!, d[p + 1]!, d[p + 2]!, d[p + 3]!);
  }
}
