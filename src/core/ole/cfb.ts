// Compound File Binary reader (MS-CFB) — the OLE2 container that the legacy
// binary Office formats live in: a `.xls` is a CFB holding a `Workbook` stream,
// a `.doc` a `WordDocument` stream, a `.ppt` a `PowerPoint Document` stream, and
// an ActiveX control's persisted state a small CFB inside `xl/activeX/*.bin`. The
// container is a tiny in-memory FAT filesystem: a header, a sector-allocation
// table (FAT) addressed through the DIFAT, a directory of named entries, and a
// secondary mini-FAT for streams below the 4096-byte cutoff. This reader turns
// the bytes into named streams; the format parsers (BIFF, …) run on those.
//
// Hardened like OpcPackage.open: every sector index is bounds-checked, every
// chain walk is capped to the sector count (cycles/overruns abort), and total
// output is bounded — a crafted container cannot loop or blow up the heap.

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

// Special FAT/mini-FAT sector values (MS-CFB §2.2).
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;

const HEADER_SIZE = 512;
const DIR_ENTRY_SIZE = 128;
const MAX_TOTAL_BYTES = 512 * 1024 * 1024; // overall output ceiling

export type CfbEntryType = 'root' | 'storage' | 'stream' | 'unknown';

export interface CfbEntry {
  readonly name: string;
  readonly type: CfbEntryType;
  readonly startSector: number;
  readonly size: number;
}

export class CfbError extends Error {}

export interface Cfb {
  /** Every directory entry, in directory order (index 0 is the root storage). */
  readonly entries: ReadonlyArray<CfbEntry>;
  /** Read a stream's bytes by entry name (case-insensitive), or undefined. */
  readStream: (name: string) => Uint8Array | undefined;
  /** Whether a stream entry with that name exists (case-insensitive). */
  hasStream: (name: string) => boolean;
}

// Cheap magic-byte probe — the first eight bytes are the CFB signature. Shared by
// the legacy readers' sniff and by OpcPackage to tell an OLE file from a ZIP.
export function isCfb(bytes: Uint8Array): boolean {
  if (bytes.length < HEADER_SIZE) return false;
  for (let i = 0; i < SIGNATURE.length; i++) {
    if (bytes[i] !== SIGNATURE[i]) return false;
  }
  return true;
}

export function openCfb(bytes: Uint8Array): Cfb {
  if (!isCfb(bytes)) throw new CfbError('not a compound file (bad signature)');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  const sectorShift = view.getUint16(30, true);
  const miniSectorShift = view.getUint16(32, true);
  if (sectorShift !== 9 && sectorShift !== 12) {
    throw new CfbError(`unsupported sector shift ${sectorShift}`);
  }
  if (miniSectorShift !== 6) throw new CfbError(`unsupported mini-sector shift ${miniSectorShift}`);
  const sectorSize = 1 << sectorShift;
  const miniSectorSize = 1 << miniSectorShift;
  const miniCutoff = view.getUint32(56, true);
  // §2.2 — only v3 (512-byte sectors) and v4 (4096) exist; the directory entry's
  // stream-size field is read per-version below (v3 ignores its high 4 bytes).
  const majorVersion = view.getUint16(26, true);

  // How many real sectors the file body can hold — every sector index must land
  // inside this, and no chain may be longer than this (the cycle guard).
  const totalSectors = Math.floor((bytes.length - HEADER_SIZE) / sectorSize);
  const sectorOffset = (sector: number): number => {
    if (sector < 0 || sector >= totalSectors) throw new CfbError(`sector ${sector} out of range`);
    return HEADER_SIZE + sector * sectorSize;
  };

  const firstDirSector = view.getUint32(48, true);
  const firstMiniFatSector = view.getUint32(60, true);
  const firstDifatSector = view.getUint32(68, true);
  const numDifatSectors = view.getUint32(72, true);

  // --- DIFAT → the list of sectors that make up the FAT ---------------------
  const fatSectorLocs: Array<number> = [];
  for (let i = 0; i < 109; i++) {
    const loc = view.getUint32(76 + i * 4, true);
    if (loc !== FREESECT && loc < totalSectors) fatSectorLocs.push(loc);
  }
  // DIFAT continuation sectors: each holds (sectorSize/4 - 1) FAT-sector
  // locations, the last 4 bytes pointing to the next DIFAT sector.
  const perDifat = sectorSize / 4 - 1;
  let difatSector = firstDifatSector;
  for (let guard = 0; difatSector !== ENDOFCHAIN && difatSector !== FREESECT; guard++) {
    if (guard > numDifatSectors || guard > totalSectors) {
      throw new CfbError('DIFAT chain too long (corrupt)');
    }
    const off = sectorOffset(difatSector);
    for (let i = 0; i < perDifat; i++) {
      const loc = view.getUint32(off + i * 4, true);
      if (loc !== FREESECT && loc < totalSectors) fatSectorLocs.push(loc);
    }
    difatSector = view.getUint32(off + perDifat * 4, true);
  }

  // --- the FAT itself: sector → next sector ---------------------------------
  const entriesPerSector = sectorSize / 4;
  const fat = new Uint32Array(fatSectorLocs.length * entriesPerSector);
  for (let s = 0; s < fatSectorLocs.length; s++) {
    const off = sectorOffset(fatSectorLocs[s]!);
    for (let i = 0; i < entriesPerSector; i++) {
      fat[s * entriesPerSector + i] = view.getUint32(off + i * 4, true);
    }
  }

  // Follow a FAT chain from `start`, returning the sector indices. Capped to the
  // sector count so a cycle or a run off the end aborts instead of hanging.
  const followChain = (start: number): Array<number> => {
    const out: Array<number> = [];
    let s = start;
    while (s !== ENDOFCHAIN && s !== FREESECT) {
      if (s >= totalSectors || out.length > totalSectors) {
        throw new CfbError('FAT chain corrupt (out of range or cyclic)');
      }
      out.push(s);
      s = fat[s] ?? ENDOFCHAIN;
    }
    return out;
  };

  // Concatenate the bytes of a sector chain (full sectors; caller truncates).
  const readChainBytes = (start: number, limit: number): Uint8Array => {
    const sectors = followChain(start);
    const total = sectors.length * sectorSize;
    if (total > limit) throw new CfbError('stream exceeds size limit');
    const out = new Uint8Array(total);
    for (let i = 0; i < sectors.length; i++) {
      const off = sectorOffset(sectors[i]!);
      out.set(bytes.subarray(off, off + sectorSize), i * sectorSize);
    }
    return out;
  };

  // --- directory: chain of 128-byte entries ---------------------------------
  const dirBytes = readChainBytes(firstDirSector, MAX_TOTAL_BYTES);
  const dirView = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
  const entries: Array<CfbEntry> = [];
  const entryCount = Math.floor(dirBytes.length / DIR_ENTRY_SIZE);
  // Raw (directory-index-keyed) view that also keeps each entry's red-black-tree
  // pointers, so the storage hierarchy can be walked (`entries` above is the
  // flat public list of allocated entries only). The sibling/child fields are
  // entry indices into this directory.
  const rawByIndex: Array<
    { entry: CfbEntry; left: number; right: number; child: number; isRoot: boolean } | undefined
  > = new Array(entryCount);
  for (let i = 0; i < entryCount; i++) {
    const base = i * DIR_ENTRY_SIZE;
    const objType = dirView.getUint8(base + 66);
    if (objType !== 1 && objType !== 2 && objType !== 5) continue; // unallocated
    const nameLen = dirView.getUint16(base + 64, true);
    const name = decodeName(dirBytes, base, nameLen);
    // §2.6.1 — in a v3 file the stream-size field is 32-bit and the high 4
    // bytes are reserved; Word and others leave non-zero garbage there, so a
    // reader MUST ignore them for v3 (using the full 64-bit value yields absurd
    // sizes that trip the size guard and reject an otherwise valid file). Only
    // v4 (4096-byte sectors) uses the full 64-bit size.
    const sizeLow = dirView.getUint32(base + 120, true);
    const sizeHigh = dirView.getUint32(base + 124, true);
    const size = majorVersion === 3 ? sizeLow : sizeHigh * 0x1_0000_0000 + sizeLow;
    const entry: CfbEntry = {
      name,
      type: objType === 5 ? 'root' : objType === 1 ? 'storage' : 'stream',
      startSector: dirView.getUint32(base + 116, true),
      size,
    };
    entries.push(entry);
    rawByIndex[i] = {
      entry,
      left: dirView.getUint32(base + 68, true),
      right: dirView.getUint32(base + 72, true),
      child: dirView.getUint32(base + 76, true),
      isRoot: objType === 5,
    };
  }

  const root = entries.find((e) => e.type === 'root');
  if (!root) throw new CfbError('compound file has no root entry');

  // The mini stream is the root entry's regular-FAT chain; small streams are
  // carved out of it by the mini-FAT. Built lazily on first small-stream read.
  let miniStream: Uint8Array | undefined;
  let miniFat: Uint32Array | undefined;
  const ensureMini = (): void => {
    if (miniStream && miniFat) return;
    miniStream = readChainBytes(root.startSector, MAX_TOTAL_BYTES).subarray(0, root.size);
    const mfBytes =
      firstMiniFatSector === ENDOFCHAIN
        ? new Uint8Array(0)
        : readChainBytes(firstMiniFatSector, MAX_TOTAL_BYTES);
    const mfView = new DataView(mfBytes.buffer, mfBytes.byteOffset, mfBytes.byteLength);
    miniFat = new Uint32Array(Math.floor(mfBytes.length / 4));
    for (let i = 0; i < miniFat.length; i++) miniFat[i] = mfView.getUint32(i * 4, true);
  };

  const readMiniStream = (entry: CfbEntry): Uint8Array => {
    ensureMini();
    const stream = miniStream!;
    const mfat = miniFat!;
    const maxMini = Math.floor(stream.length / miniSectorSize);
    const out = new Uint8Array(entry.size);
    let mini = entry.startSector;
    let written = 0;
    let guard = 0;
    while (mini !== ENDOFCHAIN && mini !== FREESECT && written < entry.size) {
      if (mini >= maxMini || guard++ > mfat.length) {
        throw new CfbError('mini-FAT chain corrupt (out of range or cyclic)');
      }
      const off = mini * miniSectorSize;
      const n = Math.min(miniSectorSize, entry.size - written);
      out.set(stream.subarray(off, off + n), written);
      written += n;
      mini = mfat[mini] ?? ENDOFCHAIN;
    }
    return out;
  };

  const readEntry = (entry: CfbEntry): Uint8Array => {
    if (entry.size === 0) return new Uint8Array(0);
    if (entry.size > MAX_TOTAL_BYTES) throw new CfbError('stream exceeds size limit');
    // Streams at or above the cutoff live in the regular FAT; smaller ones in
    // the mini stream. (The root's own bytes are never exposed as a stream.)
    if (entry.size >= miniCutoff) {
      return readChainBytes(entry.startSector, MAX_TOTAL_BYTES).subarray(0, entry.size);
    }
    return readMiniStream(entry);
  };

  // The main document's streams are the direct children of the root storage.
  // The directory is a red-black tree keyed by name; the root entry's `child`
  // points at the tree of its immediate children, while an embedded OLE object's
  // streams hang off a nested storage (e.g. ObjectPool). Collect just the root's
  // own children — following sibling links but never descending into a child
  // storage — so they take precedence over an embedded object's same-named
  // stream. Without this a flat first-wins scan can return an embedded
  // `WordDocument`/`1Table` and the parse comes up empty (MS-CFB §2.6).
  const topLevel = new Set<CfbEntry>();
  const rootRaw = rawByIndex.find((r) => r?.isRoot);
  if (rootRaw) {
    const stack = [rootRaw.child];
    const seen = new Set<number>();
    while (stack.length > 0) {
      const idx = stack.pop()!;
      if (idx >= entryCount || seen.has(idx)) continue; // NOSTREAM / out-of-range / cycle
      seen.add(idx);
      const node = rawByIndex[idx];
      if (!node) continue;
      topLevel.add(node.entry);
      stack.push(node.left, node.right); // siblings only — not node.child
    }
  }

  const byName = new Map<string, CfbEntry>();
  const addStreams = (list: Iterable<CfbEntry>): void => {
    for (const e of list) {
      if (e.type === 'stream' && !byName.has(e.name.toLowerCase())) {
        byName.set(e.name.toLowerCase(), e);
      }
    }
  };
  addStreams(topLevel); // main-document streams win the name…
  addStreams(entries); // …then any nested stream whose name is otherwise unused

  return {
    entries,
    hasStream: (name) => byName.has(name.toLowerCase()),
    readStream: (name) => {
      const e = byName.get(name.toLowerCase());
      return e ? readEntry(e) : undefined;
    },
  };
}

// Decode a directory entry's UTF-16LE name. `nameLen` counts bytes including the
// terminating NUL; an out-of-spec length is clamped to the 64-byte field.
function decodeName(dir: Uint8Array, base: number, nameLen: number): string {
  const len = Math.max(0, Math.min(64, nameLen) - 2); // drop the UTF-16 NUL
  let out = '';
  for (let i = 0; i < len; i += 2) {
    out += String.fromCharCode(dir[base + i]! | (dir[base + i + 1]! << 8));
  }
  return out;
}
