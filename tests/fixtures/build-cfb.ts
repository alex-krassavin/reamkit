// Minimal Compound File Binary (MS-CFB / OLE2) builder for tests — packs a few
// named streams into a valid v3 container (512-byte sectors, 64-byte mini
// sectors, 4096-byte cutoff) so the CFB reader and the legacy `.xls` reader have
// deterministic inputs without a checked-in binary. Streams below the cutoff go
// through the mini stream + mini-FAT; the rest through the regular FAT. The
// directory is a flat list by default (the reader's fallback scan handles it);
// pass `storages` to wire a real entry tree with embedded storages, exercising
// the reader's top-level-stream preference.

const SECTOR = 512;
const MINI = 64;
const CUTOFF = 4096;
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const FATSECT = 0xfffffffd;
const NOSTREAM = 0xffffffff;

export interface CfbStreamInput {
  readonly name: string;
  readonly data: Uint8Array;
}

// An embedded storage (e.g. an OLE ObjectPool) holding its own streams.
export interface CfbStorageInput {
  readonly name: string;
  readonly streams: ReadonlyArray<CfbStreamInput>;
}

export interface BuildCfbOptions {
  // Write 0xFFFFFFFF into the reserved high 4 bytes of every stream's 64-bit
  // size field, modelling the real Word/Office quirk: in a v3 container those
  // bytes are reserved and MUST be ignored (MS-CFB §2.6.1), yet writers leave
  // non-zero garbage there. A reader that uses the full 64-bit value sees an
  // absurd size and wrongly rejects an otherwise valid file.
  readonly garbageSizeHighDword?: boolean;
  // Embedded storages. Their streams are written to the directory BEFORE the
  // top-level streams — so a flat first-wins reader mispicks a nested same-named
  // stream — and linked under a child storage with a real directory tree
  // (root → siblings, storage → child) so a tree-aware reader prefers the
  // top-level stream. When omitted the directory stays flat (byte-identical to
  // before): the reader then falls back to its linear scan.
  readonly storages?: ReadonlyArray<CfbStorageInput>;
}

export function buildCfb(
  streams: ReadonlyArray<CfbStreamInput>,
  options: BuildCfbOptions = {},
): Uint8Array {
  const sectors: Array<Uint8Array> = [];
  const fat: Array<number> = [];

  // Store bytes across ceil(len/SECTOR) regular sectors, chained in the FAT.
  const allocRegular = (data: Uint8Array): number => {
    const n = Math.max(1, Math.ceil(data.length / SECTOR));
    const start = sectors.length;
    for (let i = 0; i < n; i++) {
      const sec = new Uint8Array(SECTOR);
      sec.set(data.subarray(i * SECTOR, i * SECTOR + SECTOR));
      sectors.push(sec);
      fat.push(i === n - 1 ? ENDOFCHAIN : start + i + 1);
    }
    return start;
  };

  // Nested (embedded-storage) streams come first in the directory; allocation is
  // keyed by stream identity, not name, so a nested stream and a top-level one
  // can share a name without colliding.
  const storages = options.storages ?? [];
  const nestedStreams = storages.flatMap((st) => st.streams);
  const allStreams = [...nestedStreams, ...streams];
  const big = allStreams.filter((s) => s.data.length >= CUTOFF);
  const small = allStreams.filter((s) => s.data.length > 0 && s.data.length < CUTOFF);

  // --- mini stream: concatenate the small streams, each padded to a mini sector.
  const miniParts: Array<number> = [];
  const miniStart = new Map<CfbStreamInput, number>();
  const miniFat: Array<number> = [];
  let miniIdx = 0;
  for (const s of small) {
    const count = Math.ceil(s.data.length / MINI);
    miniStart.set(s, miniIdx);
    for (let k = 0; k < count; k++) miniFat.push(k === count - 1 ? ENDOFCHAIN : miniIdx + k + 1);
    for (let i = 0; i < count * MINI; i++) miniParts.push(i < s.data.length ? s.data[i]! : 0);
    miniIdx += count;
  }
  const miniStream = Uint8Array.from(miniParts);

  // --- allocate the big streams + the mini stream + the mini-FAT.
  const bigStart = new Map<CfbStreamInput, number>();
  for (const s of big) bigStart.set(s, allocRegular(s.data));

  const miniStreamStart = miniStream.length > 0 ? allocRegular(miniStream) : ENDOFCHAIN;

  let firstMiniFatSector = ENDOFCHAIN;
  let numMiniFatSectors = 0;
  if (miniFat.length > 0) {
    const mfBytes = u32Array(miniFat, SECTOR, FREESECT);
    const start = allocRegular(mfBytes);
    firstMiniFatSector = start;
    numMiniFatSectors = mfBytes.length / SECTOR;
  }
  const dataStart = (s: CfbStreamInput): number =>
    s.data.length === 0
      ? ENDOFCHAIN
      : s.data.length >= CUTOFF
        ? bigStart.get(s)!
        : miniStart.get(s)!;

  // --- directory. Index plan: 0 = Root, then the nested streams, then the
  // storage entries, then the top-level streams. With no storages this is just
  // [Root, ...streams] with NOSTREAM links — byte-identical to before — and the
  // reader falls back to its flat scan. With storages we wire a real tree:
  // Root.child → the top-level sibling chain; each storage.child → its streams.
  const hasTree = storages.length > 0;
  const nestedBase = 1;
  const storageBase = nestedBase + nestedStreams.length;
  const topBase = storageBase + storages.length;
  const dir: Array<Uint8Array> = new Array(topBase + streams.length);
  const firstTop = hasTree ? (storages.length > 0 ? storageBase : topBase) : NOSTREAM;
  dir[0] = dirEntry('Root Entry', 5, miniStreamStart, miniStream.length, false, {
    child: firstTop,
  });
  // Nested streams, each right-chained within its storage.
  const storageChild: Array<number> = [];
  let ni = nestedBase;
  for (const st of storages) {
    storageChild.push(st.streams.length > 0 ? ni : NOSTREAM);
    st.streams.forEach((s, k) => {
      dir[ni] = dirEntry(s.name, 2, dataStart(s), s.data.length, options.garbageSizeHighDword, {
        right: k < st.streams.length - 1 ? ni + 1 : NOSTREAM,
      });
      ni++;
    });
  }
  // Top-level siblings in sequence: storage entries, then top-level streams,
  // right-chained so the reader visits them all from Root.child.
  const topCount = storages.length + streams.length;
  const topSeqIdx = (p: number): number =>
    p < storages.length ? storageBase + p : topBase + (p - storages.length);
  const rightOf = (p: number): number =>
    hasTree && p < topCount - 1 ? topSeqIdx(p + 1) : NOSTREAM;
  storages.forEach((st, i) => {
    dir[storageBase + i] = dirEntry(st.name, 1, ENDOFCHAIN, 0, false, {
      right: rightOf(i),
      child: storageChild[i]!,
    });
  });
  streams.forEach((s, i) => {
    dir[topBase + i] = dirEntry(
      s.name,
      2,
      dataStart(s),
      s.data.length,
      options.garbageSizeHighDword,
      {
        right: rightOf(storages.length + i),
      },
    );
  });
  // Pad the directory to a whole sector (4 entries per 512-byte sector).
  while (dir.length % 4 !== 0) dir.push(unallocatedEntry());
  const firstDirSector = allocRegular(concat(dir));

  // --- the FAT sector last: it describes every content sector plus itself.
  const fatSectorIndex = sectors.length;
  fat.push(FATSECT);
  const fatSector = u32Array(fat, SECTOR, FREESECT);
  sectors.push(fatSector); // (its own FAT entry already pushed above)

  // --- header.
  const header = new Uint8Array(SECTOR);
  const hv = new DataView(header.buffer);
  header.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
  hv.setUint16(24, 0x003e, true); // minor version
  hv.setUint16(26, 3, true); // major version (v3)
  hv.setUint16(28, 0xfffe, true); // little-endian
  hv.setUint16(30, 9, true); // sector shift → 512
  hv.setUint16(32, 6, true); // mini sector shift → 64
  hv.setUint32(44, 1, true); // number of FAT sectors
  hv.setUint32(48, firstDirSector, true);
  hv.setUint32(56, CUTOFF, true); // mini stream cutoff
  hv.setUint32(60, firstMiniFatSector, true);
  hv.setUint32(64, numMiniFatSectors, true);
  hv.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector
  hv.setUint32(72, 0, true); // number of DIFAT sectors
  for (let i = 0; i < 109; i++) hv.setUint32(76 + i * 4, i === 0 ? fatSectorIndex : FREESECT, true);

  return concat([header, ...sectors]);
}

function dirEntry(
  name: string,
  type: number,
  startSector: number,
  size: number,
  garbageHigh = false,
  links: { left?: number; right?: number; child?: number } = {},
): Uint8Array {
  const e = new Uint8Array(DIR_ENTRY_SIZE);
  const v = new DataView(e.buffer);
  for (let i = 0; i < name.length && i < 31; i++) v.setUint16(i * 2, name.charCodeAt(i), true);
  v.setUint16(64, (Math.min(name.length, 31) + 1) * 2, true); // name length incl. NUL
  v.setUint8(66, type);
  v.setUint8(67, 1); // colour (black)
  v.setUint32(68, links.left ?? NOSTREAM, true); // left sibling
  v.setUint32(72, links.right ?? NOSTREAM, true); // right sibling
  v.setUint32(76, links.child ?? NOSTREAM, true); // child
  v.setUint32(116, startSector, true);
  v.setUint32(120, size >>> 0, true);
  v.setUint32(124, garbageHigh ? 0xffffffff : Math.floor(size / 0x1_0000_0000), true);
  return e;
}

function unallocatedEntry(): Uint8Array {
  const e = new Uint8Array(DIR_ENTRY_SIZE);
  new DataView(e.buffer).setUint8(66, 0); // type = unallocated
  return e;
}

const DIR_ENTRY_SIZE = 128;

// Pack a u32 list into a byte array padded to a multiple of `align`, filling the
// tail with `fill`.
function u32Array(values: ReadonlyArray<number>, align: number, fill: number): Uint8Array {
  const count = Math.ceil((values.length * 4) / align) * (align / 4);
  const out = new Uint8Array(count * 4);
  const v = new DataView(out.buffer);
  for (let i = 0; i < count; i++) v.setUint32(i * 4, i < values.length ? values[i]! : fill, true);
  return out;
}

function concat(parts: ReadonlyArray<Uint8Array>): Uint8Array {
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}
