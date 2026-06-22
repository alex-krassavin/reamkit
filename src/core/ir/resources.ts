// IR core — content-addressed binary resource store (ir-design.md §4).
//
// IR trees are pure JSON; binary payloads (images, fonts) live here and nodes
// reference them by id. Ids derive from the bytes (content-addressed), so
// putting the same image twice deduplicates for free and equal documents get
// equal ids — which keeps IR-level diffs and byte-identical gates meaningful.
//
// The hash is a synchronous non-cryptographic 64-bit FNV-1a (two independent
// 32-bit lanes): readers are sync by contract (handoff v1 §4), which rules out
// WebCrypto digests. Collisions are handled, not assumed away: when an id is
// already taken by *different* bytes, a deterministic `~1`, `~2`, … suffix is
// probed. Ids therefore stay deterministic for a deterministic put order —
// which our readers guarantee.

/** A content-addressed id for a binary resource held in a {@link ResourceStore}. */
export type ResourceId = string & { readonly __brand: 'resource-id' };

const FNV_PRIME = 0x01000193;

// Two FNV-1a lanes with distinct offset bases, combined into 64 hash bits.
function hashBytes(bytes: Uint8Array): string {
  let a = 0x811c9dc5;
  let b = 0xcbf29ce4;
  for (const byte of bytes) {
    a = Math.imul(a ^ byte, FNV_PRIME);
    b = Math.imul(b ^ byte, FNV_PRIME);
  }
  const hexA = (a >>> 0).toString(16).padStart(8, '0');
  const hexB = (b >>> 0).toString(16).padStart(8, '0');
  return `${hexA}${hexB}`;
}

function equalBytes(x: Uint8Array, y: Uint8Array): boolean {
  if (x.length !== y.length) return false;
  for (let i = 0; i < x.length; i++) {
    if (x[i] !== y[i]) return false;
  }
  return true;
}

/**
 * A content-addressed store for the IR's binary payloads (images, fonts). IR
 * trees stay pure JSON and reference bytes by {@link ResourceId}; ids derive
 * from the bytes (a synchronous 64-bit FNV-1a), so identical payloads
 * deduplicate for free and equal documents get equal ids — which keeps IR-level
 * diffs and the byte-identical gates meaningful.
 */
export class ResourceStore {
  private readonly byId = new Map<ResourceId, Uint8Array>();

  /**
   * Store bytes and return their content-addressed id. Re-putting identical
   * bytes returns the existing id (deduplication); a hash collision with
   * different bytes probes deterministic `~n` suffixes.
   */
  put(bytes: Uint8Array): ResourceId {
    const base = `r-${hashBytes(bytes)}-${bytes.length}`;
    for (let n = 0; ; n++) {
      const id = (n === 0 ? base : `${base}~${n}`) as ResourceId;
      const existing = this.byId.get(id);
      if (existing === undefined) {
        this.byId.set(id, bytes);
        return id;
      }
      if (equalBytes(existing, bytes)) return id;
    }
  }

  /** Look up the bytes for an id, or `undefined` when absent. */
  get(id: ResourceId): Uint8Array | undefined {
    return this.byId.get(id);
  }

  /** Whether an id is present in the store. */
  has(id: ResourceId): boolean {
    return this.byId.has(id);
  }

  /** All stored ids, in insertion order. */
  ids(): ReadonlyArray<ResourceId> {
    return [...this.byId.keys()];
  }

  /** The number of distinct resources stored. */
  get size(): number {
    return this.byId.size;
  }
}
