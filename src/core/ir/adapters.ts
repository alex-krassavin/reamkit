// Adapter contracts (ir-design §7) — @experimental until they stabilize
// against three-plus adapters (handoff v1 §3a). Readers and writers are SYNC
// by contract (the bytes are already in memory; all async — font fetching,
// resource I/O — lives in the conversion facade).

import type { Feature } from '@/core/ir/features';
import type { FlowDoc } from '@/core/ir/flow';
import type { LossReport } from '@/core/ir/loss';

export interface ReadOptions {
  /** Reserved for reader-specific knobs (limits live in the readers). */
  readonly [key: string]: unknown;
}

export interface ReadResult<TDoc> {
  readonly doc: TDoc;
  readonly losses: LossReport;
}

export interface DocumentReader<TDoc = FlowDoc> {
  /** Format id: 'docx', 'xlsx', 'pdf', … */
  readonly id: string;
  /** Which tree this reader produces. */
  readonly produces: 'flow' | 'page';
  /** Feature vocabulary this reader understands (capability matrix source). */
  readonly supports: ReadonlySet<Feature>;
  /** Cheap format detection (magic bytes / container markers). */
  sniff: (bytes: Uint8Array) => boolean;
  read: (bytes: Uint8Array, opts?: ReadOptions) => ReadResult<TDoc>;
}

export interface WriteOptions {
  readonly [key: string]: unknown;
}

export interface WriteResult {
  readonly bytes: Uint8Array;
  readonly losses: LossReport;
}

export interface DocumentWriter<TDoc> {
  /** Output format id: 'pdf', 'svg', 'html', 'docx', … */
  readonly id: string;
  /** Which tree this writer consumes. */
  readonly consumes: 'flow' | 'page';
  readonly supports: ReadonlySet<Feature>;
  write: (doc: TDoc, opts?: WriteOptions) => WriteResult;
}
