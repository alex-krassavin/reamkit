// IR core (ir-design.md) — stage 1 of the v1 migration: canonical unit,
// resource store, loss reporting and the feature vocabulary. Internal for now;
// becomes part of the public adapter contract at stage 4 (@experimental).

export type { Pt } from '@/ir/units';
export {
  pt,
  twipsToPt,
  halfPtToPt,
  eighthPtToPt,
  emuToPt,
  pxToPt,
  inchToPt,
  mmToPt,
} from '@/ir/units';

export type { ResourceId } from '@/ir/resources';
export { ResourceStore } from '@/ir/resources';

export type { Feature, KnownFeature } from '@/ir/features';
export { FEATURES, featureWithin } from '@/ir/features';

export type { Loss, LossReport, LossSeverity } from '@/ir/loss';
export { ConversionLossError, formatLoss } from '@/ir/loss';

// Passthrough bag for the round-trip door (ir-design §4): a node MAY carry the
// source format's raw fragment keyed by format id (e.g. { ooxml: <fragment> }).
// Writers for other formats ignore it; a same-format writer may use it to
// reproduce constructs the neutral core does not model. v1 promises only
// cross-format fidelity — this keeps the door open without promising more.
export interface NativeBag {
  readonly [formatId: string]: unknown;
}
