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
