export { OpcPackage } from '@/core/opc/package';
export type { Relationship } from '@/core/opc/relationships';
export {
  REL_OFFICE_DOCUMENT,
  REL_CORE_PROPERTIES,
  REL_EXTENDED_PROPERTIES,
} from '@/core/opc/relationship-types';
export { parseCoreProperties } from '@/core/opc/core-properties';
export type { CoreProperties } from '@/core/opc/core-properties';
export { isOoxmlRel } from '@/core/opc/relationship-types';
export { buildOpcPackage, relsPathFor } from '@/core/opc/opc-writer';
export type { OpcPart, OpcPartRelationships, OpcWriteOptions } from '@/core/opc/opc-writer';
