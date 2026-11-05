export {
  companionOf,
  companionScale,
  fetchFontSet,
  fetchScriptFont,
  isScriptKey,
  knowsFamily,
  resolveFamilyKey,
  resolveFamilyStyle,
  clearFontCache,
} from '@/core/fonts/remote-fonts';
export { scriptForCodepoint, scriptsInFlow } from '@/core/fonts/scripts';
export type {
  Companion,
  FamilyKey,
  FamilyStyle,
  FetchFontSetOptions,
  FetchLike,
  ScriptKey,
  SubstituteKey,
} from '@/core/fonts/remote-fonts';
