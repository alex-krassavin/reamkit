/*
 * The Unicode properties `scripts/build-indic-table.ts` reads, from ICU: one
 * line per code point in the ranges given (hex pairs), tab-separated —
 *
 *   code point, Indic_Syllabic_Category, Indic_Positional_Category,
 *   General_Category, Canonical_Combining_Class, canonical decomposition
 *   (one level, space-separated), Full_Composition_Exclusion (0 or 1)
 *
 * Build and run against the ICU you have, e.g. Homebrew's icu4c:
 *
 *   cc -I$(brew --prefix icu4c)/include -L$(brew --prefix icu4c)/lib -licuuc \
 *     -o icu-dump scripts/indic-table/icu-dump.c
 *   ./icu-dump 0000 0DFF 1CD0 1CFF 2000 20FF 25C0 25FF A8E0 A8FF > indic.tsv
 */
#include <stdio.h>
#include <stdlib.h>
#include <unicode/uchar.h>
#include <unicode/unorm2.h>
#include <unicode/ustring.h>
#include <unicode/uversion.h>

int main(int argc, char **argv) {
  UErrorCode e = U_ZERO_ERROR;
  const UNormalizer2 *nfc = unorm2_getNFCInstance(&e);
  if (U_FAILURE(e)) return 1;
  printf("# ICU %s, Unicode %s\n", U_ICU_VERSION, U_UNICODE_VERSION);
  for (int a = 1; a + 1 < argc; a += 2) {
    long lo = strtol(argv[a], 0, 16), hi = strtol(argv[a + 1], 0, 16);
    for (UChar32 c = (UChar32)lo; c <= (UChar32)hi; c++) {
      if (!u_isdefined(c)) continue;
      UChar d[8];
      char dec[64] = "";
      e = U_ZERO_ERROR;
      int32_t n = unorm2_getRawDecomposition(nfc, c, d, 8, &e);
      if (n > 0) {
        int32_t i = 0, k = 0;
        UChar32 cp;
        while (i < n) {
          U16_NEXT(d, i, n, cp);
          k += sprintf(dec + k, "%s%04X", k ? " " : "", (unsigned)cp);
        }
      }
      printf("%04X\t%s\t%s\t%s\t%d\t%s\t%d\n", (unsigned)c,
             u_getPropertyValueName(UCHAR_INDIC_SYLLABIC_CATEGORY,
                                    u_getIntPropertyValue(c, UCHAR_INDIC_SYLLABIC_CATEGORY),
                                    U_LONG_PROPERTY_NAME),
             u_getPropertyValueName(UCHAR_INDIC_POSITIONAL_CATEGORY,
                                    u_getIntPropertyValue(c, UCHAR_INDIC_POSITIONAL_CATEGORY),
                                    U_LONG_PROPERTY_NAME),
             u_getPropertyValueName(UCHAR_GENERAL_CATEGORY, u_charType(c), U_SHORT_PROPERTY_NAME),
             u_getCombiningClass(c), dec, u_hasBinaryProperty(c, UCHAR_FULL_COMPOSITION_EXCLUSION));
    }
  }
  return 0;
}
