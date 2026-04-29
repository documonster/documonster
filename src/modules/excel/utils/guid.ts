/**
 * GUID helpers for OOXML.
 *
 * Office parts that need stable cross-reference ids (threaded comments,
 * chart uniqueIds, slicer caches, …) use the canonical Microsoft
 * `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` format — 32 uppercase hex
 * characters split into 8-4-4-4-12 groups, surrounded by braces.
 *
 * `synthGuid()` returns an RFC-4122 version-4-compatible value suitable
 * for those callers. We don't use `crypto.randomUUID()` directly because
 * it isn't available in every supported runtime (older Node without
 * `node:crypto` globals); `Math.random` is good enough for uniqueness
 * within a single workbook, which is the only thing Office cares about.
 */
export function synthGuid(): string {
  const hex = "0123456789ABCDEF";
  let out = "";
  for (let i = 0; i < 32; i++) {
    // Version 4: bit 48 (13th hex digit) is always '4'.
    // Variant 10xx: bit 64 (17th hex digit) is one of 8, 9, A, B.
    if (i === 12) {
      out += "4";
    } else if (i === 16) {
      out += hex[8 + Math.floor(Math.random() * 4)];
    } else {
      out += hex[Math.floor(Math.random() * 16)];
    }
    if (i === 7 || i === 11 || i === 15 || i === 19) {
      out += "-";
    }
  }
  return out;
}
