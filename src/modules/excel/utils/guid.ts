/**
 * GUID helpers for OOXML.
 *
 * Office parts needing stable cross-reference ids (threaded comments, chart
 * uniqueIds, slicer caches) use uppercase hex in 8-4-4-4-12 groups; callers add
 * the braces. `synthGuid()` delegates to {@link uuidV4} and upper-cases it
 * rather than rolling its own digits out of `Math.random()`, which skipped Web
 * Crypto even where it exists and duplicated the version/variant masking.
 */
import { uuidV4 } from "@utils/uuid";

export function synthGuid(): string {
  return uuidV4().toUpperCase();
}
