/**
 * `synthGuid` — uppercase hex, no braces. `uuidV4` returns lowercase, so
 * dropping the `toUpperCase()` would still yield a valid UUID and a silently
 * different document. Randomness tiers: `src/utils/__tests__/uuid.test.ts`.
 */

import { synthGuid } from "@excel/utils/guid";
import { describe, expect, it } from "vitest";

describe("synthGuid", () => {
  it("returns uppercase hex in 8-4-4-4-12 groups", () => {
    expect(synthGuid()).toMatch(
      /^[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}$/
    );
  });

  it("does not add braces — the caller does", () => {
    const guid = synthGuid();
    expect(guid.startsWith("{")).toBe(false);
    expect(guid.endsWith("}")).toBe(false);
    expect(guid).toHaveLength(36);
  });

  it("contains no lowercase hex", () => {
    // `uuidV4` is lowercase; this pins the `toUpperCase()` that adapts it.
    const guid = synthGuid();
    expect(guid).toBe(guid.toUpperCase());
  });

  it("differs across calls", () => {
    expect(synthGuid()).not.toBe(synthGuid());
  });
});
