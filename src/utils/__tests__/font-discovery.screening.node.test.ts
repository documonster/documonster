/**
 * A search screens faces from their tables alone and hands back the face it chose with the whole file: the outlines
 * the screening read left out are what drawing or embedding the face needs.
 */
import { describe, expect, it } from "vitest";

import { findSystemFontForCodePoints, resetFontDiscoveryCache } from "../font-discovery";

describe("system font screening", () => {
  it("returns the chosen face with its outlines", () => {
    resetFontDiscoveryCache();
    const face = findSystemFontForCodePoints(new Set([0x41, 0x61, 0x30]));
    if (face === null) {
      // A host with no TrueType face installed has nothing to check.
      return;
    }
    const glyf = face.tables.get("glyf")!;
    const outlines = face.data.subarray(glyf.offset, glyf.offset + glyf.length);
    expect(outlines.length).toBeGreaterThan(0);
    expect(outlines.some(byte => byte !== 0)).toBe(true);
  });
});
