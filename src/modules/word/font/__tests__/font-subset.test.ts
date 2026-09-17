/**
 * The Word font subsetter, on a font it can actually subset.
 *
 * Every existing test of `subsetFont` exercised a *refusal*: a font with no `glyf`,
 * an OTF, an empty string. So the real work — discarding glyph data, following
 * composite dependencies, rebuilding `loca`, writing a valid sfnt — had no coverage
 * at all, and neither did the two defects fixed alongside these tests.
 *
 * The strongest available invariant is that the output must be readable: a subset
 * font that this library's own parser rejects is a font Word may reject too.
 */

import { buildSubsettableFont } from "@test/ttf-fixture";
import { parseTtf } from "@utils/font-ttf";
import { subsetFont } from "@word/font/font-embed";
import { describe, expect, it } from "vitest";

const A = 0x41;
const B = 0x42;
const HAN = 0x4e2d;

/** Read a big-endian u32. */
function u32(data: Uint8Array, at: number): number {
  return ((data[at] << 24) | (data[at + 1] << 16) | (data[at + 2] << 8) | data[at + 3]) >>> 0;
}

/** The sfnt table directory of `font`, by tag. */
function tablesOf(font: Uint8Array): Map<string, { offset: number; length: number }> {
  const count = (font[4] << 8) | font[5];
  const tables = new Map<string, { offset: number; length: number }>();
  for (let i = 0; i < count; i++) {
    const rec = 12 + i * 16;
    const tag = String.fromCharCode(font[rec], font[rec + 1], font[rec + 2], font[rec + 3]);
    tables.set(tag, { offset: u32(font, rec + 8), length: u32(font, rec + 12) });
  }
  return tables;
}

/**
 * The whole-font checksum, which a valid sfnt fixes at `0xB1B0AFBA`.
 *
 * Computed the way the spec defines it: sum every 4-byte group of the file with the
 * `head` table's `checkSumAdjustment` field excluded, then add the stored adjustment.
 */
function wholeFontChecksum(font: Uint8Array): number {
  const head = tablesOf(font).get("head");
  const adjustmentAt = head === undefined ? -1 : head.offset + 8;
  let sum = 0;
  for (let at = 0; at < font.length; at += 4) {
    if (at === adjustmentAt) {
      continue; // the adjustment field itself is excluded
    }
    const b0 = font[at] ?? 0;
    const b1 = font[at + 1] ?? 0;
    const b2 = font[at + 2] ?? 0;
    const b3 = font[at + 3] ?? 0;
    sum = (sum + (((b0 << 24) | (b1 << 16) | (b2 << 8) | b3) >>> 0)) >>> 0;
  }
  const stored = adjustmentAt < 0 ? 0 : u32(font, adjustmentAt);
  return (sum + stored) >>> 0;
}

describe("Word font subsetting", () => {
  const font = buildSubsettableFont([A, B, HAN], "Fixture Subset");

  it("actually subsets, rather than returning the original", () => {
    const subset = subsetFont(font, "A");
    expect(subset).not.toBe(font);
    expect(subset.length).toBeLessThan(font.length);
  });

  it("produces a font this library can parse back", () => {
    // The invariant that matters: an unreadable subset is a font Word may refuse.
    const subset = subsetFont(font, "A中");
    const parsed = parseTtf(subset);
    expect(parsed.numGlyphs).toBeGreaterThan(0);
    expect(parsed.unitsPerEm).toBe(1000);
    // The characters asked for still map to a glyph.
    expect(parsed.cmap.get(A)).toBeDefined();
    expect(parsed.cmap.get(HAN)).toBeDefined();
  });

  it("keeps the glyphs a composite depends on", () => {
    // 'A' maps to glyph 1, which is a composite referencing glyph 2. Dropping the
    // component would leave a composite that draws nothing.
    const subset = subsetFont(font, "A");
    const parsed = parseTtf(subset);
    const gid = parsed.cmap.get(A)!;
    expect(parsed.glyphOffsets[gid + 1]).toBeGreaterThan(parsed.glyphOffsets[gid]);
    expect(parsed.glyphOffsets[3]).toBeGreaterThan(parsed.glyphOffsets[2]);
  });

  it("writes a whole-font checksum a validator will accept", () => {
    // `head.checkSumAdjustment` has to be recomputed after the font is rebuilt. It
    // was left at the original font's value, so every subset font carried a
    // checksum describing bytes that no longer existed.
    const subset = subsetFont(font, "A");
    expect(wholeFontChecksum(subset)).toBe(0xb1b0afba);
  });

  it("drops the digital signature it has just invalidated", () => {
    // DSIG signs the bytes of the font. Subsetting rewrites them, so keeping the
    // table ships a signature that cannot verify — worse than shipping none.
    expect(tablesOf(font).has("DSIG")).toBe(true);
    expect(tablesOf(subsetFont(font, "A")).has("DSIG")).toBe(false);
  });

  it("subsets a font whose cmap is on the Unicode platform", () => {
    // This file used to carry its own `cmap` reader that accepted only Windows
    // (platform 3) subtables, while the shared reader accepts the Unicode platform
    // too. A font publishing only a (0,3) table — the shape several macOS system
    // faces ship — was therefore declared unsubsettable and embedded whole, silently.
    const patched = new Uint8Array(font);
    const cmap = tablesOf(font).get("cmap")!;
    patched[cmap.offset + 4] = 0; // platformID = 0 (Unicode)
    patched[cmap.offset + 5] = 0;
    patched[cmap.offset + 6] = 0; // encodingID = 3 (BMP)
    patched[cmap.offset + 7] = 3;

    const subset = subsetFont(patched, "A");
    expect(subset).not.toBe(patched);
    expect(subset.length).toBeLessThan(patched.length);
    expect(parseTtf(subset).cmap.get(A)).toBeDefined();
  });

  it("still returns the original when it cannot help", () => {
    const otf = new Uint8Array([0x4f, 0x54, 0x54, 0x4f, 0, 0, 0, 0]);
    expect(subsetFont(otf, "A")).toBe(otf);
  });
});
