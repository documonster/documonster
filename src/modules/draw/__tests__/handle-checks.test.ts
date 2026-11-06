/** A wrong handle is a `DrawArgumentError`, not a `TypeError` (X10). */
import { Draw, DrawArgumentError, createRasterSurface, fontHasGlyph } from "@draw/index";
import { describe, expect, it } from "vitest";

describe("handles are checked", () => {
  it("Draw.Backend.fontHasGlyph and its deprecated forward", () => {
    for (const fn of [Draw.Backend.fontHasGlyph, fontHasGlyph]) {
      for (const bad of [42, null, {}]) {
        expect(() => fn(bad as never, 65)).toThrow(DrawArgumentError);
      }
    }
    const font = { unitsPerEm: 1000, ascent: 800, descent: -200, getOutline: () => undefined };
    expect(Draw.Backend.fontHasGlyph(font, 65)).toBe(false);
  });

  it("the deprecated createRasterSurface", () => {
    for (const bad of [42, null, {}]) {
      expect(() => createRasterSurface(bad as never)).toThrow(DrawArgumentError);
    }
  });
});
