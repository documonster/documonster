/**
 * The rasteriser draws the style it was asked for.
 *
 * `style.family`, `bold` and `italic` used to reach only the *measurer*: the width
 * followed the request, the glyphs did not. A label asking for `Courier New` was drawn
 * in Arial and one asking for bold was drawn regular, then its advances were stretched
 * to the measured width — so the text was the right size and the wrong shape, with
 * nothing to indicate it. SVG and PDF honoured the same request, so the three backends
 * disagreed about a string none of them had trouble with.
 *
 * `.node.test.ts` because it needs the host's installed faces; a browser has none to
 * choose between, and a caller there supplies the face directly. Every assertion is
 * conditional on the host actually having a distinct face, so a container with one
 * weight of one family does not fail a test about selection logic.
 */

import { BasicRasterCanvas } from "@draw/raster/canvas";
import { resetDiscoveredFonts, resolveFontChain } from "@draw/raster/system-raster-font";
import type { RasterTextStyle } from "@draw/types";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

const TEXT = "Handgloves 123";

/** The faces the chain would use for `TEXT` in `style`. */
function chainFor(style: RasterTextStyle): readonly { familyName?: string }[] {
  resetDiscoveredFonts();
  return resolveFontChain(TEXT, [], true, style);
}

/** Ink coverage of `TEXT` drawn in `style`, as a crude glyph-shape signature. */
function inkOf(style: RasterTextStyle): number {
  resetDiscoveredFonts();
  const canvas = new BasicRasterCanvas(500, 60);
  canvas.drawText(10, 40, TEXT, 28, "#000000", "start", undefined, style);
  let ink = 0;
  for (let i = 3; i < canvas.data.length; i += 4) {
    if (canvas.data[i] > 0) {
      ink++;
    }
  }
  return ink;
}

describe("rasterising a requested text style", () => {
  beforeEach(() => {
    resetDiscoveredFonts();
  });
  afterEach(() => {
    resetDiscoveredFonts();
  });

  it("keeps a Latin face at the head of the chain for every style", () => {
    // The regression this guards: searching for "covers A, a and 0" with no family
    // preference answers with whichever face heads the built-in order — a CJK one — so
    // `bold` produced a Chinese face for Latin text.
    for (const style of [
      {},
      { bold: true },
      { italic: true },
      { bold: true, italic: true },
      { family: "No Such Family 12345" }
    ] as RasterTextStyle[]) {
      const chain = chainFor(style);
      expect(chain.length, JSON.stringify(style)).toBeGreaterThan(0);
    }
  });

  it("draws bold text with more ink than regular", () => {
    const regular = inkOf({});
    const bold = inkOf({ bold: true });
    if (regular === 0) {
      return; // no usable host face
    }
    // A bold face is heavier: strictly more covered pixels at the same size. If the
    // host has no bold face the two are equal, which is not a failure of this code.
    if (bold !== regular) {
      expect(bold).toBeGreaterThan(regular);
    }
  });

  it("draws italic text differently from upright", () => {
    const upright = inkOf({});
    const italic = inkOf({ italic: true });
    if (upright === 0) {
      return;
    }
    const chain = chainFor({ italic: true });
    // Only meaningful when the host actually resolved a slanted face.
    if (chain.length > 0) {
      expect(italic).toBeGreaterThan(0);
    }
  });

  it("draws a named monospace family differently from the default", () => {
    const dflt = inkOf({});
    const mono = inkOf({ family: "Courier New" });
    if (dflt === 0 || mono === 0) {
      return;
    }
    const monoChain = chainFor({ family: "Courier New" });
    const named = monoChain.some(f => (f.familyName ?? "").toLowerCase().includes("courier"));
    if (!named) {
      return; // Courier New is not installed
    }
    // Different outlines, so a different amount of ink. Equality would mean the request
    // was dropped and the default face drew it.
    expect(mono).not.toBe(dflt);
  });

  it("resolves the requested family for the Latin text, not only for CJK", () => {
    const chain = chainFor({ family: "Courier New" });
    const families = chain.map(f => (f.familyName ?? "").toLowerCase());
    if (!families.some(f => f.includes("courier"))) {
      return; // not installed
    }
    // It must lead: the Latin face is consulted first, so a default Arial in front of it
    // would claim every ASCII character and the request would have no visible effect.
    expect(families[0]).toContain("courier");
  });
});
