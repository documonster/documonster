/**
 * How a caller may supply a font, and what each form is for.
 *
 * `RasterizeOptions.fonts` started as `Uint8Array[]`, which quietly ruled out two
 * things a caller needs. Bytes are re-parsed on every render — for a CJK face that is
 * a 43,000-entry `cmap` rebuilt each time, and a fresh object also misses the glyph
 * cache, which is keyed on outline identity. And bytes cannot name a face inside a
 * `.ttc`, so `Songti.ttc` always resolved to face 0.
 *
 * `.node.test.ts` because the cross-realm case uses `node:vm` to obtain a typed array
 * from another realm, which is the situation an iframe or a worker creates in a
 * browser — the one platform where supplying bytes is the *only* route to a non-ASCII
 * glyph, so dropping them silently is at its most damaging there.
 */

import vm from "node:vm";

import { parseRasterFont } from "@draw/raster/glyph-outline";
import type { RasterizeOptions } from "@draw/raster/surface";
import { rasterizeToRgba } from "@draw/raster/surface";
import type { DrawList } from "@draw/types";
import { buildCoverageFont } from "@test/ttf-fixture";
import { describe, expect, it } from "vitest";

const HAN_ZHONG = 0x4e2d; // 中

/** A one-label display list, so the assertions are about the font and nothing else. */
function textList(text: string): DrawList {
  return {
    width: 200,
    height: 60,
    children: [
      {
        kind: "text",
        x: 10,
        y: 40,
        lines: [{ text, dy: 0 }],
        style: { size: 24, family: "Arial", fill: { r: 0, g: 0, b: 0, a: 1 } }
      }
    ]
  };
}

/** Render 中 with `fonts`, and report what could not be drawn. */
function uncovered(fonts: RasterizeOptions["fonts"]): number[] {
  const image = rasterizeToRgba(textList("中"), { fonts, useSystemFonts: false });
  return [...image.uncoveredCodePoints];
}

describe("font sources", () => {
  it("accepts raw bytes", () => {
    expect(uncovered([buildCoverageFont([HAN_ZHONG], "Bytes")])).toEqual([]);
  });

  it("accepts an already-parsed face, so bytes are not re-parsed per render", () => {
    const parsed = parseRasterFont(buildCoverageFont([HAN_ZHONG], "Parsed"));
    expect(uncovered([parsed])).toEqual([]);
    // Reused across renders: the same object, so the glyph cache keyed on outline
    // identity still hits the second time.
    expect(uncovered([parsed])).toEqual([]);
  });

  it("accepts a face inside a collection by index", () => {
    // A single-face fixture is still a collection of one as far as the API goes; the
    // point is that the object form is accepted and the index is honoured.
    const data = buildCoverageFont([HAN_ZHONG], "Collection");
    expect(uncovered([{ data }])).toEqual([]);
    expect(uncovered([{ data, collectionIndex: 0 }])).toEqual([]);
    // An index past the end degrades to face 0 rather than throwing.
    expect(uncovered([{ data, collectionIndex: 99 }])).toEqual([]);
  });

  it("accepts a typed array from another realm", () => {
    const bytes = buildCoverageFont([HAN_ZHONG], "Foreign");
    const context = vm.createContext({});
    const foreign = vm.runInContext(`new Uint8Array(${bytes.length})`, context) as Uint8Array;
    foreign.set(bytes);

    // The premise: this is what makes a bare `instanceof` test wrong.
    expect(foreign instanceof Uint8Array).toBe(false);
    expect(ArrayBuffer.isView(foreign)).toBe(true);

    // It used to be mistaken for a `RasterFont`, fail the identity filter, and be
    // dropped without a word — the caller's font simply had no effect.
    expect(uncovered([foreign])).toEqual([]);
  });

  it("ignores a font it cannot parse without throwing", () => {
    expect(uncovered([new Uint8Array([1, 2, 3, 4])])).toEqual([HAN_ZHONG]);
    // A good font alongside a bad one still works.
    expect(uncovered([new Uint8Array([1, 2, 3, 4]), buildCoverageFont([HAN_ZHONG], "Ok")])).toEqual(
      []
    );
  });

  it("keeps a caller's own RasterFont that omits the optional members", () => {
    // `familyName` and `hasGlyph` are optional, because this interface was published
    // before they existed. Filtering on `familyName` threw such a font away.
    const real = parseRasterFont(buildCoverageFont([HAN_ZHONG], "Minimal"));
    const minimal = {
      unitsPerEm: real.unitsPerEm,
      ascent: real.ascent,
      descent: real.descent,
      getOutline: (cp: number) => real.getOutline(cp)
    };
    expect(uncovered([minimal])).toEqual([]);
  });
});
