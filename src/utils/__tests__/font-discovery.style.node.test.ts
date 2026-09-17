/**
 * Choosing a face by family, weight and slant.
 *
 * The search preferred `usWeightClass === 400` unconditionally and never looked at the
 * slant at all, so a caller asking for bold or italic text got a regular upright face.
 * For the PDF embedder that was masked — it synthesises neither, and the viewer resolves
 * the family name itself — but the rasteriser draws the outlines it is handed, so bold
 * text came out regular with nothing to indicate it.
 *
 * The candidates are **synthetic**, injected with `_setCandidatesForTest`. An earlier
 * version of this file asked the host for a bold face and skipped its assertion when the
 * answer came back at weight 400 — which is exactly what a broken selector returns, so
 * the test passed against one. A fixture that declares all four faces can distinguish
 * "the host has no bold" from "the code ignored the request".
 *
 * `.node.test.ts` because the discovery module is Node-only.
 */

import { buildCoverageFont } from "@test/ttf-fixture";
import type { FindFaceOptions } from "@utils/font-discovery";
import {
  _setCandidatesForTest,
  findSystemFontForCodePoints,
  resetFontDiscoveryCache
} from "@utils/font-discovery";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

/** Upper case, lower case and a digit. */
const LATIN = new Set([0x41, 0x61, 0x30]);
const CODE_POINTS = [0x41, 0x61, 0x30];

/** One family in all four faces, plus a second family to test ordering. */
function installFixtureFaces(): void {
  _setCandidatesForTest([
    buildCoverageFont(CODE_POINTS, "Fixture Sans", { weight: 400 }),
    buildCoverageFont(CODE_POINTS, "Fixture Sans", { weight: 700, bold: true }),
    buildCoverageFont(CODE_POINTS, "Fixture Sans", { weight: 400, italic: true }),
    buildCoverageFont(CODE_POINTS, "Fixture Sans", { weight: 700, bold: true, italic: true }),
    buildCoverageFont(CODE_POINTS, "Fixture Serif", { weight: 400 })
  ]);
}

function find(options: FindFaceOptions) {
  return findSystemFontForCodePoints(LATIN, options);
}

describe("selecting a face by style", () => {
  beforeEach(() => {
    resetFontDiscoveryCache();
    installFixtureFaces();
  });
  afterEach(() => {
    resetFontDiscoveryCache();
    _setCandidatesForTest([]);
  });

  it("prefers a regular upright face by default", () => {
    const face = find({ families: ["fixture sans"] })!;
    expect(face.weightClass).toBe(400);
    expect(face.italic).toBe(false);
  });

  it("picks the bold face when one is asked for", () => {
    const face = find({ families: ["fixture sans"], weight: 700 })!;
    expect(face.weightClass).toBe(700);
    expect(face.italic).toBe(false);
  });

  it("picks the italic face when one is asked for", () => {
    const face = find({ families: ["fixture sans"], italic: true })!;
    expect(face.italic).toBe(true);
    expect(face.weightClass).toBe(400);
  });

  it("picks the bold italic face when both are asked for", () => {
    const face = find({ families: ["fixture sans"], weight: 700, italic: true })!;
    expect(face.weightClass).toBe(700);
    expect(face.italic).toBe(true);
  });

  it("treats slant as more important than weight", () => {
    // Only a bold upright and a regular italic are on offer, and an italic is wanted at
    // a bold weight. Slant is a shape difference a reader sees immediately; a weight one
    // step off is not — so the italic must win despite the worse weight match.
    resetFontDiscoveryCache();
    _setCandidatesForTest([
      buildCoverageFont(CODE_POINTS, "Fixture Sans", { weight: 700, bold: true }),
      buildCoverageFont(CODE_POINTS, "Fixture Sans", { weight: 400, italic: true })
    ]);
    const face = find({ families: ["fixture sans"], weight: 700, italic: true })!;
    expect(face.italic).toBe(true);
    expect(face.weightClass).toBe(400);
  });

  it("honours the family order it was given", () => {
    expect(find({ families: ["fixture serif", "fixture sans"] })!.familyName).toBe("Fixture Serif");
    expect(find({ families: ["fixture sans", "fixture serif"] })!.familyName).toBe("Fixture Sans");
  });

  it("falls back past a family that is not installed", () => {
    const face = find({ families: ["no such family 12345", "fixture sans"] })!;
    expect(face.familyName).toBe("Fixture Sans");
  });

  it("returns the nearest weight when the exact one is absent", () => {
    resetFontDiscoveryCache();
    _setCandidatesForTest([
      buildCoverageFont(CODE_POINTS, "Fixture Sans", { weight: 300 }),
      buildCoverageFont(CODE_POINTS, "Fixture Sans", { weight: 900, bold: true })
    ]);
    // 700 is nearer 900 than 300, so a bold request must not settle for the light face
    // just because it is closer to the old hardcoded 400.
    expect(find({ families: ["fixture sans"], weight: 700 })!.weightClass).toBe(900);
    expect(find({ families: ["fixture sans"] })!.weightClass).toBe(300);
  });
});
