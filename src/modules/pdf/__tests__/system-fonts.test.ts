import { afterEach, describe, it, expect } from "vitest";

import {
  discoverSystemFont,
  discoverSystemFontCandidates,
  resetFontDiscoveryCache,
  _setCandidatesForTest
} from "../font/system-fonts";
import { pdf } from "../pdf";
import { buildMinimalTtf, buildTtfWithCmap } from "./ttf-test-utils";

// ===========================================================================
// Tests
// ===========================================================================

describe("System font discovery", () => {
  afterEach(() => {
    resetFontDiscoveryCache();
  });

  it("should not throw when resetting the cache", () => {
    expect(() => resetFontDiscoveryCache()).not.toThrow();
  });

  it("should return either null or a Uint8Array", () => {
    resetFontDiscoveryCache();
    const result = discoverSystemFont();
    if (result !== null) {
      expect(result).toBeInstanceOf(Uint8Array);
    } else {
      expect(result).toBeNull();
    }
  });

  it("should return a non-trivial Uint8Array if a font is found", () => {
    resetFontDiscoveryCache();
    const result = discoverSystemFont();
    if (result !== null) {
      expect(result.length).toBeGreaterThan(1000);
    }
  });

  it("should return the same result on repeated calls (caching)", () => {
    resetFontDiscoveryCache();
    const first = discoverSystemFont();
    const second = discoverSystemFont();

    if (first === null) {
      expect(second).toBeNull();
    } else {
      expect(second).toBe(first); // same reference — cached
    }
  });

  it("discoverSystemFontCandidates should return an array", () => {
    resetFontDiscoveryCache();
    const candidates = discoverSystemFontCandidates();
    expect(Array.isArray(candidates)).toBe(true);
    for (const c of candidates) {
      expect(c).toBeInstanceOf(Uint8Array);
      expect(c.length).toBeGreaterThan(1000);
    }
  });

  it("discoverSystemFont should return the first candidate", () => {
    resetFontDiscoveryCache();
    const candidates = discoverSystemFontCandidates();
    const first = discoverSystemFont();

    if (candidates.length === 0) {
      expect(first).toBeNull();
    } else {
      expect(first).toBe(candidates[0]);
    }
  });
});

describe("System font candidate iteration", () => {
  afterEach(() => {
    resetFontDiscoveryCache();
  });

  it("should skip a candidate that does not cover required chars and use the next one", async () => {
    // Candidate 1: only covers A/B (U+0041-0042) — does NOT cover ☐ (U+2610)
    const ttfNarrow = buildMinimalTtf();

    // Candidate 2: covers A/B AND ☐ (U+2610)
    // U+0041-0042 → glyphs 1-2 (delta = -0x40)
    // U+2610       → glyph 3  (delta = -0x260D)
    const ttfBroad = buildTtfWithCmap(
      [
        { start: 0x41, end: 0x42, delta: -0x40 },
        { start: 0x2610, end: 0x2610, delta: -0x260d }
      ],
      4 // .notdef + A + B + ☐
    );

    // Inject: narrow first, broad second
    _setCandidatesForTest([ttfNarrow, ttfBroad]);

    // Generate a PDF containing ☐ — the exporter must skip ttfNarrow and use ttfBroad
    const result = await pdf([["A", "☐"]]);
    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(100);

    // The PDF should contain an embedded font (TestFont) — proof that ttfBroad was used
    const pdfText = new TextDecoder("latin1").decode(result);
    expect(pdfText).toContain("TestFont");
  });

  it("should fall back to Type3 when no candidate covers the required chars", async () => {
    // Only candidate: covers A/B but NOT ☐ (U+2610)
    const ttfNarrow = buildMinimalTtf();
    _setCandidatesForTest([ttfNarrow]);

    // Generate a PDF containing ☐ — no candidate covers it, so Type3 fallback
    const result = await pdf([["A", "☐"]]);
    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(100);

    // The PDF should NOT contain "TestFont" — no embedded font was selected
    const pdfText = new TextDecoder("latin1").decode(result);
    expect(pdfText).not.toContain("TestFont");
  });

  it("should fall back to Type3 when candidate list is empty", async () => {
    _setCandidatesForTest([]);

    const result = await pdf([["Hello", "☐"]]);
    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(100);
  });

  it("should not attempt candidates when text is all WinAnsi", async () => {
    // Inject a candidate that would fail to parse if used
    const garbage = new Uint8Array([0, 1, 2, 3]);
    _setCandidatesForTest([garbage]);

    // Pure ASCII — no non-WinAnsi chars, so candidate loop should never run
    const result = await pdf([["Hello", "World"]]);
    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(100);
  });
});
