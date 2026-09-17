/**
 * The rasteriser must draw CJK, and must say so when it cannot.
 *
 * This is the regression suite for a bug that was invisible to every existing
 * test: `draw/raster` had no font discovery — a hardcoded list of Latin filenames,
 * one face, no fallback — so a PNG of a Chinese label came out **blank** while the
 * PDF of the same display list embedded a real CJK face and rendered correctly.
 * A missing glyph advanced the pen and painted nothing, with no error and no
 * warning, so the suite stayed green.
 *
 * `draw-backend-parity.test.ts` could not catch it either: it compares *geometry*
 * across backends, and a glyph that was never drawn contributes no geometry to
 * disagree about.
 *
 * Every font here is synthetic (`font-fixture.helpers.ts`). Asserting against a
 * discovered system font would pass on a macOS workstation and fail on a CI
 * container with no CJK font installed, which is a statement about the host rather
 * than the code.
 */

import { BasicRasterCanvas } from "@draw/raster/canvas";
import type { RasterFont, RasterFontSource } from "@draw/raster/glyph-outline";
import {
  buildRasterFont,
  fontHasGlyph,
  parseInjectedFonts,
  parseRasterFont
} from "@draw/raster/glyph-outline";
import { rasterizeToRgba } from "@draw/raster/surface";
import { resolveFontChain } from "@draw/raster/system-raster-font";
import type { DrawList } from "@draw/types";
import { buildCoverageFont, fontWithoutGlyf } from "@test/ttf-fixture";
import { parseTtf } from "@utils/font-ttf";
import { describe, expect, it } from "vitest";

const HAN_ZHONG = 0x4e2d; // 中
const HAN_WEN = 0x6587; // 文
const KANA_HI = 0x3072; // ひ
const HANGUL_HAN = 0xd55c; // 한
const UNCOVERABLE = 0x2a6a5; // 𪚥 — outside every fixture

/** A face that draws only ASCII letters, like a real Latin font. */
function latinFont(): Uint8Array {
  const cps: number[] = [];
  for (let cp = 0x41; cp <= 0x5a; cp++) {
    cps.push(cp);
  }
  for (let cp = 0x61; cp <= 0x7a; cp++) {
    cps.push(cp);
  }
  return buildCoverageFont(cps, "Fixture Latin");
}

/** A face that draws only the CJK code points used here. */
function cjkFont(): Uint8Array {
  return buildCoverageFont([HAN_ZHONG, HAN_WEN, KANA_HI, HANGUL_HAN], "Fixture CJK");
}

/** How many pixels of the canvas carry ink. */
function inkedPixels(canvas: BasicRasterCanvas): number {
  let count = 0;
  for (let i = 3; i < canvas.data.length; i += 4) {
    if (canvas.data[i] > 0) {
      count++;
    }
  }
  return count;
}

/**
 * Draw `text` on a fresh canvas using exactly `fonts`.
 *
 * The host's own fonts are always off: a test about an *uncovered* character has no
 * way to arrange one otherwise, since the host obligingly draws almost anything and
 * the branch under test never runs. Fonts are an argument rather than ambient state so
 * that no test can be affected by what another one registered.
 */
function render(text: string, fonts: readonly RasterFontSource[]): BasicRasterCanvas {
  const canvas = new BasicRasterCanvas(400, 60);
  canvas.setFonts(fonts, { useSystemFonts: false });
  canvas.drawText(10, 40, text, 24, "#000000", "start");
  return canvas;
}

/** Draw `text` with `fonts` and report the ink and any uncovered points. */
function draw(
  text: string,
  fonts: readonly RasterFontSource[]
): { ink: number; uncovered: number[] } {
  const canvas = render(text, fonts);
  return {
    ink: inkedPixels(canvas),
    uncovered: [...canvas.uncoveredCodePoints].sort((a, b) => a - b)
  };
}

describe("raster font fallback", () => {
  it("draws CJK when a covering face is supplied", () => {
    const { ink, uncovered } = draw("中文", [cjkFont()]);
    // The bug: this was 0.
    expect(ink).toBeGreaterThan(0);
    expect(uncovered).toEqual([]);
  });

  it("draws both scripts of a mixed label from two single-script faces", () => {
    // Neither face can draw the whole string, which is exactly the case a single
    // font cannot serve and the old code could not express.
    const fonts = [latinFont(), cjkFont()];
    const mixed = draw("AB中文", fonts);
    const latinOnly = draw("AB", fonts);
    const cjkOnly = draw("中文", fonts);

    expect(mixed.uncovered).toEqual([]);
    // Ink from both halves is present, so neither script was silently dropped.
    expect(mixed.ink).toBeGreaterThan(latinOnly.ink);
    expect(mixed.ink).toBeGreaterThan(cjkOnly.ink);
  });

  it("walks the whole chain rather than stopping at the first face", () => {
    // CJK face first: the Latin characters can only come from the second entry.
    const { ink, uncovered } = draw("AB中文", [cjkFont(), latinFont()]);
    expect(uncovered).toEqual([]);
    expect(ink).toBeGreaterThan(0);
  });

  it("reports a code point no face can draw instead of failing silently", () => {
    const { uncovered } = draw(`A中${String.fromCodePoint(UNCOVERABLE)}`, [latinFont(), cjkFont()]);
    expect(uncovered).toEqual([UNCOVERABLE]);
  });

  it("does not treat .notdef as coverage", () => {
    // 中 is *in* this face's cmap, pointing at glyph 0. `cmap.has` therefore says
    // yes while the glyph draws a `.notdef` box, so a chain that trusts `has`
    // stops at a face that cannot really draw the character.
    const font = notdefMapper();
    expect(fontHasGlyph(font, 0x41)).toBe(true);
    expect(fontHasGlyph(font, HAN_ZHONG)).toBe(false);
    expect(font.getOutline(HAN_ZHONG)).toBeUndefined();
  });

  it("falls through a face that maps the character to .notdef", () => {
    // The end-to-end form of the case above: the first face claims 中 and cannot
    // draw it, so the ink has to come from the second.
    const { ink, uncovered } = draw("中", [notdefMapper(), cjkFont()]);
    expect(uncovered).toEqual([]);
    expect(ink).toBeGreaterThan(0);
  });

  it("keeps supplied faces ahead of anything discovered", () => {
    const chain = resolveFontChain("中文", parseInjectedFonts([cjkFont()]), true);
    expect(chain[0].familyName).toBe("Fixture CJK");
  });

  it("skips a font it cannot parse instead of throwing", () => {
    const fonts = parseInjectedFonts([new Uint8Array([1, 2, 3, 4]), cjkFont()]);
    expect(fonts.map(f => f.familyName)).toEqual(["Fixture CJK"]);
    expect(draw("中", fonts).uncovered).toEqual([]);
  });

  it("degrades a font with no outlines to one that draws nothing", () => {
    // `parseTtf` rejects it — a CFF face or Apple's private `hvgl` reaches here —
    // and the rasteriser has nowhere to put an exception.
    const font = parseRasterFont(fontWithoutGlyf());
    expect(fontHasGlyph(font, 0x41)).toBe(false);
    expect(font.getOutline(0x41)).toBeUndefined();
    expect(font.unitsPerEm).toBeGreaterThan(0);
  });

  it("keeps fonts on the canvas that was given them", () => {
    // Two canvases alive at once, configured differently. Under the previous
    // implementation `fonts` went into a process-wide registry, so the second
    // configuration overwrote the first and both drew the same thing — which also
    // meant one render silently changed every later one.
    const withCjk = new BasicRasterCanvas(200, 60);
    withCjk.setFonts([cjkFont()], { useSystemFonts: false });
    const withNothing = new BasicRasterCanvas(200, 60);
    withNothing.setFonts([], { useSystemFonts: false });

    withCjk.drawText(10, 40, "中", 24, "#000000", "start");
    withNothing.drawText(10, 40, "中", 24, "#000000", "start");

    expect([...withCjk.uncoveredCodePoints]).toEqual([]);
    expect([...withNothing.uncoveredCodePoints]).toEqual([HAN_ZHONG]);
    expect(inkedPixels(withCjk)).toBeGreaterThan(0);
  });

  it("scopes rasterizeToRgba's fonts to the call that passed them", () => {
    // The leak this closes: routing the option through the process-wide registry
    // left the first render's fonts in place for the second, so two renders in one
    // process could not use different faces — and supplying a font once silently
    // changed every later render.
    const list = {
      width: 200,
      height: 60,
      children: [
        {
          kind: "text" as const,
          x: 10,
          y: 40,
          lines: [{ text: "中", dy: 0 }],
          style: { size: 24, family: "Arial", fill: { r: 0, g: 0, b: 0, a: 1 } }
        }
      ]
    };

    const withFont = rasterizeToRgba(list, { fonts: [cjkFont()], useSystemFonts: false });
    expect(withFont.uncoveredCodePoints).toEqual([]);

    // Same list, no fonts, host off: nothing should carry over from the call above.
    const withoutFont = rasterizeToRgba(list, { useSystemFonts: false });
    expect(withoutFont.uncoveredCodePoints).toEqual([HAN_ZHONG]);

    // And the order does not matter either way.
    const againWithFont = rasterizeToRgba(list, { fonts: [cjkFont()], useSystemFonts: false });
    expect(againWithFont.uncoveredCodePoints).toEqual([]);
  });

  it("never reports or draws a formatting control", () => {
    // A variation selector, a joiner, a bidi isolate and a zero-width space have no
    // glyph in any font. Looking one up made it "uncovered", so a caller was told to
    // install a font for a character that cannot have one — and whether it showed up
    // at all depended on whether the host's face happened to carry an empty glyph,
    // which made the report differ between machines.
    const fonts = [latinFont(), cjkFont()];
    const controls = [
      "\uFE0F", // VS16
      "\u200D", // ZWJ
      "\u200C", // ZWNJ
      "\u200B", // ZWSP
      "\u2060", // word joiner
      "\uFEFF", // BOM
      "\u2066", // bidi isolate
      "\u202A", // bidi embedding
      "\u{E0100}" // ideographic variation selector
    ];
    for (const control of controls) {
      expect(draw(`A${control}中`, fonts).uncovered).toEqual([]);
    }
  });

  it("puts no ink down for a formatting control", () => {
    // The report was only half of it: the control was also charged an advance and, on
    // the stroke path, drawn as `?` — inventing a visible character the caller never
    // wrote. A control on its own must produce an empty image.
    const fonts = [latinFont(), cjkFont()];
    for (const control of ["\uFE0F", "\u200D", "\u200B", "\u2060", "\uFEFF", "\u2066"]) {
      expect(inkedPixels(render(control, fonts))).toBe(0);
    }

    // Same again with no font at all, which is the ASCII-only stroke font's path and
    // the one that used to substitute `?`.
    for (const control of ["\uFE0F", "\u200B", "\u2066"]) {
      expect(inkedPixels(render(control, []))).toBe(0);
      expect(draw(control, []).uncovered).toEqual([]);
    }
  });

  it("leaves visible glyphs where they were when a control is inserted", () => {
    // The invariant that ties the two halves together. A line's glyph advances are
    // normalised to its *measured* width, so a control character the measurer charges
    // for moves every visible glyph — even once the rasteriser stops drawing it. This
    // assertion was impossible until `@utils/text-measure` stopped charging them.
    const fonts = [latinFont()];
    const baseline = columnsWithInk(render("AB", fonts));
    expect(baseline).toHaveLength(2);

    for (const control of [
      "\u200B", // zero width space
      "\u2066", // bidi isolate
      "\uFEFF", // BOM
      "\u200D", // ZWJ
      "\u0000", // NUL
      "\u0301" // combining acute — no width, though it would draw if the face had it
    ]) {
      expect(columnsWithInk(render(`A${control}B`, fonts))).toEqual(baseline);
    }
  });

  it("advances the pen for a tab without drawing or reporting it", () => {
    // Tab is width without ink: skipping it would pull the rest of the line left, and
    // treating it as a missing glyph would draw `?` and ask for a font that cannot help.
    const fonts = [latinFont()];
    const baseline = columnsWithInk(render("AB", fonts));
    const tabbed = columnsWithInk(render("A\tB", fonts));

    expect(tabbed).toHaveLength(2);
    expect(tabbed[0]).toBe(baseline[0]);
    expect(tabbed[1]).toBeGreaterThan(baseline[1]);
    expect(draw("A\tB", fonts).uncovered).toEqual([]);
    // And with no font at all, where the stroke path would have substituted `?`.
    expect(draw("A\tB", []).uncovered).toEqual([]);
    expect(inkedPixels(render("\t", []))).toBe(0);
  });

  it("does not stretch surviving text to fill a missing glyph's width", () => {
    // A missing character's assumed advance has to match what `measureText`
    // assumed, or `hScale` redistributes the difference over everything else. It
    // was a flat 0.4 em while an ideograph measures a full em, so the Latin in a
    // mixed label was visibly spread apart.
    const fonts = [latinFont()];
    const latinAlone = columnsWithInk(render("AB", fonts));
    const latinBeforeMissing = columnsWithInk(render("AB中", fonts));

    // The two 'A' glyphs must start at the same column and the 'B' must not have
    // drifted: the missing ideograph takes its own em and no more.
    expect(latinBeforeMissing[0]).toBe(latinAlone[0]);
    expect(Math.abs(latinBeforeMissing[1] - latinAlone[1])).toBeLessThanOrEqual(1);
  });
});

/**
 * A face whose `cmap` points 中 at glyph 0.
 *
 * Built by editing a parsed font rather than by assembling bytes, because
 * `readCmapFormat4` drops a glyph-0 entry as it parses — the byte-level fixture
 * cannot express this, but `buildRasterFont`'s contract still has to hold for it.
 */
function notdefMapper(): RasterFont {
  const base = parseTtf(buildCoverageFont([0x41], "Notdef Mapper"));
  return buildRasterFont({ ...base, cmap: new Map([...base.cmap, [HAN_ZHONG, 0]]) });
}

/** Leftmost inked column of each of the first two ink runs. */
function columnsWithInk(canvas: BasicRasterCanvas): number[] {
  const inked: boolean[] = [];
  for (let x = 0; x < canvas.width; x++) {
    let hit = false;
    for (let y = 0; y < canvas.height && !hit; y++) {
      if (canvas.data[(y * canvas.width + x) * 4 + 3] > 0) {
        hit = true;
      }
    }
    inked.push(hit);
  }
  const starts: number[] = [];
  for (let x = 1; x < inked.length; x++) {
    if (inked[x] && !inked[x - 1]) {
      starts.push(x);
    }
  }
  if (inked[0]) {
    starts.unshift(0);
  }
  return starts;
}

describe("reporting text the rasteriser cannot lay out", () => {
  /**
   * A missing glyph leaves a hole; text needing shaping is drawn in full, in the wrong
   * shapes and the wrong order, and looks fine to anyone who cannot read the script.
   * That is the worse failure, and it was silent — the detection existed only in the
   * PDF writer, which `draw` cannot import.
   */
  const list = (text: string): DrawList => ({
    width: 300,
    height: 40,
    children: [
      {
        kind: "text",
        x: 10,
        y: 28,
        lines: [{ text, dy: 0 }],
        style: { size: 18, fill: { r: 0, g: 0, b: 0, a: 1 } }
      }
    ]
  });

  it("stays quiet for text it renders correctly", () => {
    // The warning is only useful if it does not fire for ordinary text. Latin with
    // diacritics, CJK, Cyrillic and Greek are all laid out correctly one glyph per
    // code point.
    for (const text of [
      "Hello world",
      "Café à l'œuvre",
      "数据校验",
      "日本語テスト",
      "Привет",
      "Γειά σου"
    ]) {
      expect(rasterizeToRgba(list(text)).textWarnings, text).toEqual([]);
    }
  });

  it("reports Arabic as both unshaped and misordered", () => {
    const warnings = rasterizeToRgba(list("مرحبا بالعالم")).textWarnings;
    expect(warnings).toHaveLength(2);
    expect(warnings[0]).toContain("Arabic");
    expect(warnings[1]).toContain("right-to-left");
  });

  it("reports Hebrew as misordered but not unshaped", () => {
    const warnings = rasterizeToRgba(list("שלום עולם")).textWarnings;
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toContain("right-to-left");
  });

  it("reports Indic scripts as unshaped", () => {
    for (const [script, text] of [
      ["Devanagari", "नमस्ते"],
      ["Tamil", "வணக்கம்"],
      ["Thai", "สวัสดี"]
    ] as [string, string][]) {
      const warnings = rasterizeToRgba(list(text)).textWarnings;
      expect(warnings, script).toHaveLength(1);
      expect(warnings[0]).toContain(script);
    }
  });

  it("says supplying a font will not help, and what will", () => {
    // Unlike `uncoveredCodePoints`, this is not a coverage problem: the glyphs are all
    // there. The message has to name a route that actually works.
    const [warning] = rasterizeToRgba(list("नमस्ते")).textWarnings;
    expect(warning).toContain("SVG");
    expect(warning).toContain("DOCX");
  });

  it("reports text drawn at zero alpha, and nothing when there is no fill", () => {
    // Invisible text still cannot be laid out, so the finding holds; a text node with
    // no fill at all never reaches the canvas, so there is nothing to report.
    const invisible: DrawList = {
      width: 200,
      height: 40,
      children: [
        {
          kind: "text",
          x: 10,
          y: 28,
          lines: [{ text: "नमस्ते", dy: 0 }],
          style: { size: 18, fill: { r: 0, g: 0, b: 0, a: 0 } }
        }
      ]
    };
    expect(rasterizeToRgba(invisible).textWarnings).toHaveLength(1);

    const unfilled: DrawList = {
      width: 200,
      height: 40,
      children: [
        { kind: "text", x: 10, y: 28, lines: [{ text: "नमस्ते", dy: 0 }], style: { size: 18 } }
      ]
    };
    expect(rasterizeToRgba(unfilled).textWarnings).toEqual([]);
  });

  it("reports each script once across a whole image", () => {
    const many: DrawList = {
      width: 300,
      height: 400,
      children: Array.from({ length: 20 }, (_unused, i) => ({
        kind: "text" as const,
        x: 10,
        y: 20 + i * 18,
        lines: [{ text: "नमस्ते", dy: 0 }],
        style: { size: 14, fill: { r: 0, g: 0, b: 0, a: 1 } }
      }))
    };
    expect(rasterizeToRgba(many).textWarnings).toHaveLength(1);
  });
});
