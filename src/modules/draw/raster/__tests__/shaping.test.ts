/**
 * The rasteriser shapes complex scripts, and must not distort them while doing it.
 *
 * `drawText` makes two independent decisions about a run of text, and this suite pins
 * each of them, because they are gated on different predicates and conflating the two
 * broke a line in each direction:
 *
 * | Decision | Gate | Failure it caused |
 * | --- | --- | --- |
 * | Run the shaper | `isSimpleText` | Arabic drawn as isolated letters |
 * | Use the font's own advances | `isWellMeasuredText` | Tamil squeezed into an overlap |
 *
 * The second one is why this file exists. The glyph advances of a drawn line are
 * normally normalised to the width `@utils/text-measure` reported, so the ink fills the
 * box the layout reserved — but those tables only describe Latin, CJK, Cyrillic and
 * Greek. Tamil measures at roughly half its real width, so normalising a Tamil line
 * compresses it until the glyphs overlap. Nothing reported that: the glyphs were all
 * present, all drawn, and merely in the wrong places, which the existing suites had no
 * way to notice.
 *
 * Every font here is synthetic, for the reason given in `font-fallback.test.ts`:
 * asserting against a discovered system font would be a statement about the host.
 */

import { BasicRasterCanvas } from "@draw/raster/canvas";
import type { RasterFontSource } from "@draw/raster/glyph-outline";
import { buildCoverageFont } from "@test/ttf-fixture";
import { describe, expect, it } from "vitest";

/** Tamil base consonants — no virama, so the shaper is a no-op and only layout is tested. */
const TAMIL = "வணகம";

/** `مرحبا`, in base letters. Every one takes a contextual form when shaped. */
const ARABIC_HELLO = "\u0645\u0631\u062d\u0628\u0627";

const HEBREW_SHIN = 0x05e9; // ש — the only Hebrew letter the fixture below can draw

/** Draw `text` with exactly `fonts`, with the host's own fonts off. */
function render(text: string, fonts: readonly RasterFontSource[]): BasicRasterCanvas {
  const canvas = new BasicRasterCanvas(400, 60);
  canvas.setFonts(fonts, { useSystemFonts: false });
  canvas.drawText(10, 40, text, 24, "#000000", "start");
  return canvas;
}

/** The x of every column where ink starts after a gap — one per drawn glyph. */
function glyphStarts(canvas: BasicRasterCanvas): number[] {
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
  for (let x = 0; x < inked.length; x++) {
    if (inked[x] && !inked[x - 1]) {
      starts.push(x);
    }
  }
  return starts;
}

/** Distance from the first inked column to the last, i.e. how wide the line drew. */
function inkExtent(canvas: BasicRasterCanvas): number {
  let first = -1;
  let last = -1;
  for (let x = 0; x < canvas.width; x++) {
    for (let y = 0; y < canvas.height; y++) {
      if (canvas.data[(y * canvas.width + x) * 4 + 3] > 0) {
        if (first < 0) {
          first = x;
        }
        last = x;
        break;
      }
    }
  }
  return last - first;
}

/** A face covering `text`'s code points, every glyph an identical box at 0.7em. */
function faceFor(text: string, name: string): Uint8Array {
  return buildCoverageFont(
    [...text].map(c => c.codePointAt(0) as number),
    name
  );
}

describe("raster shaping", () => {
  it("spaces a complex script by the font's advances, not the measured width", () => {
    // The invariant, expressed so that it cannot depend on a metric the bug would also
    // corrupt: both strings are four glyphs drawn from one fixture, where every glyph
    // has the same 0.7em advance. Four glyphs at one advance occupy one width, whatever
    // script they belong to — so the two lines must draw to exactly the same extent.
    //
    // Normalising Tamil to `@utils/text-measure` breaks that: it reports roughly half
    // the real width for Tamil and the correct one for Latin, so the Tamil line is
    // squeezed to about a third of the Latin one and its glyphs overlap.
    const fonts = [faceFor(`${TAMIL}ABCD`, "Fixture Tamil+Latin")];
    const tamil = render(TAMIL, fonts);
    const latin = render("ABCD", fonts);

    expect(glyphStarts(tamil)).toHaveLength(4);
    expect(inkExtent(tamil)).toBe(inkExtent(latin));
    // And stated absolutely, so that a change making *both* wrong together still fails:
    // three advances of 0.7 × 24pt at scale 1 is 50.4px between the outer glyph origins.
    expect(inkExtent(tamil)).toBeGreaterThan(45);
  });

  it("draws Arabic through its contextual forms, so a base-letter face is not needed", () => {
    // The sharpest available statement that shaping ran: the face covers *only* the
    // presentation forms, and none of the base letters in the string. Ink can therefore
    // only appear if each letter was mapped to the form its position calls for. Drawn
    // one glyph per code point, this produces nothing at all and reports five uncovered
    // points — which is precisely the bug, and it looked like a missing font.
    const shapedForms = "\uFEE3\uFEAE\uFEA3\uFE92\uFE8E"; // initial/final forms of مرحبا
    const fonts = [
      buildCoverageFont(
        [...shapedForms].map(c => c.codePointAt(0) as number),
        "Fixture Arabic Forms"
      )
    ];

    const canvas = render(ARABIC_HELLO, fonts);
    expect(glyphStarts(canvas)).toHaveLength(5);
    expect([...canvas.uncoveredCodePoints]).toEqual([]);
  });

  it("puts a right-to-left run in visual order", () => {
    // Order is invisible when every glyph is the same box, so the fixture covers exactly
    // one letter of the four: where that single mark lands *is* the answer. `ש` is first
    // in logical order and last in visual order, so drawn correctly it sits to the right
    // of the run's midpoint. The other three letters still advance the pen — they are
    // reported uncovered rather than skipped — so the run's full width is used either way.
    const fonts = [buildCoverageFont([HEBREW_SHIN], "Fixture Hebrew Shin")];
    const canvas = render("\u05E9\u05DC\u05D5\u05DD", fonts); // שלום

    const starts = glyphStarts(canvas);
    expect(starts).toHaveLength(1);
    // Four glyphs at 0.7em × 24pt from x=10: the run spans 10 → 77, midpoint ~43.
    expect(starts[0]).toBeGreaterThan(43);
    // The three it cannot draw are named, so the caller is told what to install.
    expect([...canvas.uncoveredCodePoints].sort((a, b) => a - b)).toEqual([0x05d5, 0x05dc, 0x05dd]);
  });

  it("falls back to base letters when nothing in the chain has the contextual form", () => {
    // The regression this file exists to prevent a second time. Shaping was introduced here
    // *unconditionally*, and for a face carrying the base letters and none of the
    // presentation forms — eleven of the fifty Arabic-capable faces on one macOS host are
    // exactly that — the result was an **empty line**: 0 inked pixels, because a rasteriser
    // has no `.notdef` to fall back on. Worse, the uncovered report then named U+FE8E and
    // asked the caller to install a font for it, which nobody has.
    const arabicBase: number[] = [];
    for (let cp = 0x0600; cp <= 0x06ff; cp++) {
      arabicBase.push(cp);
    }
    const fonts = [buildCoverageFont(arabicBase, "Fixture Arabic Base Only")];
    const canvas = render(ARABIC_HELLO, fonts);

    // Every letter is drawn — unjoined, which is the point of the fallback.
    expect(glyphStarts(canvas)).toHaveLength(5);
    // And the caller is not sent after a font for a presentation form.
    expect([...canvas.uncoveredCodePoints]).toEqual([]);
  });

  it("reports what shaping still cannot do, and does not send the caller elsewhere", () => {
    // The warning is user-facing and was wrong the moment shaping landed: it told callers
    // their Arabic would "render incorrectly" and to use SVG instead, which is advice away
    // from a backend that now handles it. It has to describe this rasteriser's real gap.
    const fonts = [faceFor(ARABIC_HELLO + TAMIL, "Fixture Complex")];
    const canvas = render(`${ARABIC_HELLO} ${TAMIL}`, fonts);

    const warnings = canvas.textWarnings();
    expect(warnings.join(" ")).toContain("Arabic");
    expect(warnings.join(" ")).not.toContain("render incorrectly");
  });

  it("leaves simple text on the measured-width path", () => {
    // The converse guard, stated as the property itself: for text the measurer describes,
    // the line's width comes from the *measurement*, so widening the face's own advance by
    // 43% must not move the ink at all. Under natural widths the same pair renders 64px
    // and 86px apart.
    //
    // This needs the advance override to say anything. Written with the default fixture
    // for both faces — 0.7em, within a pixel of Arial's Latin metrics — it passed whichever
    // mode ran, and forcing natural widths for *all* text did not fail it.
    const cps = [..."ABCD"].map(c => c.codePointAt(0) as number);
    const narrow = buildCoverageFont(cps, "Fixture Latin 0.7em", {}, 700);
    const wide = buildCoverageFont(cps, "Fixture Latin 1.0em", {}, 1000);
    expect(inkExtent(render("ABCD", [wide]))).toBe(inkExtent(render("ABCD", [narrow])));

    // And a bidi isolate asks for reordering but is measured as nothing, so it must change
    // neither the glyph positions nor the count.
    const before = glyphStarts(render("ABCD", [wide]));
    expect(before).toHaveLength(4);
    expect(glyphStarts(render("AB\u2066CD", [wide]))).toEqual(before);
  });
});
