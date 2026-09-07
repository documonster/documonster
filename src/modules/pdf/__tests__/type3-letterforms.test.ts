import { describe, expect, it } from "vitest";

import { STROKE_FONT } from "../../draw/raster/stroke-font";
import { isWinAnsiCodePoint } from "../core/pdf-stream";
import { GLYPH_BBOX } from "../font/type3-font";
import { lookupGlyph, NOTDEF_GLYPH, type GlyphPen } from "../font/type3-glyphs";
import { CYRILLIC, CYRILLIC_CODE_POINTS } from "../font/type3-glyphs-cyrillic";
import { GREEK, GREEK_CODE_POINTS } from "../font/type3-glyphs-greek";
import { LATIN_EXTENDED_CODE_POINTS } from "../font/type3-glyphs-latin-ext";
import { isType3Drawable } from "../font/type3-repertoire";

/**
 * Record what a glyph draws.
 *
 * The pen is the whole interface a glyph has to the world, so capturing its calls is
 * the only way to assert about an outline without rendering one.
 */
function trace(glyph: { draw: (p: GlyphPen) => void }): {
  ops: string[];
  points: Array<readonly [number, number]>;
  /** Points grouped by subpath, so a stroke's continuous extent can be measured. */
  subpaths: Array<Array<readonly [number, number]>>;
  widths: number[];
} {
  const ops: string[] = [];
  const points: Array<readonly [number, number]> = [];
  const subpaths: Array<Array<readonly [number, number]>> = [];
  const widths: number[] = [];
  const add = (x: number, y: number): void => {
    points.push([x, y]);
    if (subpaths.length === 0) {
      subpaths.push([]);
    }
    subpaths[subpaths.length - 1].push([x, y]);
  };
  const pen: GlyphPen = {
    M: (x, y) => {
      ops.push("M");
      subpaths.push([]);
      add(x, y);
    },
    L: (x, y) => {
      ops.push("L");
      add(x, y);
    },
    C: (x1, y1, x2, y2, x3, y3) => {
      ops.push("C");
      add(x1, y1);
      add(x2, y2);
      add(x3, y3);
    },
    Z: () => ops.push("Z"),
    rect: (x, y, w, h) => {
      ops.push("rect");
      subpaths.push([]);
      add(x, y);
      add(x + w, y + h);
    },
    circle: (cx, cy, r) => {
      ops.push("circle");
      subpaths.push([]);
      add(cx - r, cy - r);
      add(cx + r, cy + r);
    },
    ellipse: (cx, cy, rx, ry) => {
      ops.push("ellipse");
      subpaths.push([]);
      add(cx - rx, cy - ry);
      add(cx + rx, cy + ry);
    },
    stroke: () => ops.push("stroke"),
    fill: () => ops.push("fill"),
    fillStroke: () => ops.push("fillStroke"),
    lineWidth: w => widths.push(w)
  };
  glyph.draw(pen);
  return { ops, points, subpaths, widths };
}

/**
 * The widest vertical gap in a glyph's ink.
 *
 * Measured over each stroke's *interval*, not over its endpoints. A stem running from
 * the baseline to the cap height is two vertices 700 units apart but continuous ink
 * between them, so a gap computed from vertices reports 700 and drowns out the 47-unit
 * gap that actually mattered — which is how the first version of the test below passed
 * while the caron over `Ť` was overlapping the `T`.
 */
function widestInkGap(glyph: { draw: (p: GlyphPen) => void }): number {
  const spans = trace(glyph)
    .subpaths.filter(path => path.length > 0)
    .map(path => {
      const ys = path.map(([, y]) => y);
      return [Math.min(...ys), Math.max(...ys)] as const;
    })
    .sort((a, b) => a[0] - b[0]);
  let gap = 0;
  let reach = spans.length > 0 ? spans[0][1] : 0;
  for (const [lo, hi] of spans.slice(1)) {
    gap = Math.max(gap, lo - reach);
    reach = Math.max(reach, hi);
  }
  return gap;
}

const ALL = [...GREEK_CODE_POINTS, ...CYRILLIC_CODE_POINTS, ...LATIN_EXTENDED_CODE_POINTS];

describe("Type3 letterforms", () => {
  it("covers the scripts a Latin document reaches for", () => {
    // The exact figures are asserted so a table cannot lose a letter unnoticed.
    expect(GREEK_CODE_POINTS.length).toBe(72);
    expect(CYRILLIC_CODE_POINTS.length).toBe(96);
    expect(LATIN_EXTENDED_CODE_POINTS.length).toBe(181);
    expect(new Set(ALL).size).toBe(349);
  });

  it("writes the languages that were boxes before", () => {
    // The point of the table, stated as text rather than as code points. Each of these
    // had every accented letter drawn as `.notdef` in a browser.
    const samples: Array<[string, string]> = [
      ["Polish", "Łódź zażółć gęślą jaźń"],
      ["Czech", "Příliš žluťoučký kůň úpěl"],
      ["Slovak", "Ľuboš ťava"],
      ["Hungarian", "Árvíztűrő tükörfúrógép"],
      ["Romanian", "Șaptezeci și cinci"],
      ["Turkish", "İstanbul şğıöçü"],
      ["Croatian", "Đačka čćžš"],
      ["Lithuanian", "Ąčęėįšųūž"],
      ["Latvian", "Ģērģis ņujorkā"],
      ["Maltese", "Ċikkulata ġbejna ħobż"],
      ["Welsh", "Ŵy ŷd"],
      ["Greek", "Γειά σου Κόσμε Δ"],
      ["Bulgarian", "Съдържание на поръчката"],
      ["Chemistry", "H₂O CO₂ Fe³⁺ m² x⁴"]
    ];
    for (const [language, text] of samples) {
      const boxed = [...text].filter(
        ch =>
          !isWinAnsiCodePoint(ch.codePointAt(0)!) && lookupGlyph(ch.codePointAt(0)!) === undefined
      );
      expect(boxed, language).toEqual([]);
    }
  });

  it("adds ink for every mark, bar and dot it composes", () => {
    // The failure a rendering test would miss: a composition whose mark silently
    // contributes nothing draws the bare base letter, which is a *different letter*.
    // `ż` without its dot is `z`, and nothing about the output looks broken.
    //
    // Compared against the stroke font rather than against `lookupGlyph(base)`: the
    // base letters are WinAnsi, so a standard-14 face draws them and they are correctly
    // absent from the Type3 tables.
    const basePoints = (ch: string) =>
      STROKE_FONT[ch.codePointAt(0)!].d.reduce((sum, stroke) => sum + stroke.length, 0);
    const pairs: Array<[number, string, string]> = [
      [0x017c, "ż", "z"],
      [0x010d, "č", "c"],
      [0x0119, "ę", "e"],
      [0x0142, "ł", "l"],
      [0x0111, "đ", "d"],
      [0x0159, "ř", "r"],
      [0x0171, "ű", "u"],
      [0x016f, "ů", "u"],
      [0x0219, "ș", "s"],
      [0x0163, "ţ", "t"]
    ];
    for (const [cp, composed, base] of pairs) {
      expect(
        trace(lookupGlyph(cp)!).points.length,
        `${composed} carries no more ink than ${base}`
      ).toBeGreaterThan(basePoints(base));
    }
  });

  it("leaves a gap between a mark and the letter it sits over", () => {
    // The defect that reached a rendered page: the caron over `Ť` came within a third of
    // a stem of the `T` bar, the ink merged, and the letter read as `†`. Every mark was
    // anchored on its top edge while the marks are different heights, so no single
    // offset could clear them — the fix is a bottom anchor and one clearance, and this
    // is what holds it.
    //
    // Measured as the largest vertical gap in the glyph's ink: a letter with a mark over
    // it must have one, and it must be wider than the pen.
    const STEM = 80;
    const flatTopped = [
      0x0164, // Ť — the one that failed
      0x0100, // Ā
      0x016a, // Ū
      0x0179, // Ź
      0x017b, // Ż
      0x0106, // Ć
      0x011e, // Ğ
      0x0170, // Ű
      0x016e, // Ů
      0x0128, // Ĩ
      0x0401, // Ё
      0x0386 // Ά
    ];
    for (const cp of flatTopped) {
      expect(
        widestInkGap(lookupGlyph(cp)!),
        `U+${cp.toString(16)} mark touches its letter`
      ).toBeGreaterThan(STEM);
    }
  });

  it("sets the Czech and Slovak apostrophe letters beside the stem, not above it", () => {
    // `ď ľ ť` take a raised apostrophe to the right. Drawn as a caron above, the mark
    // floated over an ascender with a gap wider than itself and `ť` read as `Í`.
    for (const cp of [0x010f, 0x013e, 0x0165]) {
      const glyph = lookupGlyph(cp)!;
      const { points } = trace(glyph);
      const rightmost = points.reduce((best, p) => (p[0] > best[0] ? p : best), points[0]);
      // The rightmost ink is the apostrophe: high up, and inside the advance.
      expect(rightmost[1], `U+${cp.toString(16)} apostrophe sits low`).toBeGreaterThan(600);
      expect(rightmost[0], `U+${cp.toString(16)} apostrophe outside advance`).toBeLessThanOrEqual(
        glyph.width
      );
      // And nothing floats over the ascender, which is what the caron did.
      expect(Math.max(...points.map(([, y]) => y))).toBeLessThan(900);
    }
  });

  it("keeps the dotless i dotless", () => {
    // Turkish distinguishes `i` from `ı`, so taking `i` and forgetting to drop the dot
    // would silently write the wrong letter.
    expect(trace(lookupGlyph(0x131)!).ops.filter(op => op === "M").length).toBe(1);
    expect(trace(lookupGlyph(0x130)!).ops.filter(op => op === "M").length).toBeGreaterThan(1);
  });

  it("raises a superscript above the baseline and drops a subscript below it", () => {
    // `m²` and `H₂O` are the same digit at two positions; deriving both from one
    // transform is only correct if the positions actually differ.
    const band = (cp: number) => {
      const ys = trace(lookupGlyph(cp)!).points.map(([, y]) => y);
      return { lo: Math.min(...ys), hi: Math.max(...ys) };
    };
    const sup = band(0x2074); // ⁴
    const sub = band(0x2084); // ₄
    expect(sup.lo).toBeGreaterThan(0);
    expect(sub.lo).toBeLessThan(0);
    expect(sup.hi).toBeGreaterThan(sub.hi);
    // Both are smaller than a full-size digit, and the same size as each other.
    expect(sup.hi - sup.lo).toBeCloseTo(sub.hi - sub.lo, 5);
    expect(sup.hi - sup.lo).toBeLessThan(700);
  });

  it("is reachable through the shared lookup", () => {
    // The tables are merged into `lookupGlyph`; a table nobody merged draws nothing.
    for (const cp of ALL) {
      expect(lookupGlyph(cp), `U+${cp.toString(16)}`).toBeDefined();
    }
  });

  it("draws every letter, and never the .notdef box", () => {
    // The bug this fixes was `lookupGlyph(0x394) ?? NOTDEF_GLYPH` resolving to the box.
    // A letter that traced identically to that box would be the same defect wearing a
    // glyph name, so compare against it rather than merely checking for output.
    const notdef = trace(NOTDEF_GLYPH);
    for (const cp of ALL) {
      const drawn = trace(lookupGlyph(cp)!);
      expect(drawn.ops.length, `U+${cp.toString(16)} draws nothing`).toBeGreaterThan(1);
      expect(drawn.ops).toContain("stroke");
      expect(
        JSON.stringify({ ops: drawn.ops, points: drawn.points }),
        `U+${cp.toString(16)} is the .notdef box`
      ).not.toBe(JSON.stringify({ ops: notdef.ops, points: notdef.points }));
    }
  });

  it("keeps every outline inside the box its glyph declares", () => {
    // `d1` publishes this box and a viewer may clip to it, so ink outside it is a
    // letter with its tail cut off. Asserted against the declared constant rather than
    // repeated numbers: widening one without the other is the bug this guards.
    const [llx, lly, urx, ury] = GLYPH_BBOX;
    for (const cp of ALL) {
      const glyph = lookupGlyph(cp)!;
      const { points } = trace(glyph);
      expect(glyph.width, `U+${cp.toString(16)} advance`).toBeGreaterThan(0);
      expect(glyph.width).toBeLessThanOrEqual(1000);
      for (const [x, y] of points) {
        expect(x, `U+${cp.toString(16)} x`).toBeGreaterThanOrEqual(llx);
        expect(x, `U+${cp.toString(16)} x`).toBeLessThanOrEqual(urx);
        expect(y, `U+${cp.toString(16)} y`).toBeGreaterThanOrEqual(lly);
        expect(y, `U+${cp.toString(16)} y`).toBeLessThanOrEqual(ury);
      }
    }
  });

  it("declares a box that actually contains a descender and an accent", () => {
    // A box wide enough to pass the test above by being enormous would prove nothing.
    // These are the two glyphs that forced it off `0 0 1000 1000`.
    const [, lly, , ury] = GLYPH_BBOX;
    const descender = Math.min(...trace(lookupGlyph(0x3c6)!).points.map(([, y]) => y));
    const accent = Math.max(...trace(lookupGlyph(0x401)!).points.map(([, y]) => y));
    expect(descender).toBeLessThan(0);
    expect(descender).toBeGreaterThan(lly);
    expect(accent).toBeGreaterThan(700);
    expect(accent).toBeLessThan(ury);
  });

  it("gives every letter one weight, matching the rest of the fallback", () => {
    // Monoline is the whole of the source font's design. A glyph that set two widths,
    // or none, would read as a different weight from the letter beside it.
    for (const cp of ALL) {
      expect(trace(lookupGlyph(cp)!).widths, `U+${cp.toString(16)}`).toEqual([80]);
    }
  });

  it("sits a capital on the baseline at the cap height it claims", () => {
    // Type3 glyph space is Y-up from the baseline, and the conversion from stroke-font
    // space is the one place a sign error would put every letter underground.
    for (const cp of [0x394, 0x3a0, 0x41f, 0x428]) {
      const ys = trace(lookupGlyph(cp)!).points.map(([, y]) => y);
      expect(Math.min(...ys), `U+${cp.toString(16)} baseline`).toBeCloseTo(0, 5);
      expect(Math.max(...ys), `U+${cp.toString(16)} cap height`).toBeCloseTo(700, 5);
    }
  });

  it("borrows a Latin outline rather than copying it", () => {
    // 14 Greek and 11 Cyrillic capitals *are* Latin letters. Reuse is the claim; a
    // divergence here means someone re-drew one and the two will drift.
    const sameShape = (a: number, b: number) =>
      expect(JSON.stringify(trace(lookupGlyph(a)!).points)).toBe(
        JSON.stringify(trace(lookupGlyph(b)!).points)
      );
    sameShape(0x391, 0x410); // Α and А
    sameShape(0x395, 0x415); // Ε and Е
    sameShape(0x39f, 0x41e); // Ο and О
    sameShape(0x3a1, 0x420); // Ρ and Р
    // And Greek Γ is Cyrillic Г by construction, not by coincidence of authoring.
    expect(trace(GREEK[0x393]).points.length).toBe(trace(CYRILLIC[0x413]).points.length);
  });

  it("derives Cyrillic lowercase from its capital at x-height", () => {
    // The reuse that made this table small: `в к м н п т` are small capitals. So the
    // lowercase must be shorter than the capital and no taller than the x-height.
    for (const [upper, lower] of [
      [0x41f, 0x43f], // П п
      [0x428, 0x448], // Ш ш
      [0x41c, 0x43c] // М м
    ]) {
      const capTop = Math.max(...trace(lookupGlyph(upper)!).points.map(([, y]) => y));
      const smallTop = Math.max(...trace(lookupGlyph(lower)!).points.map(([, y]) => y));
      expect(smallTop, `U+${lower.toString(16)}`).toBeLessThan(capTop);
      expect(smallTop).toBeCloseTo(((0.75 - 0.38) / 0.6) * 700, 0);
    }
  });

  it("composes an accented letter from its base plus a mark", () => {
    // `Ё` is `Е` and two dots — more strokes than the base, and taller than it.
    const base = trace(lookupGlyph(0x415)!);
    const accented = trace(lookupGlyph(0x401)!);
    expect(accented.points.length).toBeGreaterThan(base.points.length);
    expect(Math.max(...accented.points.map(([, y]) => y))).toBeGreaterThan(
      Math.max(...base.points.map(([, y]) => y))
    );
  });

  it("has no glyph for the code points Unicode leaves unassigned", () => {
    // The repertoire states Greek as three ranges precisely to skip these. A glyph here
    // would mean the ranges and the tables disagree, which the repertoire test catches
    // — this says which code points the gaps are for.
    for (const cp of [0x38b, 0x38d, 0x3a2]) {
      expect(lookupGlyph(cp), `U+${cp.toString(16)}`).toBeUndefined();
      expect(isType3Drawable(cp)).toBe(false);
    }
  });

  it("leaves the scripts it cannot shape alone", () => {
    // Polytonic Greek, the Cyrillic Supplement and Devanagari are absent on purpose:
    // glyphs without shaping would render confidently wrong, which is worse than a box
    // a caller can see and fix by supplying a font.
    for (const cp of [0x1f00, 0x0460, 0x0500, 0x0905]) {
      expect(lookupGlyph(cp), `U+${cp.toString(16)}`).toBeUndefined();
    }
  });
});
