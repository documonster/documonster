/**
 * Formatting characters take no width, and tab is the exception that does.
 *
 * Grapheme clustering already folds a variation selector or a joiner into the
 * character it modifies, so these only surface when one arrives on its own — and then
 * the measurer charged it the face's *default advance*, half an em of width for
 * something that occupies none. A zero-width space between two letters made the line
 * wider than it is, which the rasteriser then normalised its glyph advances to, moving
 * every visible glyph in the label.
 *
 * Tab is deliberately not in that set: it draws nothing but it does occupy space.
 */

import { isNonPrintingControl, isZeroWidthCodePoint } from "@utils/font-metrics";
import { measureTextWidthPx } from "@utils/text-measure";
import { describe, expect, it } from "vitest";

/** Arial has a real advance table, so this exercises the per-glyph path. */
const ARIAL = { name: "Arial", size: 14 };
/** A face with no table, which counts characters by category instead. */
const UNTABULATED = { name: "Some Unknown Face", size: 14 };

describe("measuring formatting characters", () => {
  const controls: [string, string][] = [
    ["\u200B", "zero width space"],
    ["\u200C", "zero width non-joiner"],
    ["\u200D", "zero width joiner"],
    ["\u2060", "word joiner"],
    ["\uFEFF", "byte order mark"],
    ["\u200E", "left-to-right mark"],
    ["\u202A", "bidi embedding"],
    ["\u2066", "bidi isolate"],
    ["\u0301", "combining acute"],
    ["\u0000", "NUL"],
    ["\u001F", "unit separator"]
  ];

  it("adds nothing to a line, on the tabulated path", () => {
    const baseline = measureTextWidthPx("AB", ARIAL);
    for (const [control, name] of controls) {
      expect(measureTextWidthPx(`A${control}B`, ARIAL), name).toBe(baseline);
    }
  });

  it("adds nothing to a line, on the category-average path", () => {
    // The two paths counted width differently, so both need the rule.
    const baseline = measureTextWidthPx("AB", UNTABULATED);
    for (const [control, name] of controls) {
      expect(measureTextWidthPx(`A${control}B`, UNTABULATED), name).toBe(baseline);
    }
  });

  it("measures one on its own as nothing at all", () => {
    for (const [control, name] of controls) {
      expect(measureTextWidthPx(control, ARIAL), name).toBe(0);
      expect(measureTextWidthPx(control, UNTABULATED), name).toBe(0);
    }
  });

  it("treats a line feed as a line break, not a zero-width character", () => {
    // `measureTextWidthPx` reports the widest *line*, so a line feed never reaches the
    // per-character path at all. It is in `isZeroWidthCodePoint` for the rasteriser's
    // benefit — drawing a single line, a stray newline should put no ink down.
    expect(measureTextWidthPx("A\nBB", ARIAL)).toBe(measureTextWidthPx("BB", ARIAL));
    expect(isZeroWidthCodePoint(0x0a)).toBe(true);
  });

  it("still treats a variation selector on a character as emoji presentation", () => {
    // U+FE0F asks for the emoji form, which is one em wide. That is a statement about
    // the *base* character and must survive the rule above — a lone U+FE0F is nothing,
    // but `A` + U+FE0F is not `A`.
    const plain = measureTextWidthPx("A", ARIAL);
    const emoji = measureTextWidthPx("A\uFE0F", ARIAL);
    expect(emoji).toBeGreaterThan(plain);
    expect(measureTextWidthPx("\uFE0F", ARIAL)).toBe(0);
  });

  it("keeps a ZWJ emoji sequence one glyph wide", () => {
    // Three joiners inside one cluster. Counting them would have made a four-person
    // family five and a half characters wide.
    const family = measureTextWidthPx("\u{1F469}\u200D\u{1F469}\u200D\u{1F467}", ARIAL);
    const single = measureTextWidthPx("\u{1F469}", ARIAL);
    expect(family).toBe(single);
  });

  it("gives tab width, because it draws nothing but occupies space", () => {
    const baseline = measureTextWidthPx("AB", ARIAL);
    expect(measureTextWidthPx("A\tB", ARIAL)).toBeGreaterThan(baseline);
    expect(isZeroWidthCodePoint(0x09)).toBe(false);
    expect(isNonPrintingControl(0x09)).toBe(false);
  });

  it("classifies combining marks as width-less but printing", () => {
    // The two predicates differ by exactly this set: a combining acute takes no width,
    // but a face that has a glyph for it draws one, so a rasteriser must not skip it.
    expect(isZeroWidthCodePoint(0x0301)).toBe(true);
    expect(isNonPrintingControl(0x0301)).toBe(false);
  });
});
