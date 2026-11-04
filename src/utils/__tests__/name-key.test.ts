/**
 * Excel's name comparisons, pair by pair as Excel 16 decided them: by renumbering a typed table header, refusing a
 * table or sheet rename, merging two defined names, and resolving a `LET` parameter.
 */
import { excelNameKey, sheetNameKey } from "@utils/name-key";
import { describe, expect, it } from "vitest";

// [a, b, same as names (tables, columns, defined names), same as sheets (sheets, LET parameters)]
const PAIRS: readonly (readonly [string, string, boolean, boolean])[] = [
  ["ß", "SS", true, false],
  ["ẞ", "ß", true, true],
  ["ẞ", "ss", true, false],
  ["ﬁ", "FI", true, false],
  ["ı", "I", false, false],
  ["ı", "i", false, false],
  ["i", "İ", false, false],
  ["ς", "σ", true, true],
  ["ς", "Σ", true, true],
  ["ａ", "A", true, true],
  ["ア", "あ", true, true],
  ["ｶ", "カ", true, true],
  ["é", "e", false, false],
  ["x²", "x2", true, true],
  ["µ", "Μ", true, true],
  ["Ǳ", "ǳ", true, true],
  ["\u212A", "k", true, true],
  ["Ⅰ", "ⅰ", true, true],
  ["Ⅰ", "I", true, true],
  ["ab", "AB", true, true],
  ["ゝ", "ヽ", false, false],
  ["ᾳ", "ΑΙ", false, false],
  ["e\u0301", "é", true, true]
];

describe("Excel's name comparisons", () => {
  it.each(PAIRS)("%s and %s", (a, b, asName, asSheet) => {
    expect(excelNameKey(a) === excelNameKey(b)).toBe(asName);
    expect(sheetNameKey(a) === sheetNameKey(b)).toBe(asSheet);
  });
});
