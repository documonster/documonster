import { getFontMetrics, getLayoutFontMetrics } from "@utils/font-data";
import { faceLineHeight, measureFaceText } from "@word/layout/font-face";
/**
 * Every family with metric data is measured by layout with the face's own
 * tables. Georgia and Verdana had Excel metrics but no layout entry, so a
 * Word document set in them was measured as Times New Roman with a 1.2 em
 * line, and its headings landed on the wrong page.
 */
import { describe, expect, it } from "vitest";

const FAMILIES = [
  "Calibri",
  "Arial",
  "Times New Roman",
  "Courier New",
  "Verdana",
  "Georgia",
  "Tahoma",
  "Trebuchet MS"
];

describe("layout font registry", () => {
  it("lays out every family that has metric data", () => {
    for (const family of FAMILIES) {
      expect(getFontMetrics(family), family).toBeDefined();
      expect(getLayoutFontMetrics(family), family).toBeDefined();
    }
  });

  it("spaces Georgia and Verdana lines by their hhea metrics", () => {
    // Georgia: hhea 1878 + 449 + 0 over 2048; Verdana: 2059 + 430 + 0.
    expect(faceLineHeight("Georgia", 11)).toBeCloseTo((2327 / 2048) * 11, 6);
    expect(faceLineHeight("Verdana", 9)).toBeCloseTo((2489 / 2048) * 9, 6);
  });

  it("measures with each face's own advances, bold and italic included", () => {
    // Advances from Georgia.ttf / Verdana.ttf as Word for Mac 16.114 uses them.
    expect(measureFaceText("!", "Verdana", 2048)).toBe(806);
    expect(measureFaceText("1", "Georgia", 2048)).toBe(880);
    expect(measureFaceText("a", "Georgia", 2048, true)).not.toBe(
      measureFaceText("a", "Georgia", 2048)
    );
    expect(measureFaceText("a", "Georgia", 2048, false, true)).not.toBe(
      measureFaceText("a", "Georgia", 2048)
    );
  });

  it("measures italic text in Calibri, Cambria, Arial and Times with the italic faces", () => {
    // Word lays out Cambria Italic narrower than upright: a 60-word italic
    // paragraph took 4 lines in Word and 5 measured upright (N5).
    for (const family of ["Calibri", "Cambria", "Arial", "Times New Roman"]) {
      expect(getLayoutFontMetrics(family, false, true), family).not.toBe(
        getLayoutFontMetrics(family)
      );
      expect(getLayoutFontMetrics(family, true, true), family).not.toBe(
        getLayoutFontMetrics(family, true)
      );
    }
    // Cambriai.ttf as shipped with Word for Mac 16.114.
    expect(measureFaceText("a", "Cambria", 2048, false, true)).toBe(
      getLayoutFontMetrics("Cambria", false, true)!.advances.find(
        ([s, n]) => s <= 0x61 && s + n > 0x61
      )![2]
    );
  });
});
