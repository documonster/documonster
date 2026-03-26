import { describe, it, expect } from "vitest";
import { Workbook } from "../../../index";
import {
  measureTextWidthPx,
  measureRichTextWidthPx,
  getMaxDigitWidth,
  getPixelPadding,
  pixelToCharWidth,
  charWidthToPixel,
  getLineHeightPx,
  calculateAutoFitWidth,
  calculateAutoFitHeight,
  calculateWrappedLineCount,
  pixelToPoints
} from "@excel/utils/text-metrics";
import { getCalibri11PtPixelWidth } from "@excel/utils/font-data";

// =============================================================================
// Font Data Tests
// =============================================================================

describe("Font Data - Calibri 11pt Bitmap", () => {
  it("matches rust_xlsxwriter/ClosedXML pixel widths for all A-Z", () => {
    // These values are verified against both ClosedXML (SkiaSharp) and
    // rust_xlsxwriter (empirically measured from Excel)
    const expected: Record<string, number> = {
      A: 9,
      B: 8,
      C: 8,
      D: 9,
      E: 7,
      F: 7,
      G: 9,
      H: 9,
      I: 4,
      J: 5,
      K: 8,
      L: 6,
      M: 12,
      N: 10,
      O: 10,
      P: 8,
      Q: 10,
      R: 8,
      S: 7,
      T: 7,
      U: 9,
      V: 9,
      W: 13,
      X: 8,
      Y: 7,
      Z: 7
    };

    for (const [char, px] of Object.entries(expected)) {
      expect(getCalibri11PtPixelWidth(char.codePointAt(0)!)).toBe(px);
    }
  });

  it("matches rust_xlsxwriter/ClosedXML pixel widths for all a-z", () => {
    const expected: Record<string, number> = {
      a: 7,
      b: 8,
      c: 6,
      d: 8,
      e: 8,
      f: 5,
      g: 7,
      h: 8,
      i: 4,
      j: 4,
      k: 7,
      l: 4,
      m: 12,
      n: 8,
      o: 8,
      p: 8,
      q: 8,
      r: 5,
      s: 6,
      t: 5,
      u: 8,
      v: 7,
      w: 11,
      x: 7,
      y: 7,
      z: 6
    };

    for (const [char, px] of Object.entries(expected)) {
      expect(getCalibri11PtPixelWidth(char.codePointAt(0)!)).toBe(px);
    }
  });

  it("matches pixel widths for digits 0-9", () => {
    // All Calibri digits have the same width
    for (let d = 0; d <= 9; d++) {
      expect(getCalibri11PtPixelWidth(0x30 + d)).toBe(7);
    }
  });

  it("matches pixel widths for common punctuation", () => {
    expect(getCalibri11PtPixelWidth(0x20)).toBe(3); // space
    expect(getCalibri11PtPixelWidth(0x2e)).toBe(4); // .
    expect(getCalibri11PtPixelWidth(0x2c)).toBe(4); // ,
    expect(getCalibri11PtPixelWidth(0x2d)).toBe(5); // -
    expect(getCalibri11PtPixelWidth(0x3a)).toBe(4); // :
    expect(getCalibri11PtPixelWidth(0x3b)).toBe(4); // ;
  });
});

// =============================================================================
// Unit Conversion Tests
// =============================================================================

describe("Unit Conversions", () => {
  it("calculates MDW correctly for Calibri 11pt", () => {
    // Calibri 11pt uses bitmap MDW = 7 (not outline which would give 8)
    expect(getMaxDigitWidth()).toBe(7);
    expect(getMaxDigitWidth({ name: "Calibri", size: 11 })).toBe(7);
  });

  it("calculates pixel padding correctly", () => {
    // PP = 2 * CEIL(MDW/4) + 1
    expect(getPixelPadding(7)).toBe(5); // 2 * CEIL(7/4) + 1 = 2*2+1 = 5
    expect(getPixelPadding(8)).toBe(5); // 2 * CEIL(8/4) + 1 = 2*2+1 = 5
    expect(getPixelPadding(6)).toBe(5); // 2 * CEIL(6/4) + 1 = 2*2+1 = 5
    expect(getPixelPadding(11)).toBe(7); // 2 * CEIL(11/4) + 1 = 2*3+1 = 7
  });

  it("converts pixel to char width correctly (TRUNC formula)", () => {
    // TRUNC(pixels / MDW * 256) / 256
    const mdw = 7;
    expect(pixelToCharWidth(0, mdw)).toBe(0);
    expect(pixelToCharWidth(7, mdw)).toBe(1); // 7/7*256 = 256, TRUNC/256 = 1
    expect(pixelToCharWidth(14, mdw)).toBe(2);
    expect(pixelToCharWidth(64, mdw)).toBe(Math.trunc((64 / 7) * 256) / 256);
  });

  it("converts char width to pixel correctly", () => {
    const mdw = 7;
    expect(charWidthToPixel(0, mdw)).toBe(0);
    // width >= 1: ROUND(width * MDW) + PP
    expect(charWidthToPixel(1, mdw)).toBe(Math.round(1 * 7) + 5); // 12
    expect(charWidthToPixel(9, mdw)).toBe(Math.round(9 * 7) + 5); // 68
    // width < 1: ROUND(width * (MDW + PP))
    expect(charWidthToPixel(0.5, mdw)).toBe(Math.round(0.5 * 12)); // 6
  });

  it("converts points to pixels and back", () => {
    expect(pixelToPoints(96)).toBe(72); // 96px at 96dpi = 1 inch = 72pt
    expect(pixelToPoints(15)).toBeCloseTo(11.25, 2); // Calibri 11pt line height
  });
});

// =============================================================================
// Text Width Measurement Tests
// =============================================================================

describe("Text Width Measurement", () => {
  it("measures empty string as 0", () => {
    expect(measureTextWidthPx("")).toBe(0);
    expect(measureTextWidthPx("", { name: "Calibri", size: 11 })).toBe(0);
  });

  it("measures ASCII text in Calibri 11pt using bitmap table", () => {
    // "Hello" = H(9) + e(8) + l(4) + l(4) + o(8) = 33
    expect(measureTextWidthPx("Hello")).toBe(33);
    expect(measureTextWidthPx("Hello", { name: "Calibri", size: 11 })).toBe(33);
  });

  it("measures a single character correctly", () => {
    expect(measureTextWidthPx("A")).toBe(9);
    expect(measureTextWidthPx("i")).toBe(4);
    expect(measureTextWidthPx("M")).toBe(12);
    expect(measureTextWidthPx("W")).toBe(13);
  });

  it("measures string with narrow chars accurately", () => {
    // "illili" = 4+4+4+4+4+4 = 24
    expect(measureTextWidthPx("illili")).toBe(24);
  });

  it("measures string with wide chars accurately", () => {
    // "MWmw" = 12+13+12+11 = 48
    expect(measureTextWidthPx("MWmw")).toBe(48);
  });

  it("measures multi-line text (returns widest line)", () => {
    // "AB\nW" → line 1: A(9)+B(8)=17, line 2: W(13)=13 → max = 17
    expect(measureTextWidthPx("AB\nW")).toBe(17);
  });

  it("handles CRLF line breaks", () => {
    expect(measureTextWidthPx("A\r\nB")).toBe(Math.max(9, 8));
  });

  it("applies bold multiplier for unknown bold variant", () => {
    // Calibri doesn't have a separate bold metrics table in our data,
    // so it should apply the 1.05 multiplier per character
    const normalWidth = measureTextWidthPx("Hello", { name: "Calibri", size: 11 });
    const boldWidth = measureTextWidthPx("Hello", { name: "Calibri", size: 11, bold: true });
    expect(boldWidth).toBeGreaterThan(normalWidth);
  });

  it("measures Arial text using FUnit data", () => {
    const width = measureTextWidthPx("Hello", { name: "Arial", size: 11 });
    expect(width).toBeGreaterThan(0);
    // Arial is generally wider than Calibri
    const calibriWidth = measureTextWidthPx("Hello", { name: "Calibri", size: 11 });
    expect(width).toBeGreaterThanOrEqual(calibriWidth);
  });

  it("measures text at different font sizes", () => {
    const width11 = measureTextWidthPx("Test", { name: "Arial", size: 11 });
    const width22 = measureTextWidthPx("Test", { name: "Arial", size: 22 });
    // 22pt should be roughly 2x of 11pt (not exactly due to rounding)
    expect(width22).toBeGreaterThan(width11 * 1.5);
    expect(width22).toBeLessThan(width11 * 2.5);
  });

  it("uses factor-based fallback for unknown fonts", () => {
    const width = measureTextWidthPx("Hello World", { name: "UnknownFont", size: 11 });
    expect(width).toBeGreaterThan(0);
  });
});

// =============================================================================
// Rich Text Measurement Tests
// =============================================================================

describe("Rich Text Measurement", () => {
  it("measures simple rich text", () => {
    const richText = [{ text: "Hello" }, { text: " World" }];
    const width = measureRichTextWidthPx(richText);
    const plainWidth = measureTextWidthPx("Hello World");
    // Should be the same as plain text when no font overrides
    expect(width).toBe(plainWidth);
  });

  it("handles rich text with newlines", () => {
    const richText = [{ text: "Line1\nLine2" }];
    const width = measureRichTextWidthPx(richText);
    const line1Width = measureTextWidthPx("Line1");
    const line2Width = measureTextWidthPx("Line2");
    expect(width).toBe(Math.max(line1Width, line2Width));
  });
});

// =============================================================================
// Auto-Fit Width Calculation Tests
// =============================================================================

describe("Auto-Fit Width Calculation", () => {
  it("calculates auto-fit width with padding", () => {
    const textPx = 50;
    const mdw = 7;
    const charWidth = calculateAutoFitWidth(textPx, mdw);

    // Should be > 0 and include padding
    expect(charWidth).toBeGreaterThan(0);

    // Should be greater than raw text width / MDW
    expect(charWidth).toBeGreaterThan(textPx / mdw);
  });

  it("returns 0 for empty content", () => {
    expect(calculateAutoFitWidth(0, 7)).toBe(0);
    expect(calculateAutoFitWidth(-1, 7)).toBe(0);
  });

  it("adds extra space for autofilter", () => {
    const withoutFilter = calculateAutoFitWidth(50, 7, false);
    const withFilter = calculateAutoFitWidth(50, 7, true);
    expect(withFilter).toBeGreaterThan(withoutFilter);
  });

  it("caps at maximum column width", () => {
    const charWidth = calculateAutoFitWidth(10000, 7);
    expect(charWidth).toBeLessThanOrEqual(255);
  });
});

// =============================================================================
// Auto-Fit Height Calculation Tests
// =============================================================================

describe("Auto-Fit Height Calculation", () => {
  it("calculates single-line height", () => {
    const height = calculateAutoFitHeight("Hello");
    expect(height).toBeGreaterThan(0);
    expect(height).toBeLessThan(30); // Calibri 11pt is about 15pt height
  });

  it("calculates multi-line height", () => {
    const singleLine = calculateAutoFitHeight("Hello");
    const twoLines = calculateAutoFitHeight("Hello\nWorld");
    expect(twoLines).toBeCloseTo(singleLine * 2, 0);
  });

  it("calculates wrapped text height", () => {
    const longText = "This is a very long text that should wrap to multiple lines";
    const narrowCol = 50; // 50px column width
    const height = calculateAutoFitHeight(longText, undefined, { wrapText: true }, narrowCol);
    const singleLine = calculateAutoFitHeight("X");
    expect(height).toBeGreaterThan(singleLine);
  });

  it("returns line height for empty text", () => {
    const height = calculateAutoFitHeight("");
    expect(height).toBeGreaterThan(0); // Returns single line height
  });

  it("line height is larger for larger font sizes", () => {
    const small = getLineHeightPx({ size: 11 });
    const large = getLineHeightPx({ size: 22 });
    expect(large).toBeGreaterThan(small);
  });
});

// =============================================================================
// Wrapped Line Count Tests
// =============================================================================

describe("Wrapped Line Count", () => {
  it("returns 1 for short text", () => {
    expect(calculateWrappedLineCount("Hi", 100)).toBe(1);
  });

  it("counts explicit newlines", () => {
    expect(calculateWrappedLineCount("A\nB\nC", 100)).toBe(3);
  });

  it("wraps at word boundaries like Excel", () => {
    // "Hello World" in Calibri 11pt:
    // "Hello " = H(9)+e(8)+l(4)+l(4)+o(8)+space(3) = 36px
    // "World"  = W(13)+o(8)+r(5)+l(4)+d(8) = 38px
    // In a 40px column: "Hello " fits (36px), "World" starts new line (38px)
    expect(calculateWrappedLineCount("Hello World", 40)).toBe(2);
  });

  it("does not break mid-word", () => {
    // "Spreadsheet" = 77px in Calibri 11pt
    // In a 50px column: the word doesn't fit, but Excel doesn't break mid-word
    // The word overflows on one line
    expect(calculateWrappedLineCount("Spreadsheet", 50)).toBe(1);
  });

  it("wraps multiple words correctly", () => {
    // "a b c d" with very narrow column — each word should be on its own line
    // a=7, space=3, b=8, space=3, c=6, space=3, d=8
    // "a "=10, "b "=11, "c "=9, "d"=8
    // In a 12px column: each fits individually, so 4 words = at most 4 lines
    const count = calculateWrappedLineCount("a b c d", 12);
    expect(count).toBe(4);
  });

  it("handles hyphen as break point", () => {
    // "one-two" should break after the hyphen
    const count = calculateWrappedLineCount("one-two", 30);
    // "one-" = o(8)+n(8)+e(8)+-(5) = 29px
    // "two" = t(5)+w(11)+o(8) = 24px
    // In 30px column: "one-" fits (29px), "two" on next line
    expect(count).toBe(2);
  });
});

// =============================================================================
// Worksheet Integration Tests
// =============================================================================

describe("Worksheet.autoFitColumns", () => {
  it("auto-fits a single column", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Short";
    ws.getCell("A2").value = "A much longer string value";

    ws.autoFitColumn("A");

    const col = ws.getColumn("A");
    expect(col.width).toBeGreaterThan(9); // > default width
    expect(col.bestFit).toBe(true);
  });

  it("auto-fits all columns", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Name";
    ws.getCell("B1").value = "A very long description field";
    ws.getCell("C1").value = 12345;

    ws.autoFitColumns();

    expect(ws.getColumn("A").bestFit).toBe(true);
    expect(ws.getColumn("B").bestFit).toBe(true);
    expect(ws.getColumn("C").bestFit).toBe(true);

    // Column B should be wider than A (longer text)
    expect(ws.getColumn("B").width).toBeGreaterThan(ws.getColumn("A").width!);
  });

  it("auto-fits a column range", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Col A";
    ws.getCell("B1").value = "Col B with more text";
    ws.getCell("C1").value = "Col C";

    ws.autoFitColumns("B", "C");

    // A should not be affected
    expect(ws.getColumn("A").bestFit).toBeUndefined();
    expect(ws.getColumn("B").bestFit).toBe(true);
    expect(ws.getColumn("C").bestFit).toBe(true);
  });

  it("handles empty worksheet gracefully", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    // Should not throw
    ws.autoFitColumns();
  });

  it("skips merged cells spanning multiple columns", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Short";
    ws.getCell("A2").value = "This is merged across A2:C2";
    ws.mergeCells("A2:C2");

    ws.autoFitColumn("A");

    // Width should be based on A1 "Short", not the merged cell
    const col = ws.getColumn("A");
    expect(col.bestFit).toBe(true);
  });

  it("returns this for chaining", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");
    ws.getCell("A1").value = "test";

    const result = ws.autoFitColumn("A");
    expect(result).toBe(ws);

    const result2 = ws.autoFitColumns();
    expect(result2).toBe(ws);
  });
});

describe("Worksheet.autoFitRows", () => {
  it("auto-fits a single row", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Hello";
    ws.autoFitRow(1);

    const row = ws.getRow(1);
    expect(row.height).toBeGreaterThan(0);
    expect(row.customHeight).toBe(true);
  });

  it("auto-fits all rows", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Row 1";
    ws.getCell("A2").value = "Row 2";

    ws.autoFitRows();

    expect(ws.getRow(1).customHeight).toBe(true);
    expect(ws.getRow(2).customHeight).toBe(true);
  });

  it("calculates taller height for multi-line text", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Single line";
    ws.getCell("A2").value = "Line 1\nLine 2\nLine 3";

    ws.autoFitRows();

    // Row 2 should be taller than row 1
    expect(ws.getRow(2).height).toBeGreaterThan(ws.getRow(1).height!);
  });

  it("handles empty rows gracefully", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A3").value = "Only row 3";

    // Should not throw for empty rows
    ws.autoFitRows();
  });

  it("returns this for chaining", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");
    ws.getCell("A1").value = "test";

    const result = ws.autoFitRow(1);
    expect(result).toBe(ws);

    const result2 = ws.autoFitRows();
    expect(result2).toBe(ws);
  });
});

// =============================================================================
// Integration: Combined auto-fit
// =============================================================================

describe("Auto-Fit Integration", () => {
  it("auto-fits columns and rows together", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Name";
    ws.getCell("B1").value = "Description";
    ws.getCell("A2").value = "Alice";
    ws.getCell("B2").value = "A very long description that should make column B wider";

    ws.autoFitColumns().autoFitRows();

    // Verify columns are fitted
    expect(ws.getColumn("A").bestFit).toBe(true);
    expect(ws.getColumn("B").bestFit).toBe(true);
    expect(ws.getColumn("B").width).toBeGreaterThan(ws.getColumn("A").width!);

    // Verify rows are fitted
    expect(ws.getRow(1).customHeight).toBe(true);
    expect(ws.getRow(2).customHeight).toBe(true);
  });

  it("handles number-formatted cells", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = 1234567.89;
    ws.getCell("A1").numFmt = "#,##0.00";

    ws.autoFitColumn("A");

    // Should fit to the formatted string "1,234,567.89" width, not "1234567.89"
    expect(ws.getColumn("A").bestFit).toBe(true);
    expect(ws.getColumn("A").width).toBeGreaterThan(9);
  });

  it("handles date cells", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = new Date(2024, 0, 15);
    ws.autoFitColumn("A");

    expect(ws.getColumn("A").bestFit).toBe(true);
    expect(ws.getColumn("A").width).toBeGreaterThan(0);
  });

  it("handles boolean cells", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = true;
    ws.getCell("A2").value = false;
    ws.autoFitColumn("A");

    expect(ws.getColumn("A").bestFit).toBe(true);
  });

  it("handles rich text cells", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = {
      richText: [{ text: "Bold part", font: { bold: true } }, { text: " normal part" }]
    };

    ws.autoFitColumn("A");
    expect(ws.getColumn("A").bestFit).toBe(true);
    expect(ws.getColumn("A").width).toBeGreaterThan(9);
  });

  it("handles cells with custom fonts", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Arial Text";
    ws.getCell("A1").font = { name: "Arial", size: 14 };

    ws.getCell("B1").value = "Arial Text";
    ws.getCell("B1").font = { name: "Calibri", size: 11 };

    ws.autoFitColumns();

    // Arial 14pt should be wider than Calibri 11pt
    expect(ws.getColumn("A").width).toBeGreaterThan(ws.getColumn("B").width!);
  });

  it("skips hidden rows when auto-fitting columns", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Short";
    ws.getCell("A2").value = "This is a very long hidden row value";
    ws.getRow(2).hidden = true;
    ws.getCell("A3").value = "Medium text";

    ws.autoFitColumn("A");

    // Width should be based on visible rows only (A1 "Short" or A3 "Medium text")
    const col = ws.getColumn("A");
    expect(col.bestFit).toBe(true);

    // Verify width is NOT based on the hidden row's long text
    const wb2 = new Workbook();
    const ws2 = wb2.addWorksheet("test2");
    ws2.getCell("A1").value = "Short";
    ws2.getCell("A2").value = "This is a very long hidden row value";
    ws2.getCell("A3").value = "Medium text";
    ws2.autoFitColumn("A");
    // Width with hidden row should be less than without hiding
    expect(col.width).toBeLessThan(ws2.getColumn("A").width!);
  });

  it("skips hidden columns when auto-fitting rows", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Normal";
    ws.getCell("B1").value = "Hidden\ncol\nwith\nmany\nlines";
    ws.getColumn("B").hidden = true;

    ws.autoFitRow(1);

    const row = ws.getRow(1);
    expect(row.customHeight).toBe(true);
    // Height should be based on A1 (single line), not B1 (multi-line hidden)
    const singleLineHeight = row.height!;

    const wb2 = new Workbook();
    const ws2 = wb2.addWorksheet("test2");
    ws2.getCell("A1").value = "Normal";
    ws2.getCell("B1").value = "Hidden\ncol\nwith\nmany\nlines";
    ws2.autoFitRow(1);
    // Without hiding, row should be taller
    expect(singleLineHeight).toBeLessThan(ws2.getRow(1).height!);
  });

  it("accounts for indent in column width", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "No indent";
    ws.getCell("B1").value = "No indent";
    ws.getCell("B1").alignment = { indent: 3 };

    ws.autoFitColumns();

    // Column B with indent should be wider than A with same text
    expect(ws.getColumn("B").width).toBeGreaterThan(ws.getColumn("A").width!);
  });

  it("skips shrinkToFit cells", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    ws.getCell("A1").value = "Short";
    ws.getCell("A2").value = "This very long text has shrinkToFit enabled";
    ws.getCell("A2").alignment = { shrinkToFit: true };

    ws.autoFitColumn("A");

    // Width should be based on A1 "Short" only, not the shrinkToFit cell
    const col = ws.getColumn("A");
    expect(col.bestFit).toBe(true);

    const wb2 = new Workbook();
    const ws2 = wb2.addWorksheet("test2");
    ws2.getCell("A1").value = "Short";
    ws2.getCell("A2").value = "This very long text has shrinkToFit enabled";
    ws2.autoFitColumn("A");
    // Without shrinkToFit, column should be wider
    expect(col.width).toBeLessThan(ws2.getColumn("A").width!);
  });

  it("uses Arial bold metrics directly (not multiplier)", () => {
    // Arial has a dedicated bold metrics table
    const normalWidth = measureTextWidthPx("Test", { name: "Arial", size: 11 });
    const boldWidth = measureTextWidthPx("Test", { name: "Arial", size: 11, bold: true });
    // Bold should be wider (using real bold advance widths, not 1.05x)
    expect(boldWidth).toBeGreaterThan(normalWidth);
  });
});

// =============================================================================
// Edge Cases
// =============================================================================

describe("Edge Cases", () => {
  it("handles mdw=0 gracefully", () => {
    expect(pixelToCharWidth(100, 0)).toBe(0);
    expect(charWidthToPixel(10, 0)).toBe(0);
  });

  it("handles negative mdw gracefully", () => {
    expect(pixelToCharWidth(100, -1)).toBe(0);
    expect(charWidthToPixel(10, -1)).toBe(0);
  });
});
