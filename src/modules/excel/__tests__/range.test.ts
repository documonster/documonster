import {
  type RangeData,
  rangeAbsolute,
  rangeAbsoluteBottomRight,
  rangeAbsoluteShort,
  rangeAbsoluteTopLeft,
  rangeBottom,
  rangeBr,
  rangeContains,
  rangeCreate,
  rangeExpand,
  rangeIntersects,
  rangeLeft,
  rangeRange,
  rangeRight,
  rangeShortRange,
  rangeTl,
  rangeToString,
  rangeTop
} from "@excel/range";
import { describe, it, expect } from "vitest";

describe("Range", () => {
  function check(
    d: RangeData,
    range: string,
    $range: string,
    tl: string,
    $t$l: string,
    br: string,
    $b$r: string,
    top: number,
    left: number,
    bottom: number,
    right: number,
    sheetName?: string
  ) {
    expect(rangeRange(d)).toBe(range);
    expect(rangeAbsolute(d)).toBe($range);
    expect(rangeTl(d)).toBe(tl);
    expect(rangeAbsoluteTopLeft(d)).toBe($t$l);
    expect(rangeBr(d)).toBe(br);
    expect(rangeAbsoluteBottomRight(d)).toBe($b$r);
    expect(rangeTop(d)).toBe(top);
    expect(rangeLeft(d)).toBe(left);
    expect(rangeBottom(d)).toBe(bottom);
    expect(rangeRight(d)).toBe(right);
    expect(rangeToString(d)).toBe(range);
    expect(d.sheetName).toBe(sheetName);
  }

  it("has a valid default value", () => {
    const d = rangeCreate();
    check(d, "A1:A1", "$A$1:$A$1", "A1", "$A$1", "A1", "$A$1", 1, 1, 1, 1);
  });

  it("constructs as expected", () => {
    check(rangeCreate("B5:D10"), "B5:D10", "$B$5:$D$10", "B5", "$B$5", "D10", "$D$10", 5, 2, 10, 4);
    check(rangeCreate("B10:D5"), "B5:D10", "$B$5:$D$10", "B5", "$B$5", "D10", "$D$10", 5, 2, 10, 4);
    check(rangeCreate("D5:B10"), "B5:D10", "$B$5:$D$10", "B5", "$B$5", "D10", "$D$10", 5, 2, 10, 4);
    check(rangeCreate("D10:B5"), "B5:D10", "$B$5:$D$10", "B5", "$B$5", "D10", "$D$10", 5, 2, 10, 4);

    const expectCG = (d: RangeData): void =>
      check(d, "C7:G16", "$C$7:$G$16", "C7", "$C$7", "G16", "$G$16", 7, 3, 16, 7);

    expectCG(rangeCreate("G7", "C16"));
    expectCG(rangeCreate("C7", "G16"));
    expectCG(rangeCreate("C16", "G7"));
    expectCG(rangeCreate("G16", "C7"));

    expectCG(rangeCreate(7, 3, 16, 7));
    expectCG(rangeCreate(16, 3, 7, 7));
    expectCG(rangeCreate(7, 7, 16, 3));
    expectCG(rangeCreate(16, 7, 7, 3));

    expectCG(rangeCreate([7, 3, 16, 7]));
    expectCG(rangeCreate([16, 3, 7, 7]));
    expectCG(rangeCreate([7, 7, 16, 3]));
    expectCG(rangeCreate([16, 7, 7, 3]));

    check(
      rangeCreate("$B$5:$D$10"),
      "B5:D10",
      "$B$5:$D$10",
      "B5",
      "$B$5",
      "D10",
      "$D$10",
      5,
      2,
      10,
      4
    );
    check(
      rangeCreate("blort!$B$5:$D$10"),
      "blort!B5:D10",
      "blort!$B$5:$D$10",
      "B5",
      "$B$5",
      "D10",
      "$D$10",
      5,
      2,
      10,
      4,
      "blort"
    );
  });

  it("expands properly", () => {
    const d = rangeCreate();

    rangeExpand(d, 1, 1, 1, 3);
    expect(rangeTl(d)).toBe("A1");
    expect(rangeBr(d)).toBe("C1");
    expect(rangeToString(d)).toBe("A1:C1");

    rangeExpand(d, 1, 3, 3, 3);
    expect(rangeTl(d)).toBe("A1");
    expect(rangeBr(d)).toBe("C3");
    expect(rangeToString(d)).toBe("A1:C3");
  });

  it("doesn't always include the default row/col", () => {
    const d = rangeCreate();

    rangeExpand(d, 2, 2, 4, 4);
    expect(rangeTl(d)).toBe("B2");
    expect(rangeBr(d)).toBe("D4");
    expect(rangeToString(d)).toBe("B2:D4");
  });

  it("detects intersections", () => {
    const C3F6 = rangeCreate("C3:F6");
    const x = (s: string): boolean => rangeIntersects(C3F6, rangeCreate(s));

    // touching at corners
    expect(x("A1:B2")).toBe(false);
    expect(x("G1:H2")).toBe(false);
    expect(x("A7:B8")).toBe(false);
    expect(x("G7:H8")).toBe(false);

    // Adjacent to edges
    expect(x("A1:H2")).toBe(false);
    expect(x("A1:B8")).toBe(false);
    expect(x("G1:H8")).toBe(false);
    expect(x("A7:H8")).toBe(false);

    // 1 cell margin
    expect(x("A1:H1")).toBe(false);
    expect(x("A1:A8")).toBe(false);
    expect(x("G1:G8")).toBe(false);
    expect(x("A8:G8")).toBe(false);

    // Adjacent at corners
    expect(x("A1:B3")).toBe(false);
    expect(x("A1:C2")).toBe(false);
    expect(x("F1:H2")).toBe(false);
    expect(x("G1:H3")).toBe(false);
    expect(x("A6:B8")).toBe(false);
    expect(x("A7:C8")).toBe(false);
    expect(x("F7:H8")).toBe(false);
    expect(x("G6:H8")).toBe(false);

    // Adjacent at edges
    expect(x("A4:B5")).toBe(false);
    expect(x("D1:E2")).toBe(false);
    expect(x("D7:E8")).toBe(false);
    expect(x("G4:H8")).toBe(false);

    // intersecting at corners
    expect(x("A1:C3")).toBe(true);
    expect(x("F1:H3")).toBe(true);
    expect(x("A6:C8")).toBe(true);
    expect(x("F6:H8")).toBe(true);

    // slice through middle
    expect(x("A4:H5")).toBe(true);
    expect(x("D1:E8")).toBe(true);

    // inside
    expect(x("D4:E5")).toBe(true);

    // outside
    expect(x("A1:H8")).toBe(true);
  });

  it("detects containment", () => {
    const C3F6 = rangeCreate("C3:F6");
    const c = (s: string): boolean => rangeContains(C3F6, s);

    expect(c("A1")).toBe(false);
    expect(c("B2")).toBe(false);
    expect(c("C2")).toBe(false);
    expect(c("D2")).toBe(false);
    expect(c("E2")).toBe(false);
    expect(c("F2")).toBe(false);
    expect(c("G2")).toBe(false);
    expect(c("H1")).toBe(false);
    expect(c("G3")).toBe(false);
    expect(c("G4")).toBe(false);
    expect(c("G5")).toBe(false);
    expect(c("G6")).toBe(false);
    expect(c("G7")).toBe(false);
    expect(c("H7")).toBe(false);
    expect(c("F7")).toBe(false);
    expect(c("E7")).toBe(false);
    expect(c("D7")).toBe(false);
    expect(c("C7")).toBe(false);
    expect(c("B7")).toBe(false);
    expect(c("A8")).toBe(false);
    expect(c("B6")).toBe(false);
    expect(c("B5")).toBe(false);
    expect(c("B4")).toBe(false);
    expect(c("B3")).toBe(false);

    expect(c("C3")).toBe(true);
    expect(c("D3")).toBe(true);
    expect(c("E3")).toBe(true);
    expect(c("F3")).toBe(true);
    expect(c("F4")).toBe(true);
    expect(c("F5")).toBe(true);
    expect(c("F6")).toBe(true);
    expect(c("E6")).toBe(true);
    expect(c("D6")).toBe(true);
    expect(c("C6")).toBe(true);
    expect(c("C5")).toBe(true);
    expect(c("C4")).toBe(true);
    expect(c("D4")).toBe(true);
    expect(c("E4")).toBe(true);
    expect(c("E5")).toBe(true);
    expect(c("D5")).toBe(true);

    expect(c("$A$1")).toBe(false);
    expect(c("$D$5")).toBe(true);

    expect(c("other!$A$1")).toBe(false);
    expect(c("other!$D$5")).toBe(true);

    const otherC3F6 = rangeCreate("other!C3:F6");
    const oc = (s: string): boolean => rangeContains(otherC3F6, s);
    expect(oc("$A$1")).toBe(false);
    expect(oc("$D$5")).toBe(true);
    expect(oc("other!$A$1")).toBe(false);
    expect(oc("other!$D$5")).toBe(true);
    expect(oc("blort!$A$1")).toBe(false);
    expect(oc("blort!$D$5")).toBe(false);
  });

  describe("sheet name serialisation", () => {
    it("emits a bareword for plain ASCII sheet names", () => {
      const r = rangeCreate("A1", "B2", "Sheet1");
      expect(rangeRange(r)).toBe("Sheet1!A1:B2");
    });

    it("quotes sheet names with non-bareword characters", () => {
      const r = rangeCreate("A1", "B2", "My Sheet");
      expect(rangeRange(r)).toBe("'My Sheet'!A1:B2");
    });

    it("doubles single quotes inside quoted sheet names", () => {
      const r = rangeCreate("A1", "B2", "O'Brien");
      expect(rangeRange(r)).toBe("'O''Brien'!A1:B2");
    });

    it("escapes multiple apostrophes", () => {
      const r = rangeCreate("A1", "B2", "It's a 'test'");
      expect(rangeRange(r)).toBe("'It''s a ''test'''!A1:B2");
    });

    it("absolute and short range variants also escape apostrophes", () => {
      const r = rangeCreate("A1", "B2", "O'Brien");
      expect(rangeAbsolute(r)).toBe("'O''Brien'!$A$1:$B$2");
      expect(rangeShortRange(r)).toBe("'O''Brien'!A1:B2");
      expect(rangeAbsoluteShort(r)).toBe("'O''Brien'!$A$1:$B$2");
    });
  });
});
