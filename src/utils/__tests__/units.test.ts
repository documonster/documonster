import {
  getPixelPadding,
  charWidthToPixel,
  pixelToCharWidth,
  pixelToPoints,
  pointsToPixel,
  emuToPt
} from "@utils/units";
import { describe, it, expect } from "vitest";

describe("@utils/units column-width helpers", () => {
  // Calibri 11pt default MDW.
  const MDW = 7;

  describe("getPixelPadding", () => {
    it("derives PP = 2*ceil(MDW/4)+1 (5 for MDW=7)", () => {
      expect(getPixelPadding(7)).toBe(5);
      expect(getPixelPadding(8)).toBe(5);
      expect(getPixelPadding(4)).toBe(3);
    });
  });

  describe("charWidthToPixel", () => {
    it("uses ROUND(width*MDW)+PP for width >= 1", () => {
      // Excel's canonical 8.43 default column → 64px at MDW 7.
      expect(charWidthToPixel(8.43, MDW)).toBe(Math.round(8.43 * 7) + 5);
      expect(charWidthToPixel(10, MDW)).toBe(75);
    });

    it("uses ROUND(width*(MDW+PP)) for width < 1", () => {
      expect(charWidthToPixel(0.5, MDW)).toBe(Math.round(0.5 * (7 + 5)));
    });

    it("returns 0 for zero width or non-positive MDW", () => {
      expect(charWidthToPixel(0, MDW)).toBe(0);
      expect(charWidthToPixel(10, 0)).toBe(0);
    });
  });

  describe("pixelToCharWidth", () => {
    it("is the 1/256-precision inverse direction of the stored width", () => {
      expect(pixelToCharWidth(64, MDW)).toBe(Math.trunc((64 / 7) * 256) / 256);
      expect(pixelToCharWidth(64, 0)).toBe(0);
    });
  });

  describe("pixel <-> point at 96 DPI", () => {
    it("converts px to pt with the 72/96 factor", () => {
      expect(pixelToPoints(96)).toBe(72);
      expect(pointsToPixel(72)).toBe(96);
    });
  });

  describe("emuToPt", () => {
    it("divides by 12700 EMU per point", () => {
      expect(emuToPt(12_700)).toBe(1);
      expect(emuToPt(914_400)).toBe(72);
    });
  });
});
