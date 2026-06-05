/**
 * DOCX Module - Unit Conversion Tests
 */

import { describe, it, expect } from "vitest";

import {
  cmToTwips,
  twipsToCm,
  twipsToPt,
  mmToTwips,
  emuToInches,
  emuToCm,
  ptToEmu,
  emuToPx,
  halfPointToPt,
  eighthPointToPt,
  spacingToLineMultiplier,
  tablePctToPercent
} from "../index";

describe("Unit conversions", () => {
  describe("cmToTwips", () => {
    it("should convert 1 cm to 567 twips", () => {
      expect(cmToTwips(1)).toBe(567);
    });

    it("should convert 2.54 cm (1 inch) to ~1440 twips", () => {
      expect(cmToTwips(2.54)).toBe(1440);
    });

    it("should handle zero", () => {
      expect(cmToTwips(0)).toBe(0);
    });

    it("should produce canonical A4 twips (matches A4_PAGE_WIDTH/HEIGHT)", () => {
      // A4 = 210mm × 297mm. Canonical OOXML twips are 11906 × 16838.
      expect(cmToTwips(21)).toBe(11906);
      expect(cmToTwips(29.7)).toBe(16838);
    });
  });

  describe("twipsToCm", () => {
    it("should convert 1 inch worth of twips (1440) to 2.54 cm", () => {
      expect(twipsToCm(1440)).toBeCloseTo(2.54, 6);
    });

    it("should round-trip cm → twips → cm", () => {
      // cmToTwips rounds to an integer twip, so the round-trip is accurate to
      // within half a twip (≈0.001 cm).
      expect(twipsToCm(cmToTwips(1))).toBeCloseTo(1, 2);
      expect(twipsToCm(cmToTwips(21))).toBeCloseTo(21, 2);
    });

    it("should handle zero", () => {
      expect(twipsToCm(0)).toBe(0);
    });
  });

  describe("twipsToPt", () => {
    it("should convert 20 twips to 1 pt", () => {
      expect(twipsToPt(20)).toBe(1);
    });

    it("should convert 240 twips to 12 pt", () => {
      expect(twipsToPt(240)).toBe(12);
    });
  });

  describe("mmToTwips", () => {
    it("should convert 10 mm to ~567 twips", () => {
      expect(mmToTwips(10)).toBe(567);
    });

    it("should convert 1 mm to ~57 twips", () => {
      expect(mmToTwips(1)).toBe(57);
    });
  });

  describe("emuToInches", () => {
    it("should convert 914400 EMU to 1 inch", () => {
      expect(emuToInches(914400)).toBe(1);
    });

    it("should handle zero", () => {
      expect(emuToInches(0)).toBe(0);
    });
  });

  describe("emuToCm", () => {
    it("should convert 360000 EMU to 1 cm", () => {
      expect(emuToCm(360000)).toBe(1);
    });
  });

  describe("ptToEmu", () => {
    it("should convert 1 pt to 12700 EMU", () => {
      expect(ptToEmu(1)).toBe(12700);
    });

    it("should convert 72 pt to 914400 EMU (1 inch)", () => {
      expect(ptToEmu(72)).toBe(914400);
    });
  });

  describe("emuToPx", () => {
    it("should convert 9525 EMU to 1 px at 96 DPI", () => {
      expect(emuToPx(9525)).toBe(1);
    });

    it("should convert 914400 EMU to 96 px (1 inch at 96 DPI)", () => {
      expect(emuToPx(914400)).toBeCloseTo(96, 0);
    });
  });

  describe("halfPointToPt", () => {
    it("should convert 24 half-points to 12 pt", () => {
      expect(halfPointToPt(24)).toBe(12);
    });

    it("should convert 1 half-point to 0.5 pt", () => {
      expect(halfPointToPt(1)).toBe(0.5);
    });
  });

  describe("eighthPointToPt", () => {
    it("should convert 8 eighth-points to 1 pt", () => {
      expect(eighthPointToPt(8)).toBe(1);
    });

    it("should convert 4 eighth-points to 0.5 pt", () => {
      expect(eighthPointToPt(4)).toBe(0.5);
    });
  });

  describe("spacingToLineMultiplier", () => {
    it("should convert 240 spacing to 1.0 multiplier", () => {
      expect(spacingToLineMultiplier(240)).toBe(1);
    });

    it("should convert 360 spacing to 1.5 multiplier", () => {
      expect(spacingToLineMultiplier(360)).toBe(1.5);
    });

    it("should convert 480 spacing to 2.0 multiplier", () => {
      expect(spacingToLineMultiplier(480)).toBe(2);
    });
  });

  describe("tablePctToPercent", () => {
    it("should convert 5000 to 100%", () => {
      expect(tablePctToPercent(5000)).toBe(100);
    });

    it("should convert 2500 to 50%", () => {
      expect(tablePctToPercent(2500)).toBe(50);
    });

    it("should convert 0 to 0%", () => {
      expect(tablePctToPercent(0)).toBe(0);
    });
  });
});
