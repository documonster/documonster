import {
  decodeCol,
  encodeCol,
  decodeRow,
  encodeRow,
  decodeCell,
  encodeCell,
  decodeRange,
  encodeRange
} from "@excel/utils/address";
import { describe, it, expect } from "vitest";

describe("Address utilities", () => {
  // ===========================================================================
  // Column Encoding/Decoding
  // ===========================================================================

  describe("decodeCol", () => {
    it("should decode single letter columns", () => {
      expect(decodeCol("A")).toBe(0);
      expect(decodeCol("B")).toBe(1);
      expect(decodeCol("Z")).toBe(25);
    });

    it("should decode double letter columns", () => {
      expect(decodeCol("AA")).toBe(26);
      expect(decodeCol("AB")).toBe(27);
      expect(decodeCol("AZ")).toBe(51);
      expect(decodeCol("BA")).toBe(52);
    });

    it("should handle lowercase letters", () => {
      expect(decodeCol("a")).toBe(0);
      expect(decodeCol("aa")).toBe(26);
    });
  });

  describe("encodeCol", () => {
    it("should encode single letter columns", () => {
      expect(encodeCol(0)).toBe("A");
      expect(encodeCol(1)).toBe("B");
      expect(encodeCol(25)).toBe("Z");
    });

    it("should encode double letter columns", () => {
      expect(encodeCol(26)).toBe("AA");
      expect(encodeCol(27)).toBe("AB");
      expect(encodeCol(51)).toBe("AZ");
      expect(encodeCol(52)).toBe("BA");
    });
  });

  // ===========================================================================
  // Row Encoding/Decoding
  // ===========================================================================

  describe("decodeRow", () => {
    it("should decode row strings to 0-indexed numbers", () => {
      expect(decodeRow("1")).toBe(0);
      expect(decodeRow("10")).toBe(9);
      expect(decodeRow("100")).toBe(99);
    });
  });

  describe("encodeRow", () => {
    it("should encode 0-indexed numbers to row strings", () => {
      expect(encodeRow(0)).toBe("1");
      expect(encodeRow(9)).toBe("10");
      expect(encodeRow(99)).toBe("100");
    });
  });

  // ===========================================================================
  // Cell Address Encoding/Decoding
  // ===========================================================================

  describe("decodeCell", () => {
    it("should decode cell references to CellAddress", () => {
      expect(decodeCell("A1")).toEqual({ c: 0, r: 0 });
      expect(decodeCell("B2")).toEqual({ c: 1, r: 1 });
      expect(decodeCell("AA10")).toEqual({ c: 26, r: 9 });
    });

    it("should handle lowercase references", () => {
      expect(decodeCell("a1")).toEqual({ c: 0, r: 0 });
      expect(decodeCell("b2")).toEqual({ c: 1, r: 1 });
    });
  });

  describe("encodeCell", () => {
    it("should encode CellAddress to cell references", () => {
      expect(encodeCell({ c: 0, r: 0 })).toBe("A1");
      expect(encodeCell({ c: 1, r: 1 })).toBe("B2");
      expect(encodeCell({ c: 26, r: 9 })).toBe("AA10");
    });
  });

  describe("decodeCell and encodeCell roundtrip", () => {
    it("should roundtrip correctly", () => {
      const addresses = ["A1", "B2", "Z100", "AA1", "AZ52"];
      for (const addr of addresses) {
        expect(encodeCell(decodeCell(addr))).toBe(addr);
      }
    });
  });

  // ===========================================================================
  // Range Encoding/Decoding
  // ===========================================================================

  describe("decodeRange", () => {
    it("should decode range strings", () => {
      expect(decodeRange("A1:B2")).toEqual({
        s: { c: 0, r: 0 },
        e: { c: 1, r: 1 }
      });
    });

    it("should decode single cell as range", () => {
      const result = decodeRange("A1");
      expect(result.s).toEqual({ c: 0, r: 0 });
      expect(result.e).toEqual({ c: 0, r: 0 });
    });
  });

  describe("encodeRange", () => {
    it("should encode Range object", () => {
      expect(encodeRange({ s: { c: 0, r: 0 }, e: { c: 1, r: 1 } })).toBe("A1:B2");
    });

    it("should encode two CellAddress objects", () => {
      expect(encodeRange({ c: 0, r: 0 }, { c: 1, r: 1 })).toBe("A1:B2");
    });

    it("should return single cell for same start and end", () => {
      expect(encodeRange({ c: 0, r: 0 }, { c: 0, r: 0 })).toBe("A1");
    });
  });
});
