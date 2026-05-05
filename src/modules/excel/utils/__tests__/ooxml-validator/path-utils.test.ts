/**
 * Pure-function tests for path helpers. These are deterministic,
 * dependency-free checks that make the rest of the validator's path
 * handling predictable.
 */

import {
  getExtension,
  getRelsSourceDir,
  isLegalPartName,
  isSafeResolvedPath,
  posixBasename,
  relsPathForPart,
  resolveRelTarget,
  sourcePartForRels,
  stripLeadingSlash
} from "@excel/utils/ooxml-validator/path-utils";
import { describe, expect, it } from "vitest";

describe("ooxml-validator / path-utils", () => {
  describe("stripLeadingSlash", () => {
    it("removes a single leading slash", () => {
      expect(stripLeadingSlash("/xl/workbook.xml")).toBe("xl/workbook.xml");
    });
    it("leaves non-slash paths alone", () => {
      expect(stripLeadingSlash("xl/workbook.xml")).toBe("xl/workbook.xml");
    });
  });

  describe("getExtension", () => {
    it("returns lowercased extension", () => {
      expect(getExtension("foo/bar.XML")).toBe("xml");
    });
    it("returns empty string when no extension", () => {
      expect(getExtension("foo/bar")).toBe("");
    });
  });

  describe("posixBasename", () => {
    it("returns basename", () => {
      expect(posixBasename("xl/worksheets/sheet1.xml")).toBe("sheet1.xml");
    });
    it("returns input when no slash", () => {
      expect(posixBasename("sheet1.xml")).toBe("sheet1.xml");
    });
  });

  describe("getRelsSourceDir", () => {
    it("returns empty for root rels", () => {
      expect(getRelsSourceDir("_rels/.rels")).toBe("");
    });
    it("returns xl for workbook rels", () => {
      expect(getRelsSourceDir("xl/_rels/workbook.xml.rels")).toBe("xl");
    });
    it("returns xl/worksheets for sheet rels", () => {
      expect(getRelsSourceDir("xl/worksheets/_rels/sheet1.xml.rels")).toBe("xl/worksheets");
    });
  });

  describe("resolveRelTarget", () => {
    it("resolves relative target", () => {
      expect(resolveRelTarget("xl/_rels/workbook.xml.rels", "worksheets/sheet1.xml")).toBe(
        "xl/worksheets/sheet1.xml"
      );
    });
    it("resolves absolute target", () => {
      expect(resolveRelTarget("xl/_rels/workbook.xml.rels", "/xl/styles.xml")).toBe(
        "xl/styles.xml"
      );
    });
    it("handles .. segments", () => {
      expect(
        resolveRelTarget("xl/worksheets/_rels/sheet1.xml.rels", "../drawings/drawing1.xml")
      ).toBe("xl/drawings/drawing1.xml");
    });
    it("marks escaping paths with a leading ../", () => {
      const result = resolveRelTarget("xl/_rels/workbook.xml.rels", "../../escape.xml");
      expect(result.startsWith("../")).toBe(true);
    });
  });

  describe("isSafeResolvedPath", () => {
    it("accepts safe paths", () => {
      expect(isSafeResolvedPath("xl/workbook.xml")).toBe(true);
    });
    it("rejects escape paths", () => {
      expect(isSafeResolvedPath("../escape.xml")).toBe(false);
      expect(isSafeResolvedPath("xl/../../escape.xml")).toBe(false);
    });
  });

  describe("relsPathForPart / sourcePartForRels", () => {
    it("is symmetric for workbook", () => {
      const rels = relsPathForPart("xl/workbook.xml");
      expect(rels).toBe("xl/_rels/workbook.xml.rels");
      expect(sourcePartForRels(rels)).toBe("xl/workbook.xml");
    });
    it("is symmetric for worksheet", () => {
      const rels = relsPathForPart("xl/worksheets/sheet1.xml");
      expect(rels).toBe("xl/worksheets/_rels/sheet1.xml.rels");
      expect(sourcePartForRels(rels)).toBe("xl/worksheets/sheet1.xml");
    });
    it("returns undefined for root rels", () => {
      expect(sourcePartForRels("_rels/.rels")).toBeUndefined();
    });
  });

  describe("isLegalPartName", () => {
    it("accepts normal OPC paths", () => {
      expect(isLegalPartName("xl/workbook.xml")).toBe(true);
    });
    it("rejects empty", () => {
      expect(isLegalPartName("")).toBe(false);
    });
    it("rejects backslashes", () => {
      expect(isLegalPartName("xl\\workbook.xml")).toBe(false);
    });
    it("rejects trailing slash", () => {
      expect(isLegalPartName("xl/")).toBe(false);
    });
    it("rejects dot segments", () => {
      expect(isLegalPartName("xl/./workbook.xml")).toBe(false);
      expect(isLegalPartName("xl/../workbook.xml")).toBe(false);
    });
  });
});
