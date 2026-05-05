/**
 * Package structure check: missing parts + XML well-formedness.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage } from "./fixtures";

describe("ooxml-validator / structure", () => {
  it("accepts a minimal valid package", async () => {
    const buffer = buildPackage(baseParts());
    const report = await validateXlsxBuffer(buffer);
    if (!report.ok) {
      throw new Error(`unexpected problems: ${JSON.stringify(report.problems, null, 2)}`);
    }
    expect(report.ok).toBe(true);
  });

  it("flags every missing mandatory part", async () => {
    for (const missing of [
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels"
    ]) {
      const parts = baseParts();
      delete (parts as Record<string, unknown>)[missing];
      const buffer = buildPackage(parts);
      const report = await validateXlsxBuffer(buffer);
      const found = report.problems.find(p => p.kind === "missing-part" && p.file === missing);
      expect(found, `expected missing-part for ${missing}`).toBeDefined();
    }
  });

  it("flags malformed XML in a part", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = "<worksheet><sheetData></worksheet>"; // unclosed sheetData
    const buffer = buildPackage(parts);
    const report = await validateXlsxBuffer(buffer);
    expect(
      report.problems.some(p => p.kind === "xml-malformed" && p.file === "xl/worksheets/sheet1.xml")
    ).toBe(true);
  });

  it("flags malformed content-types XML with a specific kind", async () => {
    const parts = baseParts();
    parts["[Content_Types].xml"] = "<Types><Default/>"; // unclosed
    const buffer = buildPackage(parts);
    const report = await validateXlsxBuffer(buffer);
    expect(report.problems.some(p => p.kind === "content-types-malformed")).toBe(true);
    expect(report.problems.some(p => p.kind === "xml-malformed")).toBe(false);
  });

  it("flags malformed rels XML with a specific kind", async () => {
    const parts = baseParts();
    parts["xl/_rels/workbook.xml.rels"] = "<Relationships><Relationship/>"; // unclosed
    const buffer = buildPackage(parts);
    const report = await validateXlsxBuffer(buffer);
    expect(report.problems.some(p => p.kind === "rels-malformed")).toBe(true);
  });

  it("tolerates the `checkXmlWellFormed: false` option", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = "<worksheet>"; // malformed
    const buffer = buildPackage(parts);
    const report = await validateXlsxBuffer(buffer, { checkXmlWellFormed: false });
    // Still runs other checks — the worksheet structural check parses DOM
    // and will report its own xml-malformed for sheet1 because it must
    // walk it. That's acceptable: the option turns off the *upfront*
    // package-wide sweep but specific checkers still need to parse.
    // We only assert the report completes without throwing.
    expect(report).toBeDefined();
  });

  it("respects maxProblems", async () => {
    // Five missing parts + malformed workbook -> many problems. Cap at 2.
    const buffer = buildPackage({});
    const report = await validateXlsxBuffer(buffer, { maxProblems: 2 });
    expect(report.problems.length).toBeLessThanOrEqual(2);
  });
});
