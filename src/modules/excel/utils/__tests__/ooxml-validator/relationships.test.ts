/**
 * Relationships check: duplicate ids, missing targets, wrong rel types,
 * external targets, path traversal, source-part consistency.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage, relsWith } from "./fixtures";

describe("ooxml-validator / relationships", () => {
  it("flags root rels missing officeDocument", async () => {
    const parts = baseParts();
    parts["_rels/.rels"] = relsWith([
      {
        id: "rIdX",
        type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        target: "docProps/core.xml"
      }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "root-rels-missing-officeDocument")).toBe(true);
  });

  it("flags duplicate rel Id", async () => {
    const parts = baseParts();
    parts["xl/_rels/workbook.xml.rels"] = relsWith([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        target: "worksheets/sheet1.xml"
      },
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        target: "theme/theme1.xml"
      }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "rels-duplicate-id")).toBe(true);
  });

  it("flags rel with missing target file", async () => {
    const parts = baseParts();
    parts["xl/_rels/workbook.xml.rels"] = relsWith([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        target: "worksheets/sheet1.xml"
      },
      {
        id: "rId2",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        target: "theme/missing.xml"
      }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(
      report.problems.some(
        p => p.kind === "rels-missing-target" && p.message.includes("theme/missing.xml")
      )
    ).toBe(true);
  });

  it("flags empty Target attribute", async () => {
    const parts = baseParts();
    parts["xl/_rels/workbook.xml.rels"] = relsWith([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        target: "worksheets/sheet1.xml"
      },
      {
        id: "rId2",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        target: ""
      }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "rels-empty-target")).toBe(true);
  });

  it("flags target escaping package root", async () => {
    const parts = baseParts();
    parts["xl/_rels/workbook.xml.rels"] = relsWith([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        target: "worksheets/sheet1.xml"
      },
      {
        id: "rId2",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        target: "../../secret.xml"
      }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "rels-invalid-target-path")).toBe(true);
  });

  it("flags rel type vs target-layout mismatch", async () => {
    const parts = baseParts();
    // Declare the worksheet rel to point at theme/theme1.xml.
    parts["xl/_rels/workbook.xml.rels"] = relsWith([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        target: "theme/theme1.xml"
      },
      {
        id: "rId2",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        target: "theme/theme1.xml"
      }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "rels-type-target-mismatch")).toBe(true);
  });

  it("ignores TargetMode=External", async () => {
    const parts = baseParts();
    parts["xl/_rels/workbook.xml.rels"] = relsWith([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        target: "worksheets/sheet1.xml"
      },
      {
        id: "rId2",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: "https://example.com",
        targetMode: "External"
      }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "rels-missing-target")).toBe(false);
  });

  it("flags rels file with no corresponding source part", async () => {
    const parts = baseParts();
    // Add stray rels file for non-existent part.
    parts["xl/ghost.xml.rels"] = parts["xl/_rels/workbook.xml.rels"] as string;
    // Actually that's wrong path format. OPC puts .rels under _rels/.
    // Put it in a proper _rels folder so it gets picked up.
    parts["xl/_rels/ghost.xml.rels"] = relsWith([]);
    delete (parts as Record<string, unknown>)["xl/ghost.xml.rels"];
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(
      report.problems.some(p => p.kind === "rels-source-missing" && p.message.includes("ghost.xml"))
    ).toBe(true);
  });
});
