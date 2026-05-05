/**
 * Workbook-level check: sheet wiring, name constraints, id uniqueness.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage, relsWith } from "./fixtures";

function workbookWith(sheets: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets}</sheets>
</workbook>`;
}

describe("ooxml-validator / workbook", () => {
  it("flags duplicate sheetId", async () => {
    const parts = baseParts();
    parts["xl/workbook.xml"] = workbookWith(
      `<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="1" r:id="rId1"/>`
    );
    // Add second rel target so sheet2 wiring isn't the complaint.
    parts["xl/worksheets/sheet2.xml"] = parts["xl/worksheets/sheet1.xml"] as string;
    parts["xl/_rels/workbook.xml.rels"] = relsWith([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        target: "worksheets/sheet1.xml"
      },
      {
        id: "rId2",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        target: "theme/theme1.xml"
      },
      {
        id: "rId3",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        target: "styles.xml"
      },
      {
        id: "rId4",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        target: "sharedStrings.xml"
      }
    ]);
    parts["[Content_Types].xml"] = (parts["[Content_Types].xml"] as string).replace(
      "</Types>",
      `  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`
    );
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "workbook-duplicate-sheetId")).toBe(true);
  });

  it("flags duplicate sheet r:id", async () => {
    const parts = baseParts();
    parts["xl/workbook.xml"] = workbookWith(
      `<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId1"/>`
    );
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "workbook-duplicate-sheet-rid")).toBe(true);
  });

  it("flags sheet pointing at missing rel", async () => {
    const parts = baseParts();
    parts["xl/workbook.xml"] = workbookWith(
      `<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rIdMissing"/>`
    );
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(
      report.problems.some(
        p => p.kind === "workbook-sheet-missing-rel" && p.message.includes("rIdMissing")
      )
    ).toBe(true);
  });

  it("flags sheet rel of wrong type", async () => {
    const parts = baseParts();
    parts["xl/_rels/workbook.xml.rels"] = relsWith([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        target: "theme/theme1.xml"
      },
      {
        id: "rId2",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        target: "theme/theme1.xml"
      },
      {
        id: "rId3",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        target: "styles.xml"
      },
      {
        id: "rId4",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        target: "sharedStrings.xml"
      }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "workbook-sheet-wrong-rel-type")).toBe(true);
  });

  it("flags missing sheet name", async () => {
    const parts = baseParts();
    parts["xl/workbook.xml"] = workbookWith(`<sheet sheetId="1" r:id="rId1"/>`);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "workbook-sheet-missing-name")).toBe(true);
  });

  it("flags sheet name > 31 chars", async () => {
    const parts = baseParts();
    parts["xl/workbook.xml"] = workbookWith(
      `<sheet name="ThisNameIsDefinitelyWayTooLongForExcel" sheetId="1" r:id="rId1"/>`
    );
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "workbook-sheet-name-too-long")).toBe(true);
  });

  it("flags sheet name with illegal characters", async () => {
    const parts = baseParts();
    parts["xl/workbook.xml"] = workbookWith(`<sheet name="A/B" sheetId="1" r:id="rId1"/>`);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "workbook-sheet-name-invalid-chars")).toBe(true);
  });

  it("flags duplicate sheet names (case-insensitive)", async () => {
    const parts = baseParts();
    parts["xl/workbook.xml"] = workbookWith(
      `<sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="SHEET1" sheetId="2" r:id="rId2"/>`
    );
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "workbook-sheet-name-duplicate")).toBe(true);
  });
});
