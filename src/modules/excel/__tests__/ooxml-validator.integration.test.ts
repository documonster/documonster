import { testUtils } from "@excel/__tests__/shared";
import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { makeTestDataPath } from "@test/utils";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

describe("OOXML validator", () => {
  it("validates a complex generated workbook", async () => {
    const wb = testUtils.createTestBook(new Workbook(), "xlsx", [
      "values",
      "conditionalFormatting",
      "dataValidations"
    ]);

    const stress = wb.addWorksheet("stress");

    stress.getCell("A1").value = "Hello, World!";
    stress.getCell("A2").value = { hyperlink: "https://example.com", text: "example.com" };
    stress.mergeCells("A1:B2");

    stress.getCell("C3").value = 123.45;
    stress.getCell("C3").numFmt = "0.00";
    stress.getCell("C3").note = { texts: [{ text: "This is a note" }] };

    // Table + totals
    const table = stress.addTable({
      name: "StressTable",
      ref: "A5",
      headerRow: true,
      totalsRow: true,
      style: { theme: "TableStyleMedium9", showRowStripes: true },
      columns: [
        { name: "Category", totalsRowLabel: "Totals:" },
        { name: "Value", totalsRowFunction: "sum" },
        { name: "Count", totalsRowFunction: "count" }
      ],
      rows: [
        ["A", 10, 1],
        ["A", 20, 1],
        ["B", 5, 1],
        ["B", 7, 1]
      ]
    });

    // Conditional formatting on top of the table.
    stress.addConditionalFormatting({
      ref: "B6:B9",
      rules: [
        {
          type: "cellIs",
          operator: "greaterThan",
          formulae: [10],
          style: {
            fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF00FF00" } }
          }
        }
      ]
    });

    // Legacy form control checkboxes (VML + ctrlProp).
    stress.addFormCheckbox("J2:K3", { link: "L2", text: "Option 1", checked: true });
    stress.addFormCheckbox("J4:J4", { link: "L4", text: "Option 2", checked: false });
    stress.addFormCheckbox("J5:J5", { link: "L5", text: "Option 3", checked: true });
    stress.addFormCheckbox("J6:J6", { link: "L6", text: "Option 4", checked: false });
    stress.addFormCheckbox("J7", { link: "L7", text: "Option 5", checked: true });

    // Images (media + rels + content types)
    const excelTestDataPath = makeTestDataPath(import.meta.url, "./data");
    const imageFilename = excelTestDataPath("image.png");
    const imageId = wb.addImage({ filename: imageFilename, extension: "jpeg" });
    stress.addImage(imageId, "D2:F6");
    stress.addBackgroundImage(imageId);

    // Pivot table from the Table source
    const pivotWs = wb.addWorksheet("pivot");
    pivotWs.addPivotTable({
      sourceTable: table,
      rows: ["Category"],
      columns: [],
      values: ["Value"],
      metric: "sum"
    });

    const buffer = await wb.xlsx.writeBuffer();
    const report = await validateXlsxBuffer(new Uint8Array(buffer), {
      maxProblems: 50
    });

    if (!report.ok) {
      throw new Error(`OOXML validation failed:\n${JSON.stringify(report, null, 2)}`);
    }

    expect(report.ok).toBe(true);
    expect(report.problems).toEqual([]);
  });

  it("validates a workbook with table name containing spaces (issue #91)", async () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet("test");
    sheet.addTable({
      columns: [
        { name: "A", filterButton: true },
        { name: "B", filterButton: true },
        { name: "C", filterButton: true }
      ],
      headerRow: true,
      name: "test table",
      ref: "A1",
      rows: [
        ["test", 2, "a4f"],
        ["test 2", 1, "a4f"],
        ["test 3", 6, "a4f"]
      ],
      totalsRow: false
    });

    const buffer = await wb.xlsx.writeBuffer();
    const report = await validateXlsxBuffer(new Uint8Array(buffer), {
      maxProblems: 50
    });

    if (!report.ok) {
      throw new Error(`OOXML validation failed:\n${JSON.stringify(report, null, 2)}`);
    }

    expect(report.ok).toBe(true);
    expect(report.problems).toEqual([]);
  });
});
