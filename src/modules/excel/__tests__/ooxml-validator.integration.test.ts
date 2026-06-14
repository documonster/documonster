import { testUtils } from "@excel/__tests__/shared";
import { cellSetNote, cellSetNumFmt, cellSetValue } from "@excel/cell";
import { Form, Image, Pivot, Table, Workbook, Worksheet } from "@excel/index";
import { addTable, getCell } from "@excel/worksheet";
import { makeTestDataPath } from "@test/utils";
import { describe, it } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

describe("OOXML validator", () => {
  it("validates a complex generated workbook", async () => {
    const wb = testUtils.createTestBook(Workbook.create(), "xlsx", [
      "values",
      "conditionalFormatting",
      "dataValidations"
    ]);

    const stress = Workbook.addWorksheet(wb, "stress");

    cellSetValue(getCell(stress, "A1"), "Hello, World!");
    cellSetValue(getCell(stress, "A2"), { hyperlink: "https://example.com", text: "example.com" });
    Worksheet.merge(stress, "A1:B2");

    cellSetValue(getCell(stress, "C3"), 123.45);
    cellSetNumFmt(getCell(stress, "C3"), "0.00");
    cellSetNote(getCell(stress, "C3"), { texts: [{ text: "This is a note" }] });

    // Table + totals
    const table = Table.add(stress, {
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
    Worksheet.addConditionalFormatting(stress, {
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
    Form.addCheckbox(stress, "J2:K3", { link: "L2", text: "Option 1", checked: true });
    Form.addCheckbox(stress, "J4:J4", { link: "L4", text: "Option 2", checked: false });
    Form.addCheckbox(stress, "J5:J5", { link: "L5", text: "Option 3", checked: true });
    Form.addCheckbox(stress, "J6:J6", { link: "L6", text: "Option 4", checked: false });
    Form.addCheckbox(stress, "J7", { link: "L7", text: "Option 5", checked: true });

    // Images (media + rels + content types)
    const excelTestDataPath = makeTestDataPath(import.meta.url, "./data");
    const imageFilename = excelTestDataPath("image.png");
    const imageId = Image.add(wb, { filename: imageFilename, extension: "jpeg" });
    Image.place(stress, imageId, "D2:F6");
    Image.setBackground(stress, imageId);

    // Pivot table from the Table source
    const pivotWs = Workbook.addWorksheet(wb, "pivot");
    Pivot.add(pivotWs, {
      sourceTable: table,
      rows: ["Category"],
      columns: [],
      values: ["Value"],
      metric: "sum"
    });

    const buffer = await Workbook.toXlsxBuffer(wb);
    await expectValidXlsx(buffer);
  });

  it("validates a workbook with table name containing spaces (issue #91)", async () => {
    const wb = Workbook.create();
    const sheet = Workbook.addWorksheet(wb, "test");
    addTable(sheet, {
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

    const buffer = await Workbook.toXlsxBuffer(wb);
    await expectValidXlsx(buffer);
  });
});
