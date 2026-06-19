import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { testUtils } from "@excel/__tests__/shared";
import type { WorkbookData } from "@excel/core/workbook-core";
import { addWorkbookImage } from "@excel/core/workbook-core";
import {
  addBackgroundImage,
  addFormCheckbox,
  addImage,
  addPivotTable,
  addTable
} from "@excel/core/worksheet";
import { Cell, Workbook } from "@excel/index";
import { makeTestDataPath } from "@test/utils";
import { describe, it } from "vitest";

async function assertWorkbookOoxmlOk(wb: WorkbookData): Promise<void> {
  const buffer = await Workbook.toBuffer(wb);
  await expectValidXlsx(buffer);
}

describe("OOXML validator", () => {
  it("validates multiple representative workbooks", async () => {
    // 1) Broad core features via existing shared fixtures.
    {
      const wb = testUtils.createTestBook(Workbook.create(), "xlsx", [
        "values",
        "conditionalFormatting",
        "dataValidations"
      ]);
      await assertWorkbookOoxmlOk(wb);
    }

    // 2) Images + hyperlinks.
    {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "images");

      Cell.setValue(ws, "A1", {
        hyperlink: "https://example.com",
        text: "example.com"
      });

      const excelTestDataPath = makeTestDataPath(import.meta.url, "../../__tests__/data");
      const imageFilename = excelTestDataPath("image.png");
      const imageId = addWorkbookImage(wb, { filename: imageFilename, extension: "jpeg" });

      addImage(ws, imageId, "C3:E6");
      addBackgroundImage(ws, imageId);

      await assertWorkbookOoxmlOk(wb);
    }

    // 3) Table + pivot table.
    {
      const wb = Workbook.create();
      const data = Workbook.addWorksheet(wb, "data");

      const table = addTable(data, {
        name: "SalesData",
        ref: "A1",
        columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }, { name: "E" }],
        rows: [
          ["a1", "b1", "c1", 4, 5],
          ["a1", "b2", "c1", 4, 5],
          ["a2", "b1", "c2", 14, 24],
          ["a2", "b2", "c2", 24, 35]
        ]
      });

      const pivot = Workbook.addWorksheet(wb, "pivot");
      addPivotTable(pivot, {
        sourceTable: table,
        rows: ["A"],
        columns: ["C"],
        values: ["E"],
        metric: "sum"
      });

      await assertWorkbookOoxmlOk(wb);
    }

    // 4) Legacy form control checkboxes (VML + ctrlProps).
    {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "controls");

      addFormCheckbox(ws, "J2:K3", { link: "L2", text: "Option 1", checked: true });
      addFormCheckbox(ws, "J4:J4", { link: "L4", text: "Option 2", checked: false });
      addFormCheckbox(ws, "J5:J5", { link: "L5", text: "Option 3", checked: true });
      addFormCheckbox(ws, "J6:J6", { link: "L6", text: "Option 4", checked: false });
      addFormCheckbox(ws, "J7", { link: "L7", text: "Option 5", checked: true });

      await assertWorkbookOoxmlOk(wb);
    }
  });
});
