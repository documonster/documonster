import { testUtils } from "@excel/__tests__/shared";
import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { makeTestDataPath } from "@test/utils";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

async function assertWorkbookOoxmlOk(wb: Workbook): Promise<void> {
  const buffer = await wb.xlsx.writeBuffer();
  const report = await validateXlsxBuffer(new Uint8Array(buffer), { maxProblems: 50 });

  if (!report.ok) {
    throw new Error(`OOXML validation failed:\n${JSON.stringify(report, null, 2)}`);
  }

  expect(report.ok).toBe(true);
  expect(report.problems).toEqual([]);
}

describe("OOXML validator", () => {
  it("validates multiple representative workbooks", async () => {
    // 1) Broad core features via existing shared fixtures.
    {
      const wb = testUtils.createTestBook(new Workbook(), "xlsx", [
        "values",
        "conditionalFormatting",
        "dataValidations"
      ]);
      await assertWorkbookOoxmlOk(wb);
    }

    // 2) Images + hyperlinks.
    {
      const wb = new Workbook();
      const ws = wb.addWorksheet("images");

      ws.getCell("A1").value = {
        hyperlink: "https://example.com",
        text: "example.com"
      };

      const excelTestDataPath = makeTestDataPath(import.meta.url, "./data");
      const imageFilename = excelTestDataPath("image.png");
      const imageId = wb.addImage({ filename: imageFilename, extension: "jpeg" });

      ws.addImage(imageId, "C3:E6");
      ws.addBackgroundImage(imageId);

      await assertWorkbookOoxmlOk(wb);
    }

    // 3) Table + pivot table.
    {
      const wb = new Workbook();
      const data = wb.addWorksheet("data");

      const table = data.addTable({
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

      const pivot = wb.addWorksheet("pivot");
      pivot.addPivotTable({
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
      const wb = new Workbook();
      const ws = wb.addWorksheet("controls");

      ws.addFormCheckbox("J2:K3", { link: "L2", text: "Option 1", checked: true });
      ws.addFormCheckbox("J4:J4", { link: "L4", text: "Option 2", checked: false });
      ws.addFormCheckbox("J5:J5", { link: "L5", text: "Option 3", checked: true });
      ws.addFormCheckbox("J6:J6", { link: "L6", text: "Option 4", checked: false });
      ws.addFormCheckbox("J7", { link: "L7", text: "Option 5", checked: true });

      await assertWorkbookOoxmlOk(wb);
    }
  });
});
