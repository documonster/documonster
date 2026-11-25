import { extractAll } from "@archive/unzip/extract";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { requireEntryText } from "@excel/__tests__/helpers/zip-text";
import { Stream, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

/**
 * `CT_SheetPr` is a sequence — tabColor, outlinePr, pageSetUpPr — and Excel repairs a sheet whose `<sheetPr>`
 * has them in another order. Both writers rendered `pageSetUpPr` before `outlinePr` (issue #269).
 */

const EXPECTED =
  '<sheetPr><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0" summaryRight="1"/><pageSetUpPr fitToPage="1"/></sheetPr>';

async function sheetPr(bytes: Uint8Array): Promise<string> {
  const xml = requireEntryText(await extractAll(bytes), "xl/worksheets/sheet1.xml");
  return /<sheetPr>.*?<\/sheetPr>/.exec(xml)![0];
}

describe("worksheet <sheetPr> child order", () => {
  it("Workbook.toBuffer writes outlinePr before pageSetUpPr", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    ws.properties.tabColor = { argb: "FFFF0000" };
    ws.properties.outlineProperties = { summaryBelow: false, summaryRight: true };
    ws.pageSetup = { ...ws.pageSetup, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

    const bytes = await Workbook.toBuffer(wb);
    expect(await sheetPr(bytes)).toBe(EXPECTED);
    await expectValidXlsx(bytes);
  });

  it("the streaming writer writes outlinePr before pageSetUpPr", async () => {
    const writer = new Stream.WorkbookWriter({});
    const ws = writer.addWorksheet("Sheet1", {
      properties: {
        tabColor: { argb: "FFFF0000" },
        outlineProperties: { summaryBelow: false, summaryRight: true }
      },
      pageSetup: { fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    });
    Stream.commitRow(ws.addRow([1]));

    const bytes = await writer.toBuffer();
    expect(await sheetPr(bytes)).toBe(EXPECTED);
    await expectValidXlsx(bytes);
  });
});
