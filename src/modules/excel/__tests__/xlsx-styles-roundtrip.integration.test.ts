import { getUniqueTestFilePath } from "@test/utils";
import { describe, expect, it } from "vitest";

import { Workbook } from "../../../index";
import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

describe("xlsx styles roundtrip", () => {
  it("writes and reads common formatting", async () => {
    const filename = getUniqueTestFilePath(import.meta.url);

    const wb = new Workbook();
    const ws = wb.addWorksheet("blort");

    ws.columns = [
      { header: "Col 1", key: "key", width: 25 },
      { header: "Col 2", key: "name", width: 32 },
      { header: "Col 3", key: "age" }
    ];

    ws.getColumn(9).hidden = true;
    ws.getCell("A16").value = "hidden";
    ws.getRow(16).hidden = true;

    const fonts = {
      comicSansUdB16: {
        name: "Comic Sans MS",
        family: 4,
        size: 16,
        underline: "double" as const,
        bold: true
      }
    };

    const borders = {
      thin: {
        top: { style: "thin" as const },
        left: { style: "thin" as const },
        bottom: { style: "thin" as const },
        right: { style: "thin" as const }
      }
    };

    ws.getCell("A2").value = 7;
    ws.getCell("B2").value = "Hello, World!";
    ws.getCell("B2").font = fonts.comicSansUdB16;
    ws.getCell("B2").border = borders.thin;

    ws.getCell("C2").value = -5.55;
    ws.getCell("C2").numFmt = "'£'#,##0.00;[Red]-'£'#,##0.00";

    ws.getCell("D2").value = new Date("2020-01-02T00:00:00.000Z");
    ws.getCell("D2").numFmt = "d-mmm-yyyy";

    ws.getCell("F2").value = true;
    ws.getCell("G2").value = { error: "#N/A" };

    ws.getCell("C5").value = { formula: "A2", result: 7 };

    ws.getRow(11).height = 40;
    ws.getCell(11, 1).value = "Top Left";
    ws.getCell(11, 1).alignment = { horizontal: "left", vertical: "top" };

    await wb.xlsx.writeFile(filename);
    await expectValidXlsx(new Uint8Array(await wb.xlsx.writeBuffer()));

    const wb2 = new Workbook();
    await wb2.xlsx.readFile(filename);

    const ws2 = wb2.getWorksheet("blort")!;
    expect(ws2).toBeTruthy();

    expect(ws2.getColumn(1).width).toBe(25);
    expect(ws2.getColumn(9).hidden).toBe(true);
    expect(ws2.getRow(16).hidden).toBe(true);

    expect(ws2.getCell("A2").value).toBe(7);
    expect(ws2.getCell("B2").value).toBe("Hello, World!");
    expect(ws2.getCell("B2").font).toMatchObject(fonts.comicSansUdB16);
    expect(ws2.getCell("B2").border).toMatchObject(borders.thin);

    expect(ws2.getCell("C2").value).toBe(-5.55);
    expect(ws2.getCell("C2").numFmt).toBe("'£'#,##0.00;[Red]-'£'#,##0.00");

    const d2 = ws2.getCell("D2").value;
    expect(d2 instanceof Date || typeof d2 === "number").toBe(true);
    expect(ws2.getCell("D2").numFmt).toBe("d-mmm-yyyy");

    expect(ws2.getCell("F2").value).toBe(true);
    expect((ws2.getCell("G2").value as any).error).toBe("#N/A");

    expect((ws2.getCell("C5").value as any).formula).toBe("A2");

    expect(ws2.getRow(11).height).toBe(40);
    expect(ws2.getCell(11, 1).value).toBe("Top Left");
    expect(ws2.getCell(11, 1).alignment).toMatchObject({
      horizontal: "left",
      vertical: "top"
    });
  });

  // Regression test for https://github.com/exceljs/exceljs/issues/2600
  // This tests that styled cells retain their style after roundtrip
  it("styled cells retain style after roundtrip", async () => {
    const filename = getUniqueTestFilePath(import.meta.url);

    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    ws.getCell("A1").value = "Plain text";

    ws.getCell("B1").value = "Styled text";
    ws.getCell("B1").font = { bold: true };

    await wb.xlsx.writeFile(filename);
    await expectValidXlsx(new Uint8Array(await wb.xlsx.writeBuffer()));

    const wb2 = new Workbook();
    await wb2.xlsx.readFile(filename);

    const ws2 = wb2.getWorksheet("Sheet1")!;
    expect(ws2).toBeTruthy();

    // B1 should retain its bold style
    expect(ws2.getCell("B1").font).toMatchObject({ bold: true });

    // A1 has no explicit style, so it may or may not have style info
    // (depending on whether the file format includes s="0" for default style)
    const a1Value = ws2.getCell("A1").value;
    expect(a1Value).toBe("Plain text");
  });
});
