import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { Cell, Column, Row, Workbook, Worksheet } from "@excel/index";
import { getUniqueTestFilePath } from "@test/utils";
import { describe, expect, it } from "vitest";

describe("xlsx styles roundtrip", () => {
  it("writes and reads common formatting", async () => {
    const filename = getUniqueTestFilePath(import.meta.url);

    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "blort");

    Worksheet.setColumns(ws, [
      { header: "Col 1", key: "key", width: 25 },
      { header: "Col 2", key: "name", width: 32 },
      { header: "Col 3", key: "age" }
    ]);

    Column.setHidden(ws, 9, true);
    Cell.setValue(ws, "A16", "hidden");
    Row.setHidden(ws, 16, true);

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

    Cell.setValue(ws, "A2", 7);
    Cell.setValue(ws, "B2", "Hello, World!");
    Cell.setStyle(ws, "B2", { font: fonts.comicSansUdB16 });
    Cell.setStyle(ws, "B2", { border: borders.thin });

    Cell.setValue(ws, "C2", -5.55);
    Cell.setStyle(ws, "C2", { numFmt: "'£'#,##0.00;[Red]-'£'#,##0.00" });

    Cell.setValue(ws, "D2", new Date("2020-01-02T00:00:00.000Z"));
    Cell.setStyle(ws, "D2", { numFmt: "d-mmm-yyyy" });

    Cell.setValue(ws, "F2", true);
    Cell.setValue(ws, "G2", { error: "#N/A" });

    Cell.setValue(ws, "C5", { formula: "A2", result: 7 });

    Row.setHeight(ws, 11, 40);
    Cell.setValue(ws, 11, 1, "Top Left");
    Cell.setStyle(ws, 11, 1, { alignment: { horizontal: "left", vertical: "top" } });

    await Workbook.writeFile(wb, filename);
    await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));

    const wb2 = Workbook.create();
    await Workbook.readFile(wb2, filename);

    const ws2 = Workbook.getWorksheet(wb2, "blort")!;
    expect(ws2).toBeTruthy();

    expect(Column.getWidth(ws2, 1)).toBe(25);
    expect(Column.getHidden(ws2, 9)).toBe(true);
    expect(Row.getHidden(ws2, 16)).toBe(true);

    expect(Cell.getValue(ws2, "A2")).toBe(7);
    expect(Cell.getValue(ws2, "B2")).toBe("Hello, World!");
    expect(Cell.getStyle(ws2, "B2").font).toMatchObject(fonts.comicSansUdB16);
    expect(Cell.getStyle(ws2, "B2").border).toMatchObject(borders.thin);

    expect(Cell.getValue(ws2, "C2")).toBe(-5.55);
    expect(Cell.getStyle(ws2, "C2").numFmt).toBe("'£'#,##0.00;[Red]-'£'#,##0.00");

    const d2 = Cell.getValue(ws2, "D2");
    expect(d2 instanceof Date || typeof d2 === "number").toBe(true);
    expect(Cell.getStyle(ws2, "D2").numFmt).toBe("d-mmm-yyyy");

    expect(Cell.getValue(ws2, "F2")).toBe(true);
    expect((Cell.getValue(ws2, "G2") as any).error).toBe("#N/A");

    expect((Cell.getValue(ws2, "C5") as any).formula).toBe("A2");

    expect(Row.getHeight(ws2, 11)).toBe(40);
    expect(Cell.getValue(ws2, 11, 1)).toBe("Top Left");
    expect(Cell.getStyle(ws2, 11, 1).alignment).toMatchObject({
      horizontal: "left",
      vertical: "top"
    });
  });

  // Regression test for a style-roundtrip edge case
  // This tests that styled cells retain their style after roundtrip
  it("styled cells retain style after roundtrip", async () => {
    const filename = getUniqueTestFilePath(import.meta.url);

    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    Cell.setValue(ws, "A1", "Plain text");

    Cell.setValue(ws, "B1", "Styled text");
    Cell.setStyle(ws, "B1", { font: { bold: true } });

    await Workbook.writeFile(wb, filename);
    await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));

    const wb2 = Workbook.create();
    await Workbook.readFile(wb2, filename);

    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
    expect(ws2).toBeTruthy();

    // B1 should retain its bold style
    expect(Cell.getStyle(ws2, "B1").font).toMatchObject({ bold: true });

    // A1 has no explicit style, so it may or may not have style info
    // (depending on whether the file format includes s="0" for default style)
    const a1Value = Cell.getValue(ws2, "A1");
    expect(a1Value).toBe("Plain text");
  });
});
