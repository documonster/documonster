import { readFileSync } from "node:fs";

import { extractAll } from "@archive/unzip/extract";
import { Cell, Image, Pivot, Sparkline, Table, Workbook, Worksheet } from "@excel";
import { validateXlsbBuffer } from "@excel/utils/xlsb-validator";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

/**
 * The XLSB worksheet and workbook record order is MS-XLSB's grammar (§2.1.7.62, §2.1.7.61), and three positions
 * this writer used were not: the merge cells ahead of the auto filter, the sparklines ahead of the drawing and the
 * list parts, and the pivot cache ids after `BrtFileRecover`. Each was placed from a reference workbook that did
 * not have the neighbour it ended up on the wrong side of, so the case here puts every neighbour on one sheet.
 */

const PNG = readFileSync("src/modules/excel/__tests__/data/image.png");

async function sheetRecordNames(bytes: Uint8Array): Promise<string[]> {
  const sheet = (await extractAll(bytes)).get("xl/worksheets/sheet1.bin")!.data;
  return [...iterateInterpretableRecords(sheet, "s")].map(
    r => recordSpec(r.id)?.name ?? `#${r.id}`
  );
}

function before(names: readonly string[], first: string, second: string): boolean {
  const a = names.indexOf(first);
  const b = names.indexOf(second);
  expect(a, `${first} is written`).toBeGreaterThanOrEqual(0);
  expect(b, `${second} is written`).toBeGreaterThanOrEqual(0);
  return a < b;
}

describe("XLSB worksheet record order", () => {
  it("writes the auto filter, merges, drawing, list parts and sparklines in the grammar's order", async () => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    for (let row = 1; row <= 4; row++) {
      for (let col = 1; col <= 3; col++) {
        Cell.setValue(sheet, { row, col }, row * col);
      }
    }
    Worksheet.setAutoFilter(sheet, "A1:C4");
    Worksheet.merge(sheet, "E6:F7");
    Image.place(sheet, Image.add(workbook, { buffer: PNG, extension: "png" }), {
      tl: { col: 7, row: 1 },
      ext: { width: 40, height: 40 }
    });
    Table.add(sheet, { name: "T", ref: "A10", columns: [{ name: "Q" }], rows: [[1], [2]] });
    Sparkline.add(sheet, {
      type: "line",
      sparklines: [{ cellRef: "E1", dataRef: "A1:C1" }]
    });

    const bytes = await Workbook.toBuffer(workbook, { format: "xlsb" });
    const names = await sheetRecordNames(bytes);

    expect(before(names, "BrtBeginAFilter", "BrtBeginMergeCells")).toBe(true);
    expect(before(names, "BrtDrawing", "BrtBeginSparklineGroups")).toBe(true);
    expect(before(names, "BrtBeginListParts", "BrtBeginSparklineGroups")).toBe(true);

    const report = await validateXlsbBuffer(bytes);
    expect(report.problems.filter(p => p.kind === "ordering-out-of-place")).toEqual([]);
  });
});

describe("XLSB workbook record order", () => {
  it("writes the pivot cache ids before BrtFileRecover", async () => {
    const workbook = Workbook.create();
    const source = Workbook.addWorksheet(workbook, "Data");
    Cell.setValue(source, "A1", "Region");
    Cell.setValue(source, "B1", "Units");
    Cell.setValue(source, "A2", "APAC");
    Cell.setValue(source, "B2", 10);
    Pivot.add(Workbook.addWorksheet(workbook, "Report"), {
      sourceSheet: source,
      rows: ["Region"],
      values: ["Units"],
      metric: "sum"
    });

    const bytes = await Workbook.toBuffer(workbook, { format: "xlsb", unsupported: "ignore" });
    const book = (await extractAll(bytes)).get("xl/workbook.bin")!.data;
    const names = [...iterateInterpretableRecords(book, "b")].map(
      r => recordSpec(r.id)?.name ?? `#${r.id}`
    );
    expect(before(names, "BrtBeginPivotCacheIDs", "BrtFileRecover")).toBe(true);

    const report = await validateXlsbBuffer(bytes);
    expect(report.problems.filter(p => p.kind === "ordering-out-of-place")).toEqual([]);
  });
});
