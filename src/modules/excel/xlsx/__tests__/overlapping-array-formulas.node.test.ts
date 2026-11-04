import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { extractAll } from "@archive/unzip/extract";
import { Cell, Stream, Workbook } from "@excel/index";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

/**
 * An array formula whose master sits inside an earlier array's range.
 *
 * The cell would be both a member of the outer array and the master of another, and Excel refuses the file. The XLSB
 * writer reported it and kept the outer range; the XLSX writers wrote both `<f t="array">` silently. They now apply
 * the same rule: report the inner formula, and with `unsupported: "ignore"` write its cached result alone.
 */
let dir: string;

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "documonster-xlsx-overlap-"));
});

afterEach(async () => {
  await rm(dir, { recursive: true, force: true });
});

const OUTER = { formula: "B2:B5*2", shareType: "array", ref: "D2:D5", result: 2 } as const;
const INNER = { formula: "SUM(B2:B4)", shareType: "array", ref: "D4:E4", result: 6 } as const;
const LOSS = "S!D4: array formula (overlaps the array range D2:D5 of D2)";

function overlapping(isDynamicArray: boolean) {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S");
  [1, 2, 3, 4].forEach((value, index) => Cell.setValue(sheet, `B${index + 2}`, value));
  Cell.setValue(sheet, "D2", { ...OUTER, isDynamicArray } as never);
  Cell.setValue(sheet, "D4", { ...INNER, isDynamicArray } as never);
  return workbook;
}

async function sheetXml(bytes: Uint8Array): Promise<string> {
  const parts = await extractAll(bytes);
  return new TextDecoder().decode(parts.get("xl/worksheets/sheet1.xml")!.data);
}

function expectOuterKept(xml: string): void {
  expect(xml.match(/t="array"/g)).toHaveLength(1);
  expect(xml).toMatch(/<c r="D2"[^>]*><f t="array" ref="D2:D5">B2:B5\*2<\/f>/);
  // The inner master is its cached result, with no formula and no dynamic-array metadata.
  expect(xml).toMatch(/<c r="D4"><v>6<\/v><\/c>/);
}

describe("overlapping array formulas in XLSX", () => {
  it.each([false, true])("refuses them by default (dynamic: %s)", async isDynamicArray => {
    await expect(
      Workbook.toBuffer(overlapping(isDynamicArray), { format: "xlsx" })
    ).rejects.toThrow(LOSS);
  });

  it.each([false, true])(
    "keeps the outer range when ignored (dynamic: %s)",
    async isDynamicArray => {
      const bytes = await Workbook.toBuffer(overlapping(isDynamicArray), {
        format: "xlsx",
        unsupported: "ignore"
      });
      expectOuterKept(await sheetXml(bytes));
      const read = Workbook.create();
      await Workbook.read(read, bytes);
      expect(Cell.getValue(Workbook.getWorksheet(read, "S")!, "D4")).toBe(6);
    }
  );

  it("reports them from the streaming writer and writes the cached result", async () => {
    const file = join(dir, "overlap.xlsx");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsx" });
    const sheet = writer.addWorksheet("S");
    [1, 2, 3, 4].forEach((value, index) =>
      Stream.setCellValue(sheet.getCell(`B${index + 2}`), value)
    );
    Stream.setCellValue(sheet.getCell("D2"), OUTER as never);
    Stream.setCellValue(sheet.getCell("D4"), INNER as never);
    sheet.commit();
    await writer.commit();
    expect(writer.unsupported).toEqual([LOSS]);
    expectOuterKept(await sheetXml(new Uint8Array(await readFile(file))));
  });
});

/**
 * Any two array ranges that intersect, not only a master inside an earlier range. The later array loses: its master is
 * written as its cached result, it gets no followers, and the earlier range keeps every one of its cells.
 */
describe("intersecting array ranges, in every writer", () => {
  const CASES = [
    // A range crossing an earlier one, its master outside it.
    ["a crossing range", "D2", "D2:D5", "C3", "C3:E3", ["C3", "E3"]],
    // Partial overlap of two blocks.
    ["a partial overlap", "D2", "D2:E3", "C3", "C3:D4", ["C3", "C4"]],
    // A later range covering the earlier master.
    ["a range covering an earlier master", "D2", "D2:D3", "E3", "C2:E3", ["E3", "C2"]]
  ] as const;

  async function write(
    format: "xlsx" | "xlsb",
    streamed: boolean,
    outer: readonly [string, string],
    inner: readonly [string, string]
  ): Promise<{ bytes: Uint8Array; unsupported: readonly string[] }> {
    const outerValue = { formula: "B2:B5*2", shareType: "array", ref: outer[1], result: 2 };
    const innerValue = { formula: "B2:B4+100", shareType: "array", ref: inner[1], result: 101 };
    if (!streamed) {
      const workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "S");
      [1, 2, 3, 4].forEach((value, index) => Cell.setValue(sheet, `B${index + 2}`, value));
      Cell.setValue(sheet, outer[0], outerValue as never);
      Cell.setValue(sheet, inner[0], innerValue as never);
      await expect(Workbook.toBuffer(workbook, { format })).rejects.toThrow(
        `S!${inner[0]}: array formula (overlaps the array range ${outer[1]} of ${outer[0]})`
      );
      return {
        bytes: await Workbook.toBuffer(workbook, { format, unsupported: "ignore" }),
        unsupported: [
          `S!${inner[0]}: array formula (overlaps the array range ${outer[1]} of ${outer[0]})`
        ]
      };
    }
    const file = join(dir, `intersect.${format}`);
    const writer = new Stream.WorkbookWriter({ filename: file, format });
    const sheet = writer.addWorksheet("S");
    [1, 2, 3, 4].forEach((value, index) =>
      Stream.setCellValue(sheet.getCell(`B${index + 2}`), value)
    );
    // Every case's earlier master is in an earlier row, so this is row-major order.
    Stream.setCellValue(sheet.getCell(outer[0]), outerValue as never);
    Stream.setCellValue(sheet.getCell(inner[0]), innerValue as never);
    sheet.commit();
    await writer.commit();
    return { bytes: new Uint8Array(await readFile(file)), unsupported: writer.unsupported };
  }

  describe.each([
    ["xlsx", false],
    ["xlsx", true],
    ["xlsb", false],
    ["xlsb", true]
  ] as const)("%s (streamed: %s)", (format, streamed) => {
    it.each(CASES)("%s", async (_label, outerAt, outerRef, innerAt, innerRef, innerOnly) => {
      const { bytes, unsupported } = await write(
        format,
        streamed,
        [outerAt, outerRef],
        [innerAt, innerRef]
      );
      expect(unsupported).toEqual([
        `S!${innerAt}: array formula (overlaps the array range ${outerRef} of ${outerAt})`
      ]);
      const read = Workbook.create();
      await Workbook.read(read, bytes);
      const sheet = Workbook.getWorksheet(read, "S")!;
      // The earlier array is whole: its master keeps its range and every cell in it still belongs to it.
      expect(Cell.getValue(sheet, outerAt)).toMatchObject({ shareType: "array", ref: outerRef });
      // The later one is its cached result at its master, and nothing anywhere else.
      expect(Cell.getValue(sheet, innerAt)).toBe(101);
      for (const address of innerOnly) {
        if (address !== innerAt) {
          expect(Cell.getValue(sheet, address)).toBeNull();
        }
      }
      if (format === "xlsx") {
        expect((await sheetXml(bytes)).match(/<f /g)).toHaveLength(1);
      }
    });
  });
});
