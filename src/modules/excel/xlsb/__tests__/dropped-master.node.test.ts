import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { extractAll } from "@archive/unzip/extract";
import { Cell, Stream, Workbook } from "@excel/index";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

/**
 * A shared or array master the writer drops (written as its cached value under `unsupported: "ignore"`) must not
 * leave followers behind: a `PtgExp` naming a master with no `BrtArrFmla`/`BrtShrFmla` makes Excel refuse the file.
 */
let dir: string;

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "documonster-dropped-master-"));
});

afterEach(async () => {
  await rm(dir, { recursive: true, force: true });
});

async function cellRecords(bytes: Uint8Array): Promise<string[]> {
  const parts = await extractAll(bytes);
  const out: string[] = [];
  for (const record of iterateInterpretableRecords(
    parts.get("xl/worksheets/sheet1.bin")!.data,
    "ws"
  )) {
    const name = recordSpec(record.id)?.name ?? String(record.id);
    if (/^Brt(Cell|Fmla|ArrFmla|ShrFmla)/.test(name)) {
      // The column, the first four bytes of every cell record, identifies the cell within the row.
      const column = new DataView(record.payload.buffer, record.payload.byteOffset).getUint32(
        0,
        true
      );
      out.push(
        name.startsWith("BrtCell") || name.startsWith("BrtFmla") ? `${name}@${column}` : name
      );
    }
  }
  return out;
}

describe("xlsb: a dropped shared or array master", () => {
  it("streamed: an array master refused for a sheet not yet written leaves no PtgExp followers", async () => {
    const file = join(dir, "s.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    const main = writer.addWorksheet("Main");
    for (let row = 1; row <= 3; row++) {
      Stream.setCellValue(main.getCell(`B${row}`), row);
    }
    Stream.setCellValue(main.getCell("C1"), {
      formula: "Data!A1:A3*10",
      shareType: "array",
      ref: "C1:C3",
      result: 10
    });
    main.commit();
    const data = writer.addWorksheet("Data");
    for (let row = 1; row <= 3; row++) {
      Stream.setCellValue(data.getCell(`A${row}`), row);
    }
    data.commit();
    await writer.commit();

    const records = await cellRecords(await readFile(file));
    expect(records.filter(name => name.startsWith("BrtFmla"))).toEqual([]);
    expect(records).not.toContain("BrtArrFmla");
    // C2 and C3 are still there, as values.
    expect(records.filter(name => name.endsWith("@2"))).toHaveLength(3);
  });

  it.each([
    ["array", "Nope!A1*10"],
    ["shared", "Nope!A1*10"],
    // Not another sheet, so the shared formula is kept shared up to the encoder, which refuses it.
    ["shared", "SUM(Missing[Qty])"],
    // Withheld before the encoder: an unnamed this-row reference outside every table.
    ["shared", "[@x]*2"]
  ] as const)(
    "buffered: a refused %s master (%s) leaves its followers without PtgExp",
    async (shareType, formula) => {
      const handle = Workbook.create();
      const ws = Workbook.addWorksheet(handle, "Main");
      for (let row = 1; row <= 3; row++) {
        Cell.setValue(ws, `B${row}`, row);
      }
      Cell.setValue(ws, "C1", { formula, shareType, ref: "C1:C3", result: 7 } as never);
      if (shareType === "shared") {
        Cell.setValue(ws, "C2", { sharedFormula: "C1", result: 8 });
        Cell.setValue(ws, "C3", { sharedFormula: "C1", result: 9 });
      }
      const bytes = await Workbook.toBuffer(handle, { format: "xlsb", unsupported: "ignore" });
      const records = await cellRecords(bytes);
      expect(records.filter(name => name.startsWith("BrtFmla"))).toEqual([]);
      expect(records).not.toContain("BrtArrFmla");
      expect(records).not.toContain("BrtShrFmla");
      expect(records.filter(name => name.endsWith("@2"))).toHaveLength(3);

      const back = Workbook.create();
      await Workbook.read(back, bytes);
      const sheet = Workbook.getWorksheet(back, "Main")!;
      expect(Cell.getValue(sheet, "C1")).toBe(7);
      if (shareType === "shared") {
        expect(Cell.getValue(sheet, "C3")).toBe(9);
        // The dependents' formulas are lost too, and said so.
        await expect(Workbook.toBuffer(handle, { format: "xlsb" })).rejects.toThrow(/C3: /);
      }
    }
  );
});

describe("xlsx: a dropped shared master", () => {
  it("leaves no orphan shared-formula reference behind", async () => {
    const handle = Workbook.create();
    const ws = Workbook.addWorksheet(handle, "Main");
    // An unnamed this-row reference outside every table has no file form: the master is written as its value.
    Cell.setValue(ws, "C1", {
      formula: "[@x]*2",
      shareType: "shared",
      ref: "C1:C3",
      result: 7
    } as never);
    Cell.setValue(ws, "C2", { sharedFormula: "C1", result: 8 } as never);
    Cell.setValue(ws, "C3", { sharedFormula: "C1", result: 9 } as never);
    const bytes = await Workbook.toBuffer(handle, { format: "xlsx", unsupported: "ignore" });
    const xml = new TextDecoder().decode(
      (await extractAll(bytes)).get("xl/worksheets/sheet1.xml")!.data
    );
    expect(xml).not.toContain("<f");
    expect(xml).toContain("<v>9</v>");
    await expect(Workbook.toBuffer(handle, { format: "xlsx" })).rejects.toThrow(/C3: /);
  });
});
