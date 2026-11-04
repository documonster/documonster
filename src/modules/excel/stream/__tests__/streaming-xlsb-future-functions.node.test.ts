import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { extractAll } from "@archive/unzip/extract";
import { Cell, DefinedNames, Stream, Workbook } from "@excel/index";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

/**
 * A streamed XLSB has to write the functions the `Ftab` has no id for, exactly as the buffered writer does.
 *
 * Such a call is a `PtgName` naming a hidden `_xlfn.` stub, and a `LET`/`LAMBDA` parameter is a `PtgName` naming an
 * `_xlpm.` name. The buffered writer collects those names by scanning every formula before any is encoded; a streamed
 * row is encoded before the workbook part exists, so the names it needs used to be missing from the context and every
 * such formula was reported unsupported and written as a blank cell.
 */
let dir: string;

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "documonster-sxff-"));
});

afterEach(async () => {
  await rm(dir, { recursive: true, force: true });
});

type Value =
  | number
  | { formula: string }
  | { formula: string; shareType: "shared"; ref: string }
  | {
      sharedFormula: string;
    };

const CELLS: readonly (readonly [string, Value])[] = [
  ["A1", 1],
  ["A2", 2],
  ["A3", 3],
  ["B1", 10],
  ["B2", 20],
  ["B3", 30],
  ["C1", { formula: "XLOOKUP(2,A1:A3,B1:B3)" }],
  ["C2", { formula: "SUM(FILTER(B1:B3,A1:A3>1))" }],
  ["C3", { formula: "ROWS(SEQUENCE(5))" }],
  ["C4", { formula: "LET(x,A1,y,A2,x+y)" }],
  ["C5", { formula: "LET(f,LAMBDA(n,n*2),f(A3))" }],
  ["C6", { formula: 'TEXTJOIN(",",TRUE,A1:A3)' }],
  ["C7", { formula: "_xlfn.CONCAT(A1,B1)" }],
  ["C8", { formula: "SUM(A1:A3)" }],
  [
    "D1",
    { formula: "IFS(A1>0,XLOOKUP(A1,A1:A3,B1:B3),TRUE,0)", shareType: "shared", ref: "D1:D3" }
  ],
  ["D2", { sharedFormula: "D1" }],
  ["D3", { sharedFormula: "D1" }]
];

async function streamed(file: string): Promise<readonly string[]> {
  const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
  const sheet = writer.addWorksheet("S");
  for (const [address, value] of CELLS) {
    Stream.setCellValue(sheet.getCell(address), value as never);
  }
  sheet.commit();
  // A second sheet whose rows introduce names the first did not, after the first sheet's part is closed.
  const second = writer.addWorksheet("T");
  Stream.setCellValue(second.getCell("A1"), { formula: "LET(z,S!A1,z*3)" } as never);
  Stream.setCellValue(second.getCell("A2"), { formula: 'MAXIFS(S!B1:B3,S!A1:A3,">1")' } as never);
  second.commit();
  await writer.commit();
  return writer.unsupported;
}

async function buffered(): Promise<Uint8Array> {
  const handle = Workbook.create();
  const sheet = Workbook.addWorksheet(handle, "S");
  for (const [address, value] of CELLS) {
    Cell.setValue(sheet, address, value as never);
  }
  const second = Workbook.addWorksheet(handle, "T");
  Cell.setValue(second, "A1", { formula: "LET(z,S!A1,z*3)" } as never);
  Cell.setValue(second, "A2", { formula: 'MAXIFS(S!B1:B3,S!A1:A3,">1")' } as never);
  return Workbook.toBuffer(handle, { format: "xlsb" });
}

async function formulas(bytes: Uint8Array): Promise<Record<string, unknown>> {
  const handle = Workbook.create();
  await Workbook.read(handle, bytes);
  const out: Record<string, unknown> = {};
  for (const name of ["S", "T"]) {
    const sheet = Workbook.getWorksheet(handle, name)!;
    for (const address of [...CELLS.map(([a]) => a), "A1", "A2"]) {
      out[`${name}!${address}`] = Cell.getValue(sheet, address);
    }
  }
  return out;
}

describe("streamed XLSB future functions", () => {
  it("writes them, and reads back as the buffered writer's output does", async () => {
    const file = join(dir, "f.xlsb");
    const losses = await streamed(file);
    expect(losses).toEqual([]);
    const streamedBytes = new Uint8Array(await readFile(file));
    const bufferedBytes = await buffered();
    const fromStream = await formulas(streamedBytes);
    expect(fromStream).toEqual(await formulas(bufferedBytes));
    expect(fromStream["S!C1"]).toMatchObject({ formula: expect.stringContaining("XLOOKUP") });
    expect(fromStream["S!C4"]).toMatchObject({ formula: expect.stringContaining("LET") });
    expect(fromStream["T!A1"]).toMatchObject({ formula: expect.stringContaining("LET") });
  });

  it("writes the same workbook part as the buffered writer for a single sheet", async () => {
    const file = join(dir, "one.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    const sheet = writer.addWorksheet("S");
    Stream.setCellValue(sheet.getCell("A1"), 1);
    Stream.setCellValue(sheet.getCell("B1"), { formula: "XLOOKUP(1,A1:A1,A1:A1)" } as never);
    Stream.setCellValue(sheet.getCell("C1"), { formula: "FILTER(A1:A1,A1:A1>0)" } as never);
    sheet.commit();
    await writer.commit();
    expect(writer.unsupported).toEqual([]);

    const handle = Workbook.create();
    const ws = Workbook.addWorksheet(handle, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", { formula: "XLOOKUP(1,A1:A1,A1:A1)" } as never);
    Cell.setValue(ws, "C1", { formula: "FILTER(A1:A1,A1:A1>0)" } as never);
    const reference = await Workbook.toBuffer(handle, { format: "xlsb" });

    const a = await extractAll(new Uint8Array(await readFile(file)));
    const b = await extractAll(reference);
    expect([...a.get("xl/workbook.bin")!.data]).toEqual([...b.get("xl/workbook.bin")!.data]);
  });

  it("places the stubs after the workbook's own names and its print names", async () => {
    // `PtgName` is an index, so a stub counted without the names ahead of it points at one of them instead.
    const file = join(dir, "prefix.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    DefinedNames.addFormula(writer.definedNames, "Rate", "0.2");
    const sheet = writer.addWorksheet("S");
    sheet.pageSetup.printArea = "A1:B2";
    Stream.setCellValue(sheet.getCell("A1"), 1);
    Stream.setCellValue(sheet.getCell("B1"), { formula: "XLOOKUP(1,A1:A1,A1:A1)*Rate" } as never);
    Stream.setCellValue(sheet.getCell("C1"), { formula: "LET(x,Rate,x*2)" } as never);
    sheet.commit();
    await writer.commit();
    expect(writer.unsupported).toEqual([]);

    const handle = Workbook.create();
    await Workbook.read(handle, new Uint8Array(await readFile(file)));
    const ws = Workbook.getWorksheet(handle, "S")!;
    expect(Cell.getValue(ws, "B1")).toMatchObject({ formula: "_xlfn.XLOOKUP(1,A1:A1,A1:A1)*Rate" });
    expect(Cell.getValue(ws, "C1")).toMatchObject({ formula: "_xlfn.LET(_xlpm.x,Rate,_xlpm.x*2)" });
  });

  it("reports a defined name added after a formula needing a stub was committed", async () => {
    // The committed row's index already counts the names ahead of the stub; a name added later shifts the stub.
    const file = join(dir, "late.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    const sheet = writer.addWorksheet("S");
    Stream.setCellValue(sheet.getCell("A1"), { formula: "SEQUENCE(3)" } as never);
    sheet.commit();
    DefinedNames.addFormula(writer.definedNames, "Late", "1");
    await writer.commit();
    expect(writer.unsupported.some(loss => loss.includes("defined name"))).toBe(true);
  });
});
