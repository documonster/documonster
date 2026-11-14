import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { extractAll } from "@archive/unzip/extract";
import { Cell, DefinedNames, Stream, Workbook } from "@excel/index";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { createWorkbookPartReader, readWorkbookPart } from "@excel/xlsb/read/parts";
import { recordSpec } from "@excel/xlsb/spec/records";
import { runSteps } from "@utils/yield";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

/**
 * A streamed `xl/workbook.bin` must be the buffered one.
 *
 * Both are produced by `writeWorkbookPart`, but the externals block is written only when a formula resolved an `ixti`
 * through the table the workbook part carries. The streamed rows resolved theirs through a table of their own, so the
 * part saw an unused table and left the block out — while the cells still held `PtgRef3d` indices into it. Excel
 * refuses such a file outright.
 */
let dir: string;

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "documonster-sxwb-"));
});

afterEach(async () => {
  await rm(dir, { recursive: true, force: true });
});

function workbookRecords(parts: Map<string, { data: Uint8Array }>): string[] {
  const out: string[] = [];
  for (const record of iterateInterpretableRecords(parts.get("xl/workbook.bin")!.data, "wb")) {
    const payload = [...record.payload].map(byte => byte.toString(16).padStart(2, "0")).join("");
    out.push(`${recordSpec(record.id)?.name ?? record.id} ${payload}`);
  }
  return out;
}

type Formulas = Record<string, string>;

const CASES: readonly (readonly [string, Formulas, boolean])[] = [
  ["a cross-sheet reference", { A1: "SUM(Data!A1:A10)" }, false],
  ["a cross-sheet future function", { H2: "_xlfn._xlws.SORT(Data!A1:A10)" }, false],
  ["a cross-sheet UNIQUE", { L2: "_xlfn.UNIQUE(Data!A1:A10)" }, false],
  ["a span of sheets", { A1: "SUM(Data:Main!A1)" }, false],
  ["a defined name", { A1: "Rate*2" }, true],
  ["an eta-reduced function", { E1: "_xlfn.GROUPBY(B2:B6,C2:C6,_xleta.SUM)" }, false],
  ["a local formula only", { A1: "SUM(B2:B6)" }, false]
];

describe("streamed XLSB workbook part", () => {
  it.each(CASES)("matches the buffered one for %s", async (_label, formulas, withName) => {
    const file = join(dir, "s.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    if (withName) {
      DefinedNames.addFormula(writer.definedNames, "Rate", "0.2");
    }
    const data = writer.addWorksheet("Data");
    for (let row = 1; row <= 10; row++) {
      Stream.setCellValue(data.getCell(`A${row}`), row);
    }
    data.commit();
    const main = writer.addWorksheet("Main");
    for (let row = 2; row <= 6; row++) {
      Stream.setCellValue(main.getCell(`B${row}`), row);
      Stream.setCellValue(main.getCell(`C${row}`), row * 2);
    }
    for (const [address, formula] of Object.entries(formulas)) {
      Stream.setCellValue(main.getCell(address), { formula });
    }
    main.commit();
    await writer.commit();

    const handle = Workbook.create();
    if (withName) {
      DefinedNames.addFormula(Workbook.getDefinedNames(handle), "Rate", "0.2");
    }
    const bufferedData = Workbook.addWorksheet(handle, "Data");
    for (let row = 1; row <= 10; row++) {
      Cell.setValue(bufferedData, `A${row}`, row);
    }
    const bufferedMain = Workbook.addWorksheet(handle, "Main");
    for (let row = 2; row <= 6; row++) {
      Cell.setValue(bufferedMain, `B${row}`, row);
      Cell.setValue(bufferedMain, `C${row}`, row * 2);
    }
    for (const [address, formula] of Object.entries(formulas)) {
      Cell.setValue(bufferedMain, address, { formula });
    }
    const buffered = await Workbook.toBuffer(handle, { format: "xlsb" });

    expect(writer.unsupported).toEqual([]);
    const streamedRecords = workbookRecords(await extractAll(new Uint8Array(await readFile(file))));
    const bufferedRecords = workbookRecords(await extractAll(new Uint8Array(buffered)));
    expect(streamedRecords).toEqual(bufferedRecords);
  });
});

describe("the binary workbook part's defined names", () => {
  // Decoded once the whole part is in — a name may use a later one — so a workbook defining many would decode them all
  // in one call. They are decoded in steps instead, to what the buffered reader reads.
  it("are decoded in steps, to what a whole read gives", async () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    const names = Workbook.getDefinedNames(wb);
    for (let i = 0; i < 1_000; i++) {
      DefinedNames.add(names, `S!$A$${i + 1}`, `name_${i}`);
    }
    const parts = await extractAll(await Workbook.toBuffer(wb, { format: "xlsb" }));
    const part = parts.get("xl/workbook.bin")!.data;
    const reader = createWorkbookPartReader("wb");
    for (const record of iterateInterpretableRecords(part, "wb")) {
      reader.add(record);
    }
    const steps = reader.finish();
    let taken = 0;
    let step = steps.next();
    while (!step.done) {
      taken++;
      step = steps.next();
    }
    expect(taken).toBeGreaterThan(1);
    expect(step.value).toEqual(readWorkbookPart(part, "wb"));
    expect(step.value.names).toHaveLength(1_000);
  });
});

describe("the binary workbook part's external sheet table", () => {
  // One `BrtExternSheet` record holds the whole table, and its entries were appended with a spread: a table of a few
  // hundred thousand entries — a 3.6 MB record — exceeded the call stack.
  it("is read however many entries its record holds", () => {
    const id = Array.from({ length: 0x4000 }, (_, i) => i).find(
      i => recordSpec(i)?.name === "BrtExternSheet"
    )!;
    const count = 300_000;
    const payload = new Uint8Array(4 + count * 12);
    new DataView(payload.buffer).setUint32(0, count, true);
    const reader = createWorkbookPartReader("wb");
    reader.add({ id, payload });
    expect(() => runSteps(reader.finish())).not.toThrow();
  });
});
