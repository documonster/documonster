import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { extractAll } from "@archive/unzip/extract";
import { Cell, DataValidation, DefinedNames, Stream, Workbook } from "@excel/index";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

/**
 * A streamed XLSB has to encode a cross-sheet formula, and has to *say so* when it cannot.
 *
 * Both halves were broken and each hid the other. The streamed sheet writer was given no formula context at all, so
 * `Other!A1` could not resolve its `ixti` and the row encoder wrote a `BrtCellBlank` instead — and the loss it reported
 * was stored on the worksheet in a field nothing read, so the report came back `[]`. Two formulas out of two
 * disappeared from a two-sheet workbook and every check in the repository passed.
 *
 * The context is fetched per sheet rather than captured once, because `addWorksheet` runs between sheet commits: a sheet
 * committed later resolves against more sheets than an earlier one. A reference *forward* — sheet 1 naming sheet 3 —
 * therefore still cannot resolve, which is inherent to writing forward. The test for that case asserts the report, not
 * the absence of one.
 */
let dir: string;

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "documonster-sxf-"));
});

afterEach(async () => {
  await rm(dir, { recursive: true, force: true });
});

async function cellRecords(path: string, sheet: string): Promise<readonly string[]> {
  const parts = await extractAll(new Uint8Array(await readFile(path)));
  const target = [...parts.keys()].find(name => name.endsWith(`worksheets/${sheet}`));
  expect(target, `${sheet} missing from package`).toBeDefined();
  const out: string[] = [];
  for (const record of iterateInterpretableRecords(parts.get(target!)!.data, "s")) {
    const name = recordSpec(record.id)?.name;
    if (name?.startsWith("BrtFmla") === true || name === "BrtCellBlank" || name === "BrtCellRk") {
      out.push(name);
    }
  }
  return out;
}

describe("streamed XLSB formulas", () => {
  it("encodes a reference to a sheet already committed", async () => {
    const file = join(dir, "back.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    const other = writer.addWorksheet("Other");
    Stream.setCellValue(other.getCell("A1"), 42);
    other.commit();
    const main = writer.addWorksheet("Main");
    Stream.setCellValue(main.getCell("A1"), { formula: "Other!A1" });
    Stream.setCellValue(main.getCell("B1"), { formula: "SUM(Other!A1:A1)" });
    main.commit();
    await writer.commit();

    // Both are real formula records, not blanks.
    expect(await cellRecords(file, "sheet2.bin")).toEqual(["BrtFmlaNum", "BrtFmlaNum"]);
    expect(writer.unsupported).toEqual([]);
    // And they name the right sheet. The context's sheet list kept the hole at slot 0 of the id-indexed sheet array,
    // so every `ixti` was one too high and `Other!A1` read back as `Main!A1`.
    const handle = Workbook.create();
    await Workbook.read(handle, new Uint8Array(await readFile(file)));
    const read = Workbook.getWorksheet(handle, "Main")!;
    expect(Cell.getValue(read, "A1")).toMatchObject({ formula: "Other!A1" });
    expect(Cell.getValue(read, "B1")).toMatchObject({ formula: "SUM(Other!A1:A1)" });
  });

  /**
   * A streamed sheet holds no table, so an unnamed this-row reference in it is outside any table: reported as the
   * buffered writers report it, and written with its cached value rather than as a formula Excel reads as `#REF!`.
   */
  it.each([
    ["xlsx", "[@x]*2"],
    ["xlsx", "SUM([@])"],
    ["xlsb", "[[#This Row],[x]]+1"],
    ["xlsb", "[@[x]:[y]]"]
  ] as const)(
    "reports an unnamed this-row reference in a streamed %s cell: %s",
    async (format, formula) => {
      const file = join(dir, `this-row.${format}`);
      const writer = new Stream.WorkbookWriter({ filename: file, format });
      const sheet = writer.addWorksheet("S");
      Stream.setCellValue(sheet.getCell("A1"), { formula, result: 4 } as never);
      Stream.setCellValue(sheet.getCell("A2"), { formula: "1+1", result: 2 } as never);
      sheet.commit();
      await writer.commit();
      // Named by sheet, as the buffered writers name it.
      expect(writer.unsupported).toEqual([
        "S!A1: formula (an unnamed this-row reference outside any table)"
      ]);
      // The deprecated name returns what it always did: the XLSB losses only, a row's as `A1: …`.
      expect(writer.xlsbUnsupported).toEqual(
        format === "xlsb" ? ["A1: formula (an unnamed this-row reference outside any table)"] : []
      );
      const handle = Workbook.create();
      await Workbook.read(handle, new Uint8Array(await readFile(file)));
      const read = Workbook.getWorksheet(handle, "S")!;
      expect(Cell.getValue(read, "A1")).toBe(4);
      expect(Cell.getFormula(read, "A2")).toBe("1+1");
    }
  );

  /**
   * A constant where only a reference is accepted makes a workbook Excel will not open. Streamed, it is reported and
   * left out — a cell keeps its cached value — as the buffered writers do; nothing can be refused once rows are out.
   */
  it.each(["xlsx", "xlsb"] as const)(
    "reports a non-reference in a reference-only parameter (streamed %s)",
    async format => {
      const file = join(dir, `reference-only.${format}`);
      const writer = new Stream.WorkbookWriter({ filename: file, format });
      DefinedNames.addFormula(writer.definedNames, "Bad", "OFFSET(1,0,0)");
      const sheet = writer.addWorksheet("S");
      Stream.setCellValue(sheet.getCell("A1"), {
        formula: "RANK(2,{2,2,1},0)",
        result: 3
      } as never);
      Stream.setCellValue(sheet.getCell("A2"), { formula: "IF(1,2,3)", result: 2 } as never);
      sheet.addConditionalFormatting({
        ref: "A1:A2",
        rules: [
          {
            type: "expression",
            formulae: ["COUNTIF(5,A1)>0"],
            style: { font: { bold: true } }
          } as never
        ]
      });
      DataValidation.add(sheet.dataValidations, "B1", { type: "custom", formulae: ["ROW(1)>0"] });
      sheet.commit();
      await writer.commit();
      const reported = writer.unsupported.join("\n");
      expect(reported).toContain("S!A1: formula (RANK argument 2 must be a reference)");
      expect(reported).toMatch(
        format === "xlsx"
          ? /S!A1:A2: conditional formatting formula \(COUNTIF argument 1/
          : /S!conditional formatting formula "COUNTIF\(5,A1\)>0"/
      );
      expect(reported).toMatch(/S!B1: data validation/);
      expect(reported).toMatch(/Bad: defined name/);
      const handle = Workbook.create();
      await Workbook.read(handle, new Uint8Array(await readFile(file)));
      const read = Workbook.getWorksheet(handle, "S")!;
      expect(Cell.getValue(read, "A1")).toBe(3);
      expect(Cell.getFormula(read, "A2")).toBe("IF(1,2,3)");
      const model = Workbook.getModel(handle);
      expect(model.definedNames.map(entry => entry.name)).not.toContain("Bad");
      expect(model.worksheets[0]!.conditionalFormattings ?? []).toEqual([]);
      expect(Object.keys(model.worksheets[0]!.dataValidations ?? {})).toEqual([]);
    }
  );

  it("writes a dynamic array with the metadata part its BrtCellMeta points into", async () => {
    // The metadata part is written by the package writer after the rows have gone, so the streamed rows have to say
    // that they needed it.
    const file = join(dir, "dynamic.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    const sheet = writer.addWorksheet("S");
    Stream.setCellValue(sheet.getCell("A1"), {
      formula: "SEQUENCE(3)",
      shareType: "array",
      ref: "A1:A3",
      isDynamicArray: true,
      result: 1
    } as never);
    sheet.commit();
    await writer.commit();
    expect(writer.unsupported).toEqual([]);
    const bytes = new Uint8Array(await readFile(file));
    expect((await extractAll(bytes)).has("xl/metadata.bin")).toBe(true);
    const handle = Workbook.create();
    await Workbook.read(handle, bytes);
    const cell = Workbook.getModel(handle).worksheets[0]!.rows![0]!.cells[0]!;
    expect(cell).toMatchObject({
      formula: "_xlfn.SEQUENCE(3)",
      isDynamicArray: true,
      ref: "A1:A3"
    });
  });

  it("reports a master inside an earlier array range, committed in a later row", async () => {
    // The ranges have to be remembered across rows: the master that claims D4 arrived two rows earlier.
    const file = join(dir, "overlap.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    const sheet = writer.addWorksheet("S");
    Stream.setCellValue(sheet.getCell("D2"), {
      formula: "B2:B5*2",
      shareType: "array",
      ref: "D2:D5",
      result: 0
    } as never);
    Stream.setCellValue(sheet.getCell("D4"), {
      formula: "SUM(B2:B4)",
      shareType: "array",
      ref: "D4",
      result: 0
    } as never);
    sheet.commit();
    await writer.commit();
    expect(writer.unsupported).toEqual([
      "S!D4: array formula (overlaps the array range D2:D5 of D2)"
    ]);
  });

  /**
   * An array range spanning later rows. The followers were emitted with the master's row, so each later row went out
   * early and then again when it was committed: `BrtRowHdr` out of order and twice, which MS-XLSB does not allow. They
   * are now held until their row is written, or the sheet ends.
   */
  describe("an array range over later rows", () => {
    async function rowOrder(path: string): Promise<number[]> {
      const parts = await extractAll(new Uint8Array(await readFile(path)));
      const sheet = [...parts.keys()].find(name => name.endsWith("worksheets/sheet1.bin"))!;
      return [...iterateInterpretableRecords(parts.get(sheet)!.data, "s")]
        .filter(record => recordSpec(record.id)?.name === "BrtRowHdr")
        .map(record =>
          new DataView(record.payload.buffer, record.payload.byteOffset).getUint32(0, true)
        );
    }

    function fill(set: (address: string, value: unknown) => void, extraRow: boolean): void {
      set("D2", { formula: "B2:B5*2", shareType: "array", ref: "D2:D5", result: 0 });
      if (extraRow) {
        set("B3", 5);
      }
    }

    async function cellsOf(bytes: Uint8Array): Promise<unknown> {
      const handle = Workbook.create();
      await Workbook.read(handle, bytes);
      return Workbook.getModel(handle).worksheets[0]!.rows!.map(row =>
        row.cells.map(cell => [
          cell.address,
          cell.value,
          cell.formula,
          cell.sharedFormula,
          cell.ref
        ])
      );
    }

    it.each([
      ["with a row committed inside it", true],
      ["past the last committed row", false]
    ])("writes each row once, in order, %s", async (_, extraRow) => {
      const file = join(dir, "array-rows.xlsb");
      const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
      const sheet = writer.addWorksheet("S");
      fill(
        (address, value) => Stream.setCellValue(sheet.getCell(address), value as never),
        extraRow
      );
      sheet.commit();
      await writer.commit();
      expect(writer.unsupported).toEqual([]);
      expect(await rowOrder(file)).toEqual([1, 2, 3, 4]);

      const buffered = Workbook.create();
      const target = Workbook.addWorksheet(buffered, "S");
      fill((address, value) => Cell.setValue(target, address, value as never), extraRow);
      const streamed = new Uint8Array(await readFile(file));
      expect(await cellsOf(streamed)).toEqual(
        await cellsOf(await Workbook.toBuffer(buffered, { format: "xlsb" }))
      );
    });
  });

  it("reports a forward reference instead of writing a blank cell for it", async () => {
    // The honest limit of writing forward. What must not happen is silence.
    const file = join(dir, "forward.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    const main = writer.addWorksheet("Main");
    Stream.setCellValue(main.getCell("A1"), { formula: "Later!A1" });
    main.commit();
    const later = writer.addWorksheet("Later");
    Stream.setCellValue(later.getCell("A1"), 7);
    later.commit();
    await writer.commit();

    expect(writer.unsupported).toHaveLength(1);
    expect(writer.unsupported[0]).toContain("A1");
    expect(writer.unsupported[0]).toContain("formula");
  });

  it("surfaces a row-level loss at all", async () => {
    // The field the row encoder writes into had no reader, so this is the assertion that the wiring exists — a report
    // that is computed and discarded reads as a guarantee.
    const file = join(dir, "loss.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    const main = writer.addWorksheet("Main");
    Stream.setCellValue(main.getCell("A1"), { formula: "Missing!A1" });
    main.commit();
    await writer.commit();
    expect(writer.unsupported.length).toBeGreaterThan(0);
  });

  it("keeps a same-sheet formula working", async () => {
    // Guards against a context change breaking the ordinary case.
    const file = join(dir, "local.xlsb");
    const writer = new Stream.WorkbookWriter({ filename: file, format: "xlsb" });
    const sheet = writer.addWorksheet("Only");
    Stream.setCellValue(sheet.getCell("A1"), 2);
    Stream.setCellValue(sheet.getCell("A2"), 3);
    Stream.setCellValue(sheet.getCell("A3"), { formula: "SUM(A1:A2)" });
    sheet.commit();
    await writer.commit();
    expect(await cellRecords(file, "sheet1.bin")).toEqual(["BrtCellRk", "BrtCellRk", "BrtFmlaNum"]);
    expect(writer.unsupported).toEqual([]);
  });
});
