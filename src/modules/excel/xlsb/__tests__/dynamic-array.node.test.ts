/**
 * Dynamic arrays in XLSB, against the workbook Excel itself saved in both containers.
 *
 * `data/dynamic-array.excel.{xlsx,xlsb}` hold `=SEQUENCE(3)` in A1 and `=SORT({3;1;2})` in C1, both spilling to
 * row 3, saved by Excel 16.114. Excel's XLSB form is a `BrtCellMeta` naming the `XLDAPR` cell metadata before the
 * master's `BrtFmla*`, the `BrtArrFmla` that follows it, a `BrtFmla*` with a `PtgExp` in every other cell of the
 * spill, and `xl/metadata.bin`.
 *
 * `.node` because it reads fixtures off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { Cell, Workbook } from "@excel";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { readWorkbookPart } from "@excel/xlsb/read/parts";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

const DIRECTORY = "src/modules/excel/xlsb/__tests__/data";

async function load(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(`${DIRECTORY}/${name}`));
}

/**
 * The cell-data records of a sheet part, as `name hex` lines.
 *
 * Three things are normalised, none of them part of a cell: the row-level `BrtACBegin`/`BrtACEnd` wrappers Excel puts
 * around row headers; a `PtgName` index, replaced by the name it resolves to, because the order of `BrtName` in the
 * workbook part is the workbook's business; and the ten bytes of `PtgArray` the specification calls unused, which
 * Excel fills with whatever was in memory (the same mask `operand-classes.node.test.ts` applies).
 */
function sheetDataRecords(bytes: Uint8Array, names: readonly string[]): string[] {
  const lines: string[] = [];
  let inside = false;
  for (const record of iterateInterpretableRecords(bytes, "test")) {
    const name = recordSpec(record.id)?.name ?? String(record.id);
    if (name === "BrtBeginSheetData") {
      inside = true;
      continue;
    }
    if (name === "BrtEndSheetData") {
      break;
    }
    if (!inside || name === "BrtRowHdr" || name === "BrtACBegin" || name === "BrtACEnd") {
      continue;
    }
    let hex = Buffer.from(record.payload ?? []).toString("hex");
    // `PtgName` (0x23) and its value/array forms, then a four-byte index — at a token boundary in these formulas.
    hex = hex.replace(/(23|43|63)(0[1-9a-f])000000/g, (_whole, ptg: string, index: string) => {
      return `${ptg}<${names[parseInt(index, 16) - 1]}>`;
    });
    // `PtgArray` (0x60), four used bytes, then ten unused.
    hex = hex.replace(/60(00000200)([0-9a-f]{20})/g, "60$1....................");
    lines.push(`${name} ${hex}`);
  }
  return lines;
}

/** The workbook's defined names in `BrtName` order, which `PtgName` indexes. */
function namesOf(archive: Map<string, { data: Uint8Array }>): readonly string[] {
  return (
    readWorkbookPart(archive.get("xl/workbook.bin")!.data, "xl/workbook.bin").formulaContext
      .definedNames ?? []
  );
}

async function readModel(bytes: Uint8Array) {
  const workbook = Workbook.create();
  await Workbook.read(workbook, bytes);
  return Workbook.getModel(workbook).worksheets.map(sheet => ({
    name: sheet.name,
    cells: (sheet.rows ?? []).flatMap(row =>
      row.cells.map(cell => ({
        address: cell.address,
        formula: cell.formula,
        result: cell.result,
        value: cell.value,
        shareType: cell.shareType,
        ref: cell.ref,
        isDynamicArray: cell.isDynamicArray
      }))
    )
  }));
}

describe("XLSB dynamic arrays", () => {
  it("reads Excel's XLSB into the same model as Excel's XLSX", async () => {
    const fromXlsx = await readModel(await load("dynamic-array.excel.xlsx"));
    const fromXlsb = await readModel(await load("dynamic-array.excel.xlsb"));
    expect(fromXlsb).toEqual(fromXlsx);
    expect(fromXlsb[0].cells.filter(cell => cell.isDynamicArray === true)).toHaveLength(2);
  });

  it("writes the cell records and the metadata part byte for byte as Excel does", async () => {
    const excel = await extractAll(await load("dynamic-array.excel.xlsb"));
    const workbook = Workbook.create();
    await Workbook.read(workbook, await load("dynamic-array.excel.xlsx"));
    const written = await extractAll(await Workbook.toBuffer(workbook, { format: "xlsb" }));

    const ours = sheetDataRecords(written.get("xl/worksheets/sheet1.bin")!.data, namesOf(written));
    expect(ours).toEqual(
      sheetDataRecords(excel.get("xl/worksheets/sheet1.bin")!.data, namesOf(excel))
    );
    expect(ours.filter(line => line.startsWith("BrtCellMeta"))).toHaveLength(2);
    expect(Buffer.from(written.get("xl/metadata.bin")!.data).toString("hex")).toBe(
      Buffer.from(excel.get("xl/metadata.bin")!.data).toString("hex")
    );
    const contentTypes = new TextDecoder().decode(written.get("[Content_Types].xml")!.data);
    expect(contentTypes).toContain(
      '<Override PartName="/xl/metadata.bin" ContentType="application/vnd.ms-excel.sheetMetadata"/>'
    );
    const rels = new TextDecoder().decode(written.get("xl/_rels/workbook.bin.rels")!.data);
    expect(rels).toMatch(
      /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/sheetMetadata" Target="metadata\.bin"/
    );
  });

  it("round-trips Excel's XLSB without losing the dynamic array or duplicating the metadata part", async () => {
    const original = await load("dynamic-array.excel.xlsb");
    const workbook = Workbook.create();
    await Workbook.read(workbook, original);
    // An edit, so the package is rebuilt rather than handed back as the bytes it arrived as.
    Cell.setValue(Workbook.getWorksheet(workbook, "S")!, "E5", 1);
    const bytes = await Workbook.toBuffer(workbook, { format: "xlsb" });
    const entries = await extractAll(bytes);
    expect([...entries.keys()].filter(path => /metadata/i.test(path))).toEqual(["xl/metadata.bin"]);
    const rebuilt = await readModel(bytes);
    expect(rebuilt[0]!.cells.filter(cell => cell.address !== "E5")).toEqual(
      (await readModel(original))[0]!.cells
    );
  });

  it("writes no metadata part for a workbook without dynamic arrays", async () => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    void sheet;
    const entries = await extractAll(await Workbook.toBuffer(workbook, { format: "xlsb" }));
    expect(entries.has("xl/metadata.bin")).toBe(false);
  });
});
