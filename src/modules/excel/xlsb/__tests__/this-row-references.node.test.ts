/**
 * Unnamed this-row references (`[@x]`), against the workbook Excel saved in both containers.
 *
 * `data/table-this-row.excel.{xlsx,xlsb}` hold table `Tb` (x, y) over A1:D4 with two calculated columns Excel typed as
 * `=[@x]+[@y]` and `=Tb[@x]*2`. Excel stores both — in the cells and in the table part — as
 * `Tb[[#This Row],[x]]…`, naming the table; the unnamed form is refused in a `calculatedColumnFormula` and read as
 * `#REF!` in a cell. In XLSB the reference is a `PtgList` with flags `0x0641`: one column, row type `#This Row`,
 * value class, and bit 9 (`unused` in MS-XLSB 2.5.98.52) set, which Excel sets on every this-row `PtgList`.
 *
 * `.node` because it reads fixtures off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { Cell, Table, Workbook } from "@excel";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

const DIRECTORY = "src/modules/excel/xlsb/__tests__/data";

async function load(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(`${DIRECTORY}/${name}`));
}

/** The workbook `mk-sample-b` had Excel build, authored here with unnamed this-row references. */
function authored(): ReturnType<typeof Workbook.create> {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "Sheet1");
  const rows = [
    [1, 10],
    [2, 20],
    [3, 30]
  ];
  Table.add(sheet, {
    name: "Tb",
    ref: "A1",
    headerRow: true,
    columns: [
      { name: "x" },
      { name: "y" },
      { name: "s", calculatedColumnFormula: "[@x]+[@y]" },
      { name: "d", calculatedColumnFormula: "Tb[@x]*2" }
    ],
    rows: rows.map(([x, y]) => [
      x,
      y,
      { formula: "[@x]+[@y]", result: x + y },
      { formula: "Tb[@x]*2", result: x * 2 }
    ])
  });
  return workbook;
}

function records(bytes: Uint8Array, wanted: readonly string[]): string[] {
  return [...iterateInterpretableRecords(bytes, "test")]
    .map(record => ({ name: recordSpec(record.id)?.name ?? "", payload: record.payload }))
    .filter(record => wanted.includes(record.name))
    .map(record => `${record.name} ${Buffer.from(record.payload ?? []).toString("hex")}`);
}

async function formulas(bytes: Uint8Array) {
  const workbook = Workbook.create();
  await Workbook.read(workbook, bytes);
  const model = Workbook.getModel(workbook).worksheets[0]!;
  return {
    cells: (model.rows ?? []).flatMap(row =>
      row.cells.filter(cell => cell.formula !== undefined).map(cell => [cell.address, cell.formula])
    ),
    columns: (model.tables ?? []).flatMap(table =>
      table.columns.map(column => [column.name, column.calculatedColumnFormula])
    )
  };
}

describe("unnamed this-row references", () => {
  it("are written to XLSX with the enclosing table's name, as Excel stores them", async () => {
    const excel = await extractAll(await load("table-this-row.excel.xlsx"));
    const written = await extractAll(await Workbook.toBuffer(authored()));
    const text = (archive: typeof excel, path: string) =>
      new TextDecoder().decode(archive.get(path)!.data);
    const calculated = (xml: string) =>
      [...xml.matchAll(/<calculatedColumnFormula>([^<]*)</g)].map(match => match[1]);
    const cellFormulas = (xml: string) => [...xml.matchAll(/<f>([^<]*)</g)].map(match => match[1]);

    expect(calculated(text(written, "xl/tables/table1.xml"))).toEqual(
      calculated(text(excel, "xl/tables/table1.xml"))
    );
    expect(cellFormulas(text(written, "xl/worksheets/sheet1.xml"))).toEqual(
      cellFormulas(text(excel, "xl/worksheets/sheet1.xml"))
    );
  });

  it("are written to XLSB as Excel writes them, in the cells and as BrtListCCFmla", async () => {
    const excel = await extractAll(await load("table-this-row.excel.xlsb"));
    const written = await extractAll(await Workbook.toBuffer(authored(), { format: "xlsb" }));
    const cellRecords = ["BrtFmlaNum"];
    expect(records(written.get("xl/worksheets/sheet1.bin")!.data, cellRecords)).toEqual(
      records(excel.get("xl/worksheets/sheet1.bin")!.data, cellRecords)
    );
    expect(records(written.get("xl/tables/table1.bin")!.data, ["BrtListCCFmla"])).toEqual(
      records(excel.get("xl/tables/table1.bin")!.data, ["BrtListCCFmla"])
    );
  });

  it("read from Excel's XLSB as the same formulas as from Excel's XLSX", async () => {
    const fromXlsx = await formulas(await load("table-this-row.excel.xlsx"));
    expect(fromXlsx.columns).toContainEqual(["s", "Tb[[#This Row],[x]]+Tb[[#This Row],[y]]"]);
    expect(await formulas(await load("table-this-row.excel.xlsb"))).toEqual(fromXlsx);
  });

  it("is reported outside any table rather than written broken, and the cell keeps its value", async () => {
    for (const format of ["xlsx", "xlsb"] as const) {
      const workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "S");
      Cell.setValue(sheet, "A1", { formula: "[@x]*2", result: 4 });
      await expect(Workbook.toBuffer(workbook, { format })).rejects.toThrow(
        /S!A1: formula \(an unnamed this-row reference outside any table\)/
      );
      const bytes = await Workbook.toBuffer(workbook, { format, unsupported: "ignore" });
      const back = Workbook.create();
      await Workbook.read(back, bytes);
      const cell = Workbook.getModel(back).worksheets[0]!.rows![0]!.cells[0]!;
      expect(cell.formula, format).toBeUndefined();
      expect(cell.value, format).toBe(4);
    }
  });
});
