/**
 * A formula whose result is the empty string, as Excel saves it: `<c t="str"><f>""</f><v/></c>` (Excel 16's own save
 * of `tmp/office-tests/r4-formula/r4-blank.xlsx`, sheet E, C2). It is the empty string, not "no result".
 */
import { Archive } from "@archive";
import { Cell, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

/** An XLSX whose sheet holds `cells` as written, Excel's spelling. */
async function withCells(cells: string): Promise<Uint8Array> {
  const handle = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(handle, "S"), "Z9", 1);
  const files = await Archive.unpackZip(await Workbook.toBuffer(handle));
  const entries = new Map([...files].map(([path, entry]) => [path, entry.data]));
  const sheet = new TextDecoder()
    .decode(entries.get("xl/worksheets/sheet1.xml"))
    .replace(/<sheetData>.*<\/sheetData>/, `<sheetData><row r="1">${cells}</row></sheetData>`);
  entries.set("xl/worksheets/sheet1.xml", new TextEncoder().encode(sheet));
  return Archive.packZip(entries);
}

async function sheetXml(bytes: Uint8Array): Promise<string> {
  const files = await Archive.unpackZip(bytes);
  return new TextDecoder().decode(files.get("xl/worksheets/sheet1.xml")!.data);
}

describe("an empty-string formula result", () => {
  it.each(['<c r="A1" t="str"><f>""</f><v/></c>', '<c r="A1" t="str"><f>""</f><v></v></c>'])(
    "reads %s as the empty string",
    async cell => {
      const ws = Workbook.listWorksheets(await Workbook.read(await withCells(cell)))[0]!;
      expect(Cell.getValue(ws, "A1")).toEqual({ formula: '""', result: "" });
      expect(Cell.getResult(ws, "A1")).toBe("");
    }
  );

  it("writes it back as it came, and reads it back again", async () => {
    const read = await Workbook.read(await withCells('<c r="A1" t="str"><f>B1</f><v/></c>'));
    const bytes = await Workbook.toBuffer(read);
    expect(await sheetXml(bytes)).toMatch(/<c r="A1" t="str"><f>B1<\/f><v(?:\/>|><\/v>)<\/c>/);
    const again = Workbook.listWorksheets(await Workbook.read(bytes))[0]!;
    expect(Cell.getValue(again, "A1")).toEqual({ formula: "B1", result: "" });
  });

  it.each(["xlsx", "xlsb"] as const)(
    "round-trips a result set to the empty string through %s",
    async format => {
      const handle = Workbook.create();
      Cell.setValue(Workbook.addWorksheet(handle, "S"), "A1", { formula: '""', result: "" });
      const ws = Workbook.listWorksheets(
        await Workbook.read(await Workbook.toBuffer(handle, { format }))
      )[0]!;
      expect(Cell.getValue(ws, "A1")).toEqual({ formula: '""', result: "" });
    }
  );

  it('keeps an empty string in a plain `t="str"` cell — a member of an array formula', async () => {
    const ws = Workbook.listWorksheets(
      await Workbook.read(
        await withCells(
          '<c r="A1" t="str"><f t="array" ref="A1:B1">C1:D1</f><v>p</v></c><c r="B1" t="str"><v/></c>'
        )
      )
    )[0]!;
    expect(Cell.getValue(ws, "B1")).toBe("");
  });

  it("still reads a formula with no `<v>` as having no result", async () => {
    const ws = Workbook.listWorksheets(
      await Workbook.read(await withCells('<c r="A1"><f>B1</f></c>'))
    )[0]!;
    expect(Cell.getValue(ws, "A1")).toEqual({ formula: "B1" });
  });
});
