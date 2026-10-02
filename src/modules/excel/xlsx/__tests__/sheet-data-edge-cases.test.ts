import { zip } from "@archive/create-archive";
import { extractAll } from "@archive/unzip/extract";
import { cellGetValue } from "@excel/core/cell";
import type { CellData } from "@excel/core/cell";
import { Enums } from "@excel/core/enums";
import { Cell, Workbook } from "@excel/index";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import { describe, expect, it } from "vitest";

/**
 * The full reader and the streaming reader must read the same `<sheetData>` the same way — or both refuse it.
 * Each case is a sheet our writer never produces but other producers, or damaged files, do.
 */

async function withSheetData(sheetData: string): Promise<Uint8Array> {
  const wb = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", 0);
  const archive = zip();
  for (const [name, file] of await extractAll(await Workbook.toBuffer(wb))) {
    if (file.type === "directory") {
      continue;
    }
    archive.add(
      name,
      name === "xl/worksheets/sheet1.xml"
        ? new TextDecoder()
            .decode(file.data)
            .replace(/<sheetData>.*<\/sheetData>|<sheetData\/>/s, sheetData)
        : file.data
    );
  }
  return archive.bytes();
}

async function loadWorkbook(buffer: Uint8Array) {
  const wb = Workbook.create();
  await Workbook.read(wb, buffer);
  return wb;
}

async function readFull(buffer: Uint8Array) {
  return Workbook.getWorksheet(await loadWorkbook(buffer), "Sheet1")!;
}

/** Write the workbook holding `buffer`, read that back, and return its sheet. */
async function roundTrip(buffer: Uint8Array) {
  return readFull(await Workbook.toBuffer(await loadWorkbook(buffer)));
}

/** Every cell the stream yields, as `address → value`. */
async function readStream(buffer: Uint8Array): Promise<Record<string, unknown>> {
  const reader = new WorkbookReader(buffer, { worksheets: "emit", sharedStrings: "cache" });
  const cells: Record<string, unknown> = {};
  for await (const ws of reader) {
    for await (const row of ws) {
      for (const cell of (row as { cells: (CellData | undefined)[] }).cells) {
        if (cell) {
          cells[cell.address] = cellGetValue(cell);
        }
      }
    }
  }
  return cells;
}

const sheet = (...rows: string[]) => `<sheetData>${rows.join("")}</sheetData>`;

describe("row and cell references", () => {
  it.each(["x", "12abc", "0", "-1", "1048577", "+", " "])(
    "both readers reject <row r=%j>",
    async r => {
      const buffer = await withSheetData(sheet(`<row r="${r}"><c><v>1</v></c></row>`));
      await expect(readFull(buffer)).rejects.toThrow(`Invalid row number "${r}"`);
      await expect(readStream(buffer)).rejects.toThrow(`Invalid row number "${r}"`);
    }
  );

  it.each(["007", " 7 ", "+7"])(
    "both readers accept <row r=%j>, as xsd:unsignedInt does",
    async r => {
      const buffer = await withSheetData(sheet(`<row r="${r}"><c r="B7"><v>1</v></c></row>`));
      expect(Cell.getValue(await readFull(buffer), "B7")).toBe(1);
      expect(await readStream(buffer)).toEqual({ B7: 1 });
    }
  );

  it("both readers reject a row inferred past the last one", async () => {
    const buffer = await withSheetData(
      sheet('<row r="1048576"><c><v>1</v></c></row><row><c><v>2</v></c></row>')
    );
    await expect(readFull(buffer)).rejects.toThrow("expected 1 to 1048576");
    await expect(readStream(buffer)).rejects.toThrow("expected 1 to 1048576");
  });

  it.each(["A1x", "1A", "A 1"])("both readers reject the cell reference %j", async r => {
    const buffer = await withSheetData(sheet(`<row r="1"><c r="${r}"><v>1</v></c></row>`));
    await expect(readFull(buffer)).rejects.toThrow(r);
    await expect(readStream(buffer)).rejects.toThrow(r);
  });

  it("both readers reject a cell that names another row", async () => {
    const buffer = await withSheetData(sheet('<row r="5"><c r="A7"><v>1</v></c></row>'));
    await expect(readFull(buffer)).rejects.toThrow("cell A7 is written inside row 5");
    await expect(readStream(buffer)).rejects.toThrow("cell A7 is written inside row 5");
  });

  it("both readers place an absolute reference at the cell it names", async () => {
    const buffer = await withSheetData(sheet('<row r="5"><c r="$C$5"><v>7</v></c></row>'));
    expect(Cell.getValue(await readFull(buffer), "C5")).toBe(7);
    expect(await readStream(buffer)).toEqual({ C5: 7 });
  });
});

describe("formula cached results", () => {
  it("report a boolean result's effective type as Boolean", async () => {
    const buffer = await withSheetData(
      sheet('<row r="1"><c r="A1" t="b"><f>1=1</f><v>1</v></c></row>')
    );
    expect(Cell.getEffectiveType(await readFull(buffer), "A1")).toBe(Enums.ValueType.Boolean);
  });

  it('reads a t="d" result as a date in both readers, and writes it back as one', async () => {
    const buffer = await withSheetData(
      sheet('<row r="1"><c r="A1" t="d"><f>TODAY()</f><v>2020-01-15T00:00:00</v></c></row>')
    );
    const expected = { formula: "TODAY()", result: new Date(Date.UTC(2020, 0, 15)) };
    const ws = await readFull(buffer);
    expect(Cell.getValue(ws, "A1")).toEqual(expected);
    expect((await readStream(buffer)).A1).toEqual(expected);
    expect(Cell.getValue(await roundTrip(buffer), "A1")).toEqual(expected);
  });
});

describe("shared formulas without a usable master", () => {
  it("keep the cached result of a dependent whose master never appears, and write it back", async () => {
    const buffer = await withSheetData(
      sheet('<row r="1"><c r="A1"><f t="shared" si="3"/><v>9</v></c></row>')
    );
    const ws = await readFull(buffer);
    expect(Cell.getValue(ws, "A1")).toBe(9);
    expect(await readStream(buffer)).toEqual({ A1: 9 });
    expect(Cell.getValue(await roundTrip(buffer), "A1")).toBe(9);
  });

  it("give a dependent before its master the master's formula when loading, and refuse when streaming", async () => {
    const buffer = await withSheetData(
      sheet(
        '<row r="1"><c r="A1"><f t="shared" si="0"/><v>18</v></c></row>',
        '<row r="2"><c r="A2"><f t="shared" ref="A1:A2" si="0">B2*2</f><v>2</v></c></row>'
      )
    );
    const ws = await readFull(buffer);
    expect(Cell.getValue(ws, "A1")).toEqual({ formula: "B1*2", result: 18 });
    expect(Cell.getValue(await roundTrip(buffer), "A1")).toEqual({ formula: "B1*2", result: 18 });

    await expect(readStream(buffer)).rejects.toThrow("comes before its master A2");
  });

  it("leave strings and quoted sheet names alone when sliding", async () => {
    const buffer = await withSheetData(
      sheet(
        `<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">IF(B1="B1",'Q1 C1'!C1,"x""D1")</f><v>1</v></c></row>`,
        '<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row>'
      )
    );
    const expected = `IF(B2="B1",'Q1 C1'!C2,"x""D1")`;
    expect(Cell.getFormula(await readFull(buffer), "A2")).toBe(expected);
    expect((await readStream(buffer)).A2).toMatchObject({ formula: expected });
  });

  it("slide the formula to a dependent outside the declared range, as Excel computes it", async () => {
    const buffer = await withSheetData(
      sheet(
        '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>1</v></c></row>',
        '<row r="9"><c r="A9"><f t="shared" si="0"/><v>3</v></c></row>'
      )
    );
    expect(Cell.getFormula(await readFull(buffer), "A9")).toBe("B9+1");
    expect((await readStream(buffer)).A9).toMatchObject({ formula: "B9+1", result: 3 });
  });
});
