/**
 * `ExcelMarkdown` reads what it writes — bytes as well as text, UTF-8 as `toBuffer` produces it — and takes the
 * number of tables as an option (`tables: "first" | "all"`) rather than as a second pair of names. Both bridges
 * declare `map` the way they call it: the cell's text and its 0-based column in, a cell value out, checked as
 * `Cell.setValue` checks one.
 */
import { Cell, ExcelArgumentError, ExcelInvalidValueTypeError, Workbook, Worksheet } from "@excel";
import { ExcelCsv, type ExcelCsvReadOptions } from "@excel/bridge/csv";
import { ExcelMarkdown, type ExcelMarkdownReadOptions } from "@excel/bridge/markdown";
import type { CellValueInput } from "@excel/types";
import { describe, expect, expectTypeOf, it } from "vitest";

const TWO_TABLES = "| a | b |\n| --- | --- |\n| 1 | x |\n\ntext\n\n| c |\n| --- |\n| 2 |\n";

function first(wb: Workbook.Handle) {
  return Workbook.listWorksheets(wb)[0]!;
}

describe("ExcelMarkdown.read takes bytes", () => {
  it("reads back what toBufferSync and toBuffer wrote", async () => {
    const wb = ExcelMarkdown.read(TWO_TABLES);
    for (const bytes of [ExcelMarkdown.toBufferSync(wb), await ExcelMarkdown.toBuffer(wb)]) {
      const back = ExcelMarkdown.read(bytes);
      expect(Cell.getValue(first(back), "A2")).toBe(1);
      expect(Cell.getValue(first(back), "B2")).toBe("x");
      expect(
        ExcelMarkdown.read(
          bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
        )
      ).toBeDefined();
    }
  });

  it("decodes UTF-8, a byte-order mark dropped", () => {
    const bytes = new TextEncoder().encode("\uFEFF| 名前 |\n| --- |\n| 値 |\n");
    const ws = first(ExcelMarkdown.read(bytes));
    expect([Cell.getValue(ws, "A1"), Cell.getValue(ws, "A2")]).toEqual(["名前", "値"]);
    const wb = Workbook.create();
    expect(Cell.getValue(ExcelMarkdown.addWorksheet(wb, bytes), "A1")).toBe("名前");
    expect(ExcelMarkdown.addWorksheets(wb, bytes)).toHaveLength(1);
  });

  it("still refuses what is neither text nor bytes", () => {
    for (const input of [42, null, {}, [1]]) {
      expect(() => ExcelMarkdown.read(input as never)).toThrow(
        expect.objectContaining({
          field: "input",
          message: expect.stringMatching(
            /^ExcelMarkdown\.read: input must be Markdown text or its UTF-8 bytes/
          )
        })
      );
    }
  });
});

describe("ExcelMarkdown.read({ tables })", () => {
  it("first by default, every table with 'all'", () => {
    expect(Workbook.listWorksheets(ExcelMarkdown.read(TWO_TABLES))).toHaveLength(1);
    expect(
      Workbook.listWorksheets(ExcelMarkdown.read(TWO_TABLES, { tables: "first" }))
    ).toHaveLength(1);
    const all = ExcelMarkdown.read(TWO_TABLES, { tables: "all", sheetName: "T" });
    expect(Workbook.listWorksheets(all).map(ws => Worksheet.getName(ws))).toEqual(["T", "T_2"]);
    expect(Cell.getValue(Workbook.getWorksheet(all, "T_2")!, "A2")).toBe(2);
  });

  it("is checked", () => {
    expect(() => ExcelMarkdown.read(TWO_TABLES, { tables: "some" as never })).toThrow(
      expect.objectContaining({ field: "options.tables" })
    );
    const wb = Workbook.create();
    // One sheet is what addWorksheet returns, so it reads one table; addWorksheets reads them all.
    expect(() => ExcelMarkdown.addWorksheet(wb, TWO_TABLES, { tables: "all" })).toThrow(
      ExcelArgumentError
    );
    expect(() => ExcelMarkdown.addWorksheets(wb, TWO_TABLES, { tables: "first" })).toThrow(
      ExcelArgumentError
    );
    expect(Workbook.listWorksheets(wb)).toEqual([]);
  });

  it("readAll and readAllFile are gone: one name, one option", () => {
    expect("readAll" in ExcelMarkdown).toBe(false);
  });
});

describe("map is declared as it is called", () => {
  it("text and a 0-based column in, a checked cell value out — both bridges", async () => {
    expectTypeOf<NonNullable<ExcelMarkdownReadOptions["map"]>>().toEqualTypeOf<
      (text: string, column: number) => CellValueInput
    >();
    expectTypeOf<Parameters<NonNullable<ExcelCsvReadOptions["map"]>>>().toEqualTypeOf<
      [text: string, column: number]
    >();
    expectTypeOf<
      ReturnType<NonNullable<ExcelCsvReadOptions["map"]>>
    >().toEqualTypeOf<CellValueInput>();

    const calls: [unknown, number][] = [];
    const map = (text: string, column: number): CellValueInput => {
      calls.push([text, column]);
      return column === 1 ? new Date(Date.UTC(2024, 0, 2)) : text.toUpperCase();
    };
    const ws = first(ExcelMarkdown.read("| a | b |\n| --- | --- |\n| x | y |\n", { map }));
    expect(calls).toEqual([
      ["x", 0],
      ["y", 1]
    ]);
    expect(Cell.getValue(ws, "A2")).toBe("X");
    calls.length = 0;
    await ExcelCsv.read("x,y", { map });
    expect(calls).toEqual([
      ["x", 0],
      ["y", 1]
    ]);
  });

  it("a value a cell cannot hold is refused, in both", async () => {
    const map = (): CellValueInput => Number.NaN;
    expect(() => ExcelMarkdown.read("| a |\n| --- |\n| 1 |\n", { map })).toThrow(
      ExcelInvalidValueTypeError
    );
    await expect(ExcelCsv.read("1", { map })).rejects.toThrow(ExcelInvalidValueTypeError);
  });
});
