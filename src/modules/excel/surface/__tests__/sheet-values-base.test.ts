/**
 * The sheet-wide values read the way row and column values do: `{ base }` chooses the indexing, `0` for plain
 * JavaScript indexing and `1` for the index being the row and column number. Without options `Worksheet.getValues`
 * stays 1-based, as it always was. One `ValuesOptions` type serves `Row`, `Column`, `Worksheet` and `Stream`.
 */
import { Cell, ExcelArgumentError, Row, Workbook, Worksheet } from "@excel";
import type { Column, Stream } from "@excel";
import type { ValuesOptions } from "@excel/utils/numbered-values";
import { describe, expect, expectTypeOf, it } from "vitest";

function sheet() {
  const ws = Workbook.addWorksheet(Workbook.create(), "S");
  Cell.setValue(ws, "A1", "a");
  Cell.setValue(ws, "B1", "b");
  Cell.setValue(ws, "A3", 1);
  return ws;
}

describe("Worksheet.getValues({ base })", () => {
  it("without options: 1-based, as before", () => {
    // eslint-disable-next-line no-sparse-arrays
    expect(Worksheet.getValues(sheet())).toEqual([, [, "a", "b"], , [, 1]]);
  });

  it("base 1 is the same", () => {
    expect(Worksheet.getValues(sheet(), { base: 1 })).toEqual(Worksheet.getValues(sheet()));
  });

  it("base 0: row 1 at index 0 and column A at index 0, as Row.getValues", () => {
    const ws = sheet();
    // eslint-disable-next-line no-sparse-arrays
    expect(Worksheet.getValues(ws, { base: 0 })).toEqual([["a", "b"], , [1]]);
    expect(Worksheet.getValues(ws, { base: 0 })[0]).toEqual(Row.getValues(ws, 1));
  });

  it("checks base as Row.getValues does", () => {
    expect(() => Worksheet.getValues(sheet(), { base: 2 as never })).toThrow(ExcelArgumentError);
    expect(() => Worksheet.getValues(sheet(), { base: 2 as never })).toThrow(
      /^Worksheet\.getValues: options\.base/
    );
  });

  it("one options type everywhere", () => {
    expectTypeOf<Row.ValuesOptions>().toEqualTypeOf<ValuesOptions>();
    expectTypeOf<Column.ValuesOptions>().toEqualTypeOf<ValuesOptions>();
    expectTypeOf<Parameters<typeof Worksheet.getValues>[1]>().toEqualTypeOf<
      ValuesOptions | undefined
    >();
    expectTypeOf<Parameters<typeof Stream.getRowValues>[1]>().toEqualTypeOf<
      ValuesOptions | undefined
    >();
  });
});
