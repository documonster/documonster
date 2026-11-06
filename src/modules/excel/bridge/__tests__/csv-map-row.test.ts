import { Row, Workbook, ExcelArgumentError } from "@excel";
import { ExcelCsv } from "@excel/bridge/csv";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Row.add(ws, ["name", "qty"]);
  Row.add(ws, ["a", 1]);
  Row.add(ws, ["b", 2]);
  return { wb, ws };
}

// exceljs's `formatterOptions.transform`: `row => [...row, "tail"]` wrote "name,qty,tail\na,1,tail\nb,2,tail".
describe("ExcelCsv write: mapRow", () => {
  it("turns every written row into the row returned", async () => {
    const { wb } = sheet();
    const mapRow = (values: unknown[]) => [...values, "tail"] as never;
    expect(ExcelCsv.format(wb, { mapRow })).toBe("name,qty,tail\na,1,tail\nb,2,tail");
    expect(new TextDecoder().decode(await ExcelCsv.toBuffer(wb, { mapRow }))).toBe(
      "name,qty,tail\na,1,tail\nb,2,tail"
    );
  });

  it("runs after map, and is given the row number — empty rows included", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.add(ws, ["h"]);
    Row.setValues(ws, 3, ["x"]);
    const seen: number[] = [];
    const csv = ExcelCsv.format(wb, {
      map: value => String(value).toUpperCase(),
      mapRow: (values, rowNumber) => {
        seen.push(rowNumber);
        return [rowNumber, ...values];
      }
    });
    expect(seen).toEqual([1, 2, 3]);
    expect(csv).toBe("1,H\n2\n3,X");
  });

  it("can reorder and drop fields", () => {
    const { wb } = sheet();
    expect(ExcelCsv.format(wb, { mapRow: values => [values[1]!] })).toBe("qty\n1\n2");
  });

  it("refuses a result that is not a row", () => {
    const { wb } = sheet();
    expect(() => ExcelCsv.format(wb, { mapRow: () => "x" as never })).toThrow(ExcelArgumentError);
    expect(() => ExcelCsv.format(wb, { mapRow: "x" as never })).toThrow(ExcelArgumentError);
  });
});
