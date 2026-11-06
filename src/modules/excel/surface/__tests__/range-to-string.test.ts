import { ExcelArgumentError, Range } from "@excel";
import { describe, expect, it } from "vitest";

// Expected values are what excelts 10.2.0 / exceljs 4.4.0 give for `range`, `$range`, `shortRange`, `$shortRange`.
describe("Range.toString", () => {
  it.each([
    ["B2:D5", "B2:D5", "$B$2:$D$5", "B2:D5", "$B$2:$D$5"],
    ["B2", "B2:B2", "$B$2:$B$2", "B2", "$B$2"],
    ["Sheet1!B2:D5", "Sheet1!B2:D5", "Sheet1!$B$2:$D$5", "Sheet1!B2:D5", "Sheet1!$B$2:$D$5"],
    ["Sheet1!B2", "Sheet1!B2:B2", "Sheet1!$B$2:$B$2", "Sheet1!B2", "Sheet1!$B$2"]
  ])("%s", (input, range, absolute, short, absoluteShort) => {
    const r = Range.create(input);
    expect(Range.toString(r)).toBe(range);
    expect(Range.toString(r, { absolute: true })).toBe(absolute);
    expect(Range.toString(r, { compact: true })).toBe(short);
    expect(Range.toString(r, { absolute: true, compact: true })).toBe(absoluteShort);
  });

  it("refuses options that are not", () => {
    const r = Range.create("A1");
    expect(() => Range.toString(r, { absolute: "yes" } as never)).toThrow(ExcelArgumentError);
    expect(() => Range.toString(r, 1 as never)).toThrow(ExcelArgumentError);
  });
});
