/**
 * `DateSerial` converts Excel serials that are not in a cell, without a `Date` and so without a
 * timezone. The expected serials are Excel's own: 2024-01-15 is 45306, 1900-01-01 is 1, the
 * fictitious 1900-02-29 is 60, and the 1904 epoch is 1462 days later.
 */
import { DateSerial, ExcelArgumentError } from "@excel/index";
import { describe, expect, it } from "vitest";

const midnight = { hour: 0, minute: 0, second: 0, millisecond: 0 };

describe("DateSerial", () => {
  it.each<[number, object]>([
    [45306, { year: 2024, month: 1, day: 15, ...midnight }],
    [45306.5, { year: 2024, month: 1, day: 15, ...midnight, hour: 12 }],
    [1, { year: 1900, month: 1, day: 1, ...midnight }],
    [60, { year: 1900, month: 2, day: 29, ...midnight }],
    [61, { year: 1900, month: 3, day: 1, ...midnight }]
  ])("reads %d", (serial, parts) => {
    expect(DateSerial.toParts(serial)).toEqual(parts);
    expect(DateSerial.fromParts(parts)).toBe(serial);
  });

  it("counts from 1904-01-01 under the 1904 epoch", () => {
    expect(DateSerial.toParts(0, { date1904: true })).toMatchObject({
      year: 1904,
      month: 1,
      day: 1
    });
    expect(DateSerial.fromParts({ year: 2024, month: 1, day: 15 }, { date1904: true })).toBe(
      45306 - 1462
    );
  });

  it("reads a time of day alone as a fraction of a day, on either epoch", () => {
    for (const date1904 of [false, true]) {
      expect(DateSerial.fromParts({ hour: 6 }, { date1904 })).toBe(0.25);
      expect(DateSerial.toParts(0.75, { date1904 })).toMatchObject({ hour: 18, minute: 0 });
    }
  });

  it("keeps the milliseconds through a round trip", () => {
    const parts = {
      year: 2024,
      month: 2,
      day: 29,
      hour: 23,
      minute: 59,
      second: 59,
      millisecond: 999
    };
    expect(DateSerial.toParts(DateSerial.fromParts(parts))).toEqual(parts);
  });

  it("does not depend on the host's timezone: there is no Date to have one", () => {
    // 2024-03-10 02:30 does not exist on a US clock; a serial does not care
    const parts = { year: 2024, month: 3, day: 10, hour: 2, minute: 30, second: 0, millisecond: 0 };
    expect(DateSerial.toParts(DateSerial.fromParts(parts))).toEqual(parts);
  });

  it("covers Excel's range, to the last millisecond of 9999-12-31, on either epoch", () => {
    expect(DateSerial.toParts(0)).toMatchObject({ year: 1899, month: 12, day: 31 });
    expect(DateSerial.toParts(2958465)).toMatchObject({ year: 9999, month: 12, day: 31 });
    const last = {
      year: 9999,
      month: 12,
      day: 31,
      hour: 23,
      minute: 59,
      second: 59,
      millisecond: 999
    };
    for (const date1904 of [false, true]) {
      expect(DateSerial.toParts(DateSerial.fromParts(last, { date1904 }), { date1904 })).toEqual(
        last
      );
    }
    expect(DateSerial.fromParts({ year: 1904, month: 1, day: 1 }, { date1904: true })).toBe(0);
  });

  it.each<[string, number, object?]>([
    ["a negative serial", -1],
    ["10000-01-01", 2958466],
    ["a serial that rounds to 10000-01-01", 2958466 - 1e-9],
    ["a serial past the 1904 epoch's end", 2957004, { date1904: true }],
    ["a serial far out of range", 1e9],
    ["a serial whose millisecond count overflows", 1e305]
  ])("toParts refuses %s, which is not a date to Excel", (_name, serial, options) => {
    let error: unknown;
    try {
      DateSerial.toParts(serial, options);
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(ExcelArgumentError);
    expect((error as ExcelArgumentError).field).toBe("serial");
    expect((error as Error).message).toContain("10000-01-01");
  });

  it.each<[string, () => unknown, string]>([
    [
      "a year before 1900",
      () => DateSerial.fromParts({ year: 1899, month: 12, day: 31 }),
      "parts.year"
    ],
    ["a year after 9999", () => DateSerial.fromParts({ year: 10000 }), "parts.year"],
    ["a year with no exact serial", () => DateSerial.fromParts({ year: 1e15 }), "parts.year"],
    [
      "a 1904-epoch year before 1904",
      () => DateSerial.fromParts({ year: 1903, month: 12, day: 31 }, { date1904: true }),
      "parts"
    ],
    ["a serial that is not finite", () => DateSerial.toParts(Number.NaN), "serial"],
    ["a serial that is a string", () => DateSerial.toParts("45306" as never), "serial"],
    ["options that are not an object", () => DateSerial.toParts(1, true as never), "options"],
    [
      "a month out of range",
      () => DateSerial.fromParts({ year: 2024, month: 13, day: 1 }),
      "parts.month"
    ],
    ["a field that is not an integer", () => DateSerial.fromParts({ hour: 1.5 }), "parts.hour"],
    ["an unknown field", () => DateSerial.fromParts({ year: 2024, months: 2 } as never), "parts"],
    [
      "a day past the end of its month",
      () => DateSerial.fromParts({ year: 2024, month: 2, day: 30 }),
      "parts"
    ],
    [
      "1900-02-29 under the 1904 epoch",
      () => DateSerial.fromParts({ year: 1900, month: 2, day: 29 }, { date1904: true }),
      "parts"
    ],
    ["a month without a year", () => DateSerial.fromParts({ month: 2 }), "parts"],
    ["no fields", () => DateSerial.fromParts({}), "parts"]
  ])("refuses %s", (_name, run, fieldPrefix) => {
    let error: unknown;
    try {
      run();
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(ExcelArgumentError);
    expect((error as ExcelArgumentError).field.startsWith(fieldPrefix)).toBe(true);
    expect((error as Error).message).toMatch(/^DateSerial\.(toParts|fromParts): /);
  });
});
