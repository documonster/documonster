/**
 * `parseOoxmlDate`'s readings, for the Node and the browser runs to register: every engine must read each text to the
 * same calendar value. Appending `Z` to every string read `2024-02-09Z` and `1/15/2020Z` in V8 and nowhere else —
 * Safari reads both as an invalid date — and leaving a time of day to the engine read it as local time, which one
 * inside a daylight-saving gap cannot be.
 */
import { parseOoxmlDate } from "@utils/utils";
import { describe, expect, it } from "vitest";

export function describeOoxmlDate(): void {
  describe("parseOoxmlDate", () => {
    it.each([
      ["2024-02-09", "2024-02-09T00:00:00.000Z"],
      ["2024-02-09T10:20:30", "2024-02-09T10:20:30.000Z"],
      ["2024-02-09T10:20", "2024-02-09T10:20:00.000Z"],
      ["2024-02-09T10:20:30.5", "2024-02-09T10:20:30.500Z"],
      ["2024-02-09T10:20:30Z", "2024-02-09T10:20:30.000Z"],
      ["2024-02-09T12:20:30+02:00", "2024-02-09T10:20:30.000Z"],
      ["2024-02-09 10:20:30", "2024-02-09T10:20:30.000Z"],
      ["1/15/2020", "2020-01-15T00:00:00.000Z"],
      ["Jan 15 2020", "2020-01-15T00:00:00.000Z"],
      ["1/15/2020 10:30 PM", "2020-01-15T22:30:00.000Z"],
      ["1/15/2020, 12:05:09 AM", "2020-01-15T00:05:09.000Z"],
      ["Wed Jan 15 2020 10:30:15", "2020-01-15T10:30:15.000Z"],
      ["15 January 2020 10:30", "2020-01-15T10:30:00.000Z"],
      ["Jan 15 2020 10:00 GMT+0200", "2020-01-15T08:00:00.000Z"],
      // An ISO year or month is UTC as an ISO date is: read as local, it was a day early west of Greenwich.
      ["2024", "2024-01-01T00:00:00.000Z"],
      ["2024-02", "2024-02-01T00:00:00.000Z"],
      ["2024-02-09t10:20:30", "2024-02-09T10:20:30.000Z"],
      ["2024-02-09T10:20:30z", "2024-02-09T10:20:30.000Z"],
      ["2024-02-09T12:20:30+0200", "2024-02-09T10:20:30.000Z"],
      ["2024-02-09T10:20:30.123456", "2024-02-09T10:20:30.123Z"],
      ["2024-02-09T24:00", "2024-02-10T00:00:00.000Z"],
      ["+002024-02-09", "2024-02-09T00:00:00.000Z"],
      ["0099-01-01", "0099-01-01T00:00:00.000Z"],
      [" 2024-02-09 ", "2024-02-09T00:00:00.000Z"],
      ["1/15/2020 10:30:15.1234", "2020-01-15T10:30:15.123Z"],
      ["10:30 1/15/2020", "2020-01-15T10:30:00.000Z"],
      ["January 15, 2020 at 10:30 AM", "2020-01-15T10:30:00.000Z"],
      ["2020/01/15 10:30", "2020-01-15T10:30:00.000Z"],
      // The last day of each February, and of a year — read in another year, a date could overflow into the next month.
      ["2023-02-28", "2023-02-28T00:00:00.000Z"],
      ["1900-02-28", "1900-02-28T00:00:00.000Z"],
      ["2000-02-29", "2000-02-29T00:00:00.000Z"],
      ["2024-02-29", "2024-02-29T00:00:00.000Z"],
      ["0099-02-28", "0099-02-28T00:00:00.000Z"],
      ["2023-02-28T23:59:59.999", "2023-02-28T23:59:59.999Z"],
      ["2023-12-31T24:00:00.000", "2024-01-01T00:00:00.000Z"],
      // `Date#toString()`: an instant, its zone named.
      [
        "Wed Jan 15 2020 10:30:00 GMT+1100 (Australian Eastern Daylight Time)",
        "2020-01-14T23:30:00.000Z"
      ]
    ])("reads %j", (text, iso) => {
      expect(parseOoxmlDate(text).toISOString()).toBe(iso);
    });

    it.each([
      "TODAY()",
      "13:30 PM",
      "1/15/2020 25:00",
      "10:30",
      "",
      // Engines read a bare number as different years, and a day or second out of range their own ways.
      "1",
      "12",
      "2024-02-30",
      "2024-02-09T10:30:60",
      "2024-13-01",
      // A time this cannot read, or a second one, was read as midnight.
      "1/15/2020 1:2",
      "1/15/2020 10:30 11:45",
      // A day the month does not have, February of a common year included: every February of one read as a day long.
      "2023-02-29",
      "1900-02-29",
      "2100-02-29",
      "2023-04-31",
      // The end of the day is 24:00 exactly.
      "2024-02-09T24:00:00.0001",
      "2024-02-09T24:00:01"
    ])("reads %j as an invalid date", text => {
      expect(Number.isNaN(parseOoxmlDate(text).getTime())).toBe(true);
    });
  });
}
