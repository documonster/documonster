/**
 * `parseOoxmlDate` reads text to the same calendar value in every zone the reader may run in. It read date text as
 * local time, so the value depended on the machine: where a zone skips its midnight for daylight saving (Santiago on
 * 2019-09-08), `Sep 8 2019` read as 01:00.
 */
import { parseOoxmlDate } from "@utils/utils";
import { afterEach, describe, expect, it } from "vitest";

const zone = process.env.TZ;
afterEach(() => {
  if (zone === undefined) {
    delete process.env.TZ;
  } else {
    process.env.TZ = zone;
  }
});

describe("parseOoxmlDate in every zone", () => {
  it.each(["UTC", "America/Santiago", "America/Sao_Paulo", "Asia/Kolkata", "Pacific/Kiritimati"])(
    "in %s",
    tz => {
      process.env.TZ = tz;
      const read = (text: string): string => parseOoxmlDate(text).toISOString();
      expect(read("Sep 8 2019")).toBe("2019-09-08T00:00:00.000Z");
      expect(read("9/8/2019")).toBe("2019-09-08T00:00:00.000Z");
      expect(read("Nov 4 2018")).toBe("2018-11-04T00:00:00.000Z");
      expect(read("Sep 8 2019 00:30")).toBe("2019-09-08T00:30:00.000Z");
      expect(read("1/15/2020 10:30 PM")).toBe("2020-01-15T22:30:00.000Z");
      expect(read("2019-09-08T00:30:00")).toBe("2019-09-08T00:30:00.000Z");
      // A time inside the gap is the time written, however the engine would move it.
      expect(read("2019-09-08 00:30:00")).toBe("2019-09-08T00:30:00.000Z");
      expect(read("9/8/2019 12:30 AM")).toBe("2019-09-08T00:30:00.000Z");
      expect(read("Sep 8, 2019, 12:30:15.5 AM")).toBe("2019-09-08T00:30:15.500Z");
      // Text naming its zone names an instant.
      expect(read("Jan 15 2020 10:00 +0200")).toBe("2020-01-15T08:00:00.000Z");
      expect(read("Jan 15 2020 10:00 GMT")).toBe("2020-01-15T10:00:00.000Z");
    }
  );
});
