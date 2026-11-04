/**
 * X1 — a date with no timezone in a CSV is read as the calendar value it names, so the cell holds the same serial on
 * every machine. Every other part of Excel reads a cell `Date` as UTC wall-clock time; reading `2024-03-10` as local
 * midnight put `45360.625` in the file in Tokyo and `45361.208…` in New York.
 *
 * Run under several zones, because at offset zero the wrong answer and the right one are identical.
 */
import { Archive } from "@archive";
import { Cell, Workbook } from "@excel";
import { ExcelCsv } from "@excel/bridge/csv";
import { ExcelMarkdown } from "@excel/bridge/markdown";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

const ZONES = ["UTC", "Asia/Tokyo", "America/New_York", "Pacific/Kiritimati"] as const;

let originalTz: string | undefined;
beforeEach(() => {
  originalTz = process.env.TZ;
});
afterEach(() => {
  if (originalTz === undefined) {
    delete process.env.TZ;
  } else {
    process.env.TZ = originalTz;
  }
});

async function serial(wb: Workbook.Handle, ref: string): Promise<number> {
  const files = await Archive.unpackZip(await Workbook.toBuffer(wb));
  const sheet = new TextDecoder().decode(files.get("xl/worksheets/sheet1.xml")!.data);
  const match = new RegExp(`<c r="${ref}"[^>]*><v>([^<]*)</v>`).exec(sheet);
  return Number(match?.[1]);
}

describe.each(ZONES)("in %s", zone => {
  beforeEach(() => {
    process.env.TZ = zone;
  });

  it.each([
    { text: "2024-03-10", iso: "2024-03-10T00:00:00.000Z", serial: 45361 },
    { text: "2024-03-10T14:30:00", iso: "2024-03-10T14:30:00.000Z", serial: 45361 + 14.5 / 24 },
    // A time a spring-forward transition skips in New York, which a local reading would move by an hour.
    { text: "2024-03-10T02:30:00", iso: "2024-03-10T02:30:00.000Z", serial: 45361 + 2.5 / 24 },
    { text: "03-10-2024", iso: "2024-03-10T00:00:00.000Z", serial: 45361 },
    // A stated offset is an instant, as before.
    { text: "2024-03-10T09:00:00+09:00", iso: "2024-03-10T00:00:00.000Z", serial: 45361 },
    { text: "2024-03-10T00:00:00.000Z", iso: "2024-03-10T00:00:00.000Z", serial: 45361 }
  ])("$text is the cell $iso", async ({ text, iso, serial: expected }) => {
    const wb = await ExcelCsv.read(`d\n${text}`);
    const ws = Workbook.getWorksheet(wb, 1)!;
    expect((Cell.getValue(ws, "A2") as Date).toISOString()).toBe(iso);
    expect(await serial(wb, "A2")).toBeCloseTo(expected, 9);
  });

  it("a named day-first form with a time is the calendar value too", async () => {
    const wb = await ExcelCsv.read("d\n10/03/2024 14:30:00", {
      dateFormats: ["DD/MM/YYYY HH:mm:ss"]
    });
    expect((Cell.getValue(Workbook.getWorksheet(wb, 1)!, "A2") as Date).toISOString()).toBe(
      "2024-03-10T14:30:00.000Z"
    );
  });

  it("a date written as the calendar value reads back as the same cell", async () => {
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", new Date(Date.UTC(2024, 2, 10)));
    const csv = ExcelCsv.format(wb, { dateFormat: "YYYY-MM-DD", dateUTC: true });
    expect(csv.trim()).toBe("2024-03-10");
    const back = await ExcelCsv.read(csv);
    expect((Cell.getValue(Workbook.getWorksheet(back, 1)!, "A1") as Date).toISOString()).toBe(
      "2024-03-10T00:00:00.000Z"
    );
  });

  describe("a date cell is written as the calendar value it holds", () => {
    function book() {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      Cell.setValue(ws, "A1", new Date(Date.UTC(2024, 2, 10)));
      Cell.setValue(ws, "B1", new Date(Date.UTC(2024, 2, 10, 14, 30)));
      return wb;
    }
    const local = (date: Date) =>
      `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;

    it("by default, as UTC ISO", () => {
      expect(ExcelCsv.format(book()).trim()).toBe(
        "2024-03-10T00:00:00.000Z,2024-03-10T14:30:00.000Z"
      );
      expect(ExcelMarkdown.format(book())).toContain(
        "| 2024-03-10T00:00:00.000Z | 2024-03-10T14:30:00.000Z |"
      );
    });

    it("with a dateFormat, from the UTC fields", () => {
      const options = { dateFormat: "YYYY-MM-DD HH:mm" };
      expect(ExcelCsv.format(book(), options).trim()).toBe("2024-03-10 00:00,2024-03-10 14:30");
      expect(ExcelMarkdown.format(book(), options)).toContain(
        "| 2024-03-10 00:00 | 2024-03-10 14:30 |"
      );
    });

    it("dateUTC: false still writes the host's local time", () => {
      const date = new Date(Date.UTC(2024, 2, 10));
      expect(
        ExcelCsv.format(book(), { dateFormat: "YYYY-MM-DD", dateUTC: false }).split(",")[0]
      ).toBe(local(date));
      expect(ExcelMarkdown.format(book(), { dateFormat: "YYYY-MM-DD", dateUTC: false })).toContain(
        `| ${local(date)} |`
      );
    });

    it("round-trips with a dateFormat and no other option", async () => {
      const csv = ExcelCsv.format(book(), { dateFormat: "YYYY-MM-DD HH:mm:ss" });
      const back = await ExcelCsv.read(csv, { dateFormats: ["YYYY-MM-DD HH:mm:ss"] });
      expect((Cell.getValue(Workbook.getWorksheet(back, 1)!, "B1") as Date).toISOString()).toBe(
        "2024-03-10T14:30:00.000Z"
      );
    });
  });
});
