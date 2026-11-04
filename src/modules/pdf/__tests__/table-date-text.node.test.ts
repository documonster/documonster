/**
 * `Pdf.fromTable` renders a `Date` cell as `Pdf.fromExcel` renders a date cell with no number
 * format — Excel's General date display from the UTC calendar fields — so the output does not
 * depend on the machine. It used `toLocaleDateString()`: a different text per locale, and a
 * different *day* west of UTC (`2024-01-15T00:00Z` is the 14th in New York).
 *
 * Runs under several zones because at offset zero the wrong answer and the right one agree.
 */
import { Cell, Workbook } from "@excel/index";
import { excelToPdf } from "@pdf/excel-bridge";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { Pdf } from "../index";

const ZONES = ["UTC", "America/New_York", "Asia/Shanghai", "Pacific/Kiritimati"] as const;

const MIDNIGHT = new Date(Date.UTC(2024, 0, 15));
const AFTERNOON = new Date(Date.UTC(2024, 0, 15, 13, 5, 9));

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

async function text(bytes: Uint8Array): Promise<string> {
  return (await Pdf.extract(bytes)).pages.map(page => page.text).join("\n");
}

describe.each(ZONES)("in %s", zone => {
  beforeEach(() => {
    process.env.TZ = zone;
  });

  it("renders a Date from its UTC fields: the date alone at midnight, else date and time", async () => {
    const out = await text(await Pdf.fromTable([[MIDNIGHT, { value: AFTERNOON, bold: true }]]));
    expect(out).toContain("2024-01-15");
    expect(out).toContain("2024-01-15 13:05:09");
    expect(out).not.toContain("2024-01-14");
    expect(out).not.toContain("1/15/2024");
  });

  it("matches Pdf.fromExcel for a date cell with no number format", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", MIDNIGHT);
    Cell.setValue(ws, "B1", AFTERNOON);
    const fromExcel = await text(await excelToPdf(wb));
    const fromTable = await text(await Pdf.fromTable([[MIDNIGHT, AFTERNOON]]));
    expect(fromTable.replace(/\s+/g, " ").trim()).toBe(fromExcel.replace(/\s+/g, " ").trim());
  });
});
