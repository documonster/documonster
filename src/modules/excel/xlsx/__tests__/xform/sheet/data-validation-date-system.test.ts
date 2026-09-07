/**
 * A `type: "date"` data validation bound, against the workbook's own date system.
 *
 * **A bound is a serial, exactly like a cell value, and both writers were converting it against a hard-coded
 * 1900 epoch.** In a 1904 workbook the bound therefore sat 1,462 days from the cells it constrained, so a rule
 * reading "on or after 2020-01-15" rejected that very date.
 *
 * The reason it survived so long is the reason this file asserts what it does. The read side had the *mirror
 * image* of the same omission, so a round trip through this library agreed with itself perfectly — the only
 * thing that could see the defect was comparing the bound against the cell **inside one file**. So that is the
 * assertion: `bound === cell`, in both containers, under both epochs. It needs no oracle and cannot be
 * satisfied by two mistakes cancelling.
 *
 * The XLSX half also guards against a specific way the fix can rot. The epoch reaches the bound by being
 * stamped onto the rule in `prepare`, because `prepare` and `render` run on separately constructed
 * `WorksheetXform` instances — a first attempt held it in an instance field, which could never have worked, and
 * a second forgot to call `prepare` at all. Both failures produce exactly the numbers below.
 *
 * The streaming writers reach the bound by a different route and are covered by the `.node.test.ts` sibling,
 * which needs a real file on disk. Everything here is in-memory and runs in the browser as well.
 */
import { Cell, Workbook } from "@excel";
import type { WorkbookFormat } from "@excel";
import { describe, expect, it } from "vitest";

import { boundSerial, SERIAL_1900, SERIAL_1904, WHEN } from "./data-validation-date-system.helpers";

/** A workbook with one dated cell and one date bound at the same address. */
async function build(format: WorkbookFormat, date1904: boolean): Promise<Uint8Array> {
  const workbook = Workbook.create();
  (workbook as { properties?: Record<string, unknown> }).properties = {
    ...((workbook as { properties?: Record<string, unknown> }).properties ?? {}),
    date1904
  };
  const sheet = Workbook.addWorksheet(workbook, "S");
  Cell.setValue(sheet, "A1", WHEN);
  Cell.setValidation(sheet, "A1", {
    type: "date",
    operator: "greaterThanOrEqual",
    formulae: [WHEN]
  });
  return Workbook.toBuffer(workbook, { format, unsupported: "ignore" });
}

describe.each<WorkbookFormat>(["xlsx", "xlsb"])("%s date validation bounds", format => {
  it.each([
    { date1904: false, serial: SERIAL_1900 },
    { date1904: true, serial: SERIAL_1904 }
  ])("writes the bound as $serial at date1904=$date1904", async ({ date1904, serial }) => {
    const reopened = Workbook.create();
    await Workbook.read(reopened, await build(format, date1904));
    const rule = Cell.getValidation(Workbook.getWorksheet(reopened, "S")!, "A1");
    expect(rule?.type).toBe("date");
    // Against Excel's arithmetic, so neither the reader nor the writer can drift without this failing.
    expect(boundSerial((rule as { formulae?: unknown[] }).formulae?.[0] as never, date1904)).toBe(
      serial
    );
  });

  it.each([false, true])(
    "puts the bound on the same day as the cell it constrains, at date1904=%s",
    async date1904 => {
      // The assertion that would have caught this. Both sides being wrong by the same amount is exactly what
      // hid it, and this cannot be satisfied that way: the cell's serial was always correct.
      const reopened = Workbook.create();
      await Workbook.read(reopened, await build(format, date1904));
      const sheet = Workbook.getWorksheet(reopened, "S")!;
      const rule = Cell.getValidation(sheet, "A1");
      const bound = boundSerial(
        (rule as { formulae?: unknown[] }).formulae?.[0] as never,
        date1904
      );
      const cell = Cell.getDateParts(sheet, "A1");
      expect(cell).toMatchObject({ year: 2020, month: 1, day: 15 });
      expect(bound).toBe(date1904 ? SERIAL_1904 : SERIAL_1900);
    }
  );
});

describe("the model the epoch pass walks", () => {
  it("tolerates a removed validation, which leaves a hole rather than deleting the key", async () => {
    // The validation model is sparse: `Cell.setValidation(…, undefined)` leaves the address present with an
    // `undefined` value, so `Object.values` yields holes. The first version of the pass that stamps the epoch
    // read `.type` off them and threw while *writing the workbook* — a crash on a sheet that merely once had a
    // validation. `optimiseDataValidations` filters the same holes a few lines away, which is where the shape
    // should have been learned from.
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    Cell.setValidation(sheet, "A1", {
      type: "date",
      operator: "greaterThanOrEqual",
      formulae: [WHEN]
    });
    Cell.setValidation(sheet, "A1", undefined as never);
    await expect(Workbook.toBuffer(workbook)).resolves.toBeInstanceOf(Uint8Array);
  });
});

describe("a date validation bound reads back as a Date in either container", () => {
  it.each<WorkbookFormat>(["xlsx", "xlsb"])("from %s", async format => {
    // XLSB returned the serial's *text* where XLSX returned a `Date`, so `Cell.getValidation` answered a
    // different type depending on which container the workbook arrived in — the one thing two readers of the
    // same document must not do. The earlier version of this file accepted either, which hid it.
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    Cell.setValidation(sheet, "A1", {
      type: "date",
      operator: "greaterThanOrEqual",
      formulae: [WHEN]
    });
    const reopened = Workbook.create();
    await Workbook.read(
      reopened,
      await Workbook.toBuffer(workbook, { format, unsupported: "ignore" })
    );
    const rule = Cell.getValidation(Workbook.getWorksheet(reopened, "S")!, "A1");
    const bound = (rule as { formulae?: unknown[] }).formulae?.[0];
    expect(bound).toBeInstanceOf(Date);
    expect((bound as Date).toISOString()).toBe(WHEN.toISOString());
  });
});
