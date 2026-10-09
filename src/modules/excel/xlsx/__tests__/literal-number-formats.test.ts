import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

describe("XLSX literal number formats", () => {
  it.each([
    ["\\$0.0,,\\M", 5000000000, "$5000.0M"],
    ["\\$0.0,,\\M", 1000000, "$1.0M"],
    ["\\$0.0,,\\M", 0, "$0.0M"],
    ["\\$0.0,,\\M", -1000000, "-$1.0M"],
    ['$0.0,,"M"', 5000000000, "$5000.0M"],
    ["[Green]\\$0.0,,\\M;[Red](\\$0.0,,\\M)", -1000000, "($1.0M)"],
    ["0.0\\m", 12, "12.0m"],
    ["0.0\\s", 12, "12.0s"],
    ['0" days"', 12, "12 days"],
    ["0_m", 12, "12 "],
    ["0*m", 12, "12"]
  ])(
    "preserves numeric values and formula results with %s (%s)",
    async (numFmt, value, display) => {
      let workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "Repro");
      Cell.setValue(sheet, "A1", value);
      Cell.setValue(sheet, "A2", { formula: String(value), result: value });
      Cell.setNumFmt(sheet, "A1", numFmt);
      Cell.setNumFmt(sheet, "A2", numFmt);

      // A second save/load proves the imported model remains serializable without losing the format.
      for (let pass = 0; pass < 2; pass++) {
        const loaded = Workbook.create();
        await Workbook.read(loaded, new Uint8Array(await Workbook.toBuffer(workbook)));
        const loadedSheet = Workbook.getWorksheet(loaded, "Repro")!;
        expect(Cell.getValue(loadedSheet, "A1")).toBe(value);
        expect(Cell.getValue(loadedSheet, "A2")).toEqual({ formula: String(value), result: value });
        for (const address of ["A1", "A2"]) {
          expect(Cell.getNumFmt(loadedSheet, address)).toBe(numFmt);
          expect(Cell.getDisplayText(loadedSheet, address)).toBe(display);
        }
        workbook = loaded;
      }
    }
  );

  it.each([
    ["yyyy-mm-dd", 45952, "2025-10-22"],
    ["yyyy\\-mm\\-dd", 45952, "2025-10-22"],
    ["hh:mm:ss", 0.5, "12:00:00"],
    ["[h]:mm:ss", 1.5, "36:00:00"]
  ])("still imports genuine date/time formats: %s", async (numFmt, value, display) => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "Dates");
    Cell.setValue(sheet, "A1", value);
    Cell.setValue(sheet, "A2", { formula: String(value), result: value });
    Cell.setNumFmt(sheet, "A1", numFmt);
    Cell.setNumFmt(sheet, "A2", numFmt);
    const loaded = Workbook.create();
    await Workbook.read(loaded, new Uint8Array(await Workbook.toBuffer(workbook)));
    const loadedSheet = Workbook.getWorksheet(loaded, "Dates")!;
    expect(Cell.getValue(loadedSheet, "A1")).toBeInstanceOf(Date);
    expect(Cell.getResult(loadedSheet, "A2")).toBeInstanceOf(Date);
    for (const address of ["A1", "A2"]) {
      expect(Cell.getDisplayText(loadedSheet, address)).toBe(display);
    }
  });
});

describe("XLSX elapsed-time formats", () => {
  it.each(["[h]", "[mm]", "[ss].00"])(
    "keeps a value formatted %s a number: an elapsed tag alone is a quantity",
    async numFmt => {
      const workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "Elapsed");
      Cell.setValue(sheet, "A1", 1.5);
      Cell.setNumFmt(sheet, "A1", numFmt);
      const loaded = Workbook.create();
      await Workbook.read(loaded, new Uint8Array(await Workbook.toBuffer(workbook)));
      expect(Cell.getValue(Workbook.getWorksheet(loaded, "Elapsed")!, "A1")).toBe(1.5);
    }
  );
});

describe("display text survives a round trip in either date system", () => {
  // A reader turns a date-formatted serial into a `Date` using the workbook's epoch; display has to turn it
  // back with the same epoch. Calendar fields hid a mismatch, but an elapsed format reads the whole serial,
  // and `[h]:mm:ss` showed 36 hours as 35124 in a 1904 workbook.
  const formats = ["yyyy-mm-dd", "h:mm:ss", "yyyy-mm-dd hh:mm", "[h]:mm:ss", "[mm]:ss", "[h]"];
  const cases = [false, true].flatMap(date1904 => formats.map(fmt => [date1904, fmt] as const));

  it.each(cases)("1904=%s, %s", async (date1904, numFmt) => {
    const workbook = Workbook.create();
    workbook.properties.date1904 = date1904;
    const sheet = Workbook.addWorksheet(workbook, "Epoch");
    Cell.setValue(sheet, "A1", 45306.5);
    Cell.setValue(sheet, "A2", { formula: "45306.5", result: 45306.5 });
    Cell.setValue(sheet, "A3", 1.5);
    for (const address of ["A1", "A2", "A3"]) {
      Cell.setNumFmt(sheet, address, numFmt);
    }
    const before = ["A1", "A2", "A3"].map(address => Cell.getDisplayText(sheet, address));

    const loaded = Workbook.create();
    await Workbook.read(loaded, new Uint8Array(await Workbook.toBuffer(workbook)));
    const loadedSheet = Workbook.getWorksheet(loaded, "Epoch")!;
    expect(loaded.properties.date1904).toBe(date1904);
    expect(["A1", "A2", "A3"].map(address => Cell.getDisplayText(loadedSheet, address))).toEqual(
      before
    );
  });
});

describe("display text in a 1904 workbook", () => {
  it("reads a serial in the workbook's own epoch, before and after a round trip", async () => {
    const workbook = Workbook.create();
    workbook.properties.date1904 = true;
    const sheet = Workbook.addWorksheet(workbook, "Epoch");
    Cell.setValue(sheet, "A1", 45306.5);
    Cell.setNumFmt(sheet, "A1", "yyyy-mm-dd hh:mm");
    Cell.setValue(sheet, "A2", 1.5);
    Cell.setNumFmt(sheet, "A2", "[h]:mm:ss");
    const expected = ["2028-01-16 12:00", "36:00:00"];
    expect([Cell.getDisplayText(sheet, "A1"), Cell.getDisplayText(sheet, "A2")]).toEqual(expected);

    const loaded = Workbook.create();
    await Workbook.read(loaded, new Uint8Array(await Workbook.toBuffer(workbook)));
    const loadedSheet = Workbook.getWorksheet(loaded, "Epoch")!;
    expect([
      Cell.getDisplayText(loadedSheet, "A1"),
      Cell.getDisplayText(loadedSheet, "A2")
    ]).toEqual(expected);
  });
});
