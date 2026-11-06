/**
 * The deprecated positional forms kept beside the options-object overloads behave exactly as the
 * positional signatures did — including for the placeholders plain JavaScript passes in a positional
 * slot (`null`, a numeric string, a boolean). Only a genuine options object selects the new form.
 *
 * Every expected value below was observed by running the same call against the last release
 * (c3b2407c), where the positional form was the only one.
 */
import { Anchor, Note, Stream, Workbook, Worksheet } from "@excel/index";
import { describe, expect, it } from "vitest";

/** A value the positional overload's types do not admit but plain JavaScript passes. */
function loose<T>(value: unknown): T {
  return value as T;
}

function sheet() {
  const wb = Workbook.create();
  return { wb, ws: Workbook.addWorksheet(wb, "S") };
}

function masked(protection: unknown): unknown {
  const p = protection as Record<string, unknown> | undefined;
  return (
    p && {
      ...p,
      ...(p.hashValue === undefined ? {} : { hashValue: "<hash>" }),
      ...(p.saltValue === undefined ? {} : { saltValue: "<salt>" })
    }
  );
}

describe("Worksheet.freeze / split positional placeholders", () => {
  const cases: [string, unknown, unknown, unknown, unknown][] = [
    [
      "null, 2",
      null,
      2,
      { state: "frozen", ySplit: 2, topLeftCell: "A3", activeCell: "A3" },
      { state: "split", ySplit: 2, activePane: "bottomLeft" }
    ],
    [
      "1, null",
      1,
      null,
      { state: "frozen", xSplit: 1, topLeftCell: "B1", activeCell: "B1" },
      { state: "split", xSplit: 1, activePane: "topRight" }
    ],
    [
      "undefined, 3",
      undefined,
      3,
      { state: "frozen", ySplit: 3, topLeftCell: "A4", activeCell: "A4" },
      { state: "split", ySplit: 3, activePane: "bottomLeft" }
    ],
    [
      '"2", "1"',
      "2",
      "1",
      { state: "frozen", xSplit: 2, ySplit: 1, topLeftCell: "C2", activeCell: "C2" },
      { state: "split", xSplit: 2, ySplit: 1, activePane: "bottomRight" }
    ],
    [
      "true, false",
      true,
      false,
      { state: "frozen", xSplit: 1, topLeftCell: "B1", activeCell: "B1" },
      { state: "split", xSplit: 1, activePane: "topRight" }
    ]
  ];

  for (const [label, a, b, frozen, split] of cases) {
    it(`(${label})`, () => {
      const { ws } = sheet();
      Worksheet.freeze(ws, loose<number>(a), loose<number>(b));
      expect(Worksheet.getModel(ws).views).toEqual([frozen]);
      Worksheet.split(ws, loose<number>(a), loose<number>(b));
      expect(Worksheet.getModel(ws).views).toEqual([split]);
    });
  }

  it("a null placeholder in the first slot freezes rows, as the repro reported", () => {
    const { ws } = sheet();
    Worksheet.split(ws, loose<number>(null), 300);
    expect(Worksheet.getModel(ws).views).toEqual([
      { state: "split", ySplit: 300, activePane: "bottomLeft" }
    ]);
  });
});

describe("protect positional placeholders", () => {
  it("Worksheet.protect(ws, null, options) rejects as it did", async () => {
    const { ws } = sheet();
    await expect(
      Worksheet.protect(ws, loose<string>(null), { selectLockedCells: false })
    ).rejects.toThrow(TypeError);
  });

  it("Worksheet.protect(ws, undefined | '', options)", async () => {
    const { ws } = sheet();
    await Worksheet.protect(ws, undefined, { selectLockedCells: false });
    expect(Worksheet.getModel(ws).sheetProtection).toEqual({
      sheet: true,
      selectLockedCells: false
    });
    await Worksheet.protect(ws, "", { selectLockedCells: false });
    expect(masked(Worksheet.getModel(ws).sheetProtection)).toEqual({
      sheet: true,
      algorithmName: "SHA-512",
      saltValue: "<salt>",
      spinCount: 100000,
      hashValue: "<hash>",
      selectLockedCells: false
    });
  });

  it("Workbook.protect(wb, null | undefined | '', options)", async () => {
    for (const password of [null, undefined, ""]) {
      const { wb } = sheet();
      await Workbook.protect(wb, loose<string>(password), { lockStructure: true });
      expect(Workbook.getModel(wb).protection).toEqual({ lockStructure: true });
    }
  });

  it("the streaming writers' protect(null | undefined | '', options)", async () => {
    for (const password of [null, undefined, ""]) {
      const writer = new Stream.WorkbookWriter({});
      await writer.protect(loose<string>(password), { lockStructure: true });
      expect((writer as unknown as { protection: unknown }).protection).toEqual({
        lockStructure: true
      });
    }

    const nullSheet = new Stream.WorkbookWriter({}).addWorksheet("S");
    await expect(nullSheet.protect(loose<string>(null), { formatCells: true })).rejects.toThrow(
      TypeError
    );
    // The synchronous pre-set still happens before the hash is computed.
    expect((nullSheet as unknown as { sheetProtection: unknown }).sheetProtection).toEqual({
      sheet: true
    });

    const plain = new Stream.WorkbookWriter({}).addWorksheet("S");
    await plain.protect(undefined, { formatCells: true });
    expect((plain as unknown as { sheetProtection: unknown }).sheetProtection).toEqual({
      sheet: true,
      formatCells: true
    });
  });
});

describe("Workbook.registerPerson positional placeholders", () => {
  const persons = (wb: ReturnType<typeof Workbook.create>) =>
    (wb as unknown as { _persons: { id: string }[] })._persons.map(({ id: _id, ...p }) => p);

  it("keeps a null user id and provider id as given", () => {
    let { wb } = sheet();
    Workbook.registerPerson(wb, "Ann", loose<string>(null), "AD");
    expect(persons(wb)).toEqual([{ displayName: "Ann", userId: null, providerId: "AD" }]);

    ({ wb } = sheet());
    Workbook.registerPerson(wb, "Ann", undefined, "AD");
    expect(persons(wb)).toEqual([{ displayName: "Ann", providerId: "AD" }]);

    ({ wb } = sheet());
    Workbook.registerPerson(wb, "Ann", loose<string>(5), loose<string>(null));
    expect(persons(wb)).toEqual([{ displayName: "Ann", userId: 5, providerId: null }]);
  });
});

describe("Note.create positional placeholders", () => {
  it("stores the positional author as given", () => {
    for (const author of [null, 5, ""]) {
      expect(Note.create("t", loose<string>(author))).toEqual({ note: "t", author });
    }
    const note = Note.create("t", undefined);
    expect(note).toHaveProperty("author", undefined);
    expect(note.note).toBe("t");
  });
});

describe("Anchor.create positional placeholders", () => {
  const at = (offset: unknown) => {
    const a = Anchor.create(sheet().ws, "B2", loose<number>(offset)) as unknown as {
      nativeCol: unknown;
      nativeRow: unknown;
    };
    return [a.nativeCol, a.nativeRow];
  };

  it("adds the positional offset exactly as before", () => {
    // A numeric string concatenated, as it always did: this pins the old form, it does not endorse it.
    expect(at("3")).toEqual(["23", "23"]);
    expect(at(null)).toEqual([2, 2]);
    expect(at(undefined)).toEqual([2, 2]);
    expect(at(2)).toEqual([4, 4]);
    expect(at(true)).toEqual([3, 3]);
  });
});
