/**
 * Two `Workbook` members that pair with what exists:
 *
 * - `listUnrecognizedValues(wb)` lists the values a read kept outside their setting's vocabulary as they are kept
 *   *now* — after edits and `discardUnrecognizedValues` — in the shape and paths the discard filter sees. The read
 *   report's list was the only one, a snapshot taken at the read.
 * - `verifyPassword(wb, password)` checks a password against the workbook's protection, as
 *   `Worksheet.verifyPassword` does for a sheet's.
 */
import { Archive } from "@archive";
import { ExcelArgumentError, Workbook, Worksheet } from "@excel";
import { describe, expect, it } from "vitest";

/** A workbook whose sheet XML holds two values outside their vocabularies. */
async function withUnrecognized(): Promise<Uint8Array> {
  const wb = Workbook.create();
  Workbook.addWorksheet(wb, "Report");
  Workbook.addWorksheet(wb, "Data");
  const files = await Archive.unpackZip(await Workbook.toBuffer(wb));
  const entries = new Map<string, Uint8Array>();
  for (const [path, entry] of files) {
    let data = entry.data;
    if (path === "xl/worksheets/sheet1.xml" || path === "xl/worksheets/sheet2.xml") {
      const xml = new TextDecoder()
        .decode(data)
        .replace('orientation="portrait"', 'orientation="sideways"');
      data = new TextEncoder().encode(xml);
    }
    entries.set(path, data);
  }
  return Archive.packZip(entries);
}

describe("Workbook.listUnrecognizedValues", () => {
  it("lists what is kept now, in the discard filter's shape", async () => {
    const bytes = await withUnrecognized();
    const { workbook, unrecognizedValues } = await Workbook.readWithDiagnostics(bytes);
    const listed = Workbook.listUnrecognizedValues(workbook);
    expect(listed).toEqual(unrecognizedValues);
    expect(listed).toEqual([
      { sheet: "Report", path: "pageSetup.orientation", value: "sideways" },
      { sheet: "Data", path: "pageSetup.orientation", value: "sideways" }
    ]);

    const seen: unknown[] = [];
    Workbook.discardUnrecognizedValues(workbook, {
      filter: value => {
        seen.push(value);
        return value.sheet === "Report";
      }
    });
    expect(seen).toEqual(listed);
    expect(Workbook.listUnrecognizedValues(workbook)).toEqual([
      { sheet: "Data", path: "pageSetup.orientation", value: "sideways" }
    ]);

    // Setting the field takes over from the kept text.
    Workbook.getWorksheet(workbook, "Data")!.pageSetup.orientation = "landscape";
    expect(Workbook.listUnrecognizedValues(workbook)).toEqual([]);
  });

  it("hands out copies", async () => {
    const workbook = await Workbook.read(await withUnrecognized());
    const listed = Workbook.listUnrecognizedValues(workbook) as { value: string }[];
    listed[0]!.value = "edited";
    expect(Workbook.listUnrecognizedValues(workbook)[0]!.value).toBe("sideways");
  });

  it("is empty for a workbook made here", () => {
    expect(Workbook.listUnrecognizedValues(Workbook.create())).toEqual([]);
  });
});

describe("Workbook.verifyPassword", () => {
  it("answers for the password the workbook was protected with", async () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    expect(await Workbook.verifyPassword(wb, "pw")).toBe(false);
    await Workbook.protect(wb, { password: "pw", lockStructure: true, spinCount: 10 });
    expect(await Workbook.verifyPassword(wb, "pw")).toBe(true);
    expect(await Workbook.verifyPassword(wb, "PW")).toBe(false);
    const read = await Workbook.read(await Workbook.toBuffer(wb));
    expect(await Workbook.verifyPassword(read, "pw")).toBe(true);
    Workbook.unprotect(wb);
    expect(await Workbook.verifyPassword(wb, "pw")).toBe(false);
  });

  it("is false for protection without a password", async () => {
    const wb = Workbook.create();
    await Workbook.protect(wb, { lockStructure: true });
    expect(await Workbook.verifyPassword(wb, "")).toBe(false);
  });

  it("checks its argument as Worksheet.verifyPassword does", async () => {
    const wb = Workbook.create();
    await expect(Workbook.verifyPassword(wb, 42 as never)).rejects.toThrow(ExcelArgumentError);
    await expect(Workbook.verifyPassword(wb, 42 as never)).rejects.toThrow(
      /^Workbook\.verifyPassword: password/
    );
    const ws = Workbook.addWorksheet(wb, "S");
    await expect(Worksheet.verifyPassword(ws, 42 as never)).rejects.toThrow(ExcelArgumentError);
  });
});
