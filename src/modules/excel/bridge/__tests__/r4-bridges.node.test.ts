/**
 * Round-4 review fixes in the bridges' file paths: a path is checked at the call, a file the system will not read keeps
 * its error as the cause and names the kind of file, `encoding` applies to a file on disk, and `append` appends every
 * row (`writeHeaders: false` leaves out the header row).
 */
import { createReadStream } from "node:fs";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pipeline } from "node:stream/promises";

import { ExcelArgumentError, ExcelFileError, Row, Workbook } from "@excel";
import { ExcelCsv } from "@excel/bridge/csv.node";
import { ExcelMarkdown } from "@excel/bridge/markdown.node";
import { describe, expect, it } from "vitest";

async function dir(): Promise<string> {
  return mkdtemp(join(tmpdir(), "dm-r4-bridges-"));
}

function book() {
  const wb = Workbook.create();
  Row.addMany(Workbook.addWorksheet(wb, "S"), [
    ["name", "qty"],
    ["a", 1]
  ]);
  return wb;
}

describe("X5 file paths are checked at the call, and file errors keep their cause", () => {
  it.each<[string, () => Promise<unknown>]>([
    ["Workbook.readFile", () => Workbook.readFile(42 as never)],
    ["Workbook.writeFile", () => Workbook.writeFile(book(), 42 as never)],
    ["ExcelCsv.readFile", () => ExcelCsv.readFile(42 as never)],
    ["ExcelCsv.writeFile", () => ExcelCsv.writeFile(book(), "" as never)],
    ["ExcelMarkdown.readFile", () => ExcelMarkdown.readFile(null as never)],
    ["ExcelMarkdown.writeFile", () => ExcelMarkdown.writeFile(book(), 42 as never)]
  ])("%s", async (where, run) => {
    const error = await run().then(
      () => undefined,
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf(ExcelArgumentError);
    expect((error as ExcelArgumentError).field).toBe("filename");
    expect((error as Error).message.startsWith(`${where}: `)).toBe(true);
  });

  it.each<[string, (path: string) => Promise<unknown>]>([
    ["CSV", path => ExcelCsv.readFile(path)],
    ["Markdown", path => ExcelMarkdown.readFile(path)],
    ["Excel", path => Workbook.readFile(path)]
  ])("a missing %s file names its kind and keeps the system's error", async (kind, run) => {
    const missing = join(await dir(), "missing");
    const error = await run(missing).then(
      () => undefined,
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf(ExcelFileError);
    expect((error as Error).message).toContain(`Failed to read ${kind} file`);
    expect(((error as Error).cause as { code?: string }).code).toBe("ENOENT");
  });
});

describe("X1 ExcelCsv.readFile decodes the file as options.encoding says", () => {
  it("reads a windows-1252 file", async () => {
    const path = join(await dir(), "data.csv");
    await writeFile(path, new Uint8Array([0x6e, 0x0a, 0x63, 0x61, 0x66, 0xe9, 0x0a]));
    const wb = await ExcelCsv.readFile(path, { encoding: "windows-1252" });
    expect(Row.getValues(Workbook.listWorksheets(wb)[0]!, 2)).toEqual(["café"]);
  });
});

describe("X8 ExcelCsv.writeFile checks its encoding", () => {
  it("takes Node's encodings, the WHATWG UTF spellings, and refuses the rest before the file is opened", async () => {
    const path = join(await dir(), "out.csv");
    await ExcelCsv.writeFile(book(), path, { encoding: "latin1" });
    await ExcelCsv.writeFile(book(), path, { encoding: "utf-16le" });
    expect(new TextDecoder("utf-16le").decode(await readFile(path))).toBe("name,qty\na,1");
    await expect(ExcelCsv.writeFile(book(), path, { encoding: "windows-1252" })).rejects.toThrow(
      expect.objectContaining({ field: "options.encoding" })
    );
  });
});

describe("R4-E6 ExcelCsv.writeFile({ append: true }) appends every row; writeHeaders: false leaves out the header row", () => {
  it("appends the sheet in full by default, and its data rows with writeHeaders: false", async () => {
    const path = join(await dir(), "out.csv");
    await ExcelCsv.writeFile(book(), path, { append: true });
    expect(await readFile(path, "utf8")).toBe("name,qty\na,1");
    await ExcelCsv.writeFile(book(), path, { append: true, writeHeaders: false });
    expect(await readFile(path, "utf8")).toBe("name,qty\na,1\na,1");
    await ExcelCsv.writeFile(book(), path, { append: true });
    expect(await readFile(path, "utf8")).toBe("name,qty\na,1\na,1\nname,qty\na,1");
  });

  it("does not add a blank line to a file that already ends with one", async () => {
    const path = join(await dir(), "out.csv");
    await writeFile(path, "name,qty\nz,9\n");
    await ExcelCsv.writeFile(book(), path, { append: true, writeHeaders: false });
    expect(await readFile(path, "utf8")).toBe("name,qty\nz,9\na,1");
  });

  it("appends a sheet of data rows only in full, losing none of them", async () => {
    const path = join(await dir(), "out.csv");
    await writeFile(path, "name,qty\nz,9");
    const rows = Workbook.create();
    Row.addMany(Workbook.addWorksheet(rows, "S"), [
      ["b", 2],
      ["c", 3]
    ]);
    await ExcelCsv.writeFile(rows, path, { append: true });
    expect(await readFile(path, "utf8")).toBe("name,qty\nz,9\nb,2\nc,3");
  });

  it("writes an empty existing file in full", async () => {
    const path = join(await dir(), "out.csv");
    await writeFile(path, "");
    await ExcelCsv.writeFile(book(), path, { append: true });
    expect(await readFile(path, "utf8")).toBe("name,qty\na,1");
  });
});

describe("X11 ExcelCsv.createImportStream is the Node Writable it is", () => {
  it("can be piped into", async () => {
    const path = join(await dir(), "in.csv");
    await writeFile(path, "a,b\n1,2\n");
    const wb = Workbook.create();
    const sink = ExcelCsv.createImportStream(wb, { sheetName: "In" });
    // Type-checks only because the return type includes `stream.Writable`.
    await pipeline(createReadStream(path), sink);
    expect(Row.getValues(Workbook.getWorksheet(wb, "In")!, 2)).toEqual([1, 2]);
  });
});
