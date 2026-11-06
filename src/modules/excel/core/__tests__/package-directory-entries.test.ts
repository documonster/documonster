/**
 * ZIP directory entries are not OPC parts.
 *
 * A repacked workbook may carry an entry per folder (`xl/`, `docProps/`). A part
 * name must not end with "/" (ECMA-376 Part 2 §9.1.1.1), so reading one as a part
 * and writing it back — or declaring it in `[Content_Types].xml` — yields a
 * package Excel refuses.
 */

import { extractAll } from "@archive/unzip/extract";
import { ZipArchive } from "@archive/zip/zip-archive";
import { Cell, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

async function withDirectoryEntries(bytes: Uint8Array): Promise<Uint8Array> {
  const files = await extractAll(bytes);
  const archive = new ZipArchive();
  const folders = new Set<string>();
  for (const path of files.keys()) {
    const segments = path.split("/");
    for (let i = 1; i < segments.length; i++) {
      folders.add(`${segments.slice(0, i).join("/")}/`);
    }
  }
  for (const folder of folders) {
    archive.addDirectory(folder);
  }
  for (const [path, file] of files) {
    archive.add(path, file.data);
  }
  return archive.toBuffer();
}

function sample(): Workbook.Handle {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Data");
  Cell.setValue(ws, "A1", "hello");
  return wb;
}

describe.each(["xlsx", "xlsb"] as const)("ZIP directory entries in an %s", format => {
  it("read → write keeps no directory entry and declares none", async () => {
    const input = await withDirectoryEntries(await Workbook.toBuffer(sample(), { format }));
    const inputFiles = await extractAll(input);
    expect([...inputFiles.keys()]).toEqual(expect.arrayContaining(["xl/", "docProps/"]));

    const wb = Workbook.create();
    await Workbook.read(wb, input, { format });
    // An edit, so the package is rebuilt rather than an unchanged XLSB passed through verbatim.
    Cell.setValue(Workbook.getWorksheet(wb, "Data")!, "B1", "edited");
    const output = await extractAll(await Workbook.toBuffer(wb, { format }));

    expect([...output.keys()].filter(p => p.endsWith("/"))).toEqual([]);
    expect([...output.values()].filter(f => f.type === "directory")).toEqual([]);
    const contentTypes = new TextDecoder().decode(output.get("[Content_Types].xml")!.data);
    expect(contentTypes).not.toMatch(/PartName="[^"]*\/"/);
    expect(contentTypes).toMatch(/sheet(\.binary\.macroEnabled)?\.main/);
  });
});
