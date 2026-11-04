/**
 * One write is a fixed point: reading what this library wrote and writing it again changes no part.
 *
 * Each fixture below used to move on the second write, which means the first write had said something the reader
 * then understood differently — a column hidden at width 0 written with no width and read back at the default, a
 * blank cell that overrode its row's style written as nothing, a font kept for a style nothing used. Any such
 * difference is a place where the model and the file disagree.
 *
 * `.node` because it reads fixtures off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { Column, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

const FIXTURES = [
  "src/modules/excel/__tests__/data/hyperlink-without-text.xlsx",
  "src/modules/excel/stream/__tests__/data/hyperlinks-cache.xlsx",
  "src/modules/excel/__tests__/data/malformed-comment.xlsx",
  "src/modules/excel/__tests__/data/many-defined-names.xlsx",
  "src/modules/excel/__tests__/data/content-type-element.xlsx",
  "src/modules/excel/examples/data/comments.xlsx"
];

const decoder = new TextDecoder();

async function write(bytes: Uint8Array): Promise<Uint8Array> {
  return Workbook.toBuffer(await Workbook.read(Workbook.create(), bytes));
}

async function parts(bytes: Uint8Array): Promise<Map<string, string>> {
  const entries = await extractAll(bytes);
  return new Map([...entries].map(([path, entry]) => [path, decoder.decode(entry.data)]));
}

describe("write fixed point", () => {
  it.each(FIXTURES)(
    "%s",
    async path => {
      const once = await write(new Uint8Array(await readFile(path)));
      const twice = await write(once);
      const first = await parts(once);
      const second = await parts(twice);
      expect([...second.keys()].sort()).toEqual([...first.keys()].sort());
      for (const [name, text] of first) {
        expect(second.get(name), name).toBe(text);
      }
    },
    60_000
  );
});

describe("formatting the first write used to lose", () => {
  it("keeps a row styled as a whole, and a default-styled blank under a styled row", async () => {
    const out = await parts(await write(new Uint8Array(await readFile(FIXTURES[0]!))));
    const sheet = out.get("xl/worksheets/sheet1.xml")!;
    // `<row r="2" s="150" customFormat="1"/>` in the source: a style and nothing else.
    expect(sheet).toMatch(/<row r="2" [^>]*s="\d+" customFormat="1"/);
    // `<row r="66" s="103">` holds `<c r="D66" s="29"/>`, where 29 is the default look: the cell is unformatted, and
    // written as nothing it would take the row's style.
    expect(sheet).toContain('<c r="D66" s="0"/>');
  });

  it("keeps a font's charset 0", async () => {
    const out = await parts(await write(new Uint8Array(await readFile(FIXTURES[5]!))));
    expect(out.get("xl/styles.xml")).toContain('<charset val="0"/>');
  });

  it("keeps a hidden column's zero width", async () => {
    const out = await parts(await write(new Uint8Array(await readFile(FIXTURES[3]!))));
    expect(out.get("xl/worksheets/sheet1.xml")).toMatch(
      /<col min="34" max="34" width="0" [^>]*hidden="1"/
    );
  });

  it("writes customWidth only where the source had it or a caller set a width", async () => {
    const wb = await Workbook.read(Workbook.create(), new Uint8Array(await readFile(FIXTURES[3]!)));
    // `<col min="35" max="16384" width="9.21875" style="5"/>`: a width, and no claim that the author chose it.
    let sheet = (await parts(await Workbook.toBuffer(wb))).get("xl/worksheets/sheet1.xml")!;
    expect(sheet).toMatch(/<col min="35" max="16384" width="9\.21875" style="\d+"\/>/);
    expect(sheet).toMatch(
      /<col min="1" max="1" width="4\.77734375" style="\d+" customWidth="1"\/>/
    );

    const ws = Workbook.getWorksheets(wb)[0]!;
    Column.setWidth(ws, 40, 20);
    sheet = (await parts(await Workbook.toBuffer(wb))).get("xl/worksheets/sheet1.xml")!;
    expect(sheet).toMatch(/<col min="40" max="40" width="20" style="\d+" customWidth="1"\/>/);
  });
});
