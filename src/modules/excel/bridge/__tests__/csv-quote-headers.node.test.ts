/**
 * `quoteHeaders` quotes the sheet's header row — its first row, the row `writeHeaders` names — exactly as `Csv.format`
 * quotes a header row: `true` every field, `false` (the default) only where needed, an array by column index, an
 * object by header name (the header cell's text). It had no effect. The header row's texts are also the names
 * `quoteColumns`' object form matches, as `Csv.format`'s headers are. Every writer agrees.
 */
import { mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";

import { ExcelArgumentError, Row, Workbook } from "@excel";
import { ExcelCsv } from "@excel/bridge/csv.node";
import { describe, expect, it } from "vitest";

function book(
  rows: unknown[][] = [
    ["name", "qty", "a,b"],
    ["x", 1, "y"],
    ["z", 2, "w"]
  ]
) {
  const wb = Workbook.create();
  Row.addMany(Workbook.addWorksheet(wb, "S"), rows as never);
  return wb;
}

async function everyWriter(wb: ReturnType<typeof book>, options: object): Promise<string[]> {
  const outputs: string[] = [ExcelCsv.format(wb, options)];
  outputs.push(new TextDecoder().decode(await ExcelCsv.toBuffer(wb, options)));
  outputs.push(new TextDecoder().decode(ExcelCsv.toBufferSync(wb, options)));
  let streamed = "";
  for await (const chunk of ExcelCsv.toStream(wb, options)) {
    streamed += new TextDecoder().decode(chunk as Uint8Array, { stream: true });
  }
  outputs.push(streamed);
  let written = "";
  const sink = new Writable({
    write(chunk, _encoding, callback) {
      written += chunk.toString();
      callback();
    }
  });
  await ExcelCsv.writeStream(wb, sink, options);
  outputs.push(written);
  const file = join(await mkdtemp(join(tmpdir(), "dm-quote-headers-")), "out.csv");
  await ExcelCsv.writeFile(wb, file, options);
  outputs.push(await readFile(file, "utf8"));
  return outputs;
}

describe("quoteHeaders quotes the sheet's header row, in every writer", () => {
  it.each<[string, object, string]>([
    ["true", { quoteHeaders: true }, '"name","qty","a,b"\nx,1,y\nz,2,w'],
    ["false", { quoteHeaders: false }, 'name,qty,"a,b"\nx,1,y\nz,2,w'],
    ["omitted", {}, 'name,qty,"a,b"\nx,1,y\nz,2,w'],
    ["an array, by index", { quoteHeaders: [false, true] }, 'name,"qty","a,b"\nx,1,y\nz,2,w'],
    [
      "an object, by header name",
      { quoteHeaders: { name: true, nope: true } },
      '"name",qty,"a,b"\nx,1,y\nz,2,w'
    ],
    [
      "with quoteColumns for the data rows",
      { quoteHeaders: true, quoteColumns: true },
      '"name","qty","a,b"\n"x","1","y"\n"z","2","w"'
    ],
    [
      "quoteColumns alone leaves the header row as needed",
      { quoteColumns: [true] },
      'name,qty,"a,b"\n"x",1,y\n"z",2,w'
    ],
    [
      "quoteColumns by header name",
      { quoteColumns: { qty: true } },
      'name,qty,"a,b"\nx,"1",y\nz,"2",w'
    ],
    [
      "with writeHeaders: false, which leaves the row out",
      { quoteHeaders: true, writeHeaders: false },
      "x,1,y\nz,2,w"
    ],
    [
      "with writeHeaders: false, quoteColumns by name still",
      { writeHeaders: false, quoteColumns: { name: true } },
      '"x",1,y\n"z",2,w'
    ],
    [
      "with another delimiter and line ending",
      { quoteHeaders: [true], delimiter: ";", lineEnding: "\r\n" },
      '"name";qty;a,b\r\nx;1;y\r\nz;2;w'
    ]
  ])("%s", async (_name, options, expected) => {
    expect(await everyWriter(book(), options)).toEqual(Array(6).fill(expected));
  });

  it("a sheet of the header row alone, and blank rows after it", async () => {
    expect(await everyWriter(book([["h1", "h2"]]), { quoteHeaders: true })).toEqual(
      Array(6).fill('"h1","h2"')
    );
    expect(await everyWriter(book([["h"], [], ["x"]]), { quoteHeaders: true })).toEqual(
      Array(6).fill('"h"\n\nx')
    );
  });

  it("is checked as Csv.format checks it", () => {
    expect(() => ExcelCsv.format(book(), { quoteHeaders: "yes" as never })).toThrow(
      ExcelArgumentError
    );
  });
});
