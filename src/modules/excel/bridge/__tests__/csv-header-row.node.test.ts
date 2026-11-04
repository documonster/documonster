/**
 * The sheet's first row is its header row — where `ExcelCsv.read` with `headers` puts the file's header — and
 * `writeHeaders: false` leaves it out, in every writer. `append: true` continues a file with every row, as v0.21.1 did:
 * nothing is guessed from the file's content. A line break is written first unless the file already ends with one,
 * read in the file's own encoding.
 */
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";

import { ExcelArgumentError, Row, Workbook } from "@excel";
import { ExcelCsv } from "@excel/bridge/csv.node";
import { describe, expect, it } from "vitest";

async function path(): Promise<string> {
  return join(await mkdtemp(join(tmpdir(), "dm-csv-header-")), "out.csv");
}

function book(
  rows: unknown[][] = [
    ["name", "qty"],
    ["a", 1],
    ["b", 2]
  ]
) {
  const wb = Workbook.create();
  Row.addMany(Workbook.addWorksheet(wb, "S"), rows as never);
  return wb;
}

async function collect(stream: AsyncIterable<Uint8Array>): Promise<string> {
  let text = "";
  for await (const chunk of stream) {
    text += new TextDecoder().decode(chunk, { stream: true });
  }
  return text;
}

describe("writeHeaders: false leaves out the sheet's first row", () => {
  const options = { writeHeaders: false };
  const expected = "a,1\nb,2";

  it("in format, toBuffer and toBufferSync", async () => {
    expect(ExcelCsv.format(book(), options)).toBe(expected);
    expect(new TextDecoder().decode(await ExcelCsv.toBuffer(book(), options))).toBe(expected);
    expect(new TextDecoder().decode(ExcelCsv.toBufferSync(book(), options))).toBe(expected);
  });

  it("in toStream and writeStream", async () => {
    expect(await collect(ExcelCsv.toStream(book(), options))).toBe(expected);
    let text = "";
    const sink = new Writable({
      write(chunk, _encoding, callback) {
        text += chunk.toString();
        callback();
      }
    });
    await ExcelCsv.writeStream(book(), sink, options);
    expect(text).toBe(expected);
  });

  it("in writeFile", async () => {
    const file = await path();
    await ExcelCsv.writeFile(book(), file, options);
    expect(await readFile(file, "utf8")).toBe(expected);
  });

  it("keeps the blank lines of empty rows after it, and writes everything by default", () => {
    expect(ExcelCsv.format(book([["h"], [], ["x"]]), options)).toBe("\nx");
    expect(ExcelCsv.format(book())).toBe("name,qty\na,1\nb,2");
    expect(ExcelCsv.format(book(), { writeHeaders: true })).toBe("name,qty\na,1\nb,2");
  });
});

describe("append: true appends every row", () => {
  it("guesses nothing from the file: a row equal to its first line is appended", async () => {
    const file = await path();
    await writeFile(file, "name,qty\na,1");
    await ExcelCsv.writeFile(book(), file, { append: true });
    expect(await readFile(file, "utf8")).toBe("name,qty\na,1\nname,qty\na,1\nb,2");
    await writeFile(file, "a,1\n");
    await ExcelCsv.writeFile(
      book([
        ["a", 1],
        ["c", 3]
      ]),
      file,
      { append: true }
    );
    expect(await readFile(file, "utf8")).toBe("a,1\na,1\nc,3");
  });

  it("leaves the header row out with writeHeaders: false", async () => {
    const file = await path();
    await ExcelCsv.writeFile(book(), file, { append: true });
    await ExcelCsv.writeFile(book(), file, { append: true, writeHeaders: false });
    expect(await readFile(file, "utf8")).toBe("name,qty\na,1\nb,2\na,1\nb,2");
  });

  it.each([
    ["\n", "x,9\n", "x,9\na,1\nb,2"],
    ["\r\n", "x,9\r\n", "x,9\r\na,1\r\nb,2"],
    ["\r", "x,9\r", "x,9\ra,1\rb,2"],
    ["\n", "x,9", "x,9\na,1\nb,2"]
  ])("in UTF-8, with %j line endings, after %j", async (lineEnding, existing, result) => {
    const file = await path();
    await writeFile(file, existing);
    await ExcelCsv.writeFile(book(), file, {
      append: true,
      writeHeaders: false,
      lineEnding: lineEnding as never
    });
    expect(await readFile(file, "utf8")).toBe(result);
  });

  it.each(["utf16le", "utf-16le", "ucs2"])(
    "reads the end of a %s file in that encoding",
    async encoding => {
      const file = await path();
      // `"\n"` is `0a 00` in UTF-16LE: its last byte is not a line feed.
      await writeFile(file, Buffer.from("\uFEFFname,qty\n", "utf16le"));
      await ExcelCsv.writeFile(book(), file, { append: true, writeHeaders: false, encoding });
      expect((await readFile(file)).toString("utf16le")).toBe("\uFEFFname,qty\na,1\nb,2");
      await writeFile(file, Buffer.from("name,qty", "utf16le"));
      await ExcelCsv.writeFile(book(), file, { append: true, writeHeaders: false, encoding });
      expect((await readFile(file)).toString("utf16le")).toBe("name,qty\na,1\nb,2");
    }
  );

  it("reads a latin1 file's end in latin1", async () => {
    const file = await path();
    await writeFile(file, Buffer.from("café,1\n", "latin1"));
    await ExcelCsv.writeFile(book([["thé", 2]]), file, { append: true, encoding: "latin1" });
    expect((await readFile(file)).toString("latin1")).toBe("café,1\nthé,2");
  });

  it("continues a file holding only a byte-order mark without a blank line", async () => {
    const file = await path();
    await writeFile(file, "\uFEFF");
    await ExcelCsv.writeFile(book(), file, { append: true });
    expect(await readFile(file, "utf8")).toBe("\uFEFFname,qty\na,1\nb,2");
  });

  it("writes a missing or empty file in full", async () => {
    const file = await path();
    await ExcelCsv.writeFile(book(), file, { append: true });
    expect(await readFile(file, "utf8")).toBe("name,qty\na,1\nb,2");
    await writeFile(file, "");
    await ExcelCsv.writeFile(book(), file, { append: true });
    expect(await readFile(file, "utf8")).toBe("name,qty\na,1\nb,2");
  });
});

describe("writeFile's encoding is a text encoding", () => {
  it.each(["hex", "base64", "base64url"])(
    "refuses %s, which is not a text encoding",
    async encoding => {
      const file = await path();
      await expect(ExcelCsv.writeFile(book(), file, { encoding })).rejects.toThrow(
        ExcelArgumentError
      );
    }
  );
});
