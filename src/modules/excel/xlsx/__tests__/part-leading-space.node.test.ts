/**
 * Some producers write a package part with white space before its XML declaration, where XML allows
 * none. Excel opens such a file, and so does every reader here: a part is parsed from its first `<`.
 * The XML parser itself stays strict.
 */
import { Readable } from "node:stream";

import { Archive } from "@archive/index";
import { Cell, Row, Stream, Workbook, Worksheet } from "@excel/index";
import { Xml } from "@xml/index";
import { describe, expect, it } from "vitest";

async function withLeadingSpace(bytes: Uint8Array): Promise<Uint8Array> {
  const parts = await Archive.unpackZip(bytes);
  const editor = await Archive.ZipEditor.open(bytes);
  const decoder = new TextDecoder();
  for (const [name, part] of parts) {
    if (name.endsWith(".xml") || name.endsWith(".rels")) {
      editor.set(name, "\r\n  " + decoder.decode(part.data));
    }
  }
  return editor.toBuffer();
}

describe("Excel readers take a part with white space before its XML declaration", () => {
  it("in Workbook.read, Workbook.readStream and Stream.WorkbookReader", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.addMany(ws, [
      ["a", "b"],
      [1, 2]
    ]);
    Cell.setValue(ws, "C1", { text: "link", hyperlink: "https://example.com" });
    const bytes = await withLeadingSpace(await Workbook.toBuffer(wb));
    expect(() => Xml.parse('\r\n  <?xml version="1.0"?><r/>')).toThrow();

    const read = await Workbook.read(bytes);
    const sheet = Workbook.getWorksheet(read, "S")!;
    expect(Cell.getValue(sheet, "B2")).toBe(2);
    expect(Worksheet.getName(sheet)).toBe("S");

    const streamed = await Workbook.readStream(Readable.from([bytes]));
    expect(Cell.getValue(Workbook.getWorksheet(streamed, "S")!, "A1")).toBe("a");

    const rows: unknown[] = [];
    for await (const reader of new Stream.WorkbookReader(Readable.from([bytes]))) {
      for await (const row of reader) {
        rows.push(row);
      }
    }
    expect(rows.length).toBe(2);
  });
});

describe("the parts read beside the workbook take white space before their declaration too", () => {
  const PERSONS =
    '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person displayName="Ann" id="{00000000-0000-0000-0000-000000000001}" userId="ann" providerId="None"/></personList>';

  it("keeps the threaded-comment persons", async () => {
    const bytes = await Workbook.toBuffer(Workbook.create());
    const editor = await Archive.ZipEditor.open(bytes);
    editor.set("xl/persons/person.xml", `\r\n  <?xml version="1.0"?>${PERSONS}`);
    const read = await Workbook.read(await editor.toBuffer());
    const persons = (Workbook.getModel(read) as { persons?: Array<{ displayName: string }> })
      .persons;
    expect(persons?.map(person => person.displayName)).toEqual(["Ann"]);
  });

  it("writes a theme back as it was read, its white space with it", async () => {
    const bytes = await withLeadingSpace(await Workbook.toBuffer(Workbook.create()));
    const written = await Workbook.toBuffer(await Workbook.read(bytes));
    const theme = (await Archive.unpackZip(written)).get("xl/theme/theme1.xml")!.data;
    expect(new TextDecoder().decode(theme).startsWith("\r\n  <?xml")).toBe(true);
  });

  it("in an XLSB package: its relationships and its notes' drawing", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 123);
    Cell.setNote(ws, "A1", "a note");
    const bytes = await withLeadingSpace(
      await Workbook.toBuffer(wb, { format: "xlsb", unsupported: "ignore" })
    );
    const read = await Workbook.read(bytes);
    const sheet = Workbook.getWorksheet(read, "S")!;
    expect(Cell.getValue(sheet, "A1")).toBe(123);
    expect(Cell.getNote(sheet, "A1")).toBeDefined();
  });
});
