/**
 * The xform finishing at the root's close tag ends dispatch, not validation: what follows the root still has to
 * decode and parse, as it does through `parseXformEvents`. A part with `<malformed` — or bytes that are not UTF-8 —
 * after `</worksheet>` used to read as a valid sheet.
 */
import { Archive } from "@archive/index";
import { Cell, Workbook } from "@excel/index";
import { parseXformStream } from "@excel/xlsx/xform/parse-xform";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { XmlParseError } from "@xml/errors";
import { describe, expect, it } from "vitest";

const encoder = new TextEncoder();

const STYLES =
  '<?xml version="1.0"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>';

function concat(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

describe("parseXformStream after the root closes", () => {
  it("still parses a well-formed part", async () => {
    await expect(parseXformStream(new StylesXform(), [STYLES])).resolves.toBeDefined();
    await expect(
      parseXformStream(new StylesXform(), [STYLES, "<!-- trailing -->", " \n"])
    ).resolves.toBeDefined();
  });

  it("rejects malformed text in the same chunk", async () => {
    await expect(
      parseXformStream(new StylesXform(), [STYLES + "<malformed"])
    ).rejects.toBeInstanceOf(XmlParseError);
  });

  it("rejects malformed text in a later chunk", async () => {
    await expect(
      parseXformStream(new StylesXform(), [STYLES, "<malformed"])
    ).rejects.toBeInstanceOf(XmlParseError);
    await expect(
      parseXformStream(new StylesXform(), [encoder.encode(STYLES), encoder.encode("<a></b>")])
    ).rejects.toBeInstanceOf(XmlParseError);
  });

  it("rejects bytes that are not UTF-8", async () => {
    await expect(
      parseXformStream(new StylesXform(), [concat(encoder.encode(STYLES), new Uint8Array([255]))])
    ).rejects.toBeInstanceOf(XmlParseError);
    await expect(
      parseXformStream(new StylesXform(), [encoder.encode(STYLES), new Uint8Array([255])])
    ).rejects.toBeInstanceOf(XmlParseError);
  });

  it("rejects a truncated UTF-8 sequence at the end", async () => {
    await expect(
      parseXformStream(new StylesXform(), [encoder.encode(STYLES), new Uint8Array([226, 130])])
    ).rejects.toBeInstanceOf(XmlParseError);
  });
});

describe("Workbook.read with text after a part's root", () => {
  async function withSheet(tail: Uint8Array): Promise<Uint8Array> {
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 1);
    const editor = await Archive.ZipEditor.open(await Workbook.toBuffer(wb));
    editor.set(
      "xl/worksheets/sheet1.xml",
      concat(
        encoder.encode(
          '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            "<sheetData/></worksheet>"
        ),
        tail
      )
    );
    return editor.toBuffer();
  }

  it("reads the part when nothing malformed follows", async () => {
    await expect(Workbook.read(await withSheet(new Uint8Array()))).resolves.toBeDefined();
  });

  it.each([
    ["malformed markup", encoder.encode("<malformed")],
    ["an invalid byte", new Uint8Array([255])],
    ["a truncated UTF-8 sequence", new Uint8Array([226, 130])]
  ])("rejects %s", async (_label, tail) => {
    await expect(Workbook.read(await withSheet(tail))).rejects.toThrow();
  });
});
