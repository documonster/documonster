import { extractAll } from "@archive/unzip/extract";
import { ZipArchive } from "@archive/zip/zip-archive";
import { Build, Document, DocxArgumentError, type DocxDocument, type StyleDef } from "@word/index";
import { resolveTableCellStyle } from "@word/query/style-resolve";
import { describe, expect, it } from "vitest";

/**
 * A table style's conditional formats are named in the model for what they format (`oddRowBanding`) and in the file
 * by `ST_TblStyleOverrideType` (`band1Horz`). The reader used to put the file's token in the model's field, so style
 * resolution — which matches the model's names — never applied a read style's banding, first or last column or
 * corner formats, and `wholeTable` had no name at all.
 */

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const STYLES = `<w:styles ${W}><w:style w:type="table" w:styleId="Banded"><w:name w:val="Banded"/><w:tblStylePr w:type="wholeTable"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="nwCell"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="111111"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="diagonalCell"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr></w:tblStylePr></w:style></w:styles>`;

async function readWithStyles(styles: string): Promise<DocxDocument> {
  const files = new Map<string, Uint8Array>();
  const plain = await Document.toBuffer({ body: [Build.textParagraph("x")], styles: [] });
  for (const [path, entry] of await extractAll(plain)) {
    files.set(path, entry.data);
  }
  files.set("word/styles.xml", new TextEncoder().encode(styles));
  const archive = new ZipArchive();
  for (const [path, data] of files) {
    archive.add(path, data);
  }
  return Document.read(await archive.toBuffer());
}

const cellAt = (row: number, col: number) => ({
  rowIndex: row,
  colIndex: col,
  totalRows: 4,
  totalCols: 3
});

describe("a table style's conditional formats", () => {
  it("are read by the model's names; one with a token outside the schema is kept as read, not modelled", async () => {
    const read = await readWithStyles(STYLES);
    const style = read.styles!.find(s => s.styleId === "Banded")!;
    expect(style.tableStyleConditions!.map(c => c.type)).toEqual([
      "wholeTable",
      "oddRowBanding",
      "topLeftCell"
    ]);
    expect(style.preservedXml).toEqual([expect.stringContaining('w:type="diagonalCell"')]);
  });

  it("apply when resolved: the whole table, the banding and the corner, but not the unknown one", async () => {
    const read = await readWithStyles(STYLES);
    const look = { firstRow: false, firstColumn: false, noHBand: false };
    // Row 0 is the first band; row 1 the second, which has no format of its own.
    expect(resolveTableCellStyle(read, "Banded", look, cellAt(0, 1)).shading?.fill).toBe("DDDDDD");
    expect(resolveTableCellStyle(read, "Banded", look, cellAt(1, 1)).shading?.fill).toBe("EEEEEE");
    expect(
      resolveTableCellStyle(
        read,
        "Banded",
        { ...look, firstRow: true, firstColumn: true },
        cellAt(0, 0)
      ).shading?.fill
    ).toBe("111111");
  });

  it("are written back as the file's tokens, the unknown one as read", async () => {
    const written = await Document.toBuffer(await readWithStyles(STYLES));
    const xml = new TextDecoder().decode((await extractAll(written)).get("word/styles.xml")!.data);
    const types = [...xml.matchAll(/<w:tblStylePr\b[^>]*\bw:type="([^"]+)"/g)].map(m => m[1]);
    expect(types).toEqual(["wholeTable", "band1Horz", "nwCell", "diagonalCell"]);
  });

  it("are unchanged by discarding what a read kept: nothing kept could become a condition", async () => {
    const read = await readWithStyles(STYLES);
    const before = resolveTableCellStyle(read, "Banded", undefined, cellAt(1, 1)).shading?.fill;
    Document.discardUnrecognizedValues(read);
    expect(resolveTableCellStyle(read, "Banded", undefined, cellAt(1, 1)).shading?.fill).toBe(
      before
    );
    expect(before).toBe("EEEEEE");
  });

  it("refuse a type outside the model's names when a caller sets one", async () => {
    const style: StyleDef = {
      type: "table",
      styleId: "T",
      name: "T",
      tableStyleConditions: [{ type: "band1Horz" as never }]
    };
    await expect(Document.toBuffer({ body: [], styles: [style] })).rejects.toThrow(
      DocxArgumentError
    );
  });
});

describe("a conditional format with no w:type, or an empty one", () => {
  for (const [name, attribute] of [
    ["no w:type", ""],
    ["an empty w:type", ' w:type=""']
  ] as const) {
    it(`with ${name} is kept as read: not modelled, not resolved, and written back unchanged`, async () => {
      const read = await readWithStyles(
        `<w:styles ${W}><w:style w:type="table" w:styleId="T"><w:name w:val="T"/><w:tblStylePr${attribute}><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="ABCDEF"/></w:tcPr></w:tblStylePr></w:style></w:styles>`
      );
      const style = read.styles!.find(s => s.styleId === "T")!;
      expect(style.tableStyleConditions).toBeUndefined();
      Document.discardUnrecognizedValues(read);
      expect(resolveTableCellStyle(read, "T", undefined, cellAt(1, 1)).shading).toBeUndefined();
      const xml = new TextDecoder().decode(
        (await extractAll(await Document.toBuffer(read))).get("word/styles.xml")!.data
      );
      const tag = xml.match(/<w:tblStylePr\b[^>]*>(?=<w:tcPr>)/)?.[0];
      expect(tag).toBeDefined();
      expect(/w:type=/.test(tag!)).toBe(attribute !== "");
    });
  }
});

describe("a conditional format's row properties", () => {
  it("refuse table property exceptions, which a style's w:trPr has nothing beside it to hold", async () => {
    const style: StyleDef = {
      type: "table",
      styleId: "T",
      name: "T",
      tableStyleConditions: [{ type: "firstRow", rowProperties: { tblPrEx: { indent: 123 } } }]
    };
    await expect(Document.toBuffer({ body: [], styles: [style] })).rejects.toThrow(
      DocxArgumentError
    );
    await expect(Document.toBuffer({ body: [], styles: [style] })).rejects.toThrow(/tblPrEx/);
  });
});
