import { extractAll } from "@archive/unzip/extract";
import { ZipArchive } from "@archive/zip/zip-archive";
import {
  Build,
  Document,
  DocxArgumentError,
  Query,
  type DocxDocument,
  type Table,
  type TableProperties,
  type TablePropertyExceptions
} from "@word/index";
import { TABLE_ONLY_ELEMENTS } from "@word/reader/table-properties-parsers";
import { OOXML_SEQUENCES } from "@word/writer/ooxml-sequences";
import { checkDocxElementOrder } from "@word/writer/order-check";
import { describe, expect, it } from "vitest";

/**
 * A row's table property exceptions (`w:tblPrEx`) are `CT_TblPrEx`: nine of a table's properties and a change record.
 * The model typed them as a whole table's properties, so a style or a floating position set there was accepted and
 * then dropped by the writer, and `w:tblPrExChange` was never read.
 */

const revision = { id: 7, author: "A", date: "2026-01-02T03:04:05Z" };

async function parts(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for (const [path, entry] of await extractAll(bytes)) {
    out.set(path, entry.data);
  }
  return out;
}

async function xmlParts(bytes: Uint8Array): Promise<Map<string, string>> {
  const decoded = new Map<string, string>();
  for (const [path, data] of await parts(bytes)) {
    if (path.endsWith(".xml")) {
      decoded.set(path, new TextDecoder().decode(data));
    }
  }
  return decoded;
}

function withExceptions(tblPrEx: TablePropertyExceptions): DocxDocument {
  return {
    body: [Build.table([Build.row([Build.cell("x")], { tblPrEx })], { columnWidths: [100] })]
  };
}

function exceptionsOf(doc: DocxDocument): TablePropertyExceptions | undefined {
  return (doc.body[0] as Table).rows[0]!.properties?.tblPrEx;
}

describe("a row's table property exceptions", () => {
  it("are what CT_TblPrEx holds besides its change record", () => {
    const table = OOXML_SEQUENCES["w:CT_TblPr"].flat();
    const exceptions: readonly string[] = OOXML_SEQUENCES["w:CT_TblPrEx"].flat();
    expect(table.filter(name => !exceptions.includes(name))).toEqual([...TABLE_ONLY_ELEMENTS]);
  });

  it("round-trip with a change record, in schema order", async () => {
    const tblPrEx: TablePropertyExceptions = {
      width: { value: 5000, type: "pct" },
      alignment: "center",
      indent: 120,
      borders: { top: { style: "single", size: 4 } },
      shading: { fill: "D9E2F3" },
      look: { firstRow: true },
      propertyChange: { revision, previousProperties: { alignment: "left", indent: 0 } }
    };
    const bytes = await Document.toBuffer(withExceptions(tblPrEx));
    const written = await xmlParts(bytes);
    expect(checkDocxElementOrder(written)).toEqual([]);
    expect(written.get("word/document.xml")).toContain("<w:tblPrExChange");

    const back = exceptionsOf(await Document.read(bytes))!;
    expect(back).toMatchObject({
      alignment: "center",
      indent: 120,
      shading: { fill: "D9E2F3" },
      propertyChange: {
        revision: { id: 7, author: "A" },
        previousProperties: { alignment: "left", indent: 0 }
      }
    });
  });

  it("refuse a property of the whole table, rather than drop it", async () => {
    expect(() =>
      Build.row([Build.cell("x")], { tblPrEx: { style: "Grid" } as TablePropertyExceptions })
    ).toThrow(DocxArgumentError);
    // A model typed the old way still compiles — a table's properties are assignable — and is refused on write.
    const legacy: TableProperties = { style: "Grid", alignment: "center" };
    const doc: DocxDocument = {
      body: [
        { type: "table", rows: [{ properties: { tblPrEx: legacy }, cells: [{ content: [] }] }] }
      ]
    };
    await expect(Document.toBuffer(doc)).rejects.toThrow(/tblPrEx\.style|style must be absent/);
  });

  it("refuse a change record inside a change record's previous exceptions", async () => {
    const nested: TablePropertyExceptions = {
      propertyChange: {
        revision,
        previousProperties: {
          propertyChange: { revision }
        } as TablePropertyExceptions
      }
    };
    expect(() => Build.row([Build.cell("x")], { tblPrEx: nested })).toThrow(DocxArgumentError);
    const doc: DocxDocument = {
      body: [
        { type: "table", rows: [{ properties: { tblPrEx: nested }, cells: [{ content: [] }] }] }
      ]
    };
    await expect(Document.toBuffer(doc)).rejects.toThrow(DocxArgumentError);
  });

  it("read from a file: a diagonal in their borders is set aside as a table's is", async () => {
    const files = await parts(
      await Document.toBuffer(withExceptions({ borders: { top: { style: "single", size: 4 } } }))
    );
    const body = new TextDecoder().decode(files.get("word/document.xml")!);
    const injected = body.replace(
      /(<w:tblPrEx>[\s\S]*?)<\/w:tblBorders>/,
      '$1<w:tl2br w:val="single"/></w:tblBorders>'
    );
    expect(injected).not.toBe(body);
    files.set("word/document.xml", new TextEncoder().encode(injected));
    const archive = new ZipArchive();
    for (const [path, data] of files) {
      archive.add(path, data);
    }
    const read = await Document.read(await archive.toBuffer());
    expect(exceptionsOf(read)?.borders?.tl2br).toBeUndefined();
    expect(Document.listUnrecognizedValues(read)).toEqual([
      expect.objectContaining({
        path: expect.stringMatching(/tblPrEx\.borders\.tl2br$/),
        writtenBack: false
      })
    ]);
    const written = (await xmlParts(await Document.toBuffer(read))).get("word/document.xml")!;
    expect(written).not.toContain("tl2br");
  });

  it("read from a file: a whole-table property is set aside and reported, not typed or written", async () => {
    const files = await parts(await Document.toBuffer(withExceptions({ alignment: "center" })));
    const body = new TextDecoder().decode(files.get("word/document.xml")!);
    const injected = body.replace(
      /<w:tblPrEx>([\s\S]*?)<\/w:tblPrEx>/,
      '<w:tblPrEx><w:tblStyle w:val="Grid"/>$1<w:tblPrExChange w:id="3" w:author="B"><w:tblPrEx><w:jc w:val="right"/><w:tblPrExChange w:id="4" w:author="C"><w:tblPrEx/></w:tblPrExChange></w:tblPrEx></w:tblPrExChange></w:tblPrEx>'
    );
    expect(injected).not.toBe(body);
    files.set("word/document.xml", new TextEncoder().encode(injected));
    const archive = new ZipArchive();
    for (const [path, data] of files) {
      archive.add(path, data);
    }

    const read = await Document.read(await archive.toBuffer());
    const exceptions = exceptionsOf(read)!;
    expect(exceptions).toMatchObject({
      alignment: "center",
      propertyChange: {
        revision: { id: 3, author: "B" },
        previousProperties: { alignment: "right" }
      }
    });
    expect("style" in exceptions).toBe(false);
    const kept = Document.listUnrecognizedValues(read);
    expect(kept).toEqual([
      expect.objectContaining({
        path: expect.stringMatching(/tblPrEx\.tblStyle$/),
        writtenBack: false
      }),
      expect.objectContaining({
        path: expect.stringMatching(/tblPrEx\.propertyChange\.previousProperties\.propertyChange$/),
        writtenBack: false
      })
    ]);

    const written = await xmlParts(await Document.toBuffer(read));
    const document = written.get("word/document.xml")!;
    expect(document).not.toContain("tblStyle");
    expect(document.match(/<w:tblPrExChange/g)).toHaveLength(1);
    expect(checkDocxElementOrder(written)).toEqual([]);
  });

  it("are listed, accepted and rejected as a revision", () => {
    const change = (): DocxDocument =>
      withExceptions({
        alignment: "center",
        propertyChange: { revision, previousProperties: { alignment: "left" } }
      });

    expect(Query.listRevisions(change())).toContainEqual(
      expect.objectContaining({ id: 7, type: "tablePropertyExceptionsChange", author: "A" })
    );

    const accepted = change();
    expect(Document.acceptAllRevisions(accepted)).toBe(1);
    expect(exceptionsOf(accepted)).toEqual({ alignment: "center" });

    const rejected = change();
    expect(Document.rejectAllRevisions(rejected)).toBe(1);
    expect(exceptionsOf(rejected)).toEqual({ alignment: "left" });

    const one = change();
    expect(Document.rejectRevision(one, 8)).toBe(false);
    expect(Document.rejectRevision(one, 7)).toBe(true);
    expect(exceptionsOf(one)).toEqual({ alignment: "left" });
    expect(Query.listRevisions(one)).toEqual([]);

    const kept = change();
    expect(Document.acceptRevision(kept, 7)).toBe(true);
    expect(exceptionsOf(kept)).toEqual({ alignment: "center" });
  });

  it("are replaced by the previous ones on reject, which hold the whole previous state", () => {
    // The change added a width and moved the alignment; rejecting it takes both back.
    const doc = withExceptions({
      alignment: "center",
      width: { value: 5000, type: "pct" },
      propertyChange: { revision, previousProperties: { alignment: "left" } }
    });
    Document.rejectAllRevisions(doc);
    expect(exceptionsOf(doc)).toEqual({ alignment: "left" });
  });

  it("are listed and rejected in a header too", () => {
    const header: DocxDocument = {
      body: [Build.textParagraph("body")],
      headers: new Map([
        [
          "rId10",
          {
            rId: "rId10",
            content: {
              children: [
                Build.table(
                  [
                    Build.row([Build.cell("x")], {
                      tblPrEx: {
                        alignment: "center",
                        propertyChange: { revision, previousProperties: {} }
                      }
                    })
                  ],
                  { columnWidths: [100] }
                )
              ]
            }
          }
        ]
      ]),
      sectionProperties: { headers: [{ type: "default", rId: "rId10" }] }
    };
    expect(Query.listRevisions(header)).toContainEqual(
      expect.objectContaining({ id: 7, type: "tablePropertyExceptionsChange" })
    );
    expect(Document.rejectRevision(header, 7)).toBe(true);
    const table = header.headers!.get("rId10")!.content.children[0] as Table;
    expect(table.rows[0]!.properties?.tblPrEx).toEqual({});
  });
});
