import { extractAll } from "@archive/unzip/extract";
import { ZipArchive } from "@archive/zip/zip-archive";
import {
  Build,
  Document,
  DocxArgumentError,
  Query,
  type DocxDocument,
  type Hyperlink,
  type DeletedRun,
  type InsertedRun,
  type RangeMarker,
  type Paragraph,
  type Run,
  type StructuredDocumentTag,
  type StyleDef,
  type Table
} from "@word/index";
import {
  parseTableCellProperties,
  parseTableRowProperties
} from "@word/reader/table-properties-parsers";
import { checkDocxElementOrder } from "@word/writer/order-check";
import { parseXml } from "@xml/dom";
import { describe, expect, it } from "vitest";

/**
 * A property change record (`w:pPrChange`, `w:rPrChange`, `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`) holds the
 * whole previous state, so rejecting one puts the previous properties in place of the current ones. They used to be
 * merged over them, which left a property the change had added in place after its rejection. And a run's change was
 * listed but never settled: accepting or rejecting everything left it, and settling it by its id answered `false`.
 */

const revision = { id: 5, author: "A" };

const doc = (...body: DocxDocument["body"]): DocxDocument => ({ body });
const firstParagraph = (d: DocxDocument) => d.body[0] as Paragraph;
const firstTable = (d: DocxDocument) => d.body[0] as Table;
const firstRun = (p: Paragraph) => p.children[0] as Run;

describe("rejecting a property change restores the previous properties, not a merge", () => {
  const paragraph = () =>
    doc(
      Build.paragraph([Build.text("x")], {
        style: "Heading1",
        alignment: "center",
        markRunProperties: { bold: true },
        propertyChange: { revision, previousProperties: { style: "Normal" } }
      })
    );

  it("for a paragraph, keeping its mark, by rejecting all and by id", () => {
    for (const reject of [
      (d: DocxDocument) => Document.rejectAllRevisions(d),
      (d: DocxDocument) => Document.rejectRevision(d, 5)
    ]) {
      const d = paragraph();
      reject(d);
      expect(firstParagraph(d).properties).toEqual({
        style: "Normal",
        markRunProperties: { bold: true }
      });
    }
  });

  it("for a table, a row and a cell, keeping what the change record does not cover", () => {
    const d = doc(
      Build.table(
        [
          Build.row(
            [
              Build.cell("x", {
                width: { value: 2000, type: "dxa" },
                shading: { fill: "FF0000" },
                cellMerge: { vMerge: "cont", revision: { id: 9, author: "B" } },
                propertyChange: {
                  revision,
                  previousProperties: { width: { value: 1000, type: "dxa" } }
                }
              })
            ],
            {
              cantSplit: true,
              inserted: { revision: { id: 8, author: "B" } },
              tblPrEx: { alignment: "center" },
              propertyChange: { revision, previousProperties: { tableHeader: true } }
            }
          )
        ],
        {
          columnWidths: [100],
          alignment: "center",
          propertyChange: { revision, previousProperties: { alignment: "left" } }
        }
      )
    );
    expect(Document.rejectRevision(d, 5)).toBe(true);
    const table = firstTable(d);
    expect(table.properties).toMatchObject({ alignment: "left" });
    expect(table.properties?.propertyChange).toBeUndefined();
    const row = table.rows[0]!;
    // The row's insertion and its exceptions are not part of `w:trPrChange`: they stay.
    expect(row.properties).toEqual({
      tableHeader: true,
      inserted: { revision: { id: 8, author: "B" } },
      tblPrEx: { alignment: "center" }
    });
    // Nor is the cell's merge revision part of `w:tcPrChange`.
    expect(row.cells[0]!.properties).toEqual({
      width: { value: 1000, type: "dxa" },
      cellMerge: { vMerge: "cont", revision: { id: 9, author: "B" } }
    });
  });

  it("drops a change record a read set aside inside the previous properties", () => {
    const cell = parseTableCellProperties(
      parseXml(
        `<w:tcPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:vAlign w:val="center"/><w:tcPrChange w:id="5" w:author="A"><w:tcPr><w:vAlign w:val="top"/><w:tcPrChange w:id="6" w:author="B"><w:tcPr/></w:tcPrChange></w:tcPr></w:tcPrChange></w:tcPr>`
      ).root
    );
    const d = doc(
      Build.table([Build.row([{ ...Build.cell("x"), properties: cell }])], { columnWidths: [100] })
    );
    expect(Document.listUnrecognizedValues(d)).toHaveLength(1);
    Document.rejectAllRevisions(d);
    expect(firstTable(d).rows[0]!.cells[0]!.properties).toEqual({ verticalAlign: "top" });
    expect(Document.listUnrecognizedValues(d)).toEqual([]);
  });
});

describe("a run's property change", () => {
  const run = () =>
    doc(
      Build.paragraph([
        Build.text("x", {
          bold: true,
          propertyChange: { revision, previousProperties: { italic: true } }
        })
      ])
    );

  it("is listed, and settled by accepting or rejecting all", () => {
    expect(Query.listRevisions(run())).toEqual([
      expect.objectContaining({ id: 5, type: "runPropertyChange" })
    ]);

    const accepted = run();
    expect(Document.acceptAllRevisions(accepted)).toBe(1);
    expect(firstRun(firstParagraph(accepted)).properties).toEqual({ bold: true });

    const rejected = run();
    expect(Document.rejectAllRevisions(rejected)).toBe(1);
    expect(firstRun(firstParagraph(rejected)).properties).toEqual({ italic: true });
    expect(Query.listRevisions(rejected)).toEqual([]);
  });

  it("is settled by its id", () => {
    const accepted = run();
    expect(Document.acceptRevision(accepted, 6)).toBe(false);
    expect(Document.acceptRevision(accepted, 5)).toBe(true);
    expect(firstRun(firstParagraph(accepted)).properties).toEqual({ bold: true });

    const rejected = run();
    expect(Document.rejectRevision(rejected, 5)).toBe(true);
    expect(firstRun(firstParagraph(rejected)).properties).toEqual({ italic: true });
  });

  it("is settled inside a hyperlink", () => {
    const changed = Build.text("x", {
      bold: true,
      propertyChange: { revision, previousProperties: { italic: true } }
    });
    const link: Hyperlink = { type: "hyperlink", url: "https://example.com", children: [changed] };
    const d = doc(Build.paragraph([link]));
    expect(Query.listRevisions(d)).toEqual([
      expect.objectContaining({ id: 5, type: "runPropertyChange" })
    ]);
    expect(Document.rejectAllRevisions(d)).toBe(1);
    expect(Query.listRevisions(d)).toEqual([]);
  });
});

describe("the paragraph mark's property change", () => {
  const marked = () =>
    doc(
      Build.paragraph([Build.text("x")], {
        markRunProperties: {
          bold: true,
          propertyChange: { revision, previousProperties: { italic: true } }
        }
      })
    );

  it("is listed, and settled by all and by its id", () => {
    expect(Query.listRevisions(marked())).toEqual([
      expect.objectContaining({ id: 5, type: "runPropertyChange" })
    ]);

    const accepted = marked();
    expect(Document.acceptAllRevisions(accepted)).toBe(1);
    expect(firstParagraph(accepted).properties?.markRunProperties).toEqual({ bold: true });

    const rejected = marked();
    expect(Document.rejectRevision(rejected, 5)).toBe(true);
    expect(firstParagraph(rejected).properties?.markRunProperties).toEqual({ italic: true });
    expect(Query.listRevisions(rejected)).toEqual([]);
  });
});

describe("a run's property change inside an insertion", () => {
  const inserted = (): DocxDocument => {
    const run = Build.text("x", {
      bold: true,
      propertyChange: { revision, previousProperties: { italic: true } }
    });
    const insertion: InsertedRun = { type: "insertedRun", revision: { id: 9, author: "B" }, run };
    return doc(Build.paragraph([insertion]));
  };

  it("is listed beside the insertion, and settled by its id alone", () => {
    expect(
      Query.listRevisions(inserted())
        .map(entry => entry.type)
        .sort()
    ).toEqual(["insert", "runPropertyChange"]);
    const d = inserted();
    expect(Document.rejectRevision(d, 5)).toBe(true);
    // The insertion is still there, its run now as it was before the change.
    const insertion = firstParagraph(d).children[0] as InsertedRun;
    expect(insertion.type).toBe("insertedRun");
    expect(insertion.run.properties).toEqual({ italic: true });
  });

  it("is counted once by accepting all, as it was listed", () => {
    const d = inserted();
    expect(Document.acceptAllRevisions(d)).toBe(2);
    expect(firstRun(firstParagraph(d)).properties).toEqual({ bold: true });
  });
});

describe("a change record with no previous properties", () => {
  it("restores none on reject", () => {
    const d = doc(
      Build.paragraph([Build.text("x")], { alignment: "center", propertyChange: { revision } })
    );
    Document.rejectAllRevisions(d);
    expect(firstParagraph(d).properties).toEqual({});
  });
});

describe("what a paragraph's or run's previous properties have no place for", () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

  it("is set aside on read as not written back", async () => {
    const body = `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="5" w:author="A"><w:pPr><w:jc w:val="left"/><w:rPr><w:b/></w:rPr><w:pPrChange w:id="6" w:author="B"><w:pPr/></w:pPrChange></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:rPrChange w:id="7" w:author="A"><w:rPr><w:i/><w:rPrChange w:id="8" w:author="B"><w:rPr/></w:rPrChange></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p>`;
    const read = await readBody(body);
    const paragraph = firstParagraph(read);
    expect(paragraph.properties?.propertyChange?.previousProperties).toEqual(
      expect.objectContaining({ alignment: "left" })
    );
    expect(
      Document.listUnrecognizedValues(read).map(entry => [
        entry.path.replace(/^body\[0\]\./, ""),
        entry.writtenBack
      ])
    ).toEqual([
      ["properties.propertyChange.previousProperties.markRunProperties", false],
      ["properties.propertyChange.previousProperties.propertyChange", false],
      ["children[0].properties.propertyChange.previousProperties.propertyChange", false]
    ]);
    const written = await Document.toBuffer(read);
    const xml = new TextDecoder().decode(
      (await extractAll(written)).get("word/document.xml")!.data
    );
    expect(xml.match(/<w:pPrChange/g)).toHaveLength(1);
    expect(xml.match(/<w:rPrChange/g)).toHaveLength(1);
    expect(xml).not.toMatch(/<w:pPrChange[^>]*>[\s\S]*<w:rPr>[\s\S]*<\/w:pPrChange>/);
  });

  it("is refused when built and when written, rather than written where the schema has no place for it", async () => {
    for (const previousProperties of [
      { markRunProperties: { bold: true } },
      { sectionProperties: {} },
      { paragraphInsertion: { id: 1, author: "A" } },
      { propertyChange: { revision } }
    ] as const) {
      const properties = { propertyChange: { revision, previousProperties } };
      expect(() => Build.paragraph([Build.text("x")], properties)).toThrow(DocxArgumentError);
      const paragraph: Paragraph = { type: "paragraph", properties, children: [Build.text("x")] };
      await expect(Document.toBuffer(doc(paragraph))).rejects.toThrow(DocxArgumentError);
    }
    const run = {
      propertyChange: { revision, previousProperties: { propertyChange: { revision } } }
    };
    expect(() => Build.text("x", run)).toThrow(DocxArgumentError);
    const paragraph: Paragraph = {
      type: "paragraph",
      children: [{ properties: run, content: [{ type: "text", text: "x" }] }]
    };
    await expect(Document.toBuffer(doc(paragraph))).rejects.toThrow(DocxArgumentError);
  });

  it("is refused in a style's paragraph properties, which hold no mark or section", async () => {
    const style: StyleDef = {
      type: "paragraph",
      styleId: "S",
      name: "S",
      paragraphProperties: { markRunProperties: { bold: true } }
    };
    await expect(Document.toBuffer({ body: [], styles: [style] })).rejects.toThrow(
      DocxArgumentError
    );
  });

  async function readBody(body: string): Promise<DocxDocument> {
    const files = new Map<string, Uint8Array>();
    const blank = await Document.toBuffer(doc(Build.textParagraph("")));
    for (const [path, entry] of await extractAll(blank)) {
      files.set(path, entry.data);
    }
    files.set(
      "word/document.xml",
      new TextEncoder().encode(`<w:document ${W}><w:body>${body}</w:body></w:document>`)
    );
    const archive = new ZipArchive();
    for (const [path, data] of files) {
      archive.add(path, data);
    }
    return Document.read(await archive.toBuffer());
  }
});

describe("a row's previous properties, which hold no insertion, deletion or exceptions", () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

  it("set them aside on read, as not written back", () => {
    const row = parseTableRowProperties(
      parseXml(
        `<w:trPr ${W}><w:cantSplit/><w:trPrChange w:id="5" w:author="A"><w:trPr><w:ins w:id="2" w:author="B"/><w:del w:id="3" w:author="B"/></w:trPr></w:trPrChange></w:trPr>`
      ).root
    );
    const previous = row.propertyChange!.previousProperties!;
    expect(previous.inserted).toBeUndefined();
    expect(previous.deleted).toBeUndefined();
    const d = doc(
      Build.table([{ ...Build.row([Build.cell("x")]), properties: row }], { columnWidths: [100] })
    );
    expect(
      Document.listUnrecognizedValues(d).map(entry => [entry.path, entry.writtenBack])
    ).toEqual([
      ["body[0].rows[0].properties.propertyChange.previousProperties.inserted", false],
      ["body[0].rows[0].properties.propertyChange.previousProperties.deleted", false]
    ]);
  });

  it("refuse them when built", () => {
    for (const previousProperties of [
      { inserted: { revision } },
      { deleted: { revision } },
      { tblPrEx: { alignment: "center" } }
    ] as const) {
      expect(() =>
        Build.row([Build.cell("x")], { propertyChange: { revision, previousProperties } })
      ).toThrow(DocxArgumentError);
    }
  });
});

describe("a style's paragraph properties, which hold no mark or section", () => {
  it("set them aside on read, as not written back", async () => {
    const styles = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="S"><w:name w:val="S"/><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr><w:sectPr/></w:pPr></w:style></w:styles>`;
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
    const read = await Document.read(await archive.toBuffer());
    const style = read.styles!.find(s => s.styleId === "S")!;
    expect(style.paragraphProperties).toEqual(expect.objectContaining({ alignment: "center" }));
    expect(style.paragraphProperties?.markRunProperties).toBeUndefined();
    expect(
      Document.listUnrecognizedValues(read)
        .filter(entry => entry.path.includes("paragraphProperties"))
        .map(entry => [entry.path.replace(/^styles\[\d+\]\./, ""), entry.writtenBack])
    ).toEqual([
      ["paragraphProperties.markRunProperties", false],
      ["paragraphProperties.sectionProperties", false]
    ]);
    // Written back, the style is valid: neither element is there.
    const written = await Document.toBuffer(read);
    const xml = new TextDecoder().decode((await extractAll(written)).get("word/styles.xml")!.data);
    const own = xml.slice(
      xml.indexOf('w:styleId="S"'),
      xml.indexOf("</w:style>", xml.indexOf('w:styleId="S"'))
    );
    expect(own).not.toContain("<w:rPr><w:b/>");
    expect(own).not.toContain("w:sectPr");
  });
});

describe("settling a run's property change", () => {
  it("keeps the range markers riding on the run", () => {
    const start = Build.bookmarkStart(0, "here") as RangeMarker;
    const run: Run = {
      ...Build.text("x", { bold: true, propertyChange: { revision, previousProperties: {} } }),
      rangeMarkersBefore: [start]
    };
    const d = doc(Build.paragraph([run]));
    Document.rejectAllRevisions(d);
    const children = firstParagraph(d).children;
    expect(children.some(child => "type" in child && child.type === "bookmarkStart")).toBe(true);
    expect(Query.listRevisions(d).filter(entry => entry.type === "runPropertyChange")).toEqual([]);
  });

  it("counts a deleted run's own change once, as it is listed", () => {
    const deletion: DeletedRun = {
      type: "deletedRun",
      revision: { id: 9, author: "B" },
      run: Build.text("x", { bold: true, propertyChange: { revision, previousProperties: {} } })
    };
    const d = doc(Build.paragraph([deletion]));
    expect(Query.listRevisions(d)).toHaveLength(2);
    expect(Document.acceptAllRevisions(d)).toBe(2);
    expect(firstParagraph(d).children).toEqual([]);
  });
});

describe("settling by an id two revisions share", () => {
  it("settles both, as every kind of revision does", () => {
    // Ids are unique in a file Word wrote; one that is not still has every revision with the id settled.
    const insertion: InsertedRun = {
      type: "insertedRun",
      revision,
      run: Build.text("x", { bold: true, propertyChange: { revision, previousProperties: {} } })
    };
    const d = doc(Build.paragraph([insertion]));
    expect(Document.acceptRevision(d, 5)).toBe(true);
    expect(firstRun(firstParagraph(d)).properties).toEqual({ bold: true });
    expect(Query.listRevisions(d)).toEqual([]);
  });
});

describe("properties read from a file", () => {
  it("pass the builders' checks, a date outside xsd:dateTime included, and keep that date", async () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const body = `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="5" w:author="A" w:date="2024-01-01 10:00"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>x</w:t></w:r></w:p><w:tbl><w:tblPr><w:jc w:val="center"/><w:tblPrChange w:id="6" w:author="A" w:date="yesterday"><w:tblPr/></w:tblPrChange></w:tblPr><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/><w:trPrChange w:id="7" w:author="A"><w:trPr/></w:trPrChange></w:trPr><w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/><w:tcPrChange w:id="8" w:author="A"><w:tcPr/></w:tcPrChange></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    const files = new Map<string, Uint8Array>();
    const blank = await Document.toBuffer(doc(Build.textParagraph("")));
    for (const [path, entry] of await extractAll(blank)) {
      files.set(path, entry.data);
    }
    files.set(
      "word/document.xml",
      new TextEncoder().encode(`<w:document ${W}><w:body>${body}</w:body></w:document>`)
    );
    const archive = new ZipArchive();
    for (const [path, data] of files) {
      archive.add(path, data);
    }
    const read = await Document.read(await archive.toBuffer());
    const paragraph = firstParagraph(read);
    const table = read.body[1] as Table;
    expect(paragraph.properties?.propertyChange?.revision.date).toBeUndefined();

    const rebuilt = doc(
      Build.paragraph(paragraph.children, paragraph.properties),
      Build.table(
        table.rows.map(row =>
          Build.row(
            row.cells.map(cell => Build.cell(cell.content as Paragraph[], cell.properties)),
            row.properties
          )
        ),
        { ...table.properties, columnWidths: [100] }
      )
    );
    const xml = new TextDecoder().decode(
      (await extractAll(await Document.toBuffer(rebuilt))).get("word/document.xml")!.data
    );
    expect(xml).toContain('w:date="2024-01-01 10:00"');
    expect(xml).toContain('w:date="yesterday"');
  });
});

describe("an empty revision date", () => {
  it("is kept and written back as read", async () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const files = new Map<string, Uint8Array>();
    for (const [path, entry] of await extractAll(
      await Document.toBuffer(doc(Build.textParagraph("")))
    )) {
      files.set(path, entry.data);
    }
    files.set(
      "word/document.xml",
      new TextEncoder().encode(
        `<w:document ${W}><w:body><w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="5" w:author="A" w:date=""><w:pPr/></w:pPrChange></w:pPr></w:p></w:body></w:document>`
      )
    );
    const archive = new ZipArchive();
    for (const [path, data] of files) {
      archive.add(path, data);
    }
    const read = await Document.read(await archive.toBuffer());
    const xml = new TextDecoder().decode(
      (await extractAll(await Document.toBuffer(read))).get("word/document.xml")!.data
    );
    expect(xml).toMatch(/<w:pPrChange [^>]*w:date=""/);
  });
});

describe("a revision with no date", () => {
  it("is written without a w:date", async () => {
    const d = doc(
      Build.paragraph([Build.text("x")], { alignment: "center", propertyChange: { revision } })
    );
    const xml = new TextDecoder().decode(
      (await extractAll(await Document.toBuffer(d))).get("word/document.xml")!.data
    );
    expect(xml).toMatch(/<w:pPrChange [^>]*w:author="A"/);
    expect(xml).not.toMatch(/<w:pPrChange [^>]*w:date=/);
  });
});

describe("revisions in a content control's run-level content", () => {
  const control = (): DocxDocument => ({
    body: [
      {
        type: "sdt",
        content: [
          Build.text("x", {
            bold: true,
            propertyChange: { revision, previousProperties: { italic: true } }
          }),
          {
            type: "insertedRun",
            revision: { id: 9, author: "B" },
            run: Build.text("y")
          } as InsertedRun,
          {
            type: "sdt",
            content: [Build.paragraph([], { paragraphDeletion: { id: 11, author: "C" } })]
          }
        ]
      } as StructuredDocumentTag
    ]
  });
  const ids = (d: DocxDocument) =>
    Query.listRevisions(d)
      .map(entry => entry.id)
      .sort((a, b) => a - b);

  it("are listed, and every one is settled by accepting all", () => {
    expect(ids(control())).toEqual([11, 5, 9].sort((a, b) => a - b));
    const d = control();
    expect(Document.acceptAllRevisions(d)).toBe(3);
    expect(ids(d)).toEqual([]);
    const content = (d.body[0] as StructuredDocumentTag).content;
    expect(content[0]).toMatchObject({ properties: { bold: true } });
    expect(content[1]).toMatchObject({ content: [{ type: "text", text: "y" }] });
  });

  it("are settled by id, in the document and not in a copy", () => {
    const d = control();
    expect(Document.rejectRevision(d, 5)).toBe(true);
    expect((d.body[0] as StructuredDocumentTag).content[0]).toMatchObject({
      properties: { italic: true }
    });
    expect(ids(d)).toEqual([11, 9].sort((a, b) => a - b));
    expect(Document.rejectRevision(d, 9)).toBe(true);
    expect(ids(d)).toEqual([11]);
  });
});

describe("a section's property change", () => {
  const W =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  it("keeps what CT_SectPrBase has no place for aside on read, reported and not written", async () => {
    const files = new Map<string, Uint8Array>();
    for (const [path, entry] of await extractAll(
      await Document.toBuffer(doc(Build.textParagraph("")))
    )) {
      files.set(path, entry.data);
    }
    files.set(
      "word/document.xml",
      new TextEncoder().encode(
        `<w:document ${W}><w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:sectPrChange w:id="5" w:author="A"><w:sectPr><w:headerReference w:type="default" r:id="rId99"/><w:pgSz w:w="11906" w:h="16838"/><w:sectPrChange w:id="6" w:author="B"><w:sectPr/></w:sectPrChange></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`
      )
    );
    const archive = new ZipArchive();
    for (const [path, data] of files) {
      archive.add(path, data);
    }
    const read = await Document.read(await archive.toBuffer());
    const previous = read.sectionProperties!.propertyChange!.previousProperties!;
    expect(previous.headers).toBeUndefined();
    expect(previous.propertyChange).toBeUndefined();
    expect(
      Document.listUnrecognizedValues(read).map(entry => [entry.path, entry.writtenBack])
    ).toEqual([
      ["sectionProperties.propertyChange.previousProperties.headers", false],
      ["sectionProperties.propertyChange.previousProperties.propertyChange", false]
    ]);
    const parts = new Map<string, string>();
    for (const [path, entry] of await extractAll(await Document.toBuffer(read))) {
      if (path.endsWith(".xml")) {
        parts.set(path, new TextDecoder().decode(entry.data));
      }
    }
    const change = parts
      .get("word/document.xml")!
      .match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)![0];
    expect(change).not.toContain("headerReference");
    expect(change.match(/<w:sectPrChange/g)).toHaveLength(1);
    expect(checkDocxElementOrder(parts)).toEqual([]);
  });

  it("refuses one a caller sets there", async () => {
    for (const previousProperties of [
      { headers: [{ type: "default", rId: "rId9" }] },
      { propertyChange: { revision } }
    ] as const) {
      const d: DocxDocument = {
        body: [Build.textParagraph("x")],
        sectionProperties: { propertyChange: { revision, previousProperties } }
      };
      await expect(Document.toBuffer(d)).rejects.toThrow(DocxArgumentError);
    }
  });

  it("is listed and settled, at a section break and at the document's end", () => {
    const make = (): DocxDocument => ({
      body: [
        Build.paragraph([Build.text("x")], {
          sectionProperties: {
            pageSize: { width: 12240, height: 15840 },
            propertyChange: { revision: { id: 7, author: "A" }, previousProperties: {} }
          }
        }),
        Build.textParagraph("y")
      ],
      sectionProperties: {
        headers: [{ type: "default", rId: "rId9" }],
        pageSize: { width: 11906, height: 16838 },
        propertyChange: {
          revision: { id: 8, author: "A" },
          previousProperties: { pageSize: { width: 12240, height: 15840 } }
        }
      }
    });
    expect(
      Query.listRevisions(make())
        .filter(entry => entry.type === "sectionPropertyChange")
        .map(entry => entry.id)
    ).toEqual([7, 8]);

    const accepted = make();
    expect(Document.acceptAllRevisions(accepted)).toBe(2);
    expect(Query.listRevisions(accepted)).toEqual([]);

    const rejected = make();
    expect(Document.rejectRevision(rejected, 8)).toBe(true);
    // The previous section is restored; the header references, which it does not hold, stay.
    expect(rejected.sectionProperties).toEqual({
      headers: [{ type: "default", rId: "rId9" }],
      pageSize: { width: 12240, height: 15840 }
    });
    expect(Document.rejectAllRevisions(rejected)).toBe(1);
    expect((rejected.body[0] as Paragraph).properties?.sectionProperties).toEqual({});
  });
});
