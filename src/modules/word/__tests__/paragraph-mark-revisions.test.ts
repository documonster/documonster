import { extractAll } from "@archive/unzip/extract";
import { ZipArchive } from "@archive/zip/zip-archive";
import {
  Build,
  Document,
  DocxArgumentError,
  Query,
  type DocxDocument,
  type Paragraph,
  type ParagraphProperties,
  type RangeMarker,
  type StructuredDocumentTag,
  type StyleDef,
  type Table
} from "@word/index";
import { checkDocxElementOrder } from "@word/writer/order-check";
import { describe, expect, it } from "vitest";

/**
 * A paragraph mark's own revisions (`w:pPr/w:rPr/w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`). The paragraph's
 * insertion and deletion were read and written but neither listed nor settled, and its move was dropped on read.
 * Settling one that removes the mark joins the paragraph with the next, as Word does.
 */

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const rev = (id: number) => ({ id, author: "A" });

const para = (text: string, properties?: ParagraphProperties): Paragraph =>
  Build.paragraph([Build.text(text)], properties);
const textOf = (block: unknown) =>
  (block as Paragraph).children
    .map(child =>
      "content" in child
        ? (child.content as readonly { text?: string }[]).map(c => c.text ?? "").join("")
        : ""
    )
    .join("");

async function read(body: string): Promise<DocxDocument> {
  const files = new Map<string, Uint8Array>();
  const blank = await Document.toBuffer({ body: [Build.textParagraph("")] });
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

async function documentXml(doc: DocxDocument): Promise<Map<string, string>> {
  const parts = new Map<string, string>();
  for (const [path, entry] of await extractAll(await Document.toBuffer(doc))) {
    if (path.endsWith(".xml")) {
      parts.set(path, new TextDecoder().decode(entry.data));
    }
  }
  return parts;
}

describe("a paragraph mark's revisions", () => {
  it("are read — the move included — and written back in CT_ParaRPr's order", async () => {
    const doc = await read(
      `<w:p><w:pPr><w:rPr><w:moveFrom w:id="3" w:author="A"/><w:b/></w:rPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:moveTo w:id="4" w:author="A"/></w:rPr></w:pPr><w:r><w:t>b</w:t></w:r></w:p>`
    );
    expect((doc.body[0] as Paragraph).properties).toMatchObject({
      paragraphMoveFrom: { id: 3, author: "A" },
      markRunProperties: { bold: true }
    });
    expect((doc.body[1] as Paragraph).properties?.paragraphMoveTo).toMatchObject({ id: 4 });
    const parts = await documentXml(doc);
    expect(parts.get("word/document.xml")).toMatch(/<w:rPr><w:moveFrom [^>]*\/><w:b\/><\/w:rPr>/);
    expect(checkDocxElementOrder(parts)).toEqual([]);
  });

  it("are listed", () => {
    const doc: DocxDocument = {
      body: [
        para("a", { paragraphInsertion: rev(1) }),
        para("b", { paragraphDeletion: rev(2) }),
        para("c", { paragraphMoveFrom: rev(3) }),
        para("d", { paragraphMoveTo: rev(4) }),
        para("e")
      ]
    };
    expect(Query.listRevisions(doc).map(entry => [entry.id, entry.type])).toEqual([
      [1, "insert"],
      [2, "delete"],
      [3, "moveFrom"],
      [4, "moveTo"]
    ]);
  });

  it("join the paragraph with the next when the mark goes, with the next one's properties, as Word does", () => {
    const doc: DocxDocument = {
      body: [
        para("one ", {
          alignment: "center",
          markRunProperties: { italic: true },
          paragraphDeletion: rev(2)
        }),
        para("two", { alignment: "right", markRunProperties: { bold: true } }),
        para("three")
      ]
    };
    expect(Document.acceptAllRevisions(doc)).toBe(1);
    expect(doc.body).toHaveLength(2);
    expect(textOf(doc.body[0])).toBe("one two");
    // Checked against Word: accepting the deletion of a paragraph's mark gives the joined paragraph the properties of
    // the mark that is left — the second paragraph's.
    expect((doc.body[0] as Paragraph).properties).toEqual({
      alignment: "right",
      markRunProperties: { bold: true }
    });
  });

  it("join with the next one's numbering, whichever of the two is numbered, as Word does", () => {
    const list = { numbering: { numId: 1, level: 0 } };
    const joined = (first: object, second: object, settle: "accept" | "reject"): unknown => {
      const doc: DocxDocument = { body: [para("one", first), para("two", second)] };
      if (settle === "accept") {
        Document.acceptAllRevisions(doc);
      } else {
        Document.rejectAllRevisions(doc);
      }
      return (doc.body[0] as Paragraph).properties?.numbering;
    };
    // Checked against Word: accepting a numbered paragraph's mark deletion drops its numbering; the next one's stays.
    expect(joined({ ...list, paragraphDeletion: rev(1) }, {}, "accept")).toBeUndefined();
    expect(joined({ paragraphDeletion: rev(1) }, list, "accept")).toEqual(list.numbering);
    // Rejecting an inserted mark joins the same way.
    expect(joined({ ...list, paragraphInsertion: rev(1) }, {}, "reject")).toBeUndefined();
  });

  it("keep the paragraphs apart when the mark stays", () => {
    const deleted: DocxDocument = { body: [para("a", { paragraphDeletion: rev(2) }), para("b")] };
    expect(Document.rejectAllRevisions(deleted)).toBe(1);
    expect(deleted.body.map(textOf)).toEqual(["a", "b"]);
    expect((deleted.body[0] as Paragraph).properties).toEqual({});

    const inserted: DocxDocument = { body: [para("a", { paragraphInsertion: rev(1) }), para("b")] };
    expect(Document.acceptAllRevisions(inserted)).toBe(1);
    expect(inserted.body.map(textOf)).toEqual(["a", "b"]);
    expect((inserted.body[0] as Paragraph).properties).toEqual({});
    expect(Query.listRevisions(inserted)).toEqual([]);
  });

  it("join on rejecting an insertion or a move here, and on accepting a move away", () => {
    for (const [properties, settle] of [
      [{ paragraphInsertion: rev(1) }, Document.rejectAllRevisions],
      [{ paragraphMoveTo: rev(1) }, Document.rejectAllRevisions],
      [{ paragraphMoveFrom: rev(1) }, Document.acceptAllRevisions]
    ] as const) {
      const doc: DocxDocument = { body: [para("a", properties), para("b")] };
      expect(settle(doc)).toBe(1);
      expect(doc.body.map(textOf)).toEqual(["ab"]);
      expect(Query.listRevisions(doc)).toEqual([]);
    }
  });

  it("join a run of removed marks into one paragraph, counting each", () => {
    const doc: DocxDocument = {
      body: [
        para("a", { paragraphDeletion: rev(1) }),
        para("b", { paragraphDeletion: rev(2) }),
        para("c")
      ]
    };
    expect(Document.acceptAllRevisions(doc)).toBe(2);
    expect(doc.body.map(textOf)).toEqual(["abc"]);
  });

  it("cannot join across a table: only the revision goes", () => {
    const doc: DocxDocument = {
      body: [
        para("a", { paragraphDeletion: rev(1) }),
        Build.table([Build.row([Build.cell("x")])], { columnWidths: [100] }),
        para("b")
      ]
    };
    expect(Document.acceptAllRevisions(doc)).toBe(1);
    expect(doc.body).toHaveLength(3);
    expect((doc.body[0] as Paragraph).properties).toEqual({});
  });

  it("are settled by their id alone", () => {
    const doc: DocxDocument = {
      body: [
        para("a", { paragraphDeletion: rev(1) }),
        para("b", { paragraphDeletion: rev(2) }),
        para("c")
      ]
    };
    expect(Document.acceptRevision(doc, 2)).toBe(true);
    expect(doc.body.map(textOf)).toEqual(["a", "bc"]);
    expect(Query.listRevisions(doc).map(entry => entry.id)).toEqual([1]);

    const inserted: DocxDocument = {
      body: [
        para("a", { paragraphInsertion: rev(1) }),
        para("b", { paragraphInsertion: rev(2) }),
        para("c")
      ]
    };
    expect(Document.rejectRevision(inserted, 1)).toBe(true);
    expect(inserted.body.map(textOf)).toEqual(["ab", "c"]);
    // The joined paragraph ends with the second's mark, and so carries its insertion.
    expect((inserted.body[0] as Paragraph).properties?.paragraphInsertion).toEqual(rev(2));
  });

  it("join an empty paragraph into the next one", () => {
    const doc: DocxDocument = {
      body: [
        Build.paragraph([], { alignment: "center", paragraphDeletion: rev(1) }),
        para("b", { alignment: "right" })
      ]
    };
    Document.acceptAllRevisions(doc);
    expect(doc.body.map(textOf)).toEqual(["b"]);
    expect((doc.body[0] as Paragraph).properties).toEqual({ alignment: "right" });
  });

  it("keep the range markers around the joint, and the first paragraph's identity", () => {
    const start = Build.bookmarkStart(0, "joint") as RangeMarker;
    const end = Build.bookmarkEnd(0) as RangeMarker;
    const first: Paragraph = {
      ...para("a", { paragraphDeletion: rev(1) }),
      paraId: "11111111",
      rangeMarkersAfter: [start]
    };
    const second: Paragraph = { ...para("b"), paraId: "22222222", rangeMarkersBefore: [end] };
    const doc: DocxDocument = { body: [first, second] };
    Document.acceptAllRevisions(doc);
    const joined = doc.body[0] as Paragraph;
    expect(joined.paraId).toBe("11111111");
    const kinds = joined.children.map(child => ("type" in child ? child.type : "run"));
    expect(kinds).toEqual(["run", "bookmarkStart", "bookmarkEnd", "run"]);
  });

  it("join across a section break into the next section, as Word does", () => {
    const doc: DocxDocument = {
      body: [
        para("a", {
          paragraphDeletion: rev(1),
          sectionProperties: { pageSize: { width: 12240, height: 15840 } }
        }),
        para("b")
      ]
    };
    Document.acceptAllRevisions(doc);
    expect(doc.body.map(textOf)).toEqual(["ab"]);
    expect((doc.body[0] as Paragraph).properties?.sectionProperties).toBeUndefined();
  });

  it("do not join across a content control's edge", () => {
    const control: StructuredDocumentTag = {
      type: "sdt",
      content: [para("a", { paragraphDeletion: rev(1) })]
    };
    const doc: DocxDocument = { body: [control, para("b")] };
    expect(Document.acceptAllRevisions(doc)).toBe(1);
    expect(doc.body).toHaveLength(2);
  });

  it("are settled in a table's cells and in a header", () => {
    const cell = Build.cell([para("a", { paragraphDeletion: rev(1) }), para("b")]);
    const doc: DocxDocument = {
      body: [Build.table([Build.row([cell])], { columnWidths: [100] })],
      headers: new Map([
        [
          "rId10",
          {
            rId: "rId10",
            content: { children: [para("h", { paragraphDeletion: rev(2) }), para("i")] }
          }
        ]
      ]),
      sectionProperties: { headers: [{ type: "default", rId: "rId10" }] }
    };
    expect(Document.acceptAllRevisions(doc)).toBe(2);
    expect((doc.body[0] as Table).rows[0]!.cells[0]!.content.map(textOf)).toEqual(["ab"]);
    expect(doc.headers!.get("rId10")!.content.children.map(textOf)).toEqual(["hi"]);
    expect(Query.listRevisions(doc)).toEqual([]);
  });

  it("are settled in a footnote", () => {
    const doc: DocxDocument = {
      body: [para("x")],
      footnotes: [{ id: 1, content: [para("f", { paragraphDeletion: rev(1) }), para("g")] }]
    };
    expect(Document.acceptAllRevisions(doc)).toBe(1);
    expect(doc.footnotes![0]!.content.map(textOf)).toEqual(["fg"]);
  });

  it("in the mark's previous properties are read, and written back with their ids claimed as any other's", async () => {
    const doc = await read(
      `<w:p><w:pPr><w:rPr><w:ins w:id="6" w:author="X"/><w:b/><w:rPrChange w:id="5" w:author="A"><w:rPr><w:ins w:id="6" w:author="B"/><w:i/></w:rPr></w:rPrChange></w:rPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p>`
    );
    const mark = (doc.body[0] as Paragraph).properties!.markRunProperties!;
    expect(mark.propertyChange?.previousMarkRevisions).toEqual({
      insertion: { id: 6, author: "B" }
    });
    expect(mark.propertyChange?.previousProperties).toEqual({ italic: true });
    expect(Document.listUnrecognizedValues(doc)).toEqual([]);

    const xml = (await documentXml(doc)).get("word/document.xml")!;
    // Both carried id 6 in the file; written, every revision has an id of its own.
    const ids = [...xml.matchAll(/<w:(?:ins|rPrChange) w:id="(\d+)"/g)].map(m => m[1]);
    expect(ids).toHaveLength(3);
    expect(new Set(ids).size).toBe(3);
    expect(xml).toMatch(
      /<w:rPrChange [^>]*><w:rPr><w:ins [^>]*w:author="B"\/><w:i\/><\/w:rPr><\/w:rPrChange>/
    );
  });

  it("in the mark's previous properties go with the change record when it is rejected", async () => {
    const doc = await read(
      `<w:p><w:pPr><w:rPr><w:b/><w:rPrChange w:id="5" w:author="A"><w:rPr><w:del w:id="6" w:author="B"/><w:i/></w:rPr></w:rPrChange></w:rPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p>`
    );
    expect(Document.rejectAllRevisions(doc)).toBe(1);
    const properties = (doc.body[0] as Paragraph).properties!;
    expect(properties.markRunProperties).toEqual({ italic: true });
    // The mark's current revisions are its own: the previous deletion is not made current.
    expect(properties.paragraphDeletion).toBeUndefined();
    expect(Document.listUnrecognizedValues(doc)).toEqual([]);
  });

  it("are refused on an ordinary run's change record", () => {
    expect(() =>
      Build.text("x", {
        propertyChange: { revision: rev(1), previousMarkRevisions: { insertion: rev(2) } }
      })
    ).toThrow(DocxArgumentError);
  });

  it("have no place in a style's paragraph properties", async () => {
    const style: StyleDef = {
      type: "paragraph",
      styleId: "S",
      name: "S",
      paragraphProperties: { paragraphMoveTo: rev(1) }
    };
    await expect(Document.toBuffer({ body: [], styles: [style] })).rejects.toThrow(
      DocxArgumentError
    );
  });
});

describe("joining by id, with the next paragraph's own formatting change pending", () => {
  // Settling everything settles that change first (each paragraph before its mark), so this only arises by id.
  it("keeps it on the joined paragraph, whose properties it describes, as Word does", () => {
    const doc: DocxDocument = {
      body: [
        para("a", { alignment: "center", paragraphDeletion: rev(1) }),
        para("b", {
          alignment: "right",
          propertyChange: { revision: rev(2), previousProperties: { alignment: "left" } }
        })
      ]
    };
    expect(Document.acceptRevision(doc, 1)).toBe(true);
    expect(doc.body.map(textOf)).toEqual(["ab"]);
    expect((doc.body[0] as Paragraph).properties).toEqual({
      alignment: "right",
      propertyChange: { revision: rev(2), previousProperties: { alignment: "left" } }
    });
    expect(Query.listRevisions(doc).map(entry => entry.id)).toEqual([2]);
    // Rejecting it then restores the second paragraph's previous formatting, as it does in Word.
    Document.rejectRevision(doc, 2);
    expect((doc.body[0] as Paragraph).properties).toEqual({ alignment: "left" });
  });

  it("drops the first paragraph's own pending change, with the properties the join does not keep", () => {
    const doc: DocxDocument = {
      body: [
        para("a", {
          alignment: "center",
          paragraphDeletion: rev(1),
          propertyChange: { revision: rev(3), previousProperties: { alignment: "both" } }
        }),
        para("b", {
          alignment: "right",
          propertyChange: { revision: rev(2), previousProperties: { alignment: "left" } }
        })
      ]
    };
    // Word: three revisions before, one after accepting the deletion — the second paragraph's.
    expect(Query.listRevisions(doc)).toHaveLength(3);
    Document.acceptRevision(doc, 1);
    expect(Query.listRevisions(doc).map(entry => entry.id)).toEqual([2]);
  });
});

describe("settling everything, with the next paragraph's own formatting change pending", () => {
  it("settles that change before the join, and counts every listed revision", () => {
    const doc = (): DocxDocument => ({
      body: [
        para("a", { alignment: "center", paragraphDeletion: rev(1) }),
        para("b", {
          alignment: "right",
          propertyChange: { revision: rev(2), previousProperties: { alignment: "left" } }
        })
      ]
    });
    const accepted = doc();
    expect(Document.acceptAllRevisions(accepted)).toBe(Query.listRevisions(doc()).length);
    expect(accepted.body.map(textOf)).toEqual(["ab"]);
    expect(Query.listRevisions(accepted)).toEqual([]);
  });
});
