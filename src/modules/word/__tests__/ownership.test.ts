/**
 * The Word API's ownership rule: `Document.*` edits the document it is given
 * in place; every other namespace function returning a `DocxDocument` leaves
 * its input unchanged (the result may share unchanged nodes with it), and
 * `Document.clone` gives a copy that shares nothing.
 */
import { describe, it, expect } from "vitest";

import {
  Build,
  Document,
  Font,
  Io,
  Ole,
  Query,
  Security,
  Template,
  Transform,
  Vba
} from "../index";
import type { DocxDocument, Paragraph, StructuredDocumentTag } from "../index";

const OLE2 = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);

/** A document touching every kind of value the model holds: Maps, bytes, Dates. */
function rich(): DocxDocument {
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  const at = new Date("2024-01-01T00:00:00Z");
  Document.setCoreProperties(doc, { title: "T", created: at, modified: at });
  Document.addCustomProperty(doc, "Due", { type: "date", value: at });
  Document.addHeading(doc, "Hello {{name}}");
  Document.addParagraph(doc, "Dear {{name}}, the brown fox.");
  Document.addTable(doc, [["{{name}}", "b"]]);
  Document.addImage(doc, new Uint8Array([0x89, 0x50, 0x4e, 0x47]), "png", 100, 100);
  Document.addFootnote(doc, "Note for {{name}}");
  Document.setHeader(doc, "default", { children: [Build.textParagraph("Header {{name}}")] });
  Document.setFooter(doc, "default", { children: [Build.textParagraph("Footer")] });
  Document.addParagraph(
    doc,
    Build.paragraph([
      Build.insertedRun(Build.text("red "), { id: 1, author: "A" }),
      Build.deletedRun(Build.text("brown "), { id: 2, author: "A" }),
      Build.formTextField({ name: "FullName", default: "x" }),
      Build.field(" MERGEFIELD Name ", "«Name»")
    ])
  );
  Document.addSectionBreak(doc, { pageSize: { width: 12240, height: 15840 } });
  Document.addParagraph(doc, "Second section {{name}}");
  return doc;
}

function bodyText(doc: DocxDocument): string {
  return Query.getText(doc);
}

describe("Document.* edits the document it is given", () => {
  it("replaceText", () => {
    const doc = rich();
    expect(Document.replaceText(doc, "brown fox", "grey cat")).toBe(1);
    expect(bodyText(doc)).toContain("grey cat");
  });

  it("mailMerge", () => {
    const doc = rich();
    expect(Document.mailMerge(doc, { Name: "Alice" })).toBe(1);
    expect(bodyText(doc)).toContain("Alice");
  });

  it("acceptAllRevisions / rejectAllRevisions / acceptRevision / rejectRevision", () => {
    const accepted = rich();
    expect(Document.acceptAllRevisions(accepted)).toBe(2);
    expect(Query.listRevisions(accepted)).toEqual([]);

    const rejected = rich();
    expect(Document.rejectAllRevisions(rejected)).toBe(2);
    expect(Query.listRevisions(rejected)).toEqual([]);

    const single = rich();
    expect(Document.acceptRevision(single, 1)).toBe(true);
    expect(Document.rejectRevision(single, 2)).toBe(true);
    expect(Query.listRevisions(single)).toEqual([]);
  });

  it("setCompatibilityMode", () => {
    const doc = rich();
    Document.setCompatibilityMode(doc, 14);
    expect(Query.getCompatibilityMode(doc)).toBe(14);
  });
});

describe("every other function returning a document leaves its input unchanged", () => {
  const cases: [string, (doc: DocxDocument) => unknown][] = [
    ["Template.fill", doc => Template.fill(doc, { name: "Ann" })],
    [
      "Template.fill (rich placeholders)",
      doc => Template.fill(doc, { name: "Ann" }, { richPlaceholders: true })
    ],
    [
      "Template.fillFromSource",
      doc => Template.fillFromSource(doc, Template.createJsonDataSource({ name: "Ann" }))
    ],
    ["Template.bindChartData", doc => Template.bindChartData(doc, [])],
    ["Transform.merge", doc => Transform.merge([doc, rich()])],
    ["Transform.merge (as appended)", doc => Transform.merge([rich(), doc])],
    ["Transform.split", doc => Transform.split(doc)],
    ["Transform.updateFields", doc => Transform.updateFields(doc)],
    ["Transform.updateTableOfContents", doc => Transform.updateTableOfContents(doc)],
    [
      "Transform.map",
      doc =>
        Transform.map(doc, {
          transformParagraph: p => ({ ...p, properties: { alignment: "center" } }),
          transformRun: r => ({ ...r, properties: { bold: true } })
        })
    ],
    [
      "Transform.fillFormFields",
      doc => Transform.fillFormFields(doc, new Map([["FullName", "J"]]))
    ],
    [
      "Transform.resolveDataBindings",
      doc => Transform.resolveDataBindings(doc, new Map([["{abc}", "<root><n>Bound</n></root>"]]))
    ],
    ["Security.protect", doc => Security.protect(doc, { edit: "readOnly" })],
    ["Security.unprotect", doc => Security.unprotect(doc)],
    ["Vba.add", doc => Vba.add(doc, OLE2)],
    ["Vba.remove", doc => Vba.remove(Vba.add(doc, OLE2))],
    [
      "Font.addEmbedded",
      doc =>
        Font.addEmbedded(doc, [
          {
            fontDef: { name: "Calibri", embedRegular: "rIdF1" },
            embeddedFont: { rId: "rIdF1", data: new Uint8Array([1, 2]), fileName: "font1.ttf" }
          }
        ])
    ],
    ["Ole.add", doc => Ole.add(doc, Ole.createEmbedding(OLE2, "Package"))]
  ];

  it.each(cases)("%s", async (_name, run) => {
    const doc = rich();
    const sdt: StructuredDocumentTag = {
      type: "sdt",
      properties: { dataBinding: { xpath: "/root/n", storeItemId: "{abc}" } },
      content: [Build.textParagraph("unbound")]
    };
    Document.addContent(doc, sdt);
    const before = structuredClone(doc);
    await run(doc);
    expect(doc).toEqual(before);
  });

  it("Transform.fillFormFields and Transform.resolveDataBindings return the change", () => {
    const doc = rich();
    Document.addContent(doc, {
      type: "sdt",
      properties: { dataBinding: { xpath: "/root/n", storeItemId: "{abc}" } },
      content: [Build.textParagraph("unbound")]
    });
    const filled = Transform.fillFormFields(doc, new Map([["FullName", "Jane"]]));
    expect(Query.listFormFields(filled)[0]?.value).toBe("Jane");
    const bound = Transform.resolveDataBindings(
      doc,
      new Map([["{abc}", "<root><n>Bound</n></root>"]])
    );
    expect(bodyText(bound)).toContain("Bound");
    expect(bodyText(doc)).toContain("unbound");
  });
});

describe("Document.clone", () => {
  it("copies every kind of value the model holds", async () => {
    const doc = await Io.read(await Io.toBuffer(rich()));
    const copy = Document.clone(doc);
    expect(copy).toEqual(doc);
    expect(copy).not.toBe(doc);
    expect(copy.headers).toBeInstanceOf(Map);
    expect(copy.headers).not.toBe(doc.headers);
    expect(copy.images?.[0]?.data).toBeInstanceOf(Uint8Array);
    expect(copy.images?.[0]?.data).not.toBe(doc.images?.[0]?.data);
    expect(copy.coreProperties?.created).toBeInstanceOf(Date);
    expect(copy.coreProperties?.created).not.toBe(doc.coreProperties?.created);

    // Nothing reachable from the copy is reachable from the original.
    const seen = new Set<unknown>();
    const collect = (v: unknown): void => {
      if (v === null || typeof v !== "object" || seen.has(v)) {
        return;
      }
      seen.add(v);
      const values =
        v instanceof Map ? [...v.values()] : ArrayBuffer.isView(v) ? [] : Object.values(v);
      values.forEach(collect);
    };
    collect(doc);
    const shared: unknown[] = [];
    const check = (v: unknown, visited = new Set<unknown>()): void => {
      if (v === null || typeof v !== "object" || visited.has(v)) {
        return;
      }
      visited.add(v);
      if (seen.has(v)) {
        shared.push(v);
      }
      const values =
        v instanceof Map ? [...v.values()] : ArrayBuffer.isView(v) ? [] : Object.values(v);
      values.forEach(item => check(item, visited));
    };
    check(copy);
    expect(shared).toEqual([]);
  });

  it("holds only plain data in a document read from a file, so structuredClone is exact", async () => {
    const doc = await Io.read(await Io.toBuffer(rich()));
    const odd: string[] = [];
    const walk = (v: unknown, path: string): void => {
      if (typeof v === "function" || typeof v === "symbol") {
        odd.push(path);
        return;
      }
      if (v === null || typeof v !== "object") {
        return;
      }
      const proto = Object.getPrototypeOf(v);
      if (v instanceof Map) {
        for (const [k, item] of v) {
          walk(item, `${path}.get(${String(k)})`);
        }
        return;
      }
      if (v instanceof Uint8Array || v instanceof Date) {
        return;
      }
      if (proto !== Object.prototype && proto !== Array.prototype && proto !== null) {
        odd.push(path);
        return;
      }
      for (const [k, item] of Object.entries(v)) {
        walk(item, `${path}.${k}`);
      }
    };
    walk(doc, "doc");
    expect(odd).toEqual([]);
  });

  it("isolates in-place edits of a Template.fill result from the template", () => {
    const template = rich();
    const before = structuredClone(template);
    for (const name of ["Ann", "Bob"]) {
      const letter = Document.clone(Template.fill(template, { name }));
      Document.replaceText(letter, name, "X");
      Document.mailMerge(letter, { Name: "Z" });
      Document.acceptAllRevisions(letter);
      Document.addParagraph(letter, "added");
      Document.addStyle(letter, { type: "paragraph", styleId: "S", name: "S", isDefault: true });
      Document.setCompatibilityMode(letter, 12);
      const para = letter.body[1] as Paragraph;
      (para.children as unknown[]).length = 0;
    }
    expect(template).toEqual(before);
  });

  it("isolates in-place edits of a Transform.split result from the source", () => {
    const source = rich();
    const before = structuredClone(source);
    const parts = Transform.split(source).map(part => Document.clone(part));
    expect(parts.length).toBe(2);
    for (const part of parts) {
      Document.replaceText(part, "{{name}}", "X");
      Document.addParagraph(part, "added");
      Document.addStyle(part, { type: "paragraph", styleId: "S", name: "S" });
      Document.setHeader(part, "default", { children: [Build.textParagraph("changed")] });
    }
    expect(source).toEqual(before);
  });

  it("leaves the original unchanged when the clone is edited, and vice versa", () => {
    const doc = rich();
    const copy = Document.clone(doc);
    const before = structuredClone(doc);
    Document.replaceText(copy, "brown fox", "cat");
    Document.addFootnote(copy, "more");
    expect(doc).toEqual(before);
    Document.replaceText(doc, "Dear", "Hi");
    expect(bodyText(copy)).toContain("Dear");
  });

  it("does not hand out again an id already handed out for the original", () => {
    const doc = rich();
    const id = Document.allocateBookmarkId(doc);
    const copy = Document.clone(doc);
    expect(Document.allocateBookmarkId(copy)).toBeGreaterThan(id);
  });
});
