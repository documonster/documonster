/**
 * `Document.listUnrecognizedValues(doc)` lists the values `Io.read` kept
 * because they are outside their field's vocabulary, as they are kept *now*:
 * an edit that takes over a field, or a discard, takes the value off the list.
 * Each is a `DocxUnrecognizedValue` — `{ path, value }` with the path
 * `Validation.validate` and `Document.discardUnrecognizedValues` use — and the
 * discard filter is given the same objects.
 */

import { describe, it, expect } from "vitest";

import { Document, Io, Package, Validation } from "../index";
import type { DocxDocument, Paragraph, DocxUnrecognizedValue } from "../index";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

/** A document whose paragraph reads `jc="bogus"` and whose run reads `u="squiggle"`. */
async function read(): Promise<DocxDocument> {
  const d = Document.create();
  Document.addParagraph(d, "x", { alignment: "center", run: { underline: "single" } });
  const bytes = await Io.toBuffer(d);
  const xml = decoder
    .decode((await Package.readPart(bytes, "word/document.xml"))!)
    .replace('w:val="center"', 'w:val="bogus"')
    .replace('<w:u w:val="single"/>', '<w:u w:val="squiggle"/>');
  const edited = await Package.edit(bytes, [
    { type: "replacePart", path: "word/document.xml", data: encoder.encode(xml) }
  ]);
  return Io.read(edited);
}

describe("Document.listUnrecognizedValues", () => {
  it("lists every kept value with the path Validation.validate reports", async () => {
    const doc = await read();
    const listed: DocxUnrecognizedValue[] = Document.listUnrecognizedValues(doc);
    expect(listed).toEqual([
      { path: "body[0].properties.alignment", value: "bogus" },
      { path: "body[0].children[0].properties.underline", value: "squiggle" }
    ]);
    const reported = Validation.validate(doc)
      .issues.filter(i => i.rule === "unrecognized-value")
      .map(i => i.path);
    expect(listed.map(v => v.path)).toEqual(reported);
  });

  it("lists what is kept now: an edit that takes over a field, or a discard, removes it", async () => {
    const doc = await read();
    const p = doc.body[0] as Paragraph;
    Object.assign(p.properties!, { alignment: "end" });
    expect(Document.listUnrecognizedValues(doc).map(v => v.path)).toEqual([
      "body[0].children[0].properties.underline"
    ]);
    Object.assign(p.properties!, { alignment: undefined });
    expect(Document.listUnrecognizedValues(doc)).toHaveLength(2);
    Document.discardUnrecognizedValues(doc, { filter: v => v.value === "squiggle" });
    expect(Document.listUnrecognizedValues(doc)).toEqual([
      { path: "body[0].properties.alignment", value: "bogus" }
    ]);
    Document.discardUnrecognizedValues(doc);
    expect(Document.listUnrecognizedValues(doc)).toEqual([]);
  });

  it("gives the discard filter the same DocxUnrecognizedValue it lists", async () => {
    const doc = await read();
    const seen: DocxUnrecognizedValue[] = [];
    Document.discardUnrecognizedValues(doc, {
      filter: v => {
        seen.push(v);
        return false;
      }
    });
    expect(seen).toEqual(Document.listUnrecognizedValues(doc));
  });

  it("is empty for a built document, and checks its handle", () => {
    expect(Document.listUnrecognizedValues(Document.create())).toEqual([]);
    expect(() => Document.listUnrecognizedValues(42 as never)).toThrow(/doc/);
  });
});
