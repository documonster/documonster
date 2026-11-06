/**
 * `Template.fill` returns a document whose collections are its own: editing
 * the result through `Document.*` never reaches the template, however many
 * times it is filled.
 */
import { describe, it, expect } from "vitest";

import { Build, Document, Template } from "../../index";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

function template() {
  const doc = Document.create();
  Document.addStyle(doc, { type: "paragraph", styleId: "Base", name: "Base" });
  Document.addImage(doc, PNG, "png", 9525, 9525);
  Document.addComment(doc, "A", "note");
  Document.addCustomProperty(doc, "kind", { type: "string", value: "template" });
  Document.addFont(doc, { name: "Arial" });
  Document.setHeader(doc, "default", { children: [Build.textParagraph("{{title}}")] });
  Document.addParagraph(doc, "Hello {{name}}");
  return doc;
}

function sizes(doc: ReturnType<typeof Document.create>) {
  return {
    styles: doc.styles?.length,
    images: doc.images?.length,
    comments: doc.comments?.length,
    customProperties: doc.customProperties?.length,
    fonts: doc.fonts?.length,
    headers: doc.headers?.size,
    body: doc.body.length
  };
}

describe.each([false, true])("Template.fill (richPlaceholders: %s)", richPlaceholders => {
  it("leaves the template unchanged when the result is edited", () => {
    const tpl = template();
    const before = sizes(tpl);
    for (let i = 0; i < 3; i++) {
      const filled = Template.fill(tpl, { name: `n${i}`, title: "t" }, { richPlaceholders });
      Document.addStyle(filled, { type: "paragraph", styleId: `S${i}`, name: `S${i}` });
      Document.addImage(filled, PNG, "png", 9525, 9525);
      Document.addComment(filled, "B", "more");
      Document.addCustomProperty(filled, `k${i}`, { type: "string", value: "v" });
      Document.addFont(filled, { name: `F${i}` });
      Document.setHeader(filled, "first", { children: [Build.textParagraph("x")] });
      Document.addParagraph(filled, "more");
      expect(sizes(filled)).toEqual({
        styles: before.styles! + 1,
        images: before.images! + 1,
        comments: before.comments! + 1,
        customProperties: before.customProperties! + 1,
        fonts: before.fonts! + 1,
        headers: before.headers! + 1,
        // `addImage` adds a paragraph holding the picture.
        body: before.body + 2
      });
    }
    expect(sizes(tpl)).toEqual(before);
  });
});
