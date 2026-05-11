/**
 * DOCX Module - OpenDoPE Data Binding Tests
 */

import { describe, it, expect } from "vitest";

import { resolveDataBindings } from "../query/data-binding";
import type { DocxDocument, StructuredDocumentTag, Paragraph, Run } from "../types";

function createDocWithSdt(
  sdt: StructuredDocumentTag,
  customXmlParts?: ReadonlyArray<{ itemId: string; xmlContent: string }>
): DocxDocument {
  return {
    body: [sdt],
    contentTypes: [],
    customXmlParts
  } as unknown as DocxDocument;
}

function makeBindingSdt(
  xpath: string,
  storeItemId: string,
  content?: any[]
): StructuredDocumentTag {
  return {
    type: "sdt",
    properties: {
      dataBinding: { xpath, storeItemId }
    } as any,
    content: content ?? [
      {
        type: "paragraph",
        children: [{ content: [{ type: "text", text: "PLACEHOLDER" }] } as Run]
      } as Paragraph
    ]
  } as unknown as StructuredDocumentTag;
}

describe("resolveDataBindings", () => {
  it("returns the document unchanged if no data is provided and no customXmlParts", () => {
    const sdt = makeBindingSdt("/root/name", "{abc}");
    const doc = createDocWithSdt(sdt);
    const result = resolveDataBindings(doc);
    expect(result).toBe(doc);
  });

  it("resolves simple XPath binding from data map", () => {
    const sdt = makeBindingSdt("/root/name", "{abc-123}");
    const doc = createDocWithSdt(sdt);
    const data = new Map<string, string>([["abc-123", "<root><name>John Doe</name></root>"]]);

    const result = resolveDataBindings(doc, data);
    const newSdt = result.body[0] as StructuredDocumentTag;
    const para = newSdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content[0]).toEqual({ type: "text", text: "John Doe" });
  });

  it("resolves XPath from embedded customXmlParts", () => {
    const sdt = makeBindingSdt("/root/name", "{abc}");
    const doc = createDocWithSdt(sdt, [
      { itemId: "{abc}", xmlContent: "<root><name>Jane</name></root>" }
    ]);

    const result = resolveDataBindings(doc);
    const newSdt = result.body[0] as StructuredDocumentTag;
    const para = newSdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content[0]).toEqual({ type: "text", text: "Jane" });
  });

  it("resolves attribute access XPath (/@attr)", () => {
    const sdt = makeBindingSdt("/root/element/@id", "{abc}");
    const doc = createDocWithSdt(sdt);
    const data = new Map<string, string>([
      ["abc", '<root><element id="42">content</element></root>']
    ]);

    const result = resolveDataBindings(doc, data);
    const newSdt = result.body[0] as StructuredDocumentTag;
    const para = newSdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content[0]).toEqual({ type: "text", text: "42" });
  });

  it("resolves position predicate XPath ([2])", () => {
    const sdt = makeBindingSdt("/root/item[2]", "{abc}");
    const doc = createDocWithSdt(sdt);
    const data = new Map<string, string>([
      ["abc", "<root><item>first</item><item>second</item><item>third</item></root>"]
    ]);

    const result = resolveDataBindings(doc, data);
    const newSdt = result.body[0] as StructuredDocumentTag;
    const para = newSdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content[0]).toEqual({ type: "text", text: "second" });
  });

  it("resolves last() position predicate", () => {
    const sdt = makeBindingSdt("/root/item[last()]", "{abc}");
    const doc = createDocWithSdt(sdt);
    const data = new Map<string, string>([
      ["abc", "<root><item>a</item><item>b</item><item>c</item></root>"]
    ]);

    const result = resolveDataBindings(doc, data);
    const newSdt = result.body[0] as StructuredDocumentTag;
    const para = newSdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content[0]).toEqual({ type: "text", text: "c" });
  });

  it("removes SDT (returns empty paragraph) when XPath returns no match", () => {
    const sdt = makeBindingSdt("/root/missing", "{abc}");
    const doc = createDocWithSdt(sdt);
    const data = new Map<string, string>([["abc", "<root><name>John</name></root>"]]);

    const result = resolveDataBindings(doc, data);
    const block = result.body[0];
    expect(block.type).toBe("paragraph");
    expect((block as Paragraph).children.length).toBe(0);
  });

  it("returns document unchanged when SDT has no dataBinding", () => {
    const sdt: StructuredDocumentTag = {
      type: "sdt",
      properties: {},
      content: [
        {
          type: "paragraph",
          children: [{ content: [{ type: "text", text: "static" }] } as Run]
        } as Paragraph
      ]
    } as unknown as StructuredDocumentTag;
    const doc = createDocWithSdt(sdt);
    const data = new Map<string, string>([["abc", "<root></root>"]]);

    const result = resolveDataBindings(doc, data);
    const newSdt = result.body[0] as StructuredDocumentTag;
    expect(newSdt.content).toEqual(sdt.content);
  });

  it("strips namespace prefixes in XPath", () => {
    const sdt = makeBindingSdt("/ns0:root/ns0:name", "{abc}");
    const doc = createDocWithSdt(sdt);
    const data = new Map<string, string>([
      ["abc", '<ns0:root xmlns:ns0="http://example.com"><ns0:name>NSValue</ns0:name></ns0:root>']
    ]);

    const result = resolveDataBindings(doc, data);
    const newSdt = result.body[0] as StructuredDocumentTag;
    const para = newSdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content[0]).toEqual({ type: "text", text: "NSValue" });
  });

  it("handles storeItemId in braces matching unbraced key", () => {
    const sdt = makeBindingSdt("/root/v", "{ABC}");
    const doc = createDocWithSdt(sdt);
    // Data uses lowercase no-braces key; resolveDataBindings normalizes
    const data = new Map<string, string>([["abc", "<root><v>XYZ</v></root>"]]);

    const result = resolveDataBindings(doc, data);
    const newSdt = result.body[0] as StructuredDocumentTag;
    const para = newSdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content[0]).toEqual({ type: "text", text: "XYZ" });
  });
});
