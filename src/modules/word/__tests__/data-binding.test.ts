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
  content?: StructuredDocumentTag["content"]
): StructuredDocumentTag {
  return {
    type: "sdt",
    properties: {
      dataBinding: { xpath, storeItemId }
    },
    content: content ?? [
      {
        type: "paragraph",
        children: [{ content: [{ type: "text", text: "PLACEHOLDER" }] } as Run]
      } as Paragraph
    ]
  } satisfies StructuredDocumentTag;
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
    };
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

  it("decodes XML entities in resolved values", () => {
    // Previously the regex evaluator stripped tags but never decoded
    // entities, so DOCX output would carry double-encoded text.
    const sdt = makeBindingSdt("/root/v", "{abc}");
    const doc = createDocWithSdt(sdt);
    const data = new Map<string, string>([
      ["abc", "<root><v>foo &amp; bar &lt;baz&gt;</v></root>"]
    ]);
    const result = resolveDataBindings(doc, data);
    const newSdt = result.body[0] as StructuredDocumentTag;
    const para = newSdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect((run.content[0] as { text: string }).text).toBe("foo & bar <baz>");
  });

  it("handles CDATA sections in customXml", () => {
    const sdt = makeBindingSdt("/root/v", "{abc}");
    const doc = createDocWithSdt(sdt);
    const data = new Map<string, string>([
      ["abc", "<root><v><![CDATA[<not-a-tag> & not-an-entity]]></v></root>"]
    ]);
    const result = resolveDataBindings(doc, data);
    const newSdt = result.body[0] as StructuredDocumentTag;
    const para = newSdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect((run.content[0] as { text: string }).text).toBe("<not-a-tag> & not-an-entity");
  });

  it("decodes entities in resolved attribute values", () => {
    const sdt = makeBindingSdt("/root/v/@title", "{abc}");
    const doc = createDocWithSdt(sdt);
    const data = new Map<string, string>([["abc", '<root><v title="A &amp; B"/></root>']]);
    const result = resolveDataBindings(doc, data);
    const newSdt = result.body[0] as StructuredDocumentTag;
    const para = newSdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect((run.content[0] as { text: string }).text).toBe("A & B");
  });
});

describe("resolveDataBindings and the namespaces of a binding's prefixes", () => {
  const data =
    '<p:root xmlns:p="urn:right" xmlns:q="urn:wrong"><q:value a="W" q:b="W">WRONG</q:value><p:value a="R" p:b="R">RIGHT</p:value></p:root>';
  const bound = (xpath: string, prefixMappings?: string): unknown => {
    const sdt = makeBindingSdt(xpath, "{abc}");
    const withMappings = {
      ...sdt,
      properties: { dataBinding: { xpath, storeItemId: "{abc}", prefixMappings } }
    } as StructuredDocumentTag;
    const result = resolveDataBindings(createDocWithSdt(withMappings), new Map([["abc", data]]));
    const para = (result.body[0] as StructuredDocumentTag).content[0] as Paragraph;
    return (para.children[0] as Run | undefined)?.content[0];
  };

  it("matches the element in the namespace the mapping gives a prefix, whatever prefix the data uses", () => {
    expect(bound("/x:root/x:value", 'xmlns:x="urn:right"')).toEqual({
      type: "text",
      text: "RIGHT"
    });
    expect(bound("/x:root/y:value", "xmlns:x='urn:right' xmlns:y='urn:wrong'")).toEqual({
      type: "text",
      text: "WRONG"
    });
    expect(bound("/x:root/x:value/@x:b", 'xmlns:x="urn:right"')).toEqual({
      type: "text",
      text: "R"
    });
  });

  it("matches by local name a prefix the binding does not map, as it always did", () => {
    expect(bound("/p:root/p:value")).toEqual({ type: "text", text: "WRONG" });
  });
});

describe("resolveDataBindings and the XPath it evaluates", () => {
  /** The SDT that results: its text, `null` when the SDT was removed, and the SDT itself. */
  const resolve = (
    xpath: string,
    xml: string,
    prefixMappings?: string
  ): { text: string | undefined; removed: boolean; sdt: StructuredDocumentTag | undefined } => {
    const sdt: StructuredDocumentTag = {
      ...makeBindingSdt(xpath, "{abc}"),
      properties: { dataBinding: { xpath, storeItemId: "{abc}", prefixMappings } }
    };
    const block = resolveDataBindings(createDocWithSdt(sdt), new Map([["abc", xml]])).body[0];
    if (block?.type !== "sdt") {
      return { text: undefined, removed: true, sdt: undefined };
    }
    const para = block.content[0] as Paragraph;
    const text = ((para.children[0] as Run | undefined)?.content[0] as { text?: string })?.text;
    return { text, removed: false, sdt: block };
  };

  it("matches no mapped prefix in a store whose namespaces do not resolve", () => {
    const xml = '<x:root xmlns:x="urn:WRONG"><x:value>WRONG</x:value><u:bad/></x:root>';
    expect(resolve("/p:root/p:value", xml, 'xmlns:p="urn:RIGHT"').removed).toBe(true);
    // An unmapped step is still matched by local name there.
    expect(resolve("/root/value", xml).text).toBe("WRONG");
  });

  it("reads an unprefixed attribute step as the attribute in no namespace", () => {
    const xml = '<root xmlns:x="urn:x" x:id="WRONG" id="RIGHT"/>';
    expect(resolve("/root/@id", xml).text).toBe("RIGHT");
    expect(resolve("/root/@id", '<root xmlns:x="urn:x" x:id="WRONG"/>').removed).toBe(true);
  });

  it("counts positions from 1, on the root step too", () => {
    const xml = "<root><item>A</item><item>B</item></root>";
    expect(resolve("/root/item[0]", xml).removed).toBe(true);
    expect(resolve("/root/item[2]", xml).text).toBe("B");
    expect(resolve("/root[1]/item", xml).text).toBe("A");
    expect(resolve("/root[2]/item", xml).removed).toBe(true);
  });

  it("leaves the SDT as it was for an expression it cannot evaluate", () => {
    for (const xpath of [
      "/root/count(item)",
      "/root/item[@a='1']",
      "/root | /x",
      "/root/..",
      "/root//item",
      "//item"
    ]) {
      const { sdt } = resolve(xpath, "<root><item>A</item></root>");
      expect(sdt?.content).toEqual(makeBindingSdt(xpath, "{abc}").content);
    }
    const { sdt } = resolve("/p:root", "<root/>", "xmlns:p=");
    expect(sdt?.content).toEqual(makeBindingSdt("/p:root", "{abc}").content);
  });

  it("reads prefixMappings as XML, references decoded", () => {
    const xml = '<root xmlns="urn:a&amp;b">OK</root>';
    expect(resolve("/p:root", xml, 'xmlns:p="urn:a&amp;b"').text).toBe("OK");
    expect(resolve("/p:root", xml, "xmlns:p='urn:a&#38;b'").text).toBe("OK");
  });
});
