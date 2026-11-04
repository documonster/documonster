/**
 * The functions that fill named slots with values take them the same way: a
 * plain object or a `Map`. Before, `Document.mailMerge` and `Template.fill`
 * refused a `Map` ("data must be an object, got an object") while
 * `Transform.fillFormFields` and `Transform.resolveDataBindings` refused a
 * plain object — four siblings, two shapes, each refusing the other.
 */

import { describe, it, expect } from "vitest";

import { Build, Document, DocxArgumentError, Query, Template, Transform } from "../index";
import type { DocxDocument, Paragraph, Run, StructuredDocumentTag } from "../index";

function mergeDoc(): DocxDocument {
  const d = Document.create();
  Document.addParagraph(d, Build.paragraph([Build.field(" MERGEFIELD First ")]));
  return d;
}

function formDoc(): DocxDocument {
  const d = Document.create();
  Document.addParagraph(d, Build.paragraph([Build.formTextField({ name: "Name" })]));
  return d;
}

function bindingDoc(): DocxDocument {
  const sdt: StructuredDocumentTag = {
    type: "sdt",
    properties: { dataBinding: { xpath: "/root/name", storeItemId: "{abc}" } },
    content: [
      {
        type: "paragraph",
        children: [{ content: [{ type: "text", text: "PLACEHOLDER" }] } as Run]
      } as Paragraph
    ]
  };
  const d = Document.create();
  Document.addContent(d, sdt);
  return d;
}

function templateDoc(): DocxDocument {
  const d = Document.create();
  Document.addParagraph(d, "Hello {{name}}");
  return d;
}

describe("name → value data", () => {
  it("Document.mailMerge takes a Map as it takes an object", () => {
    const viaObject = mergeDoc();
    const viaMap = mergeDoc();
    expect(Document.mailMerge(viaObject, { First: "Ann" })).toBe(1);
    expect(Document.mailMerge(viaMap, new Map([["First", "Ann"]]))).toBe(1);
    expect(viaMap.body).toEqual(viaObject.body);
  });

  it("Template.fill takes a Map as it takes an object", () => {
    const viaObject = Template.fill(templateDoc(), { name: "Ann" });
    const viaMap = Template.fill(templateDoc(), new Map([["name", "Ann"]]));
    expect(Query.getText(viaMap)).toBe(Query.getText(viaObject));
    expect(Query.getText(viaMap)).toContain("Hello Ann");
  });

  it("Transform.fillFormFields takes an object as it takes a Map", () => {
    const viaMap = Transform.fillFormFields(formDoc(), new Map([["Name", "Ann"]]));
    const viaObject = Transform.fillFormFields(formDoc(), { Name: "Ann" });
    expect(viaObject.body).toEqual(viaMap.body);
    expect(Query.listFormFields(viaObject)[0]?.value).toBe("Ann");
  });

  it("Transform.resolveDataBindings takes an object as it takes a Map", () => {
    const xml = "<root><name>Ann</name></root>";
    const viaMap = Transform.resolveDataBindings(bindingDoc(), new Map([["{abc}", xml]]));
    const viaObject = Transform.resolveDataBindings(bindingDoc(), { "{abc}": xml });
    expect(viaObject.body).toEqual(viaMap.body);
    expect(Query.getText(viaObject)).toContain("Ann");
  });

  it("each names a bad value the same way, from the parameter", () => {
    const cases: [string, () => unknown][] = [
      ['data["First"]', () => Document.mailMerge(mergeDoc(), new Map([["First", 42 as never]]))],
      ['data["First"]', () => Document.mailMerge(mergeDoc(), { First: 42 as never })],
      ['values["Name"]', () => Transform.fillFormFields(formDoc(), { Name: Number.NaN })],
      [
        'values["Name"]',
        () => Transform.fillFormFields(formDoc(), new Map([["Name", Number.NaN]]))
      ],
      [
        'data["{abc}"]',
        () => Transform.resolveDataBindings(bindingDoc(), { "{abc}": 42 as never })
      ],
      ["data", () => Document.mailMerge(mergeDoc(), 42 as never)],
      ["values", () => Transform.fillFormFields(formDoc(), 42 as never)],
      ["data", () => Template.fill(templateDoc(), 42 as never)],
      ["data", () => Transform.resolveDataBindings(bindingDoc(), [] as never)],
      ["data", () => Document.mailMerge(mergeDoc(), new Map([[1, "x"]]) as never)]
    ];
    for (const [field, call] of cases) {
      let error: unknown;
      try {
        call();
      } catch (e) {
        error = e;
      }
      expect(error, field).toBeInstanceOf(DocxArgumentError);
      expect((error as DocxArgumentError).field, field).toBe(field);
    }
  });
});
