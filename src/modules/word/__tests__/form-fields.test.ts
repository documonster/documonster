/**
 * Legacy form fields: `Transform.fillFormFields` changes what Word shows, takes
 * only values the field can hold, and `Query.listFormFields` reports what Word
 * shows — for fields built here and fields Word wrote.
 */

import { describe, it, expect } from "vitest";

import { DocxArgumentError } from "../errors";
import { Build, Document, Io, Package, Query, Transform } from "../index";
import type { DocxDocument, FieldContent, Run } from "../types";

const decoder = new TextDecoder();

function formDoc(): DocxDocument {
  const doc = Document.create();
  Document.addParagraph(
    doc,
    Build.paragraph([
      Build.formTextField({ name: "Name", default: "orig", maxLength: 8 }),
      Build.formCheckboxField({ name: "Agree" }),
      Build.formDropdownField({ name: "Plan", entries: ["Basic", "Pro"] })
    ])
  );
  return doc;
}

function value(doc: DocxDocument, name: string): unknown {
  return Query.listFormFields(doc).find(f => f.name === name)?.value;
}

/** A text and a drop-down field the way Word writes them: result runs, `w:result` for the selection. */
async function wordAuthored(): Promise<DocxDocument> {
  const doc = Document.create();
  Document.addParagraph(doc, "PLACEHOLDER");
  const bytes = await Io.toBuffer(doc);
  const xml = decoder.decode((await Package.readPart(bytes, "word/document.xml"))!);
  const fields =
    '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text1"/><w:enabled/>' +
    '<w:calcOnExit w:val="0"/><w:textInput><w:default w:val="def"/></w:textInput></w:ffData></w:fldChar></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:rPr><w:noProof/></w:rPr><w:t>typed</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
    '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Drop1"/><w:enabled/><w:calcOnExit w:val="0"/>' +
    '<w:ddList><w:result w:val="1"/><w:listEntry w:val="a"/><w:listEntry w:val="b"/></w:ddList></w:ffData></w:fldChar></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
  const replaced = xml.replace(/<w:p>(?:(?!<w:p>).)*?PLACEHOLDER[\s\S]*?<\/w:p>/, fields);
  expect(replaced).not.toBe(xml);
  return Io.read(
    await Package.edit(bytes, [
      { type: "replacePartText", path: "word/document.xml", text: replaced }
    ])
  );
}

function fieldOf(doc: DocxDocument, paragraph: number): FieldContent {
  const para = doc.body[paragraph];
  if (para?.type !== "paragraph") {
    throw new Error("not a paragraph");
  }
  for (const child of para.children) {
    for (const rc of (child as Run).content ?? []) {
      if (rc.type === "field" && rc.formField) {
        return rc;
      }
    }
  }
  throw new Error("no form field");
}

describe("Transform.fillFormFields", () => {
  it("writes a text value as the field result Word displays", async () => {
    const filled = Transform.fillFormFields(formDoc(), new Map([["Name", "Jane"]]));
    const xml = decoder.decode(
      (await Package.readPart(await Io.toBuffer(filled), "word/document.xml"))!
    );
    const result = xml.match(
      /FORMTEXT[\s\S]*?w:fldCharType="separate"[\s\S]*?w:fldCharType="end"/
    )?.[0];
    expect(result).toMatch(/<w:t[^>]*>Jane<\/w:t>/);
    expect(value(filled, "Name")).toBe("Jane");
    expect(Query.getText(filled)).toContain("Jane");
  });

  it("selects a drop-down entry by index or by its text, as the current selection", async () => {
    const byIndex = Transform.fillFormFields(formDoc(), new Map([["Plan", 1]]));
    const byText = Transform.fillFormFields(formDoc(), new Map([["Plan", "Pro"]]));
    expect(value(byIndex, "Plan")).toBe(1);
    expect(value(byText, "Plan")).toBe(1);
    const xml = decoder.decode(
      (await Package.readPart(await Io.toBuffer(byText), "word/document.xml"))!
    );
    expect(xml).toMatch(/<w:ddList><w:result w:val="1"\/>/);
    expect(value(await Io.read(await Io.toBuffer(byText)), "Plan")).toBe(1);
  });

  it("refuses a value the field cannot hold, before changing anything", () => {
    const doc = formDoc();
    const cases: [string, string | number | boolean][] = [
      ["Agree", "false"],
      ["Agree", 1],
      ["Plan", "Enterprise"],
      ["Plan", 2],
      ["Plan", -1],
      ["Plan", 0.5],
      ["Plan", true],
      ["Name", true],
      ["Name", "much too long"]
    ];
    for (const [name, v] of cases) {
      expect(
        () =>
          Transform.fillFormFields(
            doc,
            new Map<string, string | number | boolean>([
              ["Name", "ok"],
              [name, v]
            ])
          ),
        `${name}=${JSON.stringify(v)}`
      ).toThrow(DocxArgumentError);
    }
  });

  it("skips a name no field has, as it always did", () => {
    // One set of values may be applied to forms that lack some of the fields.
    const filled = Transform.fillFormFields(
      formDoc(),
      new Map([
        ["Nope", "x"],
        ["Name", "Jane"]
      ])
    );
    expect(value(filled, "Name")).toBe("Jane");
  });

  it("checks a checkbox only from a boolean", () => {
    const filled = Transform.fillFormFields(formDoc(), new Map([["Agree", true]]));
    expect(value(filled, "Agree")).toBe(true);
    expect(value(Transform.fillFormFields(filled, new Map([["Agree", false]])), "Agree")).toBe(
      false
    );
  });

  it("fills a field Word wrote, replacing its result", async () => {
    const doc = await wordAuthored();
    const filled = Transform.fillFormFields(
      doc,
      new Map<string, string | number>([
        ["Text1", "new"],
        ["Drop1", "a"]
      ])
    );
    expect(fieldOf(filled, 0).cachedValue).toBe("new");
    expect(fieldOf(filled, 0).formField).toMatchObject({ default: "def" });
    expect(value(filled, "Drop1")).toBe(0);
  });
});

describe("Query.listFormFields", () => {
  it("reports what Word shows for a field Word wrote", async () => {
    const doc = await wordAuthored();
    expect(value(doc, "Text1")).toBe("typed");
    expect(value(doc, "Drop1")).toBe(1);
  });

  it("keeps a drop-down's current selection on round trip", async () => {
    const doc = await wordAuthored();
    const again = await Io.read(await Io.toBuffer(doc));
    expect(value(again, "Drop1")).toBe(1);
  });

  it("lists the fields in every part fillFormFields fills", () => {
    const doc = Document.create();
    Document.addParagraph(doc, "x");
    Document.addFootnote(doc, [
      Build.paragraph([Build.formTextField({ name: "InNote", default: "n" })])
    ]);
    Document.addComment(doc, "A", [
      Build.paragraph([Build.formTextField({ name: "InComment", default: "c" })])
    ]);
    const names = Query.listFormFields(doc).map(f => f.name);
    expect(names).toEqual(expect.arrayContaining(["InNote", "InComment"]));
    const filled = Transform.fillFormFields(doc, new Map([["InNote", "filled"]]));
    expect(value(filled, "InNote")).toBe("filled");
  });
});

describe("Build.formTextField", () => {
  it("shows its default text until the field is filled", async () => {
    const doc = Document.create();
    Document.addParagraph(
      doc,
      Build.paragraph([Build.formTextField({ name: "F", default: "shown" })])
    );
    const xml = decoder.decode(
      (await Package.readPart(await Io.toBuffer(doc), "word/document.xml"))!
    );
    expect(xml).toMatch(/<w:t[^>]*>shown<\/w:t>/);
  });
});
