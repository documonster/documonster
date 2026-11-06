/**
 * Every consumer of fields reads a field held in pieces (see
 * `FieldContent.fieldChar`) as the one field it is: mail merge and field
 * update replace its result across the runs, wrappers and paragraphs it
 * spans, text extraction and the converters show what Word shows, the page
 * layout computes a page-number field, and the template engine keeps it.
 */

import { describe, expect, it } from "vitest";

import { Io, Layout, Query, Template, Transform } from "../index";
import { fromDocx as toHtml } from "../surface/html";
import { fromDocx as toMarkdown } from "../surface/markdown";
import type { DocxDocument } from "../types";
import { buildDocx, packageParts, partText } from "./roundtrip-integrity.helpers";

const R = (t: string): string => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const CH = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const BEGIN = CH("begin");
const SEP = CH("separate");
const END = CH("end");
const INSTR = (t: string): string =>
  `<w:r><w:instrText xml:space="preserve">${t.replace(/"/g, "&quot;")}</w:instrText></w:r>`;
const P = (inner: string, pPr = ""): string => `<w:p>${pPr}${inner}</w:p>`;
/** A begin character in the run of the text before it: the field is read in pieces. */
const SHARED_BEGIN = (before: string): string =>
  `<w:r><w:t xml:space="preserve">${before}</w:t><w:fldChar w:fldCharType="begin"/></w:r>`;

async function read(body: string, header?: string): Promise<DocxDocument> {
  return Io.read(await buildDocx({ body, ...(header !== undefined ? { header } : {}) }));
}

async function documentXml(doc: DocxDocument): Promise<string> {
  const xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
  return /<w:body>([\s\S]*)<\/w:body>/.exec(xml)?.[1] ?? xml;
}

/** Per paragraph: its text, field characters as `{b}` `{s}` `{e}`, instructions trimmed in `[…]`. */
function shape(bodyXml: string): string[] {
  return [...bodyXml.matchAll(/<w:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:p>)/g)].map(([, inner]) =>
    [
      ...(inner ?? "").matchAll(
        /<w:fldChar w:fldCharType="(\w)\w*"|<w:instrText[^>]*>([^<]*)<\/w:instrText>|<w:t[^>]*>([^<]*)<\/w:t>/g
      )
    ]
      .map(m => (m[1] ? `{${m[1]}}` : m[2] !== undefined ? `[${m[2].trim()}]` : m[3]))
      .join("")
      .replace(/\]\[/g, "")
      .replace(/&quot;/g, '"')
  );
}

const MERGEFIELD = (name: string): string =>
  BEGIN + INSTR(` MERGEFIELD ${name} `) + SEP + R(`«${name}»`) + END;

describe("mail merge — fields in pieces", () => {
  it("merges a field whose begin shares a run with text, keeping the text", async () => {
    const doc = await read(
      P(SHARED_BEGIN("Dear ") + INSTR(" MERGEFIELD Name ") + SEP + R("«Name»") + END + R("!"))
    );
    expect(Query.mailMerge(doc, { Name: "Alice" })).toBe(1);
    expect(shape(await documentXml(doc))).toEqual(["Dear Alice!"]);
  });

  it("merges a field whose result spans paragraphs into one paragraph", async () => {
    const doc = await read(
      P(R("A ") + BEGIN + INSTR(" MERGEFIELD Name ") + SEP + R("line1")) +
        P(R("line2") + END + R(" z"))
    );
    expect(Query.mailMerge(doc, { Name: "Alice" })).toBe(1);
    expect(shape(await documentXml(doc))).toEqual(["A Alice z"]);
  });

  it("merges a field inside a tracked insertion, inside it", async () => {
    const doc = await read(P(`<w:ins w:id="1" w:author="a">${MERGEFIELD("Name")}</w:ins>`));
    expect(Query.mailMerge(doc, { Name: "Alice" })).toBe(1);
    const xml = await documentXml(doc);
    expect(shape(xml)).toEqual(["Alice"]);
    expect(xml).toMatch(/<w:ins\b[^>]*>(?:(?!<\/w:ins>)[\s\S])*Alice/);
  });

  it("merges a field in pieces inside a hyperlink, inside it", async () => {
    const doc = await read(
      P(
        `<w:hyperlink w:anchor="x">${SHARED_BEGIN("Hi ")}${INSTR(" MERGEFIELD Name ")}${SEP}${R("«Name»")}${END}</w:hyperlink>`
      )
    );
    expect(Query.mailMerge(doc, { Name: "Alice" })).toBe(1);
    const xml = await documentXml(doc);
    expect(shape(xml)).toEqual(["Hi Alice"]);
    expect(xml).toMatch(/<w:hyperlink\b[^>]*>(?:(?!<\/w:hyperlink>)[\s\S])*Alice/);
  });

  it("removes an unmatched field in pieces with removeUnmatched", async () => {
    const doc = await read(
      P(SHARED_BEGIN("Dear ") + INSTR(" MERGEFIELD Other ") + SEP + R("«Other»") + END)
    );
    expect(Query.mailMerge(doc, {}, { removeUnmatched: true })).toBe(1);
    expect(shape(await documentXml(doc))).toEqual(["Dear "]);
  });

  it("keeps a bookmark inside the merged result paired", async () => {
    const doc = await read(
      P(
        SHARED_BEGIN("Dear ") +
          INSTR(" MERGEFIELD Name ") +
          SEP +
          `<w:bookmarkStart w:id="7" w:name="bm"/>` +
          R("«Name»") +
          END +
          `<w:bookmarkEnd w:id="7"/>`
      )
    );
    Query.mailMerge(doc, { Name: "Alice" });
    const xml = await documentXml(doc);
    expect(xml).toContain(`w:name="bm"`);
    expect(xml.match(/<w:bookmarkEnd w:id="7"\/>/g)).toHaveLength(1);
  });
});

describe("updateFields — fields in pieces", () => {
  it("updates NUMPAGES whose field characters share a run, keeping the field", async () => {
    const doc = await read(
      P(
        R("Page ") +
          BEGIN +
          INSTR(" PAGE ") +
          SEP +
          R("9") +
          END +
          R(" of ") +
          BEGIN +
          INSTR(" NUMPAGES ") +
          `<w:r><w:fldChar w:fldCharType="separate"/><w:t>9</w:t><w:fldChar w:fldCharType="end"/></w:r>`
      )
    );
    expect(shape(await documentXml(Transform.updateFields(doc)))).toEqual([
      "Page {b}[PAGE]{s}1{e} of {b}[NUMPAGES]{s}1{e}"
    ]);
  });

  it("replaces a result spanning paragraphs, joining them, the field kept", async () => {
    const doc = await read(
      P(SHARED_BEGIN("n=") + INSTR(" NUMPAGES ") + SEP + R("9")) + P(R("x") + END + R(" tail"))
    );
    expect(shape(await documentXml(Transform.updateFields(doc)))).toEqual([
      "n={b}[NUMPAGES]{s}1{e} tail"
    ]);
  });

  it("updates a field inside a tracked insertion", async () => {
    const doc = await read(
      P(`<w:ins w:id="1" w:author="a">${BEGIN}${INSTR(" NUMPAGES ")}${SEP}${R("9")}${END}</w:ins>`)
    );
    expect(shape(await documentXml(Transform.updateFields(doc)))).toEqual(["{b}[NUMPAGES]{s}1{e}"]);
  });

  it("gives a field with no result one", async () => {
    const doc = await read(P(SHARED_BEGIN("n=") + INSTR(" NUMPAGES ") + END));
    expect(shape(await documentXml(Transform.updateFields(doc)))).toEqual([
      "n={b}[NUMPAGES]{s}1{e}"
    ]);
  });

  it("numbers SEQ fields whole and in pieces in document order", async () => {
    const seq = (lead: string): string =>
      P(SHARED_BEGIN(lead) + INSTR(" SEQ Figure \\* ARABIC ") + SEP + R("9") + END);
    const doc = await read(
      seq("Figure ") +
        P(R("Figure ") + BEGIN + INSTR(" SEQ Figure \\* ARABIC ") + SEP + R("9") + END) +
        seq("Figure ")
    );
    expect(shape(await documentXml(Transform.updateFields(doc)))).toEqual([
      "Figure {b}[SEQ Figure \\* ARABIC]{s}1{e}",
      "Figure {b}[SEQ Figure \\* ARABIC]{s}2{e}",
      "Figure {b}[SEQ Figure \\* ARABIC]{s}3{e}"
    ]);
  });

  it("resolves REF and evaluates a formula in pieces", async () => {
    const doc = await read(
      P(`<w:bookmarkStart w:id="1" w:name="bm"/>${R("Target")}<w:bookmarkEnd w:id="1"/>`) +
        P(SHARED_BEGIN("see ") + INSTR(" REF bm ") + SEP + R("old") + END) +
        P(SHARED_BEGIN("sum ") + INSTR(" = 2 + 3 ") + SEP + R("0") + END)
    );
    const out = shape(await documentXml(Transform.updateFields(doc)));
    expect(out[1]).toBe("see {b}[REF bm]{s}Target{e}");
    expect(out[2]).toBe("sum {b}[= 2 + 3]{s}5{e}");
  });

  it("evaluates IF over what the field nested in its instruction shows", async () => {
    const ifField =
      R("Dear ") +
      BEGIN +
      INSTR(" IF ") +
      MERGEFIELD("Title") +
      INSTR(' = "Dr" "Doctor" "Friend" ') +
      SEP +
      R("Friend") +
      END +
      R(",");
    const doc = await read(P(ifField));
    expect(Query.mailMerge(doc, { Title: "Dr" })).toBe(1);
    const xml = await documentXml(Transform.updateFields(doc));
    expect(shape(xml)[0]).toMatch(/\{s\}Doctor\{e\},$/);
  });

  it("is idempotent", async () => {
    const doc = await read(
      P(SHARED_BEGIN("n=") + INSTR(" NUMPAGES ") + SEP + R("9")) + P(R("x") + END + R(" tail"))
    );
    const reread = async (d: DocxDocument): Promise<DocxDocument> => Io.read(await Io.toBuffer(d));
    const once = await reread(Transform.updateFields(doc));
    const twice = Transform.updateFields(once);
    expect(twice).toBe(once);
    expect(await documentXml(Transform.updateFields(await reread(twice)))).toBe(
      await documentXml(once)
    );
  });

  it("regenerates a TOC field in pieces outside a content control", async () => {
    const doc = await read(
      P(BEGIN + INSTR(' TOC \\o "1-2" ') + SEP) +
        P(R("Stale") + TAB + R("4")) +
        P(END) +
        P(R("Intro"), `<w:pPr><w:outlineLvl w:val="0"/></w:pPr>`) +
        P(R("Second"), `<w:pPr><w:outlineLvl w:val="1"/></w:pPr>`)
    );
    const out = shape(await documentXml(Transform.updateFields(doc)));
    expect(out.join("|")).not.toContain("Stale");
    expect(out[0]).toMatch(/^\{b\}\[TOC \\o "1-2"\]\{s\}Intro/);
    expect(out[1]).toMatch(/^Second/);
    expect(out[2]).toBe("{e}");
  });
});

const TAB = `<w:r><w:tab/></w:r>`;

describe("display — fields in pieces", () => {
  const IF_FIELD =
    R("Dear ") +
    BEGIN +
    INSTR(" IF ") +
    MERGEFIELD("Title") +
    INSTR(' = "Dr" "Doctor" "Friend" ') +
    SEP +
    R("Friend") +
    END +
    R(",");

  it("text extraction and the converters do not show a field's instruction", async () => {
    const doc = await read(P(IF_FIELD));
    expect(Query.getText(doc)).toBe("Dear Friend,");
    expect(toMarkdown(doc)).toContain("Dear Friend,");
    expect(toMarkdown(doc)).not.toContain("Title");
    expect(toHtml(doc).html.replace(/\n/g, "")).toContain("Dear Friend,");
    expect(toHtml(doc).html).not.toContain("Title");
  });

  it("the layout computes a page-number field in pieces in a header", async () => {
    const doc = await read(
      P(R("Body")),
      P(SHARED_BEGIN("Page ") + INSTR(" PAGE ") + SEP + R("9") + END)
    );
    const header = JSON.stringify(Layout.compute(doc).pages[0]!.header);
    expect(header).toContain("Page ");
    expect(header).toMatch(/"text":"1"|"text":"Page 1"/);
    expect(header).not.toContain(`"9"`);
  });
});

describe("template — fields", () => {
  it("fills placeholders around a field, keeping it and the text after it", async () => {
    const doc = await read(
      P(R("Hi {{name}}, page ") + BEGIN + INSTR(" PAGE ") + SEP + R("3") + END + R(" of {{n}}"))
    );
    const out = Template.fill(doc, { name: "Ann", n: "5" });
    expect(shape(await documentXml(out))).toEqual(["Hi Ann, page {b}[PAGE]{s}3{e} of 5"]);
  });

  it("keeps a field in pieces and leaves placeholders inside it alone", async () => {
    const doc = await read(
      P(SHARED_BEGIN("{{a}} ") + INSTR(" MERGEFIELD Name ") + SEP + R("{{b}}") + END + R(" {{a}}"))
    );
    const out = Template.fill(doc, { a: "X", b: "Y" }, { strict: false });
    expect(shape(await documentXml(out))).toEqual(["X {b}[MERGEFIELD Name]{s}{{b}}{e} X"]);
  });
});

describe("Query.listFields", () => {
  it("lists a field in pieces as the one field it is", async () => {
    const doc = await read(
      P(SHARED_BEGIN("Dear ") + INSTR(" MERGEFIELD ") + INSTR("Name ") + SEP + R("«Name»")) +
        P(R("more") + END) +
        P(BEGIN + INSTR(" PAGE ") + SEP + R("3") + END)
    );
    expect(
      Query.listFields(doc).map(f => [f.type, f.instruction.trim(), f.result, f.split, f.part])
    ).toEqual([
      ["MERGEFIELD", "MERGEFIELD Name", "«Name»\nmore", true, "body"],
      ["PAGE", "PAGE", "3", false, "body"]
    ]);
  });
});

describe("fields in pieces inside wrappers", () => {
  it("updateFields updates one inside a hyperlink, inside it", async () => {
    const doc = await read(
      P(
        `<w:hyperlink w:anchor="x">${SHARED_BEGIN("n=")}${INSTR(" NUMPAGES ")}${SEP}${R("9")}${END}</w:hyperlink>`
      )
    );
    const xml = await documentXml(Transform.updateFields(doc));
    expect(shape(xml)).toEqual(["n={b}[NUMPAGES]{s}1{e}"]);
    expect(xml).toMatch(/<w:hyperlink\b[^>]*>(?:(?!<\/w:hyperlink>)[\s\S])*>1</);
  });

  it("text extraction hides the instruction of one inside a tracked insertion", async () => {
    const doc = await read(
      P(
        `<w:ins w:id="1" w:author="a">${R("Dear ")}${BEGIN}${INSTR(" IF ")}${MERGEFIELD("Title")}` +
          `${INSTR(' = "Dr" "Doctor" "Friend" ')}${SEP}${R("Friend")}${END}</w:ins>`
      )
    );
    expect(Query.getText(doc)).toBe("Dear Friend");
  });
});
