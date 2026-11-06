/**
 * A block-level SDT inside a container whose model holds blocks but no SDT —
 * a table cell, a header — keeps its content whatever kind of control it
 * is. A table of contents used to lose everything: its SDT was read as a TOC,
 * which a cell cannot hold, and dropped. It now keeps its entries and its
 * field, read field character by field character.
 */

import { describe, expect, it } from "vitest";

import { Io, Query } from "../index";
import type { Table } from "../types";
import { buildDocx, packageParts, partText } from "./roundtrip-integrity.helpers";

const R = (t: string): string => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const TAB = `<w:r><w:tab/></w:r>`;
const BEGIN = `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`;
const SEP = `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
const END = `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const INSTR = (t: string): string =>
  `<w:r><w:instrText xml:space="preserve">${t.replace(/"/g, "&quot;")}</w:instrText></w:r>`;
const sdt = (pr: string, inner: string): string =>
  `<w:sdt><w:sdtPr>${pr}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const GALLERY = (name: string): string =>
  `<w:docPartObj><w:docPartGallery w:val="${name}"/><w:docPartUnique/></w:docPartObj>`;
const CELL = (inner: string): string =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>`;

/** A TOC whose entries are plain runs, the field opening in the first entry. */
const PLAIN_TOC =
  `<w:p>${BEGIN}${INSTR(' TOC \\o "1-3" \\h \\z \\u ')}${SEP}${R("One")}${TAB}${R("1")}</w:p>` +
  `<w:p>${R("Two")}${TAB}${R("2")}</w:p>` +
  `<w:p>${R("Three")}${TAB}${R("3")}${END}</w:p>`;

/** A TOC the way Word writes one: hyperlinked entries with PAGEREF fields. */
const WORD_TOC =
  `<w:p>${BEGIN}${INSTR(' TOC \\o "1-3" \\h \\z \\u ')}${SEP}</w:p>` +
  `<w:p><w:hyperlink w:anchor="_Toc1">${R("Intro")}${TAB}` +
  `${BEGIN}${INSTR(" PAGEREF _Toc1 \\h ")}${SEP}${R("9")}${END}</w:hyperlink></w:p>` +
  `<w:p>${END}</w:p>`;

function paragraphTexts(xml: string): string[] {
  return [...xml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>|<w:p\b[^>]*\/>/g)]
    .map(m =>
      [...(m[1] ?? "").matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>/g)]
        .map(([, text]) => text ?? "\t")
        .join("")
    )
    .filter(t => t !== "");
}

async function cellTexts(source: Uint8Array): Promise<string[][]> {
  const out: string[][] = [];
  let buffer = source;
  for (let i = 0; i < 2; i++) {
    buffer = await Io.toBuffer(await Io.read(buffer));
    const xml = partText(await packageParts(buffer), "word/document.xml");
    out.push(paragraphTexts(/<w:tc>([\s\S]*?)<\/w:tc>/.exec(xml)?.[1] ?? ""));
  }
  return out;
}

describe("an SDT in a table cell", () => {
  it("keeps a TOC's entries, paragraph by paragraph", async () => {
    const source = await buildDocx({
      body: CELL(sdt(GALLERY("Table of Contents"), PLAIN_TOC)) + `<w:p>${R("after")}</w:p>`
    });
    const doc = await Io.read(source);
    const cell = (doc.body[0] as Table).rows[0]!.cells[0]!;
    expect(cell.content.map(b => b.type)).toEqual(["paragraph", "paragraph", "paragraph"]);
    const [once, twice] = await cellTexts(source);
    expect(once).toEqual(["One\t1", "Two\t2", "Three\t3"]);
    expect(twice).toEqual(once);
    // Its field too, its characters where they were.
    const xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    const cellXml = /<w:tc>([\s\S]*?)<\/w:tc>/.exec(xml)![1]!;
    expect(
      [...cellXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map(([p]) =>
        [...p.matchAll(/w:fldCharType="(\w+)"/g)].map(m => m[1])
      )
    ).toEqual([["begin", "separate"], [], ["end"]]);
    expect(cellXml.replace(/&quot;/g, '"')).toContain('TOC \\o "1-3" \\h \\z \\u');
  });

  it("keeps a Word TOC's hyperlinked entries and their page references", async () => {
    const source = await buildDocx({ body: CELL(sdt(GALLERY("Table of Contents"), WORD_TOC)) });
    const [once, twice] = await cellTexts(source);
    expect(once).toEqual(["Intro\t9"]);
    expect(twice).toEqual(once);
    const doc = await Io.read(source);
    const anchors: string[] = [];
    Query.each(doc, {
      enterHyperlink(h) {
        anchors.push(h.anchor ?? "");
      }
    });
    expect(anchors).toEqual(["_Toc1"]);
  });

  it.each([
    ["another gallery", GALLERY("Page Numbers (Bottom of Page)")],
    ["an alias", `<w:alias w:val="A"/>`],
    [
      "a checkbox",
      `<w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
        `<w14:checked w14:val="1"/></w14:checkbox>`
    ]
  ])("keeps the content of %s", async (_name, pr) => {
    const source = await buildDocx({ body: CELL(sdt(pr, `<w:p>${R("CELLTEXT")}</w:p>`)) });
    const [once, twice] = await cellTexts(source);
    expect(once).toEqual(["CELLTEXT"]);
    expect(twice).toEqual(once);
  });

  it("flattens nested SDTs and keeps runs standing directly in the content", async () => {
    const inner = sdt(`<w:alias w:val="inner"/>`, `<w:p>${R("NESTED")}</w:p>`);
    const source = await buildDocx({
      body: CELL(sdt(`<w:alias w:val="outer"/>`, `<w:p>${R("FIRST")}</w:p>${inner}${R("LOOSE")}`))
    });
    const [once, twice] = await cellTexts(source);
    expect(once).toEqual(["FIRST", "NESTED", "LOOSE"]);
    expect(twice).toEqual(once);
  });

  it.each([
    ["a tag and an alias", `<w:alias w:val="A"/><w:tag w:val="T"/>`],
    ["a gallery entry", GALLERY("Cover Pages")],
    ["a TOC gallery", GALLERY("Table of Contents")]
  ])("keeps the control around its blocks: %s", async (_name, pr) => {
    const source = await buildDocx({
      body: CELL(
        `<w:p>${R("before")}</w:p>${sdt(pr, `<w:p>${R("IN")}</w:p>`)}<w:p>${R("after")}</w:p>`
      )
    });
    const doc = await Io.read(source);
    const cell = (doc.body[0] as Table).rows[0]!.cells[0]!;
    expect(cell.contentControls?.map(c => [c.start, c.end])).toEqual([[1, 2]]);
    let buffer = source;
    for (let i = 0; i < 2; i++) {
      buffer = await Io.toBuffer(await Io.read(buffer));
      const cellXml = /<w:tc>([\s\S]*?)<\/w:tc>/.exec(
        partText(await packageParts(buffer), "word/document.xml")
      )![1]!;
      const shell = /<w:sdt>([\s\S]*?)<\/w:sdt>/.exec(cellXml)?.[1] ?? "";
      expect(paragraphTexts(cellXml)).toEqual(["before", "IN", "after"]);
      expect(paragraphTexts(shell)).toEqual(["IN"]);
      for (const part of pr.match(/<w:\w+ w:val="[^"]*"\/>|<w:docPartUnique\/>/g)!) {
        expect(shell).toContain(part);
      }
    }
  });

  it("keeps nested controls nested", async () => {
    const inner = sdt(`<w:alias w:val="inner"/>`, `<w:p>${R("NESTED")}</w:p>`);
    const source = await buildDocx({
      body: CELL(sdt(`<w:alias w:val="outer"/>`, `<w:p>${R("FIRST")}</w:p>${inner}`))
    });
    const out = partText(
      await packageParts(await Io.toBuffer(await Io.read(source))),
      "word/document.xml"
    );
    expect(
      /<w:tc>([\s\S]*?)<\/w:tc>/.exec(out)![1]!.replace(/<w:p\b[\s\S]*?<\/w:p>/g, "P")
    ).toMatch(
      /<w:sdt><w:sdtPr>(?:(?!<\/w:sdtPr>)[\s\S])*"outer"[\s\S]*?<w:sdtContent>P<w:sdt><w:sdtPr>(?:(?!<\/w:sdtPr>)[\s\S])*"inner"[\s\S]*?<w:sdtContent>P<\/w:sdtContent><\/w:sdt><\/w:sdtContent><\/w:sdt>/
    );
  });
});

describe("a body SDT holding a gallery entry", () => {
  it("keeps its gallery", async () => {
    const source = await buildDocx({
      body: sdt(GALLERY("Cover Pages"), `<w:p>${R("cover")}</w:p>`) + `<w:p>${R("after")}</w:p>`
    });
    const doc = await Io.read(source);
    expect(doc.body[0]).toMatchObject({
      type: "sdt",
      properties: { docPart: { gallery: "Cover Pages", unique: true } }
    });
    const out = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    expect(out).toContain(GALLERY("Cover Pages"));
  });
});

describe("an SDT in a header", () => {
  it("keeps a TOC's entries", async () => {
    const source = await buildDocx({
      body: `<w:p>${R("body")}</w:p>`,
      header: sdt(GALLERY("Table of Contents"), PLAIN_TOC)
    });
    const doc = await Io.read(source);
    const header = [...(doc.headers?.values() ?? [])][0];
    const texts = (header?.content.children ?? []).map(b =>
      b.type === "paragraph"
        ? b.children
            .flatMap(c => ("content" in c && Array.isArray(c.content) ? c.content : []))
            .map(c => (c.type === "text" ? c.text : c.type === "tab" ? "\t" : ""))
            .join("")
        : ""
    );
    expect(texts).toEqual(["One\t1", "Two\t2", "Three\t3"]);
  });
});

describe("a TOC nested in another SDT", () => {
  it("keeps its entries", async () => {
    const source = await buildDocx({
      body: sdt(`<w:alias w:val="outer"/>`, sdt(GALLERY("Table of Contents"), PLAIN_TOC))
    });
    const out = partText(
      await packageParts(await Io.toBuffer(await Io.read(source))),
      "word/document.xml"
    );
    expect(paragraphTexts(/<w:body>([\s\S]*)<\/w:body>/.exec(out)![1]!)).toEqual([
      "One\t1",
      "Two\t2",
      "Three\t3"
    ]);
  });
});
