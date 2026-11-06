/**
 * A table of contents keeps its field instruction exactly as it was read:
 * Word's own default is `TOC \o "1-3" \h \z \u`, and the switches the model
 * does not name (`\z`, `\u`, `\n`, `\w`, …) are as much a part of the field
 * as the ones it does.
 */

import { describe, expect, it } from "vitest";

import { Document, Io, Transform } from "../index";
import type { BodyContent, TableOfContents } from "../types";
import {
  buildDocx,
  packageParts,
  partText,
  tocBlockInstructions
} from "./roundtrip-integrity.helpers";

const R = (t: string): string => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const BEGIN = `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`;
const SEP = `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
const END = `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const INSTR = (t: string): string =>
  `<w:r><w:instrText xml:space="preserve">${t.replace(/"/g, "&quot;")}</w:instrText></w:r>`;
const PAGEREF = (bm: string, n: number): string =>
  `${BEGIN}${INSTR(` PAGEREF ${bm} \\h `)}${SEP}${R(String(n))}${END}`;
const ENTRY = (level: number, text: string, bm: string, n: number): string =>
  `<w:p><w:pPr><w:pStyle w:val="TOC${level}"/></w:pPr><w:hyperlink w:anchor="${bm}">` +
  `${R(text)}<w:r><w:tab/></w:r>${PAGEREF(bm, n)}</w:hyperlink></w:p>`;
const SDT = (inner: string): string =>
  `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>` +
  `<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
/** A TOC block the way Word writes one: the field opens in a paragraph of its own. */
const TOC_BLOCK = (instruction: string, entries = ENTRY(1, "Intro", "_Toc1", 1)): string =>
  SDT(`<w:p>${BEGIN}${INSTR(instruction)}${SEP}</w:p>${entries}<w:p>${END}</w:p>`);
const TAIL =
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="0" w:name="_Toc1"/>` +
  `${R("Intro")}<w:bookmarkEnd w:id="0"/></w:p>`;

/** The visible text of `w:body`: its `w:t` text and tabs (not field instructions), a pilcrow per paragraph. */
function bodyText(xml: string): string {
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(xml)![1]!;
  return [...body.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<\/w:p>/g)]
    .map(([token, text]) => text ?? (token === "<w:tab/>" ? "\t" : "¶"))
    .join("");
}

async function roundTrip(source: Uint8Array): Promise<string> {
  return partText(
    await packageParts(await Io.toBuffer(await Io.read(source))),
    "word/document.xml"
  );
}

const VARIANTS: Record<string, string> = {
  "Word's default": ' TOC \\o "1-3" \\h \\z \\u ',
  "an \\o without a range": " TOC \\o ",
  "a table of figures": ' TOC \\h \\z \\c "Figure" ',
  "every switch": [
    " TOC",
    '\\a "Figure"',
    "\\b Chapter1",
    '\\c "Table"',
    '\\d "-"',
    "\\f a",
    "\\h",
    '\\l "1-2"',
    '\\n "2-3"',
    '\\o "1-3"',
    '\\p "-"',
    "\\s chapter",
    '\\t "Title,1;Subtitle,2"',
    "\\u",
    "\\w",
    "\\x",
    "\\z "
  ].join(" "),
  "unusual spacing": 'TOC  \\o  "1-4"\\h',
  "an instruction split across runs": "split"
};

describe("TOC field instruction", () => {
  it.each(Object.entries(VARIANTS))("survives a round trip verbatim: %s", async (_name, value) => {
    const instruction = value === "split" ? ' TOC \\o "1-3" \\h \\z ' : value;
    const head =
      value === "split"
        ? `<w:p>${BEGIN}${INSTR(" TO")}${INSTR('C \\o "1-3"')}${INSTR(" \\h \\z ")}${SEP}</w:p>`
        : `<w:p>${BEGIN}${INSTR(instruction)}${SEP}</w:p>`;
    const body =
      SDT(
        `${head}${ENTRY(1, "Intro", "_Toc1", 1)}${ENTRY(2, "Details", "_Toc2", 2)}<w:p>${END}</w:p>`
      ) + TAIL;
    const source = await buildDocx({ body });
    const once = await roundTrip(source);
    expect(tocBlockInstructions(once)).toEqual([instruction]);
    expect(bodyText(once)).toBe(bodyText(`<w:body>${body}</w:body>`));
    const twice = await roundTrip(await Io.toBuffer(await Io.read(source)));
    expect(tocBlockInstructions(twice)).toEqual([instruction]);
  });

  it("keeps each of two TOCs' instructions", async () => {
    const a = " TOC \\o \\z ";
    const b = ' TOC \\h \\c "Table" \\u ';
    const body = TOC_BLOCK(a) + TOC_BLOCK(b, ENTRY(1, "T1", "_Toc2", 2)) + TAIL;
    const out = await roundTrip(await buildDocx({ body }));
    expect(tocBlockInstructions(out)).toEqual([a, b]);
    expect(bodyText(out)).toBe(bodyText(`<w:body>${body}</w:body>`));
  });

  it("reads the options and the verbatim instruction together", async () => {
    const instruction = ' TOC \\o "1-3" \\h \\z \\u ';
    const doc = await Io.read(await buildDocx({ body: TOC_BLOCK(instruction) + TAIL }));
    const toc = doc.body[0] as TableOfContents;
    expect(toc).toMatchObject({
      type: "tableOfContents",
      instruction,
      headingStyleRange: "1-3",
      hyperlink: true
    });
  });

  it("writes the changed options and keeps the switches they do not model", async () => {
    const doc = await Io.read(
      await buildDocx({ body: TOC_BLOCK(' TOC \\o "1-3" \\h \\z \\u \\n "1-1" ') + TAIL })
    );
    const toc = doc.body[0] as TableOfContents;
    (doc.body as BodyContent[])[0] = { ...toc, headingStyleRange: "2-4", hyperlink: false };
    const xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    expect(tocBlockInstructions(xml)).toEqual([' TOC \\o "2-4" \\z \\u \\n "1-1" ']);
  });

  it("replaces a bare \\o once a range is set, and keeps it otherwise", async () => {
    const doc = await Io.read(await buildDocx({ body: TOC_BLOCK(" TOC \\o \\z ") + TAIL }));
    const toc = doc.body[0] as TableOfContents;
    (doc.body as BodyContent[])[0] = { ...toc, hyperlink: true };
    let xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    expect(tocBlockInstructions(xml)).toEqual([" TOC \\h \\o \\z "]);
    (doc.body as BodyContent[])[0] = { ...toc, headingStyleRange: "1-2" };
    xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    expect(tocBlockInstructions(xml)).toEqual([' TOC \\o "1-2" \\z ']);
  });

  it("Document.addTableOfContents takes an instruction and its options from it", async () => {
    const doc = Document.create();
    Document.addTableOfContents(doc, { instruction: ' TOC \\o "1-2" \\z \\u ' });
    expect(doc.body[0]).toMatchObject({ headingStyleRange: "1-2" });
    expect((doc.body[0] as TableOfContents).hyperlink).toBeFalsy();
    const xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    expect(tocBlockInstructions(xml)).toEqual([' TOC \\o "1-2" \\z \\u ']);
  });

  it("Document.addTableOfContents without an instruction writes what it always did", async () => {
    const doc = Document.create();
    Document.addTableOfContents(doc);
    const xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    expect(tocBlockInstructions(xml)).toEqual([' TOC \\o "1-3" \\h ']);
  });
});

/** Per field in `xml`, in order: its field characters and instruction, as one string. */
function fieldShapes(xml: string): string[] {
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(xml)?.[1] ?? xml;
  return [
    ...body.matchAll(
      /<w:fldChar w:fldCharType="(\w+)"|<w:instrText[^>]*>([^<]*)<\/w:instrText>|<w:fldSimple\b/g
    )
  ].map(m =>
    m[1] ? m[1] : m[2] !== undefined ? `instr:${m[2].replace(/&quot;/g, '"').trim()}` : "fldSimple"
  );
}

/** `shapes` without the `PAGEREF` fields regenerated entries carry. */
function withoutPageRefs(shapes: readonly string[]): string[] {
  const out: string[] = [];
  for (let i = 0; i < shapes.length; i++) {
    if (shapes[i] === "begin" && shapes[i + 1]?.startsWith("instr:PAGEREF")) {
      i += 3;
      continue;
    }
    out.push(shapes[i]!);
  }
  return out;
}

describe("a TOC block written as it was read", () => {
  it("does not invent a field for a hand-written table of contents", async () => {
    const body =
      SDT(
        `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>${R("Contents")}</w:p>` +
          `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>${R("Manual one")}<w:r><w:tab/></w:r>${R("1")}</w:p>`
      ) + TAIL;
    const once = await roundTrip(await buildDocx({ body }));
    expect(fieldShapes(once)).toEqual([]);
    expect(bodyText(once)).toBe(bodyText(`<w:body>${body}</w:body>`));
    const doc = await Io.read(await buildDocx({ body }));
    expect((doc.body[0] as TableOfContents).field).toBe("none");
  });

  it("keeps the text in the run holding the field's begin character", async () => {
    const body =
      SDT(
        `<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:t>X</w:t></w:r>${INSTR(" TOC \\o ")}${SEP}` +
          `${R("Intro")}${END}</w:p>`
      ) + TAIL;
    const once = await roundTrip(await buildDocx({ body }));
    expect(bodyText(once)).toBe(bodyText(`<w:body>${body}</w:body>`));
    // One TOC field, as read: none added around it.
    expect(fieldShapes(once)).toEqual(["begin", "instr:TOC \\o", "separate", "end"]);
    expect(await roundTrip(await Io.toBuffer(await Io.read(await buildDocx({ body }))))).toBe(once);
  });

  it("reads a block-level fldSimple TOC as its field and entries", async () => {
    const instruction = ' TOC \\o "1-3" \\h ';
    const body =
      SDT(
        `<w:fldSimple w:instr="${instruction.replace(/"/g, "&quot;")}">` +
          `<w:p>${R("SimpleEntry")}<w:r><w:tab/></w:r>${R("1")}</w:p>` +
          `<w:p>${R("Second")}<w:r><w:tab/></w:r>${R("2")}</w:p></w:fldSimple>`
      ) + TAIL;
    const doc = await Io.read(await buildDocx({ body }));
    const toc = doc.body[0] as TableOfContents;
    expect(toc.instruction).toBe(instruction);
    expect(toc.cachedParagraphs).toHaveLength(2);
    const once = await roundTrip(await buildDocx({ body }));
    expect(tocBlockInstructions(once)).toEqual([instruction]);
    expect(bodyText(once)).toContain("SimpleEntry\t1¶Second\t2¶");
    expect(bodyText(once)).not.toContain("Update this table");
  });

  it("reads a block-level fldSimple outside a TOC as the field it is", async () => {
    const body =
      `<w:fldSimple w:instr=" INDEX \\c 2 "><w:p>${R("Apple, 3")}</w:p><w:p>${R("Pear, 4")}</w:p></w:fldSimple>` +
      `<w:p>${R("after")}</w:p>`;
    const once = await roundTrip(await buildDocx({ body }));
    expect(bodyText(once)).toBe("Apple, 3¶Pear, 4¶after¶");
    expect(fieldShapes(once)).toEqual(["begin", "instr:INDEX \\c 2", "separate", "end"]);
  });

  it("gives regenerated entries the field the old ones held in pieces", async () => {
    const body =
      SDT(
        `<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:t>X</w:t></w:r>${INSTR(" TOC \\o \\z ")}${SEP}` +
          `${R("Stale")}${END}</w:p>`
      ) + TAIL;
    const doc = Transform.updateTableOfContents(await Io.read(await buildDocx({ body })));
    const toc = doc.body[0] as TableOfContents;
    expect(toc.field).toBeUndefined();
    const xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    expect(tocBlockInstructions(xml)).toEqual([" TOC \\o \\z "]);
    expect(withoutPageRefs(fieldShapes(xml))).toEqual([
      "begin",
      "instr:TOC \\o \\z",
      "separate",
      "end"
    ]);
    expect(bodyText(xml)).not.toContain("Stale");
  });
});

describe("a TOC in Word's own layout", () => {
  /** Word's gallery TOC: a heading, the field opening in the first entry, closing after the last. */
  const NATIVE =
    SDT(
      `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>${R("Contents")}</w:p>` +
        `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>${BEGIN}${INSTR(' TOC \\o "1-3" \\h \\z \\u ')}${SEP}` +
        `<w:hyperlink w:anchor="_Toc1" w:history="1">${R("Intro")}<w:r><w:tab/></w:r>${PAGEREF("_Toc1", 1)}</w:hyperlink></w:p>` +
        ENTRY(2, "Details", "_Toc2", 2) +
        `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p>`
    ) + TAIL;

  /** The TOC block's paragraphs, each with its field characters and visible text. */
  const shape = (xml: string): string[] => {
    const sdt = /<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/.exec(xml)![1]!;
    return [...sdt.matchAll(/<w:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:p>)/g)].map(([p]) =>
      [
        /<w:pStyle w:val="([^"]+)"/.exec(p)?.[1] ?? "-",
        /<w:spacing [^>]*\/>/.test(p) ? "spacing" : "",
        ...[...p.matchAll(/w:fldCharType="(\w+)"|<w:t[^>]*>([^<]*)<|<w:b\/>/g)].map(
          m => m[1] ?? m[2] ?? "bold"
        )
      ].join(" ")
    );
  };

  it("is written back with the same paragraphs, the field where it was", async () => {
    const once = await roundTrip(await buildDocx({ body: NATIVE }));
    expect(shape(once)).toEqual(shape(`<w:body>${NATIVE}</w:body>`));
    expect(tocBlockInstructions(once)).toEqual([' TOC \\o "1-3" \\h \\z \\u ']);
    expect(bodyText(once)).toBe(bodyText(`<w:body>${NATIVE}</w:body>`));
  });

  it("takes changed options into its field where it is", async () => {
    const doc = await Io.read(await buildDocx({ body: NATIVE }));
    const toc = doc.body[0] as TableOfContents;
    expect(toc.field).toBe("entries");
    (doc.body as BodyContent[])[0] = { ...toc, headingStyleRange: "1-2" };
    const xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    expect(tocBlockInstructions(xml)).toEqual([' TOC \\o "1-2" \\h \\z \\u ']);
    expect(shape(xml)).toEqual(shape(`<w:body>${NATIVE}</w:body>`));
  });

  it("gets the writer's own field when its entries are regenerated", async () => {
    const doc = Transform.updateTableOfContents(await Io.read(await buildDocx({ body: NATIVE })));
    const toc = doc.body[0] as TableOfContents;
    expect(toc.field).toBeUndefined();
    const xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    expect(tocBlockInstructions(xml)).toEqual([' TOC \\o "1-3" \\h \\z \\u ']);
    expect(withoutPageRefs(fieldShapes(xml))).toEqual([
      "begin",
      'instr:TOC \\o "1-3" \\h \\z \\u',
      "separate",
      "end"
    ]);
    // `\h`: each entry links to the bookmark its page number refers to.
    expect(xml).toMatch(
      /<w:hyperlink w:anchor="(_Toc\d+)"[^>]*>(?:(?!<\/w:hyperlink>)[\s\S])*PAGEREF \1 \\h/
    );
  });

  it("leaves the builder's TOC output as it was", async () => {
    const d = Document.create();
    Document.addTableOfContents(d);
    Document.addHeading(d, "Intro", 1);
    const first = await Io.toBuffer(d);
    const xml = partText(await packageParts(first), "word/document.xml");
    const sdt = (x: string): string => /<w:sdtContent>[\s\S]*?<\/w:sdtContent>/.exec(x)![0]!;
    // The field opens in the first paragraph, beside its text, and closes in
    // one of its own — as Word writes a table of contents.
    expect(sdt(xml)).toBe(
      "<w:sdtContent>" +
        `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
        `<w:r><w:t xml:space="preserve">Update this table of contents by right-clicking and selecting &apos;Update Field&apos;</w:t></w:r></w:p>` +
        `<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
        "</w:sdtContent>"
    );
    // Read back, it keeps those paragraphs.
    expect(sdt(await roundTrip(first))).toBe(sdt(xml));
  });
});
