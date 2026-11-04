/**
 * A complex field that is more than one run of field characters around a
 * result — one spanning paragraphs (a TOC, INDEX or BIBLIOGRAPHY without an
 * SDT, as LibreOffice and older Word write them), one nesting another, one
 * whose field characters share a run with text — is read field character by
 * field character: each paragraph keeps its own runs, and the begin,
 * separate and end characters stay in the paragraphs and at the places they
 * were. A field read whole used to be moved to its last paragraph, with the
 * runs of the paragraphs it spanned merged into one.
 */

import { describe, expect, it } from "vitest";

import { extractParagraphText } from "../core/text-utils";
import { Build, Document, Io, Streaming, Transform } from "../index";
import type { BodyContent, FieldContent, Paragraph, Run, Table } from "../types";
import { buildDocx, packageParts, partText } from "./roundtrip-integrity.helpers";

const R = (t: string, rPr = ""): string => `<w:r>${rPr}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const TAB = `<w:r><w:tab/></w:r>`;
const CH = (type: string, extra = ""): string =>
  `<w:r><w:fldChar w:fldCharType="${type}"/>${extra}</w:r>`;
const BEGIN = CH("begin");
const SEP = CH("separate");
const END = CH("end");
const INSTR = (t: string): string =>
  `<w:r><w:instrText xml:space="preserve">${t.replace(/"/g, "&quot;")}</w:instrText></w:r>`;
const P = (inner: string): string => `<w:p>${inner}</w:p>`;

/**
 * Per paragraph, its direct runs as tokens: field characters, instruction
 * text (adjacent pieces joined, outer spaces trimmed — a field read whole
 * keeps its instruction trimmed), text and tabs.
 */
function fieldStructure(xml: string): string[][] {
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(xml)?.[1] ?? xml;
  return [...body.matchAll(/<w:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:p>)/g)].map(([, inner]) => {
    const tokens: string[] = [];
    const runs = (inner ?? "").replace(/<w:hyperlink\b[\s\S]*?<\/w:hyperlink>/g, "<link/>");
    for (const m of runs.matchAll(
      /<w:fldChar w:fldCharType="(\w+)"|<w:instrText[^>]*>([^<]*)<\/w:instrText>|<w:t[^>]*>([^<]*)<\/w:t>|<w:tab\/>|<link\/>/g
    )) {
      if (m[1]) {
        tokens.push(m[1]);
      } else if (m[2] !== undefined) {
        const text = m[2].replace(/&quot;/g, '"');
        const last = tokens.length - 1;
        if (last >= 0 && tokens[last]!.startsWith("instr:")) {
          tokens[last] += text;
        } else {
          tokens.push(`instr:${text}`);
        }
      } else if (m[3] !== undefined) {
        tokens.push(`t:${m[3]}`);
      } else {
        tokens.push(m[0] === "<link/>" ? "link" : "tab");
      }
    }
    return tokens.map(t => (t.startsWith("instr:") ? `instr:${t.slice(6).trim()}` : t));
  });
}

async function documentXml(source: Uint8Array): Promise<string> {
  return partText(await packageParts(source), "word/document.xml");
}

/** Read → write twice; the field structure of both writes, and of the source. */
async function roundTrips(
  body: string
): Promise<{ input: string[][]; once: string[][]; twice: string[][] }> {
  const source = await buildDocx({ body });
  const once = await Io.toBuffer(await Io.read(source));
  const twice = await Io.toBuffer(await Io.read(once));
  return {
    input: fieldStructure(`<w:body>${body}</w:body>`),
    once: fieldStructure(await documentXml(once)),
    twice: fieldStructure(await documentXml(twice))
  };
}

const TOC_INSTR = ' TOC \\o "1-3" \\h \\z \\u ';

const CASES: Record<string, string> = {
  "a TOC spanning paragraphs, entries as plain runs":
    P(BEGIN + INSTR(TOC_INSTR) + SEP + R("Entry one") + TAB + R("1")) +
    P(R("Entry two") + TAB + R("2")) +
    P(R("Entry three") + TAB + R("3") + END),
  "a TOC whose field characters have paragraphs of their own":
    P(BEGIN + INSTR(TOC_INSTR) + SEP) +
    P(R("Entry one") + TAB + R("1")) +
    P(R("Entry two") + TAB + R("2")) +
    P(END),
  "a field whose separator is in the next paragraph":
    P(R("lead ") + BEGIN + INSTR(" INDEX \\c 2 ")) + P(SEP + R("Apple, 3")) + P(R("Pear, 4") + END),
  "a BIBLIOGRAPHY with a nested field in its result":
    P(BEGIN + INSTR(" BIBLIOGRAPHY ") + SEP + R("Smith, 2020. ")) +
    P(BEGIN + INSTR(" PAGE ") + SEP + R("7") + END + R(" tail") + END),
  "fields nested in one paragraph": P(
    R("Dear ") +
      BEGIN +
      INSTR(" IF ") +
      BEGIN +
      INSTR(" MERGEFIELD Title ") +
      SEP +
      R("«Title»") +
      END +
      INSTR(' = "Dr" "Doctor" "Friend" ') +
      SEP +
      R("Friend") +
      END +
      R(",")
  ),
  "text in the run holding the begin character": P(
    CH("begin", `<w:t>X</w:t>`) + INSTR(" PAGE ") + SEP + R("1") + END + R(" after")
  ),
  "text in the run holding the end character": P(
    BEGIN + INSTR(" PAGE ") + SEP + `<w:r><w:t>Y</w:t><w:fldChar w:fldCharType="end"/></w:r>`
  )
};

describe("a complex field read field character by field character", () => {
  it.each(Object.entries(CASES))("keeps its structure: %s", async (_name, body) => {
    const { input, once, twice } = await roundTrips(body);
    expect(once).toEqual(input);
    expect(twice).toEqual(input);
  });

  it("keeps each paragraph's own runs and text in the model", async () => {
    const doc = await Io.read(
      await buildDocx({ body: CASES["a TOC spanning paragraphs, entries as plain runs"]! })
    );
    const paragraphs = doc.body.filter((b): b is Paragraph => b.type === "paragraph");
    expect(paragraphs.map(p => extractParagraphText(p))).toEqual([
      "Entry one\t1",
      "Entry two\t2",
      "Entry three\t3"
    ]);
    const pieces = paragraphs.map(p =>
      p.children
        .flatMap(c => ("content" in c && Array.isArray(c.content) ? (c as Run).content : []))
        .filter((c): c is FieldContent => c.type === "field")
        .map(f => [f.fieldChar, f.instruction])
    );
    expect(pieces).toEqual([
      [
        ["begin", TOC_INSTR],
        ["separate", ""]
      ],
      [],
      [["end", ""]]
    ]);
  });

  it("keeps a field spanning the paragraphs of a table cell", async () => {
    const cell = CASES["a TOC spanning paragraphs, entries as plain runs"]!;
    const body = `<w:tbl><w:tr><w:tc>${cell}</w:tc></w:tr></w:tbl>${P(R("after"))}`;
    const { input, once, twice } = await roundTrips(body);
    expect(once).toEqual(input);
    expect(twice).toEqual(input);
    const doc = await Io.read(await buildDocx({ body }));
    expect((doc.body[0] as Table).rows[0]!.cells[0]!.content).toHaveLength(3);
  });

  it("keeps a field spanning the paragraphs of a header", async () => {
    const header = P(BEGIN + INSTR(" STYLEREF 1 ") + SEP + R("Chapter")) + P(R("One") + END);
    const source = await buildDocx({ body: P(R("body")), header });
    const out = await packageParts(await Io.toBuffer(await Io.read(source)));
    const headerPath = [...out.keys()].find(k => /^word\/header\d+\.xml$/.test(k))!;
    expect(fieldStructure(partText(out, headerPath))).toEqual(fieldStructure(header));
  });

  it("is regenerated by Transform.updateFields, its field kept", async () => {
    // No heading to list: the entries go, the field's characters stay.
    const body = CASES["a TOC spanning paragraphs, entries as plain runs"]!;
    const doc = await Io.read(await buildDocx({ body }));
    const updated = Transform.updateFields(doc);
    expect(fieldStructure(await documentXml(await Io.toBuffer(updated)))).toEqual([
      ["begin", 'instr:TOC \\o "1-3" \\h \\z \\u', "separate"],
      ["end"]
    ]);
  });

  it.each(Object.entries(CASES))(
    "is read the same by the streaming reader: %s",
    async (_name, body) => {
      const source = await buildDocx({ body });
      const streamed: BodyContent[] = [];
      for await (const block of Streaming.createReader(source)) {
        streamed.push(block);
      }
      expect(streamed).toEqual((await Io.read(source)).body);
    }
  );

  it("still reads a field whole past proofing marks and page-break records", async () => {
    const doc = await Io.read(
      await buildDocx({
        body: P(
          BEGIN +
            `<w:proofErr w:type="spellStart"/>` +
            INSTR(" MERGEFIELD Nmae ") +
            `<w:proofErr w:type="spellEnd"/>` +
            SEP +
            `<w:r><w:lastRenderedPageBreak/><w:t>«Nmae»</w:t></w:r>` +
            END
        )
      })
    );
    const run = (doc.body[0] as Paragraph).children[0] as Run;
    expect(run.content).toEqual([
      { type: "field", instruction: "MERGEFIELD Nmae", cachedValue: "«Nmae»" }
    ]);
  });

  it("reads the line breaks the writer writes in a result back into the field", async () => {
    const field: FieldContent = {
      type: "field",
      instruction: "INDEX \\c 2",
      cachedValue: "gadget 3\nwidget 2"
    };
    const doc = Document.create();
    Document.addParagraph(doc, Build.paragraph([{ content: [field] }]));
    const read = await Io.read(await Io.toBuffer(doc));
    expect(((read.body[0] as Paragraph).children[0] as Run).content).toEqual([field]);
  });

  it("still reads a field held by one paragraph whole", async () => {
    const doc = await Io.read(
      await buildDocx({ body: P(BEGIN + INSTR(" PAGE ") + SEP + R("4") + END) })
    );
    const run = (doc.body[0] as Paragraph).children[0] as Run;
    expect(run.content).toEqual([{ type: "field", instruction: "PAGE", cachedValue: "4" }]);
  });
});
