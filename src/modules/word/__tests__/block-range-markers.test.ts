/**
 * Range markers (bookmarks, comment ranges, …) that sit between blocks rather
 * than inside a paragraph — Word starts a bookmark spanning a table before the
 * `w:tbl`, between rows, or between cells. They must be read, kept in place,
 * written back where they were, and understood by every consumer that cares
 * about bookmarks.
 */

import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { describe, expect, it } from "vitest";

import { utf8Decoder, utf8Encoder } from "../core/internal-utils";
import { WordHtml } from "../html";
import {
  Build,
  Convert,
  Document,
  Io,
  Layout,
  Query,
  Streaming,
  Transform,
  Validation
} from "../index";
import { WordMarkdown } from "../markdown";
import type { BodyContent, DocxDocument, Paragraph, RangeMarker, Table } from "../types";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

async function entriesOf(buffer: Uint8Array): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for await (const entry of unzip(buffer).entries()) {
    out.set(entry.path.replace(/^\/+/, ""), await entry.bytes());
  }
  return out;
}

/** A real package whose `word/document.xml` body is `bodyXml`. */
async function docxWithBody(bodyXml: string): Promise<Uint8Array> {
  const doc = Document.create();
  Document.addContent(doc, Build.textParagraph("placeholder"));
  const entries = await entriesOf(await Io.toBuffer(doc));
  entries.set(
    "word/document.xml",
    utf8Encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:document xmlns:w="${W_NS}"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`
    )
  );
  const archive = zip();
  for (const [p, d] of entries) {
    archive.add(p, d);
  }
  return archive.toBuffer();
}

async function documentXml(doc: DocxDocument): Promise<string> {
  const entries = await entriesOf(await Io.toBuffer(doc));
  return utf8Decoder.decode(entries.get("word/document.xml")!);
}

/** The body of `document.xml` with paragraph/run/table formatting noise removed. */
function bodyShape(xml: string): string {
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(xml)![1]!;
  return body
    .replace(/<w:(pPr|rPr|tblPr|tblGrid|tcPr|trPr|sectPr)\b[\s\S]*?<\/w:\1>/g, "")
    .replace(/<w:(pPr|rPr|tblPr|tblGrid|tcPr|trPr|sectPr)\b[^>]*\/>/g, "")
    .replace(/<w:t\b[^>]*>/g, "<w:t>")
    .replace(/\s(w14:paraId|w14:textId|w:rsid\w*)="[^"]*"/g, "");
}

function p(text: string, extra = ""): string {
  return `<w:p><w:r><w:t>${text}</w:t></w:r>${extra}</w:p>`;
}

function tc(text: string): string {
  return `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
}

const GRID = `<w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>`;

/** `block` carrying `markers` as the range markers before it. */
function before<T extends BodyContent>(block: T, ...markers: unknown[]): T {
  return { ...block, rangeMarkersBefore: markers as RangeMarker[] };
}

/** Every range marker in the body, in document order, with where it sits. */
function markerMap(doc: {
  readonly body: readonly BodyContent[];
  readonly bodyRangeMarkersAfter?: readonly RangeMarker[];
}): string[] {
  const { body } = doc;
  const out: string[] = [];
  const describe = (m: RangeMarker): string =>
    `${m.type}#${"id" in m ? m.id : ""}${m.type === "bookmarkStart" ? `:${m.name}` : ""}`;
  const fromParagraph = (para: Paragraph, where: string): void => {
    para.children.forEach((c, i) => {
      if ("type" in c && /Range|bookmark/.test(c.type)) {
        out.push(`${where}.children[${i}] ${describe(c as RangeMarker)}`);
      }
    });
  };
  const fromTable = (table: Table, where: string): void => {
    table.rows.forEach((row, r) => {
      for (const m of row.rangeMarkersBefore ?? []) {
        out.push(`${where}.rows[${r}].before ${describe(m)}`);
      }
      row.cells.forEach((cell, c) => {
        for (const m of cell.rangeMarkersBefore ?? []) {
          out.push(`${where}.rows[${r}].cells[${c}].before ${describe(m)}`);
        }
        cell.content.forEach((block, k) => {
          const at = `${where}.rows[${r}].cells[${c}].content[${k}]`;
          if (block.type === "paragraph") {
            fromParagraph(block, at);
          } else {
            fromTable(block, at);
          }
        });
      });
      for (const m of row.rangeMarkersAfter ?? []) {
        out.push(`${where}.rows[${r}].after ${describe(m)}`);
      }
    });
    for (const m of table.rangeMarkersAfter ?? []) {
      out.push(`${where}.after ${describe(m)}`);
    }
  };
  body.forEach((block, i) => {
    for (const m of block.rangeMarkersBefore ?? []) {
      out.push(`body[${i}].before ${describe(m)}`);
    }
    if (block.type === "paragraph") {
      fromParagraph(block, `body[${i}]`);
    } else if (block.type === "table") {
      fromTable(block, `body[${i}]`);
    }
  });
  for (const m of doc.bodyRangeMarkersAfter ?? []) {
    out.push(`body.after ${describe(m)}`);
  }
  return out;
}

const SPANNING_TABLE = [
  p("Intro"),
  `<w:bookmarkStart w:id="0" w:name="tbl"/>`,
  `<w:tbl>${GRID}<w:tr>${tc("A1")}${tc("B1")}</w:tr></w:tbl>`,
  p("After", `<w:bookmarkEnd w:id="0"/>`)
].join("");

const BLOCK_TYPES = new Set([
  "paragraph",
  "table",
  "floatingImage",
  "tableOfContents",
  "math",
  "textBox",
  "checkBox",
  "drawingShape",
  "opaqueDrawing",
  "chart",
  "chartEx",
  "altChunk",
  "sdt"
]);

describe("block-level range markers — reading", () => {
  it("adds no entries to doc.body: every entry is still a block", async () => {
    const xml = [
      `<w:commentRangeStart w:id="0"/>`,
      SPANNING_TABLE,
      `<w:moveToRangeStart w:id="7" w:author="A" w:name="m"/>`,
      `<w:sdt><w:sdtPr/><w:sdtContent>${p("In SDT")}</w:sdtContent></w:sdt>`,
      `<w:moveToRangeEnd w:id="7"/><w:commentRangeEnd w:id="0"/>`,
      `<w:bookmarkStart w:id="5" w:name="tail"/><w:bookmarkEnd w:id="5"/>`
    ].join("");
    const bytes = await docxWithBody(xml);
    const doc = await Io.read(bytes);
    expect(doc.body.map(b => b.type)).toEqual(["paragraph", "table", "paragraph", "sdt"]);
    for (const block of doc.body) {
      expect(BLOCK_TYPES.has(block.type)).toBe(true);
    }
    for await (const item of Streaming.createReader(bytes)) {
      expect(BLOCK_TYPES.has(item.type)).toBe(true);
    }
    expect(markerMap(doc)).toEqual([
      "body[0].before commentRangeStart#0",
      "body[1].before bookmarkStart#0:tbl",
      "body[2].children[1] bookmarkEnd#0",
      "body[3].before moveToRangeStart#7",
      "body.after moveToRangeEnd#7",
      "body.after commentRangeEnd#0",
      "body.after bookmarkStart#5:tail",
      "body.after bookmarkEnd#5"
    ]);
  });

  it("keeps a bookmark that starts before a table and ends in a later paragraph", async () => {
    const doc = await Io.read(await docxWithBody(SPANNING_TABLE));
    expect(doc.body.map(b => b.type)).toEqual(["paragraph", "table", "paragraph"]);
    expect(doc.body[1]!.rangeMarkersBefore).toEqual([
      { type: "bookmarkStart", id: 0, name: "tbl" }
    ]);
    expect(markerMap(doc)).toEqual([
      "body[1].before bookmarkStart#0:tbl",
      "body[2].children[1] bookmarkEnd#0"
    ]);
  });

  it("keeps bookmarks between rows and between cells, column bookmarks included", async () => {
    const xml = [
      `<w:tbl>${GRID}`,
      `<w:tr>${tc("A1")}${tc("B1")}</w:tr>`,
      `<w:bookmarkStart w:id="1" w:name="row2"/>`,
      `<w:tr><w:bookmarkStart w:id="2" w:name="col1" w:colFirst="0" w:colLast="0"/>${tc("A2")}`,
      `<w:bookmarkStart w:id="3" w:name="b2"/>${tc("B2")}<w:bookmarkEnd w:id="2"/></w:tr>`,
      `<w:bookmarkEnd w:id="1"/><w:bookmarkEnd w:id="3"/>`,
      `</w:tbl>`,
      p("After")
    ].join("");
    const doc = await Io.read(await docxWithBody(xml));
    expect(markerMap(doc)).toEqual([
      "body[0].rows[1].before bookmarkStart#1:row2",
      "body[0].rows[1].cells[0].before bookmarkStart#2:col1",
      "body[0].rows[1].cells[1].before bookmarkStart#3:b2",
      "body[0].rows[1].after bookmarkEnd#2",
      "body[0].after bookmarkEnd#1",
      "body[0].after bookmarkEnd#3"
    ]);
    const col = (doc.body[0] as Table).rows[1]!.cells[0]!.rangeMarkersBefore![0]!;
    expect(col).toEqual({ type: "bookmarkStart", id: 2, name: "col1", colFirst: 0, colLast: 0 });
  });

  it("keeps markers at the very end of the body", async () => {
    const xml = [
      p("Text", `<w:bookmarkStart w:id="3" w:name="end"/>`),
      `<w:bookmarkEnd w:id="3"/>`,
      `<w:bookmarkStart w:id="4" w:name="tail"/><w:bookmarkEnd w:id="4"/>`
    ].join("");
    const doc = await Io.read(await docxWithBody(xml));
    expect(markerMap(doc)).toEqual([
      "body[0].children[1] bookmarkStart#3:end",
      "body.after bookmarkEnd#3",
      "body.after bookmarkStart#4:tail",
      "body.after bookmarkEnd#4"
    ]);
  });

  it("keeps comment, move and custom-XML range markers between blocks", async () => {
    const xml = [
      `<w:commentRangeStart w:id="0"/>`,
      p("First"),
      p("Second"),
      `<w:commentRangeEnd w:id="0"/>`,
      `<w:moveFromRangeStart w:id="7" w:author="A" w:date="2026-01-01T00:00:00Z" w:name="move1"/>`,
      p("Moved"),
      `<w:moveFromRangeEnd w:id="7"/>`,
      `<w:customXmlInsRangeStart w:id="8" w:author="A"/><w:customXmlInsRangeEnd w:id="8"/>`
    ].join("");
    const doc = await Io.read(await docxWithBody(xml));
    expect(markerMap(doc)).toEqual([
      "body[0].before commentRangeStart#0",
      "body[2].before commentRangeEnd#0",
      "body[2].before moveFromRangeStart#7",
      "body.after moveFromRangeEnd#7",
      "body.after customXmlInsRangeStart#8",
      "body.after customXmlInsRangeEnd#8"
    ]);
    expect(doc.body[2]!.rangeMarkersBefore![1]).toEqual({
      type: "moveFromRangeStart",
      id: 7,
      author: "A",
      date: "2026-01-01T00:00:00Z",
      name: "move1"
    });
  });

  it("folds markers between the blocks of a cell into its paragraphs", async () => {
    const xml = [
      `<w:tbl>${GRID}<w:tr>`,
      `<w:tc><w:bookmarkStart w:id="5" w:name="incell"/>${p("One")}<w:bookmarkEnd w:id="6"/>${p("Two")}<w:bookmarkEnd w:id="5"/></w:tc>`,
      `${tc("B1")}</w:tr></w:tbl>`,
      p("x", `<w:bookmarkStart w:id="6" w:name="prev"/>`)
    ].join("");
    const doc = await Io.read(await docxWithBody(xml));
    expect(markerMap(doc)).toEqual([
      "body[0].rows[0].cells[0].content[0].children[0] bookmarkStart#5:incell",
      "body[0].rows[0].cells[0].content[0].children[2] bookmarkEnd#6",
      "body[0].rows[0].cells[0].content[1].children[1] bookmarkEnd#5",
      "body[1].children[1] bookmarkStart#6:prev"
    ]);
  });

  it("the streaming reader yields the same markers as the batch reader", async () => {
    const bytes = await docxWithBody(
      `${SPANNING_TABLE}<w:bookmarkStart w:id="9" w:name="tail"/><w:bookmarkEnd w:id="9"/>`
    );
    const reader = Streaming.createReader(bytes);
    const body: BodyContent[] = [];
    for await (const item of reader) {
      body.push(item);
    }
    const streamed = { body, bodyRangeMarkersAfter: reader.bodyRangeMarkersAfter };
    expect(markerMap(streamed)).toEqual(markerMap(await Io.read(bytes)));
    expect(markerMap(streamed)).toContain("body.after bookmarkStart#9:tail");
  });
});

describe("block-level range markers — writing", () => {
  it("round-trips every marker to the same place", async () => {
    const xml = [
      `<w:commentRangeStart w:id="0"/>`,
      SPANNING_TABLE,
      `<w:tbl>${GRID}<w:tr>${tc("A1")}${tc("B1")}</w:tr>`,
      `<w:bookmarkStart w:id="1" w:name="row2"/>`,
      `<w:tr><w:bookmarkStart w:id="2" w:name="col1" w:colFirst="0" w:colLast="0"/>${tc("A2")}${tc("B2")}<w:bookmarkEnd w:id="2"/></w:tr>`,
      `<w:bookmarkEnd w:id="1"/></w:tbl>`,
      `<w:commentRangeEnd w:id="0"/>`,
      p("Tail", `<w:bookmarkStart w:id="3" w:name="end"/>`),
      `<w:bookmarkEnd w:id="3"/>`
    ].join("");
    const first = await Io.read(await docxWithBody(xml));
    const written = await documentXml(first);
    const second = await Io.read(await Io.toBuffer(first));
    expect(markerMap(second)).toEqual(markerMap(first));
    expect(markerMap(first)).toHaveLength(10);

    const shape = bodyShape(written);
    expect(shape).toContain(`<w:bookmarkStart w:id="0" w:name="tbl"/><w:tbl>`);
    expect(shape).toContain(`<w:t>After</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`);
    expect(shape).toContain(`</w:tr><w:bookmarkStart w:id="1" w:name="row2"/><w:tr>`);
    expect(shape).toContain(
      `<w:tr><w:bookmarkStart w:id="2" w:name="col1" w:colFirst="0" w:colLast="0"/><w:tc>`
    );
    expect(shape).toContain(
      `</w:tc><w:bookmarkEnd w:id="2"/></w:tr><w:bookmarkEnd w:id="1"/></w:tbl>`
    );
    expect(shape).toMatch(/^<w:commentRangeStart w:id="0"\/><w:p>/);
    expect(shape).toContain(`</w:tbl><w:commentRangeEnd w:id="0"/><w:p>`);
    // The final section properties still close the body, after the last marker.
    expect(written).toMatch(/<w:bookmarkEnd w:id="3"\/><w:sectPr[\s>/]/);
  });

  it("pairs a block-level start with its end through the shared bookmark id map", async () => {
    // Two bookmarks share id 0 in the model: one inside the first paragraph,
    // one starting before the table. The writer renumbers the second start,
    // and the end in the last paragraph must follow it rather than be dropped.
    const xml = [
      p("Intro", `<w:bookmarkStart w:id="0" w:name="a"/><w:bookmarkEnd w:id="0"/>`),
      SPANNING_TABLE.slice(p("Intro").length)
    ].join("");
    const doc = await Io.read(await docxWithBody(xml));
    const written = bodyShape(await documentXml(doc));
    const startId = /<w:bookmarkStart w:id="(\d+)" w:name="tbl"\/><w:tbl>/.exec(written)?.[1];
    expect(startId).toBeDefined();
    expect(startId).not.toBe("0");
    expect(written).toContain(`<w:t>After</w:t></w:r><w:bookmarkEnd w:id="${startId}"/></w:p>`);
    const reread = await Io.read(await Io.toBuffer(doc));
    expect(Query.getBookmark(reread, "tbl")?.between).toBe("before");
  });

  it("the streaming writer writes block-level markers in place", async () => {
    const stream = Streaming.createWriter({
      bodyRangeMarkersAfter: [
        Build.bookmarkStart(5, "tail") as RangeMarker,
        Build.bookmarkEnd(5) as RangeMarker
      ]
    });
    stream.add(Build.textParagraph("Intro"));
    stream.add(before(Build.table([Build.row([Build.cell("A1")])]), Build.bookmarkStart(4, "tbl")));
    stream.add(Build.paragraph([Build.text("After"), Build.bookmarkEnd(4)]));
    const bytes = await stream.finalize();
    const xml = utf8Decoder.decode((await entriesOf(bytes)).get("word/document.xml")!);
    expect(bodyShape(xml)).toContain(`<w:bookmarkStart w:id="4" w:name="tbl"/><w:tbl>`);
    expect(bodyShape(xml)).toContain(`<w:bookmarkEnd w:id="4"/></w:p>`);
    const doc = await Io.read(bytes);
    expect(markerMap(doc)).toEqual([
      "body[1].before bookmarkStart#4:tbl",
      "body[2].children[1] bookmarkEnd#4",
      "body.after bookmarkStart#5:tail",
      "body.after bookmarkEnd#5"
    ]);
  });
});

describe("block-level range markers — consumers", () => {
  it("Query.getBookmark finds a bookmark that starts between blocks", async () => {
    const doc = await Io.read(await docxWithBody(SPANNING_TABLE));
    expect(Query.getBookmark(doc, "tbl")).toEqual({
      bookmark: { type: "bookmarkStart", id: 0, name: "tbl" },
      paragraphIndex: 1,
      childIndex: 0,
      between: "before",
      markerIndex: 0
    });
    expect(Query.getBookmark(doc, "missing")).toBeUndefined();

    const tail = await Io.read(
      await docxWithBody(
        p("A", `<w:bookmarkStart w:id="1" w:name="inside"/><w:bookmarkEnd w:id="1"/>`) +
          `<w:bookmarkStart w:id="2" w:name="end"/><w:bookmarkEnd w:id="2"/>`
      )
    );
    // A bookmark inside a paragraph keeps the shape it always had.
    expect(Query.getBookmark(tail, "inside")).toEqual({
      bookmark: { type: "bookmarkStart", id: 1, name: "inside" },
      paragraphIndex: 0,
      childIndex: 1
    });
    expect(Query.getBookmark(tail, "end")).toMatchObject({
      paragraphIndex: 0,
      childIndex: 0,
      between: "after",
      markerIndex: 0
    });
  });

  it("resolves the page of a bookmark that starts before a table", () => {
    const doc = Document.create();
    Document.addContent(doc, Build.textParagraph("Page one"));
    Document.addContent(doc, Build.paragraph([Build.text("Page two")], { pageBreakBefore: true }));
    Document.addContent(
      doc,
      before(Build.table([Build.row([Build.cell("A1")])]), Build.bookmarkStart(0, "tbl"))
    );
    Document.addContent(doc, Build.paragraph([Build.text("After"), Build.bookmarkEnd(0)]));
    expect(Layout.getPageMap(doc).bookmarkPages.get("tbl")).toBe(2);
    expect(Layout.compute(doc).bookmarkPages.get("tbl")).toBe(2);
  });

  it("resolves the page of a bookmark that starts between table rows", () => {
    const rows = Array.from({ length: 80 }, (_, i) => Build.row([Build.cell(`Row ${i}`)]));
    const marked = rows.map((row, i) =>
      i === 70
        ? { ...row, rangeMarkersBefore: [Build.bookmarkStart(0, "late") as RangeMarker] }
        : row
    );
    const doc = Document.create();
    Document.addContent(doc, { ...Build.table(rows), rows: marked });
    const full = Layout.compute(doc);
    expect(full.pages.length).toBeGreaterThan(1);
    expect(full.bookmarkPages.get("late")).toBeGreaterThan(1);
  });

  it("a block-level marker occupies no space in the layout", () => {
    const plain = Document.create();
    const marked = Document.create();
    for (let i = 0; i < 60; i++) {
      Document.addContent(plain, Build.textParagraph(`Line ${i}`));
      Document.addContent(
        marked,
        before(
          Build.textParagraph(`Line ${i}`),
          Build.bookmarkEnd(i - 1),
          Build.bookmarkStart(i, `b${i}`)
        )
      );
    }
    expect(Layout.compute(marked).pages.length).toBe(Layout.compute(plain).pages.length);
    expect(Layout.getPageMap(marked).pageCount).toBe(Layout.getPageMap(plain).pageCount);
  });

  it("HTML writes the anchor where the range begins, never between rows", async () => {
    const doc = await Io.read(
      await docxWithBody(
        [
          SPANNING_TABLE,
          `<w:tbl>${GRID}<w:tr>${tc("A1")}${tc("B1")}</w:tr><w:bookmarkStart w:id="1" w:name="row2"/><w:tr>${tc("A2")}${tc("B2")}</w:tr><w:bookmarkEnd w:id="1"/></w:tbl>`
        ].join("")
      )
    );
    const { html } = WordHtml.fromDocx(doc);
    expect(html).toMatch(/<td[^>]*>\s*<p[^>]*>\s*<a id="bookmark-tbl" data-name="tbl"><\/a>/);
    expect(html).toMatch(
      /<td[^>]*>\s*<p[^>]*>\s*<a id="bookmark-row2" data-name="row2"><\/a>\s*(?:<span[^>]*>\s*)?A2/
    );
    expect(html).not.toMatch(/<\/tr>\s*<a /);
    expect(html).not.toContain("Unknown element type");
  });

  it("Markdown keeps a list together across a block-level marker", async () => {
    const doc = await WordMarkdown.toDocx("- one\n- two");
    const plain = WordMarkdown.fromDocx(doc);
    const [one, two] = doc.body;
    const marked: DocxDocument = {
      ...doc,
      body: [one!, before(two!, Build.bookmarkStart(0, "mid"))],
      bodyRangeMarkersAfter: [Build.bookmarkEnd(0) as RangeMarker]
    };
    expect(WordMarkdown.fromDocx(marked)).toBe(plain);
  });

  it("ODT moves block-level bookmarks into paragraphs and adds no empty ones", async () => {
    const doc = await Io.read(
      await docxWithBody(
        [
          SPANNING_TABLE,
          `<w:tbl>${GRID}<w:tr>${tc("A1")}${tc("B1")}</w:tr><w:bookmarkStart w:id="1" w:name="row2"/><w:tr>${tc("A2")}${tc("B2")}</w:tr><w:bookmarkEnd w:id="1"/></w:tbl>`,
          p("Tail"),
          `<w:bookmarkStart w:id="2" w:name="tail"/><w:bookmarkEnd w:id="2"/>`
        ].join("")
      )
    );
    const odt = await Convert.toOdt(doc);
    const content = utf8Decoder.decode((await entriesOf(odt)).get("content.xml")!);
    expect(content).not.toMatch(/<text:p\/>/);
    for (const name of ["tbl", "row2", "tail"]) {
      expect(content).toContain(`<text:bookmark-start text:name="${name}"/>`);
      expect(content).toContain(`<text:bookmark-end text:name="${name}"/>`);
    }
    const back = await Convert.fromOdt(odt);
    const names = markerMap(back).filter(m => m.includes("bookmarkStart"));
    expect(names.map(m => m.split(":").pop())).toEqual(["tbl", "row2", "tail"]);
  });

  it("validation accepts block-level markers", async () => {
    const doc = await Io.read(await docxWithBody(SPANNING_TABLE));
    const result = Validation.validate(doc);
    expect(result.issues.filter(i => i.rule === "unknown-body-type")).toEqual([]);
  });

  it("accepting all revisions removes move ranges between blocks and keeps bookmarks of deleted rows", async () => {
    const xml = [
      `<w:moveFromRangeStart w:id="7" w:author="A" w:name="m"/>`,
      p("Moved"),
      `<w:moveFromRangeEnd w:id="7"/>`,
      `<w:tbl>${GRID}`,
      `<w:bookmarkStart w:id="1" w:name="gone"/>`,
      `<w:tr><w:trPr><w:del w:id="9" w:author="A"/></w:trPr>${tc("A1")}${tc("B1")}</w:tr>`,
      `<w:tr>${tc("A2")}${tc("B2")}</w:tr>`,
      `<w:bookmarkEnd w:id="1"/></w:tbl>`
    ].join("");
    const doc = await Io.read(await docxWithBody(xml));
    Query.acceptAllRevisions(doc);
    expect(markerMap(doc)).toEqual([
      "body[1].rows[0].before bookmarkStart#1:gone",
      "body[1].after bookmarkEnd#1"
    ]);
  });

  it("the TOC/REF field engine takes a block-level bookmark's text from the paragraph it opens", async () => {
    const doc = await Io.read(
      await docxWithBody(
        [
          p("Intro"),
          `<w:bookmarkStart w:id="0" w:name="target"/>`,
          p("Target text", `<w:bookmarkEnd w:id="0"/>`),
          `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF target </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>stale</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
        ].join("")
      )
    );
    const updated = Transform.updateFields(doc);
    const fields = (updated.body[2] as Paragraph).children.flatMap(c =>
      "content" in c && !("type" in c) ? c.content.filter(x => x.type === "field") : []
    );
    expect(fields.map(f => (f as { cachedValue?: string }).cachedValue)).toEqual(["Target text"]);
  });
});
