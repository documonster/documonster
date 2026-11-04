/**
 * Range markers in repeated content.
 *
 * A template loop or a repeating section copies its content once per item, and
 * a copy of a bookmark is a second bookmark with the same name — which Word
 * rejects — while a copy of a comment range anchors one comment twice and a
 * copy of a revision range records one edit twice. The copies together must
 * hold each range once, with every start still paired with its end:
 *
 * - a permission range starts on the first copy and ends on the last, so one
 *   editable region spans every item;
 * - every other range (bookmark, comment and its reference, move, custom XML)
 *   stays on the first copy only;
 * - a range crossing the boundary of the content (one end outside) keeps a
 *   start inside on the first copy and an end inside on the last, so it still
 *   covers every item.
 */

import { unzip } from "@archive/read-archive";
import { describe, expect, it } from "vitest";

import { utf8Decoder } from "../core/internal-utils";
import { Document, Io, Query, Template } from "../index";
import type {
  BodyContent,
  DocxDocument,
  Paragraph,
  ParagraphChild,
  RangeMarker,
  Run,
  StructuredDocumentTag,
  Table,
  TableCell
} from "../types";

const ITEMS = ["a", "b", "c"];

function run(text: string, rangeMarkersBefore?: RangeMarker[]): Run {
  return {
    ...(rangeMarkersBefore ? { rangeMarkersBefore } : {}),
    content: [{ type: "text", text }]
  };
}

/**
 * A paragraph whose text is `text`, wrapped in one range of every kind: a
 * bookmark, a comment range with its reference, a move range, a custom-XML
 * revision range and a permission range.
 */
function markedParagraph(text: string): Paragraph {
  const children: ParagraphChild[] = [
    { type: "bookmarkStart", id: 1, name: "item" },
    { type: "commentRangeStart", id: 0 },
    { type: "moveToRangeStart", id: 5, author: "A", name: "move1" },
    { type: "customXmlInsRangeStart", id: 6, author: "A" },
    run(text, [{ type: "permStart", id: 9, edGrp: "everyone" }]),
    { type: "customXmlInsRangeEnd", id: 6 },
    { type: "moveToRangeEnd", id: 5 },
    { type: "commentRangeEnd", id: 0 },
    { type: "bookmarkEnd", id: 1 },
    { type: "commentReference", id: 0 }
  ];
  return { type: "paragraph", children, rangeMarkersAfter: [{ type: "permEnd", id: 9 }] };
}

const p = (text: string): Paragraph => ({ type: "paragraph", children: [run(text)] });
const cell = (content: Paragraph): TableCell => ({ content: [content] });

/** What every repeated case must come to: each range once, permissions spanning all items. */
const EXPECTED = [
  "bookmarkStart:1",
  "commentRangeStart:0",
  "moveToRangeStart:5",
  "customXmlInsRangeStart:6",
  "permStart:9",
  "customXmlInsRangeEnd:6",
  "moveToRangeEnd:5",
  "commentRangeEnd:0",
  "bookmarkEnd:1",
  "permEnd:9"
];

function docOf(body: BodyContent[]): DocxDocument {
  const doc = Document.create();
  for (const block of body) {
    Document.addContent(doc, block);
  }
  return doc;
}

function markers(doc: DocxDocument): string[] {
  const out: string[] = [];
  Query.each(doc, {
    visitRangeMarker: m => void out.push(`${m.type}:${(m as { id: number }).id}`)
  });
  return out;
}

function commentReferences(doc: DocxDocument): number {
  let count = 0;
  Query.each(doc, {
    enterParagraph(para) {
      count += para.children.filter(c => "type" in c && c.type === "commentReference").length;
    }
  });
  return count;
}

async function documentXml(doc: DocxDocument): Promise<string> {
  for await (const entry of unzip(await Io.toBuffer(doc)).entries()) {
    if (entry.path.replace(/^\/+/, "") === "word/document.xml") {
      return utf8Decoder.decode(await entry.bytes());
    }
  }
  throw new Error("no document.xml");
}

/** The written package has one `item` bookmark and one paired permission range. */
async function expectWrittenOnce(doc: DocxDocument): Promise<void> {
  const xml = await documentXml(doc);
  const names = [...xml.matchAll(/<w:bookmarkStart\b[^>]*w:name="([^"]*)"/g)].map(m => m[1]);
  expect(names.filter(n => n === "item")).toHaveLength(1);
  expect(new Set(names).size).toBe(names.length);
  const bookmarkIds = [...xml.matchAll(/<w:bookmarkStart\b[^>]*w:id="(\d+)"/g)].map(m => m[1]);
  for (const id of bookmarkIds) {
    expect(xml.match(new RegExp(`<w:bookmarkEnd w:id="${id}"`, "g"))).toHaveLength(1);
  }
  const perms = [...xml.matchAll(/<w:perm(Start|End) w:id="(\d+)"/g)].map(m => m[1]);
  expect(perms).toEqual(["Start", "End"]);
}

async function expectEachRangeOnce(doc: DocxDocument): Promise<void> {
  expect(markers(doc)).toEqual(EXPECTED);
  expect(commentReferences(doc)).toBe(1);
  await expectWrittenOnce(doc);
}

/** The text of every paragraph, tables included. */
function texts(doc: DocxDocument): string[] {
  const out: string[] = [];
  Query.each(doc, {
    enterParagraph(para) {
      out.push(
        para.children
          .flatMap(c => ("content" in c && !("type" in c) ? c.content : []))
          .map(c => (c.type === "text" ? c.text : ""))
          .join("")
      );
    }
  });
  return out;
}

describe("range markers in repeated content", () => {
  it("block {{#each}}", async () => {
    const doc = docOf([p("{{#each xs}}"), markedParagraph("{{.}}"), p("{{/each}}")]);
    const filled = Template.fill(doc, { xs: ITEMS });
    expect(texts(filled)).toEqual(ITEMS);
    await expectEachRangeOnce(filled);
  });

  it("block {{#each}} with a single item keeps everything", async () => {
    const doc = docOf([p("{{#each xs}}"), markedParagraph("{{.}}"), p("{{/each}}")]);
    await expectEachRangeOnce(Template.fill(doc, { xs: ["only"] }));
  });

  it("table row {{#each}} on one row", async () => {
    const table: Table = {
      type: "table",
      rows: [{ cells: [cell(markedParagraph("{{#each xs}}{{.}}{{/each}}"))] }]
    };
    const filled = Template.fill(docOf([table]), { xs: ITEMS });
    expect((filled.body[0] as Table).rows).toHaveLength(3);
    await expectEachRangeOnce(filled);
  });

  it("table row {{#each}} over several rows", async () => {
    const table: Table = {
      type: "table",
      rows: [
        { cells: [cell(p("{{#each xs}}"))] },
        { cells: [cell(markedParagraph("{{.}}"))] },
        { cells: [cell(p("{{/each}}"))] }
      ]
    };
    // Each row used to be filled on its own, so the first one read as an
    // unclosed inline loop and a loop over several rows always threw.
    const filled = Template.fill(docOf([table]), { xs: ITEMS });
    expect((filled.body[0] as Table).rows).toHaveLength(9);
    expect(texts(filled)).toEqual(ITEMS.flatMap(x => ["", x, ""]));
    await expectEachRangeOnce(filled);
  });

  it("table {{#cols}}", async () => {
    const table: Table = {
      type: "table",
      rows: [{ cells: [cell(p("Name")), cell(markedParagraph("{{#cols xs}}{{.}}{{/cols}}"))] }]
    };
    const filled = Template.fill(docOf([table]), { xs: ITEMS }, { richPlaceholders: true });
    expect((filled.body[0] as Table).rows[0]!.cells).toHaveLength(4);
    await expectEachRangeOnce(filled);
  });

  it("a permission range crossing the loop's end still spans every item", async () => {
    const opened: Paragraph = {
      type: "paragraph",
      children: [run("{{.}}", [{ type: "permStart", id: 3, edGrp: "everyone" }])]
    };
    const doc = docOf([
      p("{{#each xs}}"),
      opened,
      p("{{/each}}"),
      {
        type: "paragraph",
        children: [run("tail")],
        rangeMarkersAfter: [{ type: "permEnd", id: 3 }]
      }
    ]);
    const filled = Template.fill(doc, { xs: ITEMS });
    expect(markers(filled)).toEqual(["permStart:3", "permEnd:3"]);
    expect(texts(filled).indexOf("a")).toBe(0);
  });

  describe("a range crossing the boundary of the repeated content covers every copy", () => {
    const KINDS: [string, RangeMarker, RangeMarker][] = [
      ["bookmark", { type: "bookmarkStart", id: 1, name: "x" }, { type: "bookmarkEnd", id: 1 }],
      ["comment", { type: "commentRangeStart", id: 0 }, { type: "commentRangeEnd", id: 0 }],
      [
        "move",
        { type: "moveToRangeStart", id: 5, author: "A", name: "m" },
        { type: "moveToRangeEnd", id: 5 }
      ],
      [
        "custom XML",
        { type: "customXmlInsRangeStart", id: 6, author: "A" },
        { type: "customXmlInsRangeEnd", id: 6 }
      ],
      ["permission", { type: "permStart", id: 9, edGrp: "everyone" }, { type: "permEnd", id: 9 }]
    ];

    /** Range markers, and `item` for each item's run, in document order. */
    function sequence(doc: DocxDocument): string[] {
      const out: string[] = [];
      Query.each(doc, {
        visitRangeMarker: m => void out.push(m.type),
        enterRun(r) {
          const text = r.content.map(c => (c.type === "text" ? c.text : "")).join("");
          if (text !== "" && text !== "head" && text !== "tail") {
            out.push("item");
          }
        }
      });
      return out;
    }

    /** The item paragraph, with `start` before its text or `end` after it. */
    function item(text: string, inside: { start?: RangeMarker; end?: RangeMarker }): Paragraph {
      return {
        type: "paragraph",
        children: [run(text, inside.start ? [inside.start] : undefined)],
        ...(inside.end ? { rangeMarkersAfter: [inside.end] } : {})
      };
    }

    type Shape = (inner: Paragraph) => {
      body: BodyContent[];
      fill: (d: DocxDocument) => DocxDocument;
    };
    const SHAPES: [string, Shape][] = [
      [
        "block {{#each}}",
        inner => ({
          body: [p("{{#each xs}}"), inner, p("{{/each}}")],
          fill: d => Template.fill(d, { xs: ITEMS })
        })
      ],
      [
        "row {{#each}}",
        inner => ({
          body: [
            {
              type: "table",
              rows: [
                { cells: [cell(p("{{#each xs}}"))] },
                { cells: [cell(inner)] },
                { cells: [cell(p("{{/each}}"))] }
              ]
            }
          ],
          fill: d => Template.fill(d, { xs: ITEMS })
        })
      ],
      [
        "{{#cols}}",
        inner => ({
          body: [{ type: "table", rows: [{ cells: [cell(inner)] }] }],
          fill: d => Template.fill(d, { xs: ITEMS }, { richPlaceholders: true })
        })
      ],
      [
        "repeating section",
        inner => ({
          body: [
            {
              type: "sdt",
              properties: {
                repeatingSection: {},
                dataBinding: { storeItemId: "{AAA}", xpath: "/root/item" }
              },
              content: [inner]
            }
          ],
          fill: d =>
            Query.resolveDataBindings(
              d,
              new Map([["{AAA}", "<root><item>a</item><item>b</item><item>c</item></root>"]])
            )
        })
      ]
    ];

    const textFor = (shape: string): string =>
      shape === "{{#cols}}" ? "{{#cols xs}}{{.}}{{/cols}}" : "{{.}}";

    for (const [shape, build] of SHAPES) {
      for (const [kind, start, end] of KINDS) {
        const ends = (seq: string[]): string[] =>
          seq.filter(s => s === start.type || s === end.type);

        it(`${shape}: ${kind} starting before and ending inside`, () => {
          const { body, fill } = build(item(textFor(shape), { end }));
          const head: Paragraph = { type: "paragraph", children: [run("head")] };
          const seq = sequence(fill(docOf([{ ...head, rangeMarkersAfter: [start] }, ...body])));
          expect(ends(seq)).toEqual([start.type, end.type]);
          expect(seq.filter(s => s === "item")).toHaveLength(3);
          expect(seq.indexOf(end.type)).toBeGreaterThan(seq.lastIndexOf("item"));
        });

        it(`${shape}: ${kind} starting inside and ending after`, () => {
          const { body, fill } = build(item(textFor(shape), { start }));
          const tail: Paragraph = { type: "paragraph", children: [run("tail")] };
          const seq = sequence(fill(docOf([...body, { ...tail, rangeMarkersAfter: [end] }])));
          expect(ends(seq)).toEqual([start.type, end.type]);
          expect(seq.filter(s => s === "item")).toHaveLength(3);
          expect(seq.indexOf(start.type)).toBeLessThan(seq.indexOf("item"));
        });
      }
    }

    it("a comment reference follows its end onto the last copy", () => {
      const inner: Paragraph = {
        type: "paragraph",
        children: [
          run("{{.}}"),
          { type: "commentRangeEnd", id: 0 },
          { type: "commentReference", id: 0 }
        ]
      };
      const head: Paragraph = {
        type: "paragraph",
        children: [{ type: "commentRangeStart", id: 0 }, run("head")]
      };
      const filled = Template.fill(docOf([head, p("{{#each xs}}"), inner, p("{{/each}}")]), {
        xs: ITEMS
      });
      expect(commentReferences(filled)).toBe(1);
      const last = filled.body[filled.body.length - 1] as Paragraph;
      expect(last.children.some(c => "type" in c && c.type === "commentReference")).toBe(true);
    });
  });

  it("data-binding repeating section", async () => {
    const sdt: StructuredDocumentTag = {
      type: "sdt",
      properties: {
        repeatingSection: {},
        dataBinding: { storeItemId: "{AAA}", xpath: "/root/item" }
      },
      content: [markedParagraph("row")]
    };
    const resolved = Query.resolveDataBindings(
      docOf([sdt]),
      new Map([["{AAA}", "<root><item>1</item><item>2</item><item>3</item></root>"]])
    );
    expect((resolved.body[0] as StructuredDocumentTag).content).toHaveLength(3);
    await expectEachRangeOnce(resolved);
  });
});
