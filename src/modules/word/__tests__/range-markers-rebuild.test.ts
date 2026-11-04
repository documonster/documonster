/**
 * Range markers survive code that rebuilds runs, paragraphs and blocks.
 *
 * A range permission rides on the paragraph child after it, on a paragraph's
 * end, or on the block after it; a bookmark is a child of its own. Code that
 * merges runs, replaces a paragraph, removes blocks or regenerates content
 * builds new nodes, and a node built from `properties` and `content` alone
 * drops those fields — leaving a `w:permStart` with no `w:permEnd`, or the
 * reverse. Each case below asserts the markers are still there, in order, and
 * that the written `document.xml` still pairs every id.
 */

import { unzip } from "@archive/read-archive";
import { describe, expect, it } from "vitest";

import { utf8Decoder } from "../core/internal-utils";
import { Document, Io, Package, Query, Template, Transform } from "../index";
import type {
  BodyContent,
  DocxDocument,
  ImageDef,
  Paragraph,
  RangeMarker,
  Run,
  StructuredDocumentTag,
  Table
} from "../types";

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89
]);

const start = (id: number): RangeMarker => ({ type: "permStart", id, edGrp: "everyone" });
const end = (id: number): RangeMarker => ({ type: "permEnd", id });

function run(text: string, rangeMarkersBefore?: RangeMarker[]): Run {
  return {
    ...(rangeMarkersBefore ? { rangeMarkersBefore } : {}),
    content: [{ type: "text", text }]
  };
}

function para(children: Paragraph["children"], extra: Partial<Paragraph> = {}): Paragraph {
  return { type: "paragraph", children, ...extra };
}

function docOf(body: BodyContent[], extra: Partial<DocxDocument> = {}): DocxDocument {
  const doc = Document.create();
  for (const block of body) {
    Document.addContent(doc, block);
  }
  return Object.assign(doc, extra);
}

/** The plain text of a paragraph's runs. */
function textOf(p: Paragraph): string {
  return p.children
    .flatMap(c => ("content" in c && !("type" in c) ? c.content : []))
    .map(c => (c.type === "text" ? c.text : ""))
    .join("");
}

/** Every range marker in document order, as the walker reports them. */
function markers(doc: DocxDocument): string[] {
  const out: string[] = [];
  Query.each(doc, {
    visitRangeMarker: m => void out.push(`${m.type}:${(m as { id: number }).id}`)
  });
  return out;
}

/** The permission markers of the written `document.xml`, as `start:id` / `end:id`. */
async function writtenPerms(doc: DocxDocument): Promise<string[]> {
  for await (const entry of unzip(await Io.toBuffer(doc)).entries()) {
    if (entry.path.replace(/^\/+/, "") === "word/document.xml") {
      const xml = utf8Decoder.decode(await entry.bytes());
      return [...xml.matchAll(/<w:perm(Start|End) w:id="(\d+)"/g)].map(
        m => `${m[1] === "Start" ? "start" : "end"}:${m[2]}`
      );
    }
  }
  throw new Error("no document.xml");
}

/** Each id opens once and closes once, in that order. */
function expectPaired(perms: readonly string[], ids: readonly number[]): void {
  for (const id of ids) {
    const s = perms.indexOf(`start:${id}`);
    const e = perms.indexOf(`end:${id}`);
    expect(s, `permStart ${id}`).toBeGreaterThanOrEqual(0);
    expect(e, `permEnd ${id}`).toBeGreaterThan(s);
    expect(perms.filter(p => p.endsWith(`:${id}`))).toHaveLength(2);
  }
}

describe("range markers survive rebuilt content", () => {
  it("Template.fill: a placeholder split across runs keeps the markers riding on them", async () => {
    // {{name}} split over three runs, a permission opening on the second and
    // closing on the third: the runs merge into one.
    const doc = docOf([
      para([run("Hi {{na"), run("me", [start(1)]), run("}}!", [end(1)])], {
        rangeMarkersBefore: [start(2)],
        rangeMarkersAfter: [end(2)]
      })
    ]);
    const filled = Template.fill(doc, { name: "Ada" });
    expect(Query.listParagraphs(filled)[0]).toBeDefined();
    expect(markers(filled)).toEqual(["permStart:2", "permStart:1", "permEnd:1", "permEnd:2"]);
    expectPaired(await writtenPerms(filled), [1, 2]);
  });

  it("Template.fill: a block #if removed keeps the markers of what it removed", async () => {
    const doc = docOf([
      para([run("{{#if show}}", [start(1)])]),
      para([run("hidden", [start(2)])], { rangeMarkersAfter: [end(2)] }),
      para([run("{{/if}}")], { rangeMarkersAfter: [end(1)] }),
      para([run("after")])
    ]);
    const filled = Template.fill(doc, { show: false });
    expect(markers(filled)).toEqual(["permStart:1", "permStart:2", "permEnd:2", "permEnd:1"]);
    expect(Query.listParagraphs(filled).map(p => textOf(p))).toEqual(["after"]);
    expectPaired(await writtenPerms(filled), [1, 2]);
  });

  it("Template.fill: an #each with no items keeps its body's markers", async () => {
    const doc = docOf([
      para([run("{{#each items}}")], { rangeMarkersBefore: [start(1)] }),
      para([run("{{this}}")]),
      para([run("{{/each}}")]),
      para([run("tail", [end(1)])])
    ]);
    const filled = Template.fill(doc, { items: [] });
    expect(markers(filled)).toEqual(["permStart:1", "permEnd:1"]);
    expectPaired(await writtenPerms(filled), [1]);
  });

  it("Template.fill: an image placeholder keeps the paragraph's markers", async () => {
    const doc = docOf([
      para([run("{{%logo}}", [start(1)])], {
        rangeMarkersBefore: [start(2)],
        rangeMarkersAfter: [end(1)]
      }),
      para([run("tail")], { rangeMarkersBefore: [end(2)] })
    ]);
    const image: ImageDef = { data: PNG, mediaType: "png", fileName: "logo.png" };
    const filled = Template.fill(
      doc,
      { logo: { image, width: 914400, height: 457200 } },
      { richPlaceholders: true }
    );
    const first = filled.body[0] as Paragraph;
    expect(first.rangeMarkersBefore).toEqual([start(2)]);
    expect(first.children[0]!.rangeMarkersBefore).toEqual([start(1)]);
    expect(first.rangeMarkersAfter).toEqual([end(1)]);
    expectPaired(await writtenPerms(filled), [1, 2]);
  });

  it("Template.fill: a rich-text placeholder keeps the paragraph's markers", async () => {
    const doc = docOf([para([run("{{&rt}}", [start(1)])], { rangeMarkersAfter: [end(1)] })]);
    const filled = Template.fill(doc, { rt: [run("rich")] }, { richPlaceholders: true });
    expect(markers(filled)).toEqual(["permStart:1", "permEnd:1"]);
    expect(textOf(filled.body[0] as Paragraph)).toBe("rich");
    expectPaired(await writtenPerms(filled), [1]);
  });

  it("Template.fill: a sub-document placeholder keeps the markers around what replaced it", async () => {
    const doc = docOf([para([run("{{>sub}}", [start(1)])], { rangeMarkersAfter: [end(1)] })]);
    const sub: BodyContent[] = [para([run("one")]), para([run("two")])];
    const filled = Template.fill(doc, { sub }, { richPlaceholders: true });
    expect(markers(filled)).toEqual(["permStart:1", "permEnd:1"]);
    expect(Query.listParagraphs(filled).map(p => textOf(p))).toEqual(["one", "two"]);
    expectPaired(await writtenPerms(filled), [1]);
  });

  it("Template.fill: a table row loop with no items keeps the rows' markers", async () => {
    const cell = (text: string, rangeMarkersBefore?: RangeMarker[]) => ({
      ...(rangeMarkersBefore ? { rangeMarkersBefore } : {}),
      content: [para([run(text)])]
    });
    const table: Table = {
      type: "table",
      rows: [
        { cells: [cell("Head")] },
        { cells: [cell("{{#each rows}}{{this}}{{/each}}", [start(1)])] },
        { cells: [cell("Foot")], rangeMarkersBefore: [end(1)] }
      ]
    };
    const filled = Template.fill(docOf([table]), { rows: [] });
    expect(markers(filled)).toEqual(["permStart:1", "permEnd:1"]);
    expect((filled.body[0] as Table).rows).toHaveLength(2);
    expectPaired(await writtenPerms(filled), [1]);
  });

  it("Template.fill: a #cols column loop with no items keeps the columns' markers", async () => {
    const cell = (text: string, rangeMarkersBefore?: RangeMarker[]) => ({
      ...(rangeMarkersBefore ? { rangeMarkersBefore } : {}),
      content: [para([run(text)])]
    });
    const table: Table = {
      type: "table",
      rows: [
        {
          cells: [cell("Name"), cell("{{#cols months}}{{this}}{{/cols}}", [start(1)]), cell("Sum")],
          rangeMarkersAfter: [end(1)]
        }
      ]
    };
    const filled = Template.fill(docOf([table]), { months: [] }, { richPlaceholders: true });
    expect((filled.body[0] as Table).rows[0]!.cells).toHaveLength(2);
    expect(markers(filled)).toEqual(["permStart:1", "permEnd:1"]);
    expectPaired(await writtenPerms(filled), [1]);
  });

  it("Transform.updateTableOfContents keeps the markers of the entries it regenerates", () => {
    const doc = docOf([
      {
        type: "tableOfContents",
        headingStyleRange: "1-3",
        cachedParagraphs: [para([run("stale", [start(1)])], { rangeMarkersAfter: [end(1)] })]
      },
      para([run("Chapter")], { properties: { style: "Heading1" } })
    ]);
    const updated = Transform.updateTableOfContents(doc);
    // The heading gains the `_Toc` bookmark its entry's page number refers to.
    expect(markers(updated)).toEqual([
      "permStart:1",
      "permEnd:1",
      "bookmarkStart:1",
      "bookmarkEnd:1"
    ]);
  });

  it("Package.patchTemplate: paragraph and image patches keep the replaced paragraph's markers", async () => {
    const doc = docOf([
      para([run("{{block}}", [start(1)])], { rangeMarkersAfter: [end(1)] }),
      para([run("{{pic}}", [start(2)])], { rangeMarkersAfter: [end(2)] })
    ]);
    const bytes = await Io.toBuffer(doc);
    const image: ImageDef = { data: PNG, mediaType: "png", fileName: "pic.png" };
    const out = await Package.patch(bytes, [
      { placeholder: "{{block}}", content: { type: "paragraph", children: [para([run("x")])] } },
      { placeholder: "{{pic}}", content: { type: "image", image, width: 10, height: 10 } }
    ]);
    const patched = await Io.read(out);
    expect(markers(patched)).toEqual(["permStart:1", "permEnd:1", "permStart:2", "permEnd:2"]);
    expectPaired(await writtenPerms(patched), [1, 2]);
  });

  it("Query.replaceText across runs leaves every marker on its run", async () => {
    const doc = docOf([
      para([run("foo b", [start(1)]), run("ar baz", [end(1)])], {
        rangeMarkersAfter: [end(2)],
        rangeMarkersBefore: [start(2)]
      })
    ]);
    expect(Query.replaceText(doc, "bar", "qux")).toBe(1);
    expect(textOf(doc.body[0] as Paragraph)).toBe("foo qux baz");
    expect(markers(doc)).toEqual(["permStart:2", "permStart:1", "permEnd:1", "permEnd:2"]);
    expectPaired(await writtenPerms(doc), [1, 2]);
  });

  it("Query.mailMerge leaves every marker on its run", async () => {
    const doc = docOf([
      para([
        {
          rangeMarkersBefore: [start(1)],
          content: [{ type: "field", instruction: " MERGEFIELD name " } as never]
        },
        run(".", [end(1)])
      ])
    ]);
    Query.mailMerge(doc, { name: "Ada" });
    expect(markers(doc)).toEqual(["permStart:1", "permEnd:1"]);
    expectPaired(await writtenPerms(doc), [1]);
  });

  it("Query.resolveDataBindings keeps the markers of the content it replaces or removes", async () => {
    const binding = { storeItemId: "{AAA}", xpath: "/root/name" };
    const bound: StructuredDocumentTag = {
      type: "sdt",
      properties: { dataBinding: binding },
      content: [para([run("old", [start(1)])], { rangeMarkersAfter: [end(1)] })]
    };
    const removed: StructuredDocumentTag = {
      type: "sdt",
      rangeMarkersBefore: [start(2)],
      properties: { dataBinding: { storeItemId: "{AAA}", xpath: "/root/missing" } },
      content: [para([run("gone")], { rangeMarkersAfter: [end(2)] })]
    };
    const doc = docOf([bound, removed]);
    const resolved = Query.resolveDataBindings(
      doc,
      new Map([["{AAA}", "<root><name>Ada</name></root>"]])
    );
    expect(markers(resolved)).toEqual(["permStart:1", "permEnd:1", "permStart:2", "permEnd:2"]);
    expect(textOf((resolved.body[0] as StructuredDocumentTag).content[0] as Paragraph)).toBe("Ada");
    expectPaired(await writtenPerms(resolved), [1, 2]);
  });

  it("Transform.map: a removed run or block leaves its markers behind", async () => {
    const doc = docOf([
      para([run("keep"), run("drop", [start(1)]), run("keep")]),
      para([run("gone")], { rangeMarkersBefore: [start(2)] }),
      para([run("end")], { rangeMarkersAfter: [end(1), end(2)] })
    ]);
    const mapped = Transform.map(doc, {
      transformRun: r => (textOf(para([r])) === "drop" ? null : r),
      transformParagraph: p => (textOf(p) === "gone" ? null : p)
    });
    expect(markers(mapped)).toEqual(["permStart:1", "permStart:2", "permEnd:1", "permEnd:2"]);
    expectPaired(await writtenPerms(mapped), [1, 2]);
  });

  it("Transform.split by page break keeps the markers of a dropped break-only paragraph", async () => {
    const doc = docOf([
      para([run("one", [start(1)])]),
      para([{ rangeMarkersBefore: [end(1)], content: [{ type: "break", breakType: "page" }] }]),
      para([run("two")])
    ]);
    const [first] = Transform.split(doc, { by: "pageBreak" });
    expect(markers(first!)).toEqual(["permStart:1", "permEnd:1"]);
    expectPaired(await writtenPerms(first!), [1]);
  });

  it("Document.removeContent moves the markers before the removed block to the next one", async () => {
    const doc = docOf([
      para([run("a")], { rangeMarkersBefore: [start(1)] }),
      para([run("b")], { rangeMarkersBefore: [start(2)] }),
      para([run("c")], { rangeMarkersAfter: [end(1)] })
    ]);
    const b = doc.body[1];
    const removed = Document.removeContent(doc, 1);
    expect(removed).toBe(b);
    expect(removed?.rangeMarkersBefore).toBeUndefined();
    expect(doc.body[1]!.rangeMarkersBefore).toEqual([start(2)]);
    const c = doc.body[1];
    expect(Document.removeContent(doc, 1)).toBe(c);
    expect(doc.bodyRangeMarkersAfter).toEqual([start(2)]);
    expect(markers(doc)).toEqual(["permStart:1", "permStart:2"]);
  });
});
