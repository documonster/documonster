/**
 * Bookmark ids are reconciled across every paragraph-rendering path of a
 * package: body, tables, SDTs, text boxes, the TOC's cached entries, headers,
 * notes, comments, the streaming writer, preserved raw XML and
 * `Package.edit`.
 */
import { unzip } from "@archive/read-archive";
import { describe, it, expect } from "vitest";

import { Build, Document, Io, Package, Streaming, Template, Transform } from "../../index";
import type { Paragraph, ParagraphChild } from "../../index";

async function partsOf(bytes: Uint8Array): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for await (const entry of unzip(bytes).entries()) {
    if (entry.path.endsWith(".xml") || entry.path.endsWith(".rels")) {
      out.set(entry.path, new TextDecoder().decode(await entry.bytes()));
    }
  }
  return out;
}

/** Every `w:bookmarkStart` id in the package, part by part, in document order. */
function startIds(parts: Map<string, string>): string[] {
  const ids: string[] = [];
  for (const [path, xml] of parts) {
    if (path.startsWith("word/") && !path.includes("glossary")) {
      ids.push(...[...xml.matchAll(/<w:bookmarkStart [^>]*\bw:id="(-?\d+)"/g)].map(m => m[1]!));
    }
  }
  return ids;
}

/** `[id, name]` pairs and the ids of ends, in order, from one part. */
function marks(xml: string): string[] {
  return [
    ...xml.matchAll(/<w:bookmark(Start|End) [^>]*?w:id="(-?\d+)"(?: w:name="([^"]*)")?/g)
  ].map(m => (m[1] === "Start" ? `S${m[2]}:${m[3]}` : `E${m[2]}`));
}

const mark = (name: string, id = 0): Paragraph =>
  Build.paragraph([Build.bookmarkStart(id, name), Build.text(name), Build.bookmarkEnd(id)]);

describe("bookmark ids in every part", () => {
  it("are unique across body, text box, TOC, table, footnote, comment and header", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, mark("body"));
    Document.addTextBox(doc, [mark("box")]);
    Document.addTableOfContents(doc, { cachedParagraphs: [mark("toc")] });
    Document.addTable(doc, {
      type: "table",
      rows: [{ cells: [{ content: [mark("cell")] }] }]
    });
    Document.addFootnote(doc, [mark("note")]);
    Document.addComment(doc, "A", [mark("comment")]);
    Document.setHeader(doc, "default", { children: [mark("head")] });

    const ids = startIds(await partsOf(await Io.toBuffer(doc)));
    expect(ids).toHaveLength(7);
    expect(new Set(ids).size).toBe(7);
  });

  it("pair each end with its own start inside a text box", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, mark("body"));
    Document.addTextBox(doc, [mark("box")]);
    const xml = (await partsOf(await Io.toBuffer(doc))).get("word/document.xml")!;
    expect(marks(xml)).toEqual(["S0:body", "E0", "S1:box", "E1"]);
  });

  it("give a text box hyperlink the relationship the packager registered", async () => {
    const doc = Document.create();
    Document.addTextBox(doc, [
      Build.paragraph([Build.hyperlink("link", { url: "https://example.com/box" })])
    ]);
    const parts = await partsOf(await Io.toBuffer(doc));
    const xml = parts.get("word/document.xml")!;
    const rels = parts.get("word/_rels/document.xml.rels")!;
    const linkRId = /<w:hyperlink [^>]*r:id="([^"]+)"/.exec(xml)?.[1];
    expect(linkRId).toBeDefined();
    expect(rels).toMatch(new RegExp(`Id="${linkRId}"[^>]*example\\.com/box`));
  });

  it("are unique in the streaming writer", async () => {
    const stream = Streaming.createWriter({
      headers: new Map([["default", { content: { children: [mark("head")] } }]]),
      footnotes: [{ id: 1, content: [mark("note")] }]
    });
    stream.add(mark("body"));
    stream.add({ type: "textBox", content: [mark("box")] });
    stream.add({ type: "tableOfContents", cachedParagraphs: [mark("toc")] });
    const ids = startIds(await partsOf(await stream.finalize()));
    expect(ids).toHaveLength(5);
    expect(new Set(ids).size).toBe(5);
  });

  it("survive Transform.merge of two documents that both use id 0", async () => {
    const a = Document.create();
    Document.addParagraph(a, mark("a"));
    const b = Document.create();
    Document.addParagraph(b, mark("b"));
    const merged = Transform.merge([a, b]);
    const xml = (await partsOf(await Io.toBuffer(merged))).get("word/document.xml")!;
    const seq = marks(xml);
    expect(seq.filter(s => s.startsWith("S"))).toHaveLength(2);
    expect(new Set(startIds(new Map([["word/document.xml", xml]]))).size).toBe(2);
    // Each end directly follows and closes its own start.
    expect(seq[1]).toBe(`E${seq[0]!.slice(1).split(":")[0]}`);
    expect(seq[3]).toBe(`E${seq[2]!.slice(1).split(":")[0]}`);
  });
});

describe("bookmarks in preserved raw XML", () => {
  const raw = (rawXml: string): ParagraphChild =>
    ({ type: "opaqueParagraphChild", rawXml }) as ParagraphChild;

  /** Each start id once, and every end closing a start written before it. */
  function expectPaired(seq: readonly string[]): void {
    const open = new Set<string>();
    const seen = new Set<string>();
    for (const s of seq) {
      const id = s.slice(1).split(":")[0]!;
      if (s.startsWith("S")) {
        expect(seen.has(id)).toBe(false);
        seen.add(id);
        open.add(id);
      } else {
        expect(open.has(id)).toBe(true);
        open.delete(id);
      }
    }
    expect([...open]).toEqual([]);
  }

  it("are renumbered like model bookmarks when their id clashes", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, mark("model", 4));
    Document.addParagraph(
      doc,
      Build.paragraph([raw('<w:bookmarkStart w:id="4" w:name="raw"/><w:bookmarkEnd w:id="4"/>')])
    );
    const seq = marks((await partsOf(await Io.toBuffer(doc))).get("word/document.xml")!);
    expect(seq).toEqual(["S4:model", "E4", "S5:raw", "E5"]);
  });

  it("close a model bookmark from inside raw XML", async () => {
    // `second` opens in the model and closes inside a preserved `w:bdo`; the
    // earlier bookmark sharing its id must not take the raw end.
    const doc = Document.create();
    Document.addParagraph(doc, mark("first", 1));
    Document.addParagraph(
      doc,
      Build.paragraph([
        Build.bookmarkStart(1, "second"),
        Build.text("b"),
        raw('<w:bdo w:val="rtl"><w:r><w:t>c</w:t></w:r><w:bookmarkEnd w:id="1"/></w:bdo>')
      ])
    );
    const seq = marks((await partsOf(await Io.toBuffer(doc))).get("word/document.xml")!);
    expect(seq).toEqual(["S1:first", "E1", "S2:second", "E2"]);
  });

  it("are closed by a model end, and the model's later bookmark moves off their id", async () => {
    const doc = Document.create();
    Document.addParagraph(
      doc,
      Build.paragraph([
        raw('<w:bdo w:val="rtl"><w:bookmarkStart w:id="2" w:name="inRaw"/></w:bdo>'),
        Build.text("d"),
        Build.bookmarkEnd(2)
      ])
    );
    Document.addParagraph(doc, mark("dupe", 2));
    const seq = marks((await partsOf(await Io.toBuffer(doc))).get("word/document.xml")!);
    expect(seq).toEqual(["S2:inRaw", "E2", "S3:dupe", "E3"]);
  });

  it("drop a raw end that closes nothing, in either form", async () => {
    const doc = Document.create();
    Document.addParagraph(
      doc,
      Build.paragraph([
        raw('<w:bdo w:val="rtl"><w:bookmarkEnd w:id="7"/><w:r><w:t>x</w:t></w:r></w:bdo>'),
        raw('<w:bdo w:val="rtl"><w:bookmarkEnd w:id="8"></w:bookmarkEnd></w:bdo>')
      ])
    );
    const xml = (await partsOf(await Io.toBuffer(doc))).get("word/document.xml")!;
    expect(xml).not.toContain("bookmarkEnd");
    expect(xml).toContain('<w:bdo w:val="rtl"><w:r><w:t>x</w:t></w:r></w:bdo>');
    expect(xml).toContain('<w:bdo w:val="rtl"></w:bdo>');
  });

  it("repeat the first alternative's ids in the other alternatives of mc:AlternateContent", async () => {
    // A consumer keeps one alternative, so the Fallback's bookmark is the
    // Choice's bookmark, not a second one.
    const box = (tag: string) =>
      `<${tag}><w:txbxContent><w:p><w:bookmarkStart w:id="0" w:name="box"/><w:bookmarkEnd w:id="0"/></w:p></w:txbxContent></${tag}>`;
    const doc = Document.create();
    Document.addParagraph(doc, mark("body", 0));
    Document.addParagraph(
      doc,
      Build.paragraph([
        raw(
          `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${box("wps:txbx")}</mc:Choice>` +
            `<mc:Fallback>${box("v:textbox")}</mc:Fallback></mc:AlternateContent></w:r>`
        )
      ])
    );
    const seq = marks((await partsOf(await Io.toBuffer(doc))).get("word/document.xml")!);
    expect(seq).toEqual(["S0:body", "E0", "S1:box", "E1", "S1:box", "E1"]);
  });

  it("pair across documents merged with Transform.merge", async () => {
    const a = Document.create();
    Document.addParagraph(a, mark("docA", 0));
    const b = Document.create();
    Document.addParagraph(
      b,
      Build.paragraph([
        Build.bookmarkStart(0, "docB"),
        Build.text("B "),
        raw('<w:bdo w:val="rtl"><w:r><w:t>rtl</w:t></w:r><w:bookmarkEnd w:id="0"/></w:bdo>')
      ])
    );
    const seq = marks(
      (await partsOf(await Io.toBuffer(Transform.merge([a, b])))).get("word/document.xml")!
    );
    expect(seq).toEqual(["S0:docA", "E0", "S1:docB", "E1"]);
  });

  it("keep one start and one end when a template loop repeats a raw end", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "{{#each items}}");
    Document.addParagraph(
      doc,
      Build.paragraph([
        Build.bookmarkStart(1, "lb"),
        Build.text("{{name}} "),
        raw('<w:bdo w:val="rtl"><w:r><w:t>rtl</w:t></w:r><w:bookmarkEnd w:id="1"/></w:bdo>')
      ])
    );
    Document.addParagraph(doc, "{{/each}}");
    const filled = Template.fill(doc, { items: [{ name: "a" }, { name: "b" }, { name: "c" }] });
    const seq = marks((await partsOf(await Io.toBuffer(filled))).get("word/document.xml")!);
    expectPaired(seq);
    expect(seq.filter(s => s.startsWith("S"))).toHaveLength(1);
  });

  it("are renumbered in the streaming writer against what was already streamed", async () => {
    const stream = Streaming.createWriter();
    stream.add(mark("model", 3));
    stream.add(
      Build.paragraph([raw('<w:bookmarkStart w:id="3" w:name="raw"/><w:bookmarkEnd w:id="3"/>')])
    );
    const seq = marks((await partsOf(await stream.finalize())).get("word/document.xml")!);
    expect(seq).toEqual(["S3:model", "E3", "S4:raw", "E4"]);
  });

  it("pair range permissions the same way", async () => {
    const doc = Document.create();
    Document.addParagraph(
      doc,
      Build.paragraph([
        raw('<w:permStart w:id="1" w:edGrp="everyone"/>'),
        Build.text("p1"),
        raw('<w:permEnd w:id="1"/>')
      ])
    );
    Document.addParagraph(
      doc,
      Build.paragraph([
        raw('<w:permStart w:id="1" w:ed="bob"/>'),
        Build.text("p2"),
        raw('<w:permEnd w:id="1"/>')
      ])
    );
    const xml = (await partsOf(await Io.toBuffer(doc))).get("word/document.xml")!;
    const perms = [...xml.matchAll(/<w:perm(Start|End) w:id="(\d+)"/g)].map(m => m[1]! + m[2]!);
    expect(perms).toEqual(["Start1", "End1", "Start2", "End2"]);
  });
});

describe("overlapping and orphaned bookmarks", () => {
  it("pair overlapping ranges that share an id by the end's name", async () => {
    const doc = Document.create();
    Document.addParagraph(
      doc,
      Build.paragraph([
        Build.bookmarkStart(0, "first"),
        Build.text("a"),
        Build.bookmarkStart(0, "second"),
        Build.text("b"),
        Build.bookmarkEnd(0, { name: "first" }),
        Build.text("c"),
        Build.bookmarkEnd(0, { name: "second" })
      ])
    );
    const xml = (await partsOf(await Io.toBuffer(doc))).get("word/document.xml")!;
    expect(marks(xml)).toEqual(["S0:first", "S1:second", "E0", "E1"]);
  });

  it("drop an end that closes nothing rather than letting it close another bookmark", async () => {
    const doc = Document.create();
    Document.addParagraph(
      doc,
      Build.paragraph([
        Build.bookmarkStart(0, "a"),
        Build.bookmarkEnd(0),
        Build.bookmarkStart(0, "b"),
        // Orphan whose id equals the id `b` is renumbered to.
        Build.bookmarkEnd(1),
        Build.text("x"),
        Build.bookmarkEnd(0)
      ])
    );
    const xml = (await partsOf(await Io.toBuffer(doc))).get("word/document.xml")!;
    expect(marks(xml)).toEqual(["S0:a", "E0", "S1:b", "E1"]);
  });
});

describe("Package.edit", () => {
  it("gives new body bookmarks ids the untouched parts do not use", async () => {
    const doc = Document.create();
    Document.setHeader(doc, "default", {
      children: [mark("head", 0), mark("head2", 1)]
    });
    Document.addParagraph(doc, "old");
    const edited = await Package.edit(await Io.toBuffer(doc), [
      { type: "replaceBody", body: [mark("new", 0), mark("new2", 1)] }
    ]);
    const ids = startIds(await partsOf(edited));
    expect(ids).toHaveLength(4);
    expect(new Set(ids).size).toBe(4);
  });

  it("gives new header bookmarks ids the body does not use", async () => {
    const doc = Document.create();
    Document.setHeader(doc, "default", { children: [Build.textParagraph("h")] });
    Document.addParagraph(doc, mark("body", 0));
    const edited = await Package.edit(await Io.toBuffer(doc), [
      { type: "replaceHeader", path: "word/header1.xml", children: [mark("h", 0), mark("h2", 0)] }
    ]);
    const ids = startIds(await partsOf(edited));
    expect(ids).toHaveLength(3);
    expect(new Set(ids).size).toBe(3);
  });
});
