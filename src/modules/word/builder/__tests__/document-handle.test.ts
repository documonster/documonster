/**
 * `DocxDocument` is the document tree itself: `Document.create()` and
 * `Io.read()` return the same kind of object, and `Document.*` helpers edit
 * either one without id collisions.
 */
import { unzip } from "@archive/read-archive";
import { describe, it, expect } from "vitest";

import { WordHtml } from "../../html";
import {
  Build,
  Document,
  DocxDuplicateIdError,
  DocxArgumentError,
  DocxWriteError,
  Io,
  Layout,
  Stream
} from "../../index";
import type { DocxDocument } from "../../index";
import { finalizeDocument } from "../document-finalize";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

async function parts(bytes: Uint8Array): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for await (const entry of unzip(bytes).entries()) {
    out.set(entry.path, new TextDecoder().decode(await entry.bytes()));
  }
  return out;
}

function populate(doc: DocxDocument): void {
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "第一章 概述", 1);
  Document.addParagraph(doc, "Body text 正文");
  Document.addTable(doc, [
    ["a", "b"],
    ["c", "d"]
  ]);
  Document.addBulletList(doc, ["x", "y"]);
  Document.addNumberedList(doc, ["1", "2"]);
  Document.addImage(doc, PNG, "png", 9525, 9525);
  Document.addFootnote(doc, "fn");
  Document.addEndnote(doc, "en");
  Document.addComment(doc, "me", "note", { date: "2020-01-01T00:00:00Z" });
  Document.setHeader(doc, "default", { children: [Build.textParagraph("Header")] });
  Document.setCoreProperties(doc, { title: "t", created: new Date(0), modified: new Date(0) });
}

describe("DocxDocument", () => {
  it("create returns a plain document tree with default page geometry", () => {
    const doc = Document.create();
    expect(doc.body).toEqual([]);
    expect(doc.sectionProperties?.pageSize).toEqual({ width: 12240, height: 15840 });
    expect(doc.styles).toBeUndefined();
  });

  it("writes the same package with or without the deprecated build step", async () => {
    const a = Document.create();
    populate(a);
    const b = Document.create();
    populate(b);
    const viaBuild = await parts(await Io.toBuffer(Document.build(a)));
    const direct = await parts(await Io.toBuffer(b));
    expect([...direct.keys()].sort()).toEqual([...viaBuild.keys()].sort());
    for (const [path, xml] of viaBuild) {
      expect(direct.get(path), path).toBe(xml);
    }
  });

  it("does not modify the document when it is written", async () => {
    const doc = Document.create();
    Document.useDefaultStyles(doc);
    Document.addParagraph(doc, "简体中文内容");
    await Io.toBuffer(doc);
    expect(doc.docDefaults?.runProperties?.language?.eastAsia).toBeUndefined();
  });

  it("derives the East Asian language for a created document when written", async () => {
    const doc = Document.create();
    Document.useDefaultStyles(doc);
    Document.addParagraph(doc, "简体中文内容");
    const reread = await Io.read(await Io.toBuffer(doc));
    expect(reread.docDefaults?.runProperties?.language?.eastAsia).toBe("zh-CN");
  });

  it("does not derive an East Asian language for a read document", async () => {
    const source = Document.create();
    Document.setDocDefaults(source, { runProperties: { language: { val: "en-US" } } });
    Document.addParagraph(source, "English only");
    const read = await Io.read(await Io.toBuffer(source));
    expect(read.docDefaults?.runProperties?.language?.eastAsia).toBeUndefined();

    Document.addParagraph(read, "追加的中文段落");
    const again = await Io.read(await Io.toBuffer(read));
    expect(again.docDefaults?.runProperties?.language?.eastAsia).toBeUndefined();
  });

  describe("deriveEastAsianDefaults survives every way of copying a created document", () => {
    function chinese(): DocxDocument {
      const doc = Document.create();
      Document.useDefaultStyles(doc);
      Document.addParagraph(doc, "简体中文内容");
      return doc;
    }
    async function eastAsiaOf(doc: DocxDocument): Promise<string | undefined> {
      const reread = await Io.read(await Io.toBuffer(doc));
      return reread.docDefaults?.runProperties?.language?.eastAsia;
    }

    it("is a visible field on a created document and absent on a read one", async () => {
      const doc = chinese();
      expect(doc.deriveEastAsianDefaults).toBe(true);
      expect(Object.keys(doc)).toContain("deriveEastAsianDefaults");
      const read = await Io.read(await Io.toBuffer(doc));
      expect(read.deriveEastAsianDefaults).toBeUndefined();
    });

    it("spread", async () => {
      expect(await eastAsiaOf({ ...chinese(), customProperties: [] })).toBe("zh-CN");
    });

    it("structuredClone", async () => {
      expect(await eastAsiaOf(structuredClone(chinese()))).toBe("zh-CN");
    });

    it("JSON round trip", async () => {
      const json = JSON.parse(JSON.stringify(chinese())) as DocxDocument;
      expect(await eastAsiaOf(json)).toBe("zh-CN");
    });

    it("merge of a created and a read document, in either order", async () => {
      const english = Document.create();
      Document.addParagraph(english, "English only");
      const read = await Io.read(await Io.toBuffer(english));
      const created = chinese();
      expect(await eastAsiaOf({ ...read, ...created })).toBe("zh-CN");
      expect(await eastAsiaOf({ ...created, ...read, body: created.body })).toBe("zh-CN");
    });

    it("is not leaked onto a read document by Object.assign of something else", async () => {
      const english = Document.create();
      Document.addParagraph(english, "English only");
      const read = await Io.read(await Io.toBuffer(english));
      Object.assign(read, { body: [...read.body, Build.textParagraph("追加的中文段落")] });
      expect(await eastAsiaOf(read)).toBeUndefined();
    });

    it("can be cleared on a created document or set on a read one", async () => {
      expect(await eastAsiaOf({ ...chinese(), deriveEastAsianDefaults: false })).toBeUndefined();
      const read = await Io.read(await Io.toBuffer(Document.create()));
      Document.addParagraph(read, "简体中文内容");
      expect(await eastAsiaOf({ ...read, deriveEastAsianDefaults: true })).toBe("zh-CN");
    });

    it("is never written into the package", async () => {
      const files = await parts(await Io.toBuffer(chinese()));
      for (const [path, xml] of files) {
        expect(xml, path).not.toContain("deriveEastAsianDefaults");
      }
    });
  });

  it("lays out and renders a created document as the deprecated build step did", () => {
    const doc = Document.create();
    Document.useDefaultStyles(doc);
    Document.addParagraph(doc, "简体中文内容 mixed with Latin");
    const built = Document.build(doc);
    expect(built.docDefaults?.runProperties?.language?.eastAsia).toBe("zh-CN");
    expect(Layout.compute(doc)).toEqual(Layout.compute(built));
    expect(WordHtml.fromDocx(doc)).toEqual(WordHtml.fromDocx(built));
  });

  it("finalizing an already finalized document is a no-op", () => {
    for (const text of ["简体中文内容", "English only"]) {
      const doc = Document.create();
      Document.addParagraph(doc, text);
      const once = finalizeDocument(doc);
      expect(once).not.toBe(doc);
      expect(once.deriveEastAsianDefaults).toBeUndefined();
      expect(finalizeDocument(once)).toBe(once);
    }
  });

  it("Document.build keeps its old shape: a new object, no empty collections", () => {
    const doc = Document.create();
    Document.addParagraph(doc, "x");
    (doc as { styles?: unknown[] }).styles = [];
    const built = Document.build(doc);
    expect(built).not.toBe(doc);
    expect(built.styles).toBeUndefined();
    expect(built.body).toBe(doc.body);
    expect(built.deriveEastAsianDefaults).toBeUndefined();
  });

  it("reading through the helpers creates no empty collection", () => {
    const doc = Document.create();
    Document.getContentCount(doc);
    Document.getContent(doc, 0);
    Document.removeContent(doc, 0);
    Document.addParagraph(doc, "x");
    Document.addSectionBreak(doc, {});
    expect(Object.keys(doc).sort()).toEqual(
      ["body", "deriveEastAsianDefaults", "sectionProperties"].sort()
    );
    Document.addFootnote(doc, "fn");
    expect(doc.footnotes).toHaveLength(1);
    expect(doc.endnotes).toBeUndefined();
    expect(doc.comments).toBeUndefined();
    expect(doc.images).toBeUndefined();
  });

  describe("ids stay unique when direct edits are interleaved with helper calls", () => {
    it("footnotes, endnotes and comments", () => {
      const doc = Document.create();
      const a = Document.addFootnote(doc, "a");
      (doc.footnotes as { id: number; content: [] }[]).push({ id: a + 1, content: [] });
      const b = Document.addFootnote(doc, "b");
      expect(b).toBe(a + 2);

      const e = Document.addEndnote(doc, "e");
      (doc.endnotes as { id: number; content: [] }[]).push({ id: e + 5, content: [] });
      expect(Document.addEndnote(doc, "f")).toBe(e + 6);

      const c = Document.addComment(doc, "me", "c");
      (doc.comments as { id: number; author: string; content: [] }[]).push({
        id: c + 1,
        author: "x",
        content: []
      });
      expect(Document.addComment(doc, "me", "d")).toBe(c + 2);

      // Collections replaced wholesale are seen too.
      (doc as { footnotes?: unknown }).footnotes = [{ id: 40, content: [] }];
      expect(Document.addFootnote(doc, "g")).toBe(41);

      // So is an entry replaced by assigning a new array, at any position.
      const notes = doc.footnotes ?? [];
      (doc as { footnotes?: unknown }).footnotes = [{ id: 99, content: [] }, ...notes.slice(1)];
      expect(Document.addFootnote(doc, "h")).toBe(100);
      const comments = doc.comments ?? [];
      (doc as { comments?: unknown }).comments = [
        { id: 50, author: "x", content: [] },
        ...comments.slice(1)
      ];
      expect(Document.addComment(doc, "me", "i")).toBe(51);
    });

    it("an edited copy of the document is read afresh", () => {
      const doc = Document.create();
      for (let i = 0; i < 300; i++) {
        Document.addFootnote(doc, "n");
      }
      const notes = doc.footnotes ?? [];
      const copy: DocxDocument = {
        ...doc,
        footnotes: [{ ...notes[0]!, id: 301 }, ...notes.slice(1)]
      };
      expect(Document.addFootnote(copy, "new")).toBe(302);
      const ids = (copy.footnotes ?? []).map(f => f.id);
      expect(new Set(ids).size).toBe(ids.length);
    });

    it("an id handed out and then removed by the caller is not reused", () => {
      const doc = Document.create();
      const a = Document.addFootnote(doc, "a");
      (doc as { footnotes?: unknown }).footnotes = [];
      expect(Document.addFootnote(doc, "b")).toBe(a + 1);
    });

    it("numbering: abstract definitions and instances", () => {
      const doc = Document.create();
      Document.addBulletList(doc, ["x"]);
      (doc.numberingInstances as { numId: number; abstractNumId: number }[]).push({
        numId: 9,
        abstractNumId: 0
      });
      (doc.abstractNumberings as { abstractNumId: number; levels: [] }[]).push({
        abstractNumId: 7,
        levels: []
      });
      Document.addNumberedList(doc, ["y"]);
      const instances = doc.numberingInstances ?? [];
      expect(instances.at(-1)?.numId).toBe(10);
      expect(doc.abstractNumberings?.at(-1)?.abstractNumId).toBe(8);
      const ids = instances.map(n => n.numId);
      expect(new Set(ids).size).toBe(ids.length);
    });

    it("images, and drawing ids reconciled when written", async () => {
      const doc = Document.create();
      const first = Document.addImage(doc, PNG, "png", 9525, 9525);
      (doc.images as { data: Uint8Array; mediaType: "png"; fileName: string; rId: string }[]).push({
        data: PNG,
        mediaType: "png",
        fileName: "image5.png",
        rId: "__img_5"
      });
      Document.addParagraph(
        doc,
        Build.paragraph([
          {
            content: [
              { type: "image", rId: "__img_5", width: 1, height: 1, drawingId: first.drawingId + 1 }
            ]
          }
        ])
      );
      const second = Document.addImage(doc, PNG, "png", 9525, 9525);
      expect(second.rId).toBe("__img_6");
      expect(first.drawingId).toBeLessThan(second.drawingId);
      const names = (doc.images ?? []).map(i => i.fileName);
      expect(new Set(names).size).toBe(names.length);
      expect(Document.addFloatingImage(doc, PNG, "png", 9525, 9525)).toBe("__img_7");

      // The direct edit reused the id the helper handed out next; the package
      // still carries one `wp:docPr` id per drawing.
      const xml = (await parts(await Io.toBuffer(doc))).get("word/document.xml")!;
      const docPr = [...xml.matchAll(/<wp:docPr [^>]*\bid="(\d+)"/g)].map(m => m[1]);
      expect(docPr).toHaveLength(4);
      expect(new Set(docPr).size).toBe(docPr.length);
    });

    it("bookmark ids are seeded from the tree, and clashes are reconciled when written", async () => {
      const doc = Document.create();
      Document.setHeader(doc, "default", {
        children: [Build.paragraph([Build.bookmarkStart(20, "h"), Build.bookmarkEnd(20)])]
      });
      const first = Document.allocateBookmarkId(doc);
      expect(first).toBe(21);
      // Pushed directly after the seed, reusing the id the helper hands out next.
      const clash = first + 1;
      Document.addParagraph(
        doc,
        Build.paragraph([
          Build.bookmarkStart(clash, "direct"),
          Build.text("d"),
          Build.bookmarkEnd(clash)
        ])
      );
      const next = Document.allocateBookmarkId(doc);
      expect(next).toBe(clash);
      Document.addParagraph(
        doc,
        Build.paragraph([
          Build.bookmarkStart(next, "helper"),
          Build.text("h"),
          Build.bookmarkEnd(next)
        ])
      );

      const written = await Io.read(await Io.toBuffer(doc));
      const pairs = new Map<string, { start: number; end?: number }>();
      let open: string | undefined;
      for (const block of [
        ...written.body,
        ...[...(written.headers?.values() ?? [])].flatMap(h => h.content.children)
      ]) {
        if (block.type !== "paragraph") {
          continue;
        }
        for (const child of block.children) {
          if ("type" in child && child.type === "bookmarkStart") {
            pairs.set(child.name, { start: child.id });
            open = child.name;
          } else if ("type" in child && child.type === "bookmarkEnd" && open) {
            pairs.get(open)!.end = child.id;
          }
        }
      }
      expect([...pairs.keys()].sort()).toEqual(["direct", "h", "helper"]);
      const starts = [...pairs.values()].map(p => p.start);
      expect(new Set(starts).size).toBe(starts.length);
      for (const p of pairs.values()) {
        expect(p.end).toBe(p.start);
      }
    });

    it("a written document carries no duplicate ids", async () => {
      const doc = Document.create();
      populate(doc);
      (doc.footnotes as { id: number; content: unknown[] }[]).push({
        id: Document.addFootnote(doc, "x") + 1,
        content: [Build.textParagraph("direct")]
      });
      Document.addFootnote(doc, "after");
      const written = await Io.read(await Io.toBuffer(doc));
      const ids = (written.footnotes ?? []).map(f => f.id);
      expect(new Set(ids).size).toBe(ids.length);
      expect(ids.length).toBe((doc.footnotes ?? []).length);
    });

    describe("an in-place edit outside the contract fails at write instead of duplicating", () => {
      type Entry = Record<string, unknown>;
      const swapFirst = (list: readonly unknown[] | undefined, patch: Entry): void => {
        const entries = list as Entry[];
        entries[0] = { ...entries[0], ...patch };
      };
      async function expectDuplicate(
        doc: DocxDocument,
        collection: string,
        id: number | string
      ): Promise<void> {
        const error = await Io.toBuffer(doc).then(
          () => undefined,
          (e: unknown) => e
        );
        expect(error).toBeInstanceOf(DocxDuplicateIdError);
        expect(error).toBeInstanceOf(DocxWriteError);
        expect(error).toMatchObject({ collection, id });
      }

      it("footnotes, the reported case", async () => {
        const doc = Document.create();
        for (let i = 0; i < 300; i++) {
          Document.addFootnote(doc, "n");
        }
        swapFirst(doc.footnotes, { id: 301 });
        const id = Document.addFootnote(doc, "new");
        const ids = (doc.footnotes ?? []).map(f => f.id);
        // Either the helper saw the edit, or the write refuses the document.
        if (new Set(ids).size === ids.length) {
          expect(id).not.toBe(301);
        } else {
          await expectDuplicate(doc, "footnotes", 301);
        }
      });

      it("endnotes, comments and numbering", async () => {
        for (const [collection, fill, patch, id] of [
          ["endnotes", (d: DocxDocument) => Document.addEndnote(d, "e"), { id: 201 }, 201],
          ["comments", (d: DocxDocument) => Document.addComment(d, "me", "c"), { id: 200 }, 200],
          [
            "numberingInstances",
            (d: DocxDocument) => Document.addNumberedList(d, ["x"]),
            { numId: 201 },
            201
          ]
        ] as const) {
          const doc = Document.create();
          for (let i = 0; i < 200; i++) {
            fill(doc);
          }
          swapFirst(doc[collection], patch);
          fill(doc);
          await expectDuplicate(doc, collection, id);
        }

        const doc = Document.create();
        Document.addBulletList(doc, ["x"]);
        Document.addNumberedList(doc, ["y"]);
        swapFirst(doc.abstractNumberings, { abstractNumId: 2, levels: [] });
        Document.addBulletList(doc, ["z"]);
        await expectDuplicate(doc, "abstractNumberings", 2);
      });

      it("images", async () => {
        const doc = Document.create();
        for (let i = 0; i < 200; i++) {
          Document.addImage(doc, PNG, "png", 9525, 9525);
        }
        swapFirst(doc.images, {
          rId: "__img_201",
          fileName: "image201.png",
          data: new Uint8Array([...PNG, 7])
        });
        expect(Document.addImage(doc, PNG, "png", 9525, 9525).rId).toBe("__img_201");
        await expectDuplicate(doc, "images", "__img_201");
      });

      it("a note id the writer reserves for its separators", async () => {
        const doc = Document.create();
        (doc as { footnotes?: unknown }).footnotes = [{ id: 0, content: [] }];
        await expectDuplicate(doc, "footnotes", 0);
        (doc as { footnotes?: unknown }).footnotes = [
          { id: -1, type: "separator", content: [] },
          { id: 0, content: [] }
        ];
        await expect(Io.toBuffer(doc)).resolves.toBeInstanceOf(Uint8Array);
      });

      it("the streaming writer refuses the same models", () => {
        expect(
          () =>
            new Stream.DocumentWriter({
              comments: [
                { id: 3, author: "a", content: [] },
                { id: 3, author: "b", content: [] }
              ]
            })
        ).toThrow(DocxDuplicateIdError);
      });

      it("two different pictures under one image key", async () => {
        const doc = Document.create();
        const other = new Uint8Array([...PNG, 9]);
        (doc as { images?: unknown }).images = [
          { data: PNG, mediaType: "png", fileName: "a.png", rId: "rId1" },
          { data: other, mediaType: "png", fileName: "b.png", rId: "rId1" }
        ];
        await expectDuplicate(doc, "images", "rId1");
        (doc as { images?: unknown }).images = [
          { data: PNG, mediaType: "png", fileName: "a.png", rId: "rId1" },
          { data: other, mediaType: "png", fileName: "b.png", rId: "rId2", aliasRIds: ["rId1"] }
        ];
        await expectDuplicate(doc, "images", "rId1");
        // The same picture listed twice under one key is not ambiguous.
        (doc as { images?: unknown }).images = [
          { data: PNG, mediaType: "png", fileName: "a.png", rId: "rId1" },
          { data: PNG, mediaType: "png", fileName: "a.png", rId: "rId1" }
        ];
        await expect(Io.toBuffer(doc)).resolves.toBeInstanceOf(Uint8Array);
      });
    });
  });

  it("addListDefinition registers a list a paragraph can refer to", () => {
    const doc = Document.create();
    const bullet = Document.addListDefinition(doc, "bullet", { numId: 5 });
    const numbered = Document.addListDefinition(doc, "numbered");
    expect(bullet).toBe(5);
    expect(numbered).toBe(6);
    expect(doc.abstractNumberings?.map(a => a.levels[0]?.format)).toEqual(["bullet", "decimal"]);
    expect(() => Document.addListDefinition(doc, "bullet", { numId: 5 })).toThrow(
      DocxArgumentError
    );
    // Reuses the definition rather than defining it again.
    Document.addBulletList(doc, ["x"]);
    expect(doc.abstractNumberings).toHaveLength(2);
  });

  it("gives new content ids that do not collide with a read document's", async () => {
    const source = Document.create();
    populate(source);
    const bookmarkId = Document.allocateBookmarkId(source);
    Document.addParagraph(
      source,
      Build.paragraph([
        Build.bookmarkStart(bookmarkId, "existing"),
        Build.text("marked"),
        Build.bookmarkEnd(bookmarkId)
      ])
    );
    const read = await Io.read(await Io.toBuffer(source));

    const before = {
      footnotes: (read.footnotes ?? []).map(f => f.id),
      endnotes: (read.endnotes ?? []).map(e => e.id),
      comments: (read.comments ?? []).map(c => c.id),
      numIds: (read.numberingInstances ?? []).map(n => n.numId),
      abstractIds: (read.abstractNumberings ?? []).map(a => a.abstractNumId),
      imageNames: (read.images ?? []).map(i => i.fileName)
    };

    const footnote = Document.addFootnote(read, "new fn");
    const endnote = Document.addEndnote(read, "new en");
    const comment = Document.addComment(read, "me", "new comment");
    const bookmark = Document.allocateBookmarkId(read);
    Document.addImage(read, PNG, "png", 9525, 9525);
    Document.addNumberedList(read, ["new list"]);
    Document.addParagraph(read, "Appended");

    expect(before.footnotes).not.toContain(footnote);
    expect(before.endnotes).not.toContain(endnote);
    expect(before.comments).not.toContain(comment);
    expect(bookmark).toBeGreaterThan(bookmarkId);
    const newNumId = read.numberingInstances!.at(-1)!.numId;
    expect(before.numIds).not.toContain(newNumId);
    const newImage = read.images!.at(-1)!.fileName;
    expect(before.imageNames).not.toContain(newImage);

    const written = await Io.read(await Io.toBuffer(read));
    const ids = (written.footnotes ?? []).map(f => f.id);
    expect(new Set(ids).size).toBe(ids.length);
    const commentIds = (written.comments ?? []).map(c => c.id);
    expect(new Set(commentIds).size).toBe(commentIds.length);
    expect(written.images?.length).toBe((source.images?.length ?? 0) + 1);
  });
});
