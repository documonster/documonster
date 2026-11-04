/**
 * Bookmark ids are reconciled as a package is written: every start gets a
 * unique id and every end closes the start it closed in the model.
 */
import { describe, it, expect } from "vitest";

import { Build, Document, Io } from "../../index";
import {
  createBookmarkIdMap,
  createIdGenerators,
  createRevisionIdMap,
  prepareRawMarkup,
  renumberRawRangeIds,
  renumberRawRevisionIds
} from "../render-context";

describe("renumberRawRangeIds", () => {
  it("closes a renumbered model start from an end inside raw XML", () => {
    const bookmarkIds = createBookmarkIdMap();
    bookmarkIds.start(1, "first");
    bookmarkIds.end(1);
    expect(bookmarkIds.start(1, "second")).toBe(2);
    expect(renumberRawRangeIds(`<w:bdo><w:bookmarkEnd w:id="1"/></w:bdo>`, { bookmarkIds })).toBe(
      `<w:bdo><w:bookmarkEnd w:id="2"/></w:bdo>`
    );
  });

  it("moves a raw range off reserved and written ids, and leaves the model's pairing alone", () => {
    const bookmarkIds = createBookmarkIdMap([3]);
    expect(bookmarkIds.start(3, "model")).toBe(4);
    const raw = `<w:bookmarkStart w:id="3" w:name="raw"/><w:bookmarkEnd w:id="3" w:name="raw"/>`;
    expect(renumberRawRangeIds(raw, { bookmarkIds })).toBe(
      `<w:bookmarkStart w:id="5" w:name="raw"/><w:bookmarkEnd w:id="5" w:name="raw"/>`
    );
    expect(bookmarkIds.end(3, "model")).toBe(4);
  });
});

describe("prepareRawMarkup", () => {
  it("remaps image ids, reconciles range ids and allocates drawing ids in one pass", () => {
    const bookmarkIds = createBookmarkIdMap();
    bookmarkIds.start(0, "taken");
    const ids = createIdGenerators({ docPrId: 7 });
    const raw =
      `<w:bookmarkStart w:id="0" w:name="raw"/>` +
      `<wp:docPr id="1" name="Picture 1"/><a:blip r:embed="rId9"/>` +
      `<w:bookmarkEnd w:id="0"/>`;
    expect(
      prepareRawMarkup(raw, {
        bookmarkIds,
        imageRemap: new Map([["rId9", "rId2"]]),
        nextDocPrId: ids.nextDocPrId
      })
    ).toBe(
      `<w:bookmarkStart w:id="1" w:name="raw"/>` +
        `<wp:docPr id="7" name="Picture 1"/><a:blip r:embed="rId2"/>` +
        `<w:bookmarkEnd w:id="1"/>`
    );
  });
});

describe("createIdGenerators", () => {
  it("hands out drawing ids around the reserved ones", () => {
    const ids = createIdGenerators({ reservedDocPrIds: [1, 3, 4] });
    expect([ids.nextDocPrId(), ids.nextDocPrId(), ids.nextDocPrId()]).toEqual([2, 5, 6]);
  });
});

describe("createBookmarkIdMap", () => {
  it("keeps unique ids unchanged", () => {
    const map = createBookmarkIdMap();
    expect([map.start(0), map.end(0), map.start(5), map.end(5)]).toEqual([0, 0, 5, 5]);
  });

  it("renumbers a repeated id and pairs each end with its own start", () => {
    const map = createBookmarkIdMap();
    expect(map.start(0)).toBe(0);
    expect(map.end(0)).toBe(0);
    expect(map.start(0)).toBe(1);
    expect(map.end(0)).toBe(1);
    // A later bookmark that legitimately used 1 moves above everything seen.
    expect(map.start(1)).toBe(2);
    expect(map.end(1)).toBe(2);
  });

  it("closes the most recent open start for nested duplicates", () => {
    const map = createBookmarkIdMap();
    const outer = map.start(3);
    const inner = map.start(3);
    expect(map.end(3)).toBe(inner);
    expect(map.end(3)).toBe(outer);
    expect(inner).not.toBe(outer);
  });

  it("drops an end without a start", () => {
    expect(createBookmarkIdMap().end(9)).toBeUndefined();
  });

  it("keeps an end without a start when its id is reserved", () => {
    expect(createBookmarkIdMap([9]).end(9)).toBe(9);
  });

  it("moves a start off a reserved id", () => {
    const map = createBookmarkIdMap([0, 7, 20]);
    expect(map.start(0)).toBe(21);
    expect(map.end(0)).toBe(21);
    expect(map.start(3)).toBe(3);
    expect(map.start(20)).toBe(22);
  });

  it("pairs overlapping duplicates by name, and falls back to the latest start", () => {
    const map = createBookmarkIdMap();
    const a = map.start(0, "a");
    const b = map.start(0, "b");
    expect(map.end(0, "a")).toBe(a);
    expect(map.end(0, "missing")).toBe(b);
  });

  it("keeps each start's name when a single open start gains and loses company", () => {
    const map = createBookmarkIdMap();
    const a = map.start(4, "a");
    const b = map.start(4, "b");
    const c = map.start(4, "c");
    expect(map.end(4, "b")).toBe(b);
    expect(map.end(4, "c")).toBe(c);
    // Down to one open start again; a further duplicate must still find "a" by name.
    const d = map.start(4, "d");
    expect(map.end(4, "a")).toBe(a);
    expect(map.end(4)).toBe(d);
    expect(map.end(4)).toBeUndefined();
    expect(new Set([a, b, c, d]).size).toBe(4);
  });

  it("closes a start under an id far above the ones before it", () => {
    const map = createBookmarkIdMap();
    expect(map.start(100_000, "far")).toBe(100_000);
    expect(map.end(100_000)).toBe(100_000);
    expect(map.end(100_000)).toBeUndefined();
  });
});

describe("createRevisionIdMap", () => {
  it("keeps the first use of an id and moves every repeat above all ids seen", () => {
    const map = createRevisionIdMap([10]);
    expect([map.claim(1), map.claim(1), map.claim(10), map.claim(12), map.claim(-3)]).toEqual([
      1, 11, 12, 13, -3
    ]);
  });
});

describe("renumberRawRevisionIds", () => {
  it("renumbers every revision element in raw XML, and nothing else", () => {
    const revisionIds = createRevisionIdMap();
    revisionIds.claim(1);
    const raw =
      `<w:ins w:id="1" w:author="A"><w:r><w:instrText>x</w:instrText></w:r></w:ins>` +
      `<w:rPr><w:del w:id="1" w:author="A"/></w:rPr>` +
      `<w:pPrChange w:id="1" w:author="A"><w:pPr/></w:pPrChange>` +
      `<w:cellIns w:id="1" w:author="A"/>` +
      `<w:moveFromRangeStart w:id="1" w:name="m"/><w:delText>t</w:delText><w:insideH w:val="x"/>`;
    expect(renumberRawRevisionIds(raw, revisionIds)).toBe(
      `<w:ins w:id="2" w:author="A"><w:r><w:instrText>x</w:instrText></w:r></w:ins>` +
        `<w:rPr><w:del w:id="3" w:author="A"/></w:rPr>` +
        `<w:pPrChange w:id="4" w:author="A"><w:pPr/></w:pPrChange>` +
        `<w:cellIns w:id="5" w:author="A"/>` +
        `<w:moveFromRangeStart w:id="1" w:name="m"/><w:delText>t</w:delText><w:insideH w:val="x"/>`
    );
  });

  it("leaves XML without revisions alone", () => {
    const raw = `<w:r><w:t>plain</w:t></w:r>`;
    expect(renumberRawRevisionIds(raw, createRevisionIdMap())).toBe(raw);
  });
});

describe("written packages", () => {
  it("carry unique bookmark ids across body and header", async () => {
    const doc = Document.create();
    const mark = (name: string) =>
      Build.paragraph([Build.bookmarkStart(0, name), Build.text(name), Build.bookmarkEnd(0)]);
    Document.setHeader(doc, "default", { children: [mark("head")] });
    Document.addParagraph(doc, mark("a"));
    Document.addParagraph(doc, mark("b"));
    const read = await Io.read(await Io.toBuffer(doc));
    const ids: number[] = [];
    for (const block of [
      ...read.body,
      ...[...(read.headers?.values() ?? [])].flatMap(h => h.content.children)
    ]) {
      if (block.type === "paragraph") {
        const [start, , end] = block.children as { type?: string; id?: number }[];
        if (start?.type === "bookmarkStart") {
          expect(end?.id).toBe(start.id);
          ids.push(start.id!);
        }
      }
    }
    expect(ids).toHaveLength(3);
    expect(new Set(ids).size).toBe(3);
  });
});

describe("createBookmarkIdMap — ids outside the common range", () => {
  it("handles negative and very large ids as it does small ones", () => {
    const map = createBookmarkIdMap();
    expect([map.start(-5), map.start(-5)]).toEqual([-5, 0]);
    expect(map.start(2 ** 30)).toBe(2 ** 30);
    expect(map.start(2 ** 30)).toBe(2 ** 30 + 1);
    expect([map.end(2 ** 30), map.end(2 ** 30), map.end(-5), map.end(-5)]).toEqual([
      2 ** 30 + 1,
      2 ** 30,
      0,
      -5
    ]);
  });
});
