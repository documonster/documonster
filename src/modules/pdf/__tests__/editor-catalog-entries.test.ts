/**
 * A full rewrite builds a new catalog. These tests pin which document-level entries it carries
 * over — bookmarks, named destinations, page labels and the rest — and how the ones that name
 * pages follow `removePage` and `splitPages` (see `builder/catalog-rewrite.ts`).
 */
import { PdfDocument } from "@pdf/reader/pdf-document";
import { decodePdfStringBytes, isPdfArray, isPdfRef } from "@pdf/reader/pdf-parser";
import type { PdfDictValue, PdfObject } from "@pdf/reader/pdf-parser";
import { describe, expect, it, vi } from "vitest";

import { Pdf } from "../index";
import { buildRawPdf } from "./raw-pdf.helpers";
import { expectSplitMatchesReference } from "./split-pages.helpers";

const content = (text: string): string => {
  const ops = `BT /F1 12 Tf 72 700 Td (${text}) Tj ET`;
  return `<< /Length ${ops.length} >>\nstream\n${ops}\nendstream`;
};
const page = (contents: number, extra = ""): string =>
  `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents ${contents} 0 R /Resources << /Font << /F1 9 0 R >> >> ${extra} >>`;

/**
 * Four pages (objects 3–6) with every catalog entry a rewrite has to think about.
 *
 * Outline:
 * - `Chapter 1` → page 0, open
 *   - `Section 1.1` → page 1
 *   - `Section 1.2` → named destination `sec12` (page 2, through the name tree)
 * - `Chapter 2` → page 3 through a `GoTo` action, closed
 *   - `Section 2.1` → `[3 /Fit]`, a page given by index
 * - `Website` → a URI action
 */
function fixture(openAction = "[5 0 R /Fit]"): Uint8Array {
  return buildRawPdf([
    /* 1 */ `<< /Type /Catalog /Pages 2 0 R /Outlines 14 0 R /Names << /Dests 21 0 R /JavaScript 22 0 R >>
      /Dests << /old [4 0 R /Fit] /gone [5 0 R /Fit] >>
      /PageLabels << /Nums [0 << /S /r >> 2 << /S /D /P (A-) /St 5 >>] >>
      /OpenAction ${openAction} /ViewerPreferences << /DisplayDocTitle true /PrintPageRange [1 2 4 4] >>
      /PageLayout /TwoColumnLeft /PageMode /UseOutlines /Lang (en-US) /MarkInfo << /Marked true >>
      /StructTreeRoot 23 0 R /OCProperties << /OCGs [24 0 R] /D << /Order [24 0 R] >> >>
      /OutputIntents [<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (sRGB) >>]
      /URI << /Base (https://example.com/) >> /Perms << /UR3 25 0 R >> /Version /1.7 >>`,
    /* 2 */ "<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R 6 0 R] /Count 4 >>",
    /* 3 */ page(10, "/StructParents 0"),
    /* 4 */ page(11, "/StructParents 1"),
    /* 5 */ page(12, "/StructParents 2"),
    /* 6 */ page(13, "/StructParents 3"),
    /* 7 */ "null",
    /* 8 */ "null",
    /* 9 */ "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    /* 10 */ content("P0"),
    /* 11 */ content("P1"),
    /* 12 */ content("P2"),
    /* 13 */ content("P3"),
    /* 14 */ "<< /Type /Outlines /First 15 0 R /Last 20 0 R /Count 5 >>",
    /* 15 */ "<< /Title (Chapter 1) /Parent 14 0 R /Next 18 0 R /First 16 0 R /Last 17 0 R /Count 2 /Dest [3 0 R /Fit] >>",
    /* 16 */ "<< /Title (Section 1.1) /Parent 15 0 R /Next 17 0 R /Dest [4 0 R /XYZ 0 792 0] >>",
    /* 17 */ "<< /Title (Section 1.2) /Parent 15 0 R /Prev 16 0 R /Dest (sec12) >>",
    /* 18 */ "<< /Title (Chapter 2) /Parent 14 0 R /Prev 15 0 R /Next 20 0 R /First 19 0 R /Last 19 0 R /Count -1 /A << /S /GoTo /D [6 0 R /Fit] >> /C [1 0 0] >>",
    /* 19 */ "<< /Title (Section 2.1) /Parent 18 0 R /Dest [3 /Fit] >>",
    /* 20 */ "<< /Title (Website) /Parent 14 0 R /Prev 18 0 R /A << /S /URI /URI (https://example.com/) >> >>",
    /* 21 */ "<< /Names [(p0) [3 0 R /Fit] (sec12) << /D [5 0 R /Fit] >>] >>",
    /* 22 */ "<< /Names [(init) << /S /JavaScript /JS (1;) >>] >>",
    /* 23 */ "<< /Type /StructTreeRoot /K [] >>",
    /* 24 */ "<< /Type /OCG /Name (Layer) >>",
    /* 25 */ "<< /Type /Sig >>"
  ]);
}

/** Run `body` with the clock stopped: a rewrite stamps /CreationDate when the original has none. */
async function withFixedClock(body: () => Promise<unknown>): Promise<void> {
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(new Date("2026-01-02T03:04:05Z"));
  try {
    await body();
  } finally {
    vi.useRealTimers();
  }
}

interface Outline {
  title: string;
  page: number | undefined;
  children: Outline[];
}

async function outline(bytes: Uint8Array): Promise<Outline[]> {
  const strip = (items: readonly { title: string; pageIndex?: number; children: any[] }[]) =>
    items.map((item): Outline => ({
      title: item.title,
      page: item.pageIndex,
      children: strip(item.children)
    }));
  return strip((await Pdf.extract(bytes)).bookmarks);
}

async function texts(bytes: Uint8Array): Promise<string[]> {
  return (await Pdf.extract(bytes)).pages.map(p => p.text.trim());
}

/** The catalog of a written file, with a reader to resolve what it points at. */
function catalogOf(bytes: Uint8Array): { doc: PdfDocument; catalog: PdfDictValue } {
  const doc = new PdfDocument(bytes);
  return { doc, catalog: doc.getCatalog() };
}

/** An object as plain data: refs resolved (once), strings decoded, names prefixed with `/`. */
function plain(doc: PdfDocument, value: PdfObject | undefined, seen = new Set<string>()): unknown {
  if (isPdfRef(value)) {
    const key = `${value.objNum}:${value.gen}`;
    if (seen.has(key)) {
      return "<cycle>";
    }
    return plain(doc, doc.deref(value), new Set([...seen, key]));
  }
  if (value instanceof Uint8Array) {
    return decodePdfStringBytes(value);
  }
  if (typeof value === "string") {
    return `/${value}`;
  }
  if (isPdfArray(value)) {
    return value.map(v => plain(doc, v, seen));
  }
  if (value instanceof Map) {
    return Object.fromEntries([...value].map(([k, v]) => [k, plain(doc, v, seen)]));
  }
  return value ?? null;
}

/** Each written page's object, by index, to compare destinations with. */
function pageNums(doc: PdfDocument): number[] {
  return doc.getPagesWithObjInfo().map(p => p.objNum);
}

/** The `/Count`s of the outline, depth first. */
function counts(doc: PdfDocument, catalog: PdfDictValue): Array<[string, unknown]> {
  const out: Array<[string, unknown]> = [];
  const root = doc.derefDict(catalog.get("Outlines"))!;
  out.push(["root", root.get("Count")]);
  const walk = (first: PdfObject | undefined) => {
    for (
      let item = doc.derefDict(first ?? null);
      item;
      item = doc.derefDict(item.get("Next") ?? null)
    ) {
      out.push([decodePdfStringBytes(item.get("Title") as Uint8Array), item.get("Count") ?? null]);
      walk(item.get("First"));
    }
  };
  walk(root.get("First"));
  return out;
}

const FULL_OUTLINE: Outline[] = [
  {
    title: "Chapter 1",
    page: 0,
    children: [
      { title: "Section 1.1", page: 1, children: [] },
      { title: "Section 1.2", page: 2, children: [] }
    ]
  },
  { title: "Chapter 2", page: 3, children: [{ title: "Section 2.1", page: 3, children: [] }] },
  { title: "Website", page: undefined, children: [] }
];

describe("PdfEditor full rewrite — the outline", () => {
  it("the fixture reads as intended", async () => {
    expect(await outline(fixture())).toEqual(FULL_OUTLINE);
    expect(await texts(fixture())).toEqual(["P0", "P1", "P2", "P3"]);
  });

  it("toBuffer() keeps every bookmark", async () => {
    const out = await Pdf.Editor.read(fixture()).toBuffer();
    expect(await outline(out)).toEqual(FULL_OUTLINE);
    const { doc, catalog } = catalogOf(out);
    expect(counts(doc, catalog)).toEqual([
      ["root", 5],
      ["Chapter 1", 2],
      ["Section 1.1", null],
      ["Section 1.2", null],
      ["Chapter 2", -1],
      ["Section 2.1", null],
      ["Website", null]
    ]);
  });

  it("removePage drops a bookmark to the removed page and renumbers the others", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(1);
    const out = await editor.toBuffer();
    expect(await texts(out)).toEqual(["P0", "P2", "P3"]);
    expect(await outline(out)).toEqual([
      { title: "Chapter 1", page: 0, children: [{ title: "Section 1.2", page: 1, children: [] }] },
      // `[3 /Fit]` named the fourth page by index; it is the third now.
      { title: "Chapter 2", page: 2, children: [{ title: "Section 2.1", page: 2, children: [] }] },
      { title: "Website", page: undefined, children: [] }
    ]);
    const { doc, catalog } = catalogOf(out);
    expect(counts(doc, catalog)).toEqual([
      ["root", 4],
      ["Chapter 1", 1],
      ["Section 1.2", null],
      ["Chapter 2", -1],
      ["Section 2.1", null],
      ["Website", null]
    ]);
  });

  it("a parent whose page is removed stays, without a destination, while a child survives", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(0);
    const out = await editor.toBuffer();
    expect(await outline(out)).toEqual([
      {
        title: "Chapter 1",
        page: undefined,
        children: [
          { title: "Section 1.1", page: 0, children: [] },
          { title: "Section 1.2", page: 1, children: [] }
        ]
      },
      { title: "Chapter 2", page: 2, children: [{ title: "Section 2.1", page: 2, children: [] }] },
      { title: "Website", page: undefined, children: [] }
    ]);
    const { doc, catalog } = catalogOf(out);
    const first = doc.derefDict(doc.derefDict(catalog.get("Outlines"))!.get("First"))!;
    expect(first.has("Dest")).toBe(false);
    expect(first.has("A")).toBe(false);
    // Its colour and other entries are kept.
    const second = doc.derefDict(first.get("Next"))!;
    expect(plain(doc, second.get("C"))).toEqual([1, 0, 0]);
    expect(second.get("Count")).toBe(-1);
  });

  it("a bookmark without surviving children or page is dropped", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(3);
    const out = await editor.toBuffer();
    expect((await outline(out)).map(b => b.title)).toEqual(["Chapter 1", "Website"]);
  });

  it("outline items are linked consistently: /Parent, /Prev, /Next, /First, /Last", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(1).removePage(3);
    const out = await editor.toBuffer();
    const doc = new PdfDocument(out);
    const rootRef = doc.getCatalog().get("Outlines");
    const check = (parentRef: PdfObject | undefined) => {
      const parent = doc.derefDict(parentRef ?? null)!;
      const first = parent.get("First");
      const last = parent.get("Last");
      let prev: PdfObject | undefined;
      let ref = first;
      let n = 0;
      while (ref) {
        const item = doc.derefDict(ref)!;
        expect(item.get("Parent")).toEqual(parentRef);
        expect(item.get("Prev")).toEqual(prev);
        if (!item.get("Next")) {
          expect(ref).toEqual(last);
        }
        if (item.get("First")) {
          check(ref);
        }
        prev = ref;
        ref = item.get("Next");
        n++;
      }
      return n;
    };
    expect(check(rootRef)).toBe(2);
  });

  it("an outline with a /Next cycle is read once and written without one", async () => {
    const bodies = [
      "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      page(4),
      content("P0"),
      "<< /Type /Outlines /First 6 0 R /Last 7 0 R /Count 2 >>",
      "<< /Title (A) /Parent 5 0 R /Next 7 0 R /Dest [3 0 R /Fit] /First 7 0 R /Last 7 0 R /Count 1 >>",
      "<< /Title (B) /Parent 5 0 R /Next 6 0 R /Dest [3 0 R /Fit] >>",
      "null",
      "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    ];
    const out = await Pdf.Editor.read(buildRawPdf(bodies)).toBuffer();
    expect(await outline(out)).toEqual([
      { title: "A", page: 0, children: [] },
      { title: "B", page: 0, children: [] }
    ]);
  });

  it("splitPages keeps the bookmarks that lead to the part's page, with their ancestors", async () => {
    const parts = await Pdf.Editor.read(fixture()).splitPages();
    expect(await Promise.all(parts.map(outline))).toEqual([
      [{ title: "Chapter 1", page: 0, children: [] }],
      [
        {
          title: "Chapter 1",
          page: undefined,
          children: [{ title: "Section 1.1", page: 0, children: [] }]
        }
      ],
      [
        {
          title: "Chapter 1",
          page: undefined,
          children: [{ title: "Section 1.2", page: 0, children: [] }]
        }
      ],
      [{ title: "Chapter 2", page: 0, children: [{ title: "Section 2.1", page: 0, children: [] }] }]
    ]);
  });

  it("a part that is an added page has no outline", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.addPage();
    const parts = await editor.splitPages([4]);
    expect(catalogOf(parts[0]).catalog.has("Outlines")).toBe(false);
  });

  it("bookmark titles of an encrypted document are encrypted again and read back", async () => {
    const builder = new Pdf.Builder();
    for (let i = 0; i < 3; i++) {
      builder.addPage().drawText(`P${i}`, { x: 72, y: 700 });
    }
    builder.addBookmark("Ünïcode 一", 0).addBookmark("Two", 1).addBookmark("Three", 2);
    builder.setEncryption({ userPassword: "u", ownerPassword: "o" });
    const editor = Pdf.Editor.read(await builder.toBuffer(), { password: "u" });
    editor.removePage(1);
    const out = await editor.toBuffer();
    const read = await Pdf.extract(out, { password: "u" });
    expect(read.bookmarks.map(b => [b.title, b.pageIndex])).toEqual([
      ["Ünïcode 一", 0],
      ["Three", 1]
    ]);
  });
});

describe("PdfEditor full rewrite — named destinations", () => {
  it("are pruned to the pages written, and keep resolving", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(1);
    const out = await editor.toBuffer();
    const { doc, catalog } = catalogOf(out);
    const pages = pageNums(doc);
    const names = plain(doc, catalog.get("Names")) as Record<string, any>;
    // `p0` stays on page 0; `sec12` follows page 2, now page 1.
    const tree = doc.derefDict(doc.derefDict(catalog.get("Names"))!.get("Dests"))!;
    const leaf = doc.deref(tree.get("Names")) as PdfObject[];
    expect(leaf.length).toBe(4);
    expect(decodePdfStringBytes(leaf[0] as Uint8Array)).toBe("p0");
    expect((doc.deref(leaf[1]) as PdfObject[])[0]).toMatchObject({ objNum: pages[0] });
    expect(decodePdfStringBytes(leaf[2] as Uint8Array)).toBe("sec12");
    expect((doc.derefDict(leaf[3])!.get("D") as PdfObject[])[0]).toMatchObject({
      objNum: pages[1]
    });
    // The other name trees are copied as they were.
    expect(names.JavaScript).toEqual({ Names: ["init", { S: "/JavaScript", JS: "1;" }] });
    // The /Dests dictionary loses `old`, whose page was removed, and keeps `gone`, whose was not.
    const dests = doc.derefDict(catalog.get("Dests"))!;
    expect([...dests.keys()]).toEqual(["gone"]);
    expect((dests.get("gone") as PdfObject[])[0]).toMatchObject({ objNum: pages[1] });
  });

  it("an entry with no destination left is dropped, and with it an empty tree", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(0).removePage(2);
    const out = await editor.toBuffer();
    const { doc, catalog } = catalogOf(out);
    const names = doc.derefDict(catalog.get("Names"))!;
    expect(names.has("Dests")).toBe(false);
    expect(names.has("JavaScript")).toBe(true);
  });
});

describe("PdfEditor full rewrite — other catalog entries", () => {
  it("toBuffer() keeps every entry it can, and drops /Perms", async () => {
    const out = await Pdf.Editor.read(fixture()).toBuffer();
    const { doc, catalog } = catalogOf(out);
    expect([...catalog.keys()].sort()).toEqual(
      [
        "Type",
        "Pages",
        "Outlines",
        "Names",
        "Dests",
        "PageLabels",
        "OpenAction",
        "ViewerPreferences",
        "PageLayout",
        "PageMode",
        "Lang",
        "MarkInfo",
        "StructTreeRoot",
        "OCProperties",
        "OutputIntents",
        "URI",
        "Version"
      ].sort()
    );
    expect(plain(doc, catalog.get("PageLabels"))).toEqual({
      Nums: [0, { S: "/r" }, 2, { S: "/D", P: "A-", St: 5 }]
    });
    expect(plain(doc, catalog.get("ViewerPreferences"))).toEqual({
      DisplayDocTitle: true,
      PrintPageRange: [1, 2, 4, 4]
    });
    expect(plain(doc, catalog.get("Lang"))).toBe("en-US");
    expect(plain(doc, catalog.get("PageLayout"))).toBe("/TwoColumnLeft");
    expect(plain(doc, catalog.get("OCProperties"))).toEqual({
      OCGs: [{ Type: "/OCG", Name: "Layer" }],
      D: { Order: [{ Type: "/OCG", Name: "Layer" }] }
    });
    // One OCG, shared by /OCGs and /Order, as in the original.
    const oc = doc.derefDict(catalog.get("OCProperties"))!;
    expect((oc.get("OCGs") as PdfObject[])[0]).toEqual(
      (doc.derefDict(oc.get("D"))!.get("Order") as PdfObject[])[0]
    );
    expect((doc.deref(catalog.get("OpenAction")) as PdfObject[])[0]).toMatchObject({
      objNum: pageNums(doc)[2]
    });
    // Each page keeps its key into the structure tree's parent tree.
    expect(doc.getPagesWithObjInfo().map(p => p.dict.get("StructParents"))).toEqual([0, 1, 2, 3]);
  });

  it("removePage relabels, renumbers the print range, and drops the structure tree", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(1);
    const out = await editor.toBuffer();
    const { doc, catalog } = catalogOf(out);
    // Pages keep their labels: i, A-5, A-6.
    expect(plain(doc, catalog.get("PageLabels"))).toEqual({
      Nums: [0, { S: "/r" }, 1, { S: "/D", P: "A-", St: 5 }]
    });
    // Pages 1, 2 and 4 (1-based) are to be printed; page 2 is gone and page 4 is now page 3.
    const prefs = doc.derefDict(catalog.get("ViewerPreferences"))!;
    expect(plain(doc, prefs.get("PrintPageRange"))).toEqual([1, 1, 3, 3]);
    expect(catalog.has("StructTreeRoot")).toBe(false);
    expect(catalog.has("MarkInfo")).toBe(false);
    expect(doc.getPagesWithObjInfo().some(p => p.dict.has("StructParents"))).toBe(false);
    // The open action follows its page, now the second.
    expect((doc.deref(catalog.get("OpenAction")) as PdfObject[])[0]).toMatchObject({
      objNum: pageNums(doc)[1]
    });
    expect(plain(doc, catalog.get("Lang"))).toBe("en-US");
  });

  it("an open action to a removed page is dropped", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(2);
    const { catalog } = catalogOf(await editor.toBuffer());
    expect(catalog.has("OpenAction")).toBe(false);
  });

  it("a page label continues onto an added page", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(3);
    editor.addPage();
    const { doc, catalog } = catalogOf(await editor.toBuffer());
    // i, ii, A-5, then the added page is A-6: one range covers both.
    expect(plain(doc, catalog.get("PageLabels"))).toEqual({
      Nums: [0, { S: "/r" }, 2, { S: "/D", P: "A-", St: 5 }]
    });
  });

  it("each part of splitPages carries its page's label and the entries that do not name pages", async () => {
    const parts = await Pdf.Editor.read(fixture()).splitPages();
    const labels = parts.map(part => {
      const { doc, catalog } = catalogOf(part);
      return plain(doc, catalog.get("PageLabels"));
    });
    expect(labels).toEqual([
      { Nums: [0, { S: "/r" }] },
      { Nums: [0, { S: "/r", St: 2 }] },
      { Nums: [0, { S: "/D", P: "A-", St: 5 }] },
      { Nums: [0, { S: "/D", P: "A-", St: 6 }] }
    ]);
    const { doc, catalog } = catalogOf(parts[2]);
    expect(plain(doc, catalog.get("Lang"))).toBe("en-US");
    expect((doc.deref(catalog.get("OpenAction")) as PdfObject[])[0]).toMatchObject({
      objNum: pageNums(doc)[0]
    });
    expect(catalogOf(parts[0]).catalog.has("OpenAction")).toBe(false);
    expect(catalog.has("StructTreeRoot")).toBe(false);
  });

  it("a name with a delimiter in it survives the rewrite", async () => {
    const bodies = [
      "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /Dests << /Chapter#201 [3 0 R /Fit] >> >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      page(4),
      content("P0"),
      "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
      "<< /Title (One) /Parent 5 0 R /Dest /Chapter#201 >>",
      "null",
      "null",
      "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    ];
    const out = await Pdf.Editor.read(buildRawPdf(bodies)).toBuffer();
    expect(await outline(out)).toEqual([{ title: "One", page: 0, children: [] }]);
    expect([...catalogOf(out).catalog.keys()]).toContain("Dests");
  });
});

describe("PdfEditor full rewrite — article threads (ISO 32000-1 §12.4.3)", () => {
  /**
   * Four pages (objects 3–6). Thread `A` runs through one bead on every page, in page order;
   * thread `B` is one bead on page 1. Each page lists its beads in `/B`.
   */
  function threaded(): Uint8Array {
    const bead = (thread: number, page: number, next: number, prev: number, first = false) =>
      `<< /Type /Bead ${first ? `/T ${thread} 0 R ` : ""}/N ${next} 0 R /V ${prev} 0 R /P ${page} 0 R /R [0 0 10 10] >>`;
    return buildRawPdf([
      /* 1 */ "<< /Type /Catalog /Pages 2 0 R /Threads [30 0 R 35 0 R] >>",
      /* 2 */ "<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R 6 0 R] /Count 4 >>",
      /* 3 */ page(10, "/B [31 0 R]"),
      /* 4 */ page(11, "/B [32 0 R 36 0 R]"),
      /* 5 */ page(12, "/B [33 0 R]"),
      /* 6 */ page(13, "/B [34 0 R]"),
      ...Array.from({ length: 2 }, () => "null"),
      /* 9 */ "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
      /* 10 */ content("P0"),
      /* 11 */ content("P1"),
      /* 12 */ content("P2"),
      /* 13 */ content("P3"),
      ...Array.from({ length: 16 }, () => "null"),
      /* 30 */ "<< /Type /Thread /F 31 0 R /I << /Title (A) >> >>",
      /* 31 */ bead(30, 3, 32, 34, true),
      /* 32 */ bead(30, 4, 33, 31),
      /* 33 */ bead(30, 5, 34, 32),
      /* 34 */ bead(30, 6, 31, 33),
      /* 35 */ "<< /Type /Thread /F 36 0 R /I << /Title (B) >> >>",
      /* 36 */ bead(35, 4, 36, 36, true)
    ]);
  }

  /**
   * Each thread of a written file as its title and the pages of its beads in reading order,
   * after checking that the beads form a ring as §12.4.3 requires: `/N` and `/V` agree, the
   * last bead leads back to the first, and the first bead's `/T` is the thread.
   */
  function threadsOf(bytes: Uint8Array): Array<[string, number[]]> {
    const { doc, catalog } = catalogOf(bytes);
    const pages = pageNums(doc);
    const list = doc.deref(catalog.get("Threads") ?? null);
    if (!isPdfArray(list)) {
      return [];
    }
    return list.map(threadRef => {
      expect(isPdfRef(threadRef)).toBe(true);
      const thread = doc.derefDict(threadRef)!;
      const first = thread.get("F") as { objNum: number };
      expect(doc.derefDict(first as PdfObject)!.get("T")).toEqual(threadRef);
      const order: number[] = [];
      let at = first;
      do {
        const bead = doc.derefDict(at as PdfObject)!;
        order.push(pages.indexOf((bead.get("P") as { objNum: number }).objNum));
        const next = bead.get("N") as { objNum: number };
        expect((doc.derefDict(next as PdfObject)!.get("V") as { objNum: number }).objNum).toBe(
          at.objNum
        );
        at = next;
      } while (at.objNum !== first.objNum && order.length < 10);
      const title = doc.derefDict(thread.get("I") ?? null)?.get("Title") as Uint8Array;
      return [decodePdfStringBytes(title), order];
    });
  }

  /** Each written page's `/B`, as the pages its beads say they are on. */
  function beadLists(bytes: Uint8Array): number[][] {
    const doc = new PdfDocument(bytes);
    const pages = pageNums(doc);
    return doc.getPagesWithObjInfo().map(({ dict }) => {
      const list = doc.deref(dict.get("B") ?? null);
      return isPdfArray(list)
        ? list.map(ref =>
            pages.indexOf((doc.derefDict(ref)!.get("P") as { objNum: number }).objNum)
          )
        : [];
    });
  }

  const removing = async (...pages: number[]): Promise<Uint8Array> => {
    const editor = Pdf.Editor.read(threaded());
    for (const p of pages) {
      editor.removePage(p);
    }
    return editor.toBuffer();
  };

  it("toBuffer() keeps every thread and every page's /B", async () => {
    const out = await Pdf.Editor.read(threaded()).toBuffer();
    expect(threadsOf(out)).toEqual([
      ["A", [0, 1, 2, 3]],
      ["B", [1]]
    ]);
    expect(beadLists(out)).toEqual([[0], [1, 1], [2], [3]]);
  });

  it("a middle page removed: its beads go, the rest are relinked, a thread left empty goes", async () => {
    const out = await removing(1);
    expect(threadsOf(out)).toEqual([["A", [0, 1, 2]]]);
    expect(beadLists(out)).toEqual([[0], [1], [2]]);
  });

  it("the first page removed: the thread starts at the next bead, which gains /T", async () => {
    const out = await removing(0);
    expect(threadsOf(out)).toEqual([
      ["A", [0, 1, 2]],
      ["B", [0]]
    ]);
    expect(beadLists(out)).toEqual([[0, 0], [1], [2]]);
  });

  it("the last page removed: the last bead left leads back to the first", async () => {
    const out = await removing(3);
    expect(threadsOf(out)).toEqual([
      ["A", [0, 1, 2]],
      ["B", [1]]
    ]);
  });

  it("every page of a thread removed: the thread goes, and no bead of it is written", async () => {
    const out = await removing(0, 1, 2);
    expect(threadsOf(out)).toEqual([["A", [0]]]);
    expect(new TextDecoder("latin1").decode(out).match(/\/Type\s*\/Bead/g)).toHaveLength(1);
  });

  it("pages added and copied after the originals leave the threads as they are", async () => {
    const editor = Pdf.Editor.read(threaded());
    editor.addPage();
    editor.copyPagesFrom(threaded(), { pageIndices: [1] });
    const out = await editor.toBuffer();
    expect(threadsOf(out)).toEqual([
      ["A", [0, 1, 2, 3]],
      ["B", [1]]
    ]);
    // A copied page brings no bead of another document's thread.
    expect(beadLists(out)).toEqual([[0], [1, 1], [2], [3], [], []]);
  });

  it("each part of splitPages keeps the beads on its page, as a thread of their own", async () => {
    const parts = await Pdf.Editor.read(threaded()).splitPages();
    expect(parts.map(threadsOf)).toEqual([
      [["A", [0]]],
      [
        ["A", [0]],
        ["B", [0]]
      ],
      [["A", [0]]],
      [["A", [0]]]
    ]);
    expect(parts.map(beadLists)).toEqual([[[0]], [[0, 0]], [[0]], [[0]]]);
  });

  it("each part is what removing every other page writes", async () => {
    await withFixedClock(async () => {
      await expectSplitMatchesReference(threaded());
      await expectSplitMatchesReference(threaded(), { edit: editor => void editor.removePage(1) });
    });
  });

  it("a part that is an added page has no threads", async () => {
    const editor = Pdf.Editor.read(threaded());
    editor.addPage();
    const [part] = await editor.splitPages([4]);
    expect(catalogOf(part).catalog.has("Threads")).toBe(false);
  });
});

describe("PdfEditor full rewrite — /ViewerPreferences /PrintPageRange (ISO 32000-1 §12.2)", () => {
  /** Four pages; `range` is the print range, written as given. */
  const ranged = (range: string): Uint8Array =>
    buildRawPdf([
      `<< /Type /Catalog /Pages 2 0 R /ViewerPreferences << /DisplayDocTitle true /PrintPageRange ${range} >> >>`,
      "<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R 6 0 R] /Count 4 >>",
      page(10),
      page(11),
      page(12),
      page(13),
      "null",
      "null",
      "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
      content("P0"),
      content("P1"),
      content("P2"),
      content("P3")
    ]);
  const printRange = async (
    range: string,
    edit: (editor: ReturnType<typeof Pdf.Editor.read>) => void
  ): Promise<unknown> => {
    const editor = Pdf.Editor.read(ranged(range));
    edit(editor);
    const { doc, catalog } = catalogOf(await editor.toBuffer());
    return plain(doc, catalog.get("ViewerPreferences"));
  };

  it.each([
    ["the first page removed", "[1 2 4 4]", [0], [1, 1, 3, 3]],
    ["the last page removed", "[1 2 4 4]", [3], [1, 2]],
    ["a middle page inside a range removed", "[1 3]", [1], [1, 2]],
    ["every page of one range removed", "[1 2 4 4]", [0, 1], [2, 2]],
    ["ranges that meet once a page is gone are merged", "[1 1 3 4]", [1], [1, 3]],
    ["an odd trailing number is ignored", "[1 2 4]", [0], [1, 1]],
    ["entries that are not numbers are skipped with their pair", "[(x) 2 3 4]", [0], [2, 3]]
  ])("%s", async (_, range, removed, expected) => {
    expect(
      await printRange(range, editor => {
        for (const p of removed) {
          editor.removePage(p);
        }
      })
    ).toEqual({ DisplayDocTitle: true, PrintPageRange: expected });
  });

  it("no page of the range left: the key is dropped, the other preferences kept", async () => {
    expect(
      await printRange("[2 2]", editor => {
        editor.removePage(1);
      })
    ).toEqual({ DisplayDocTitle: true });
  });

  it("an added page after the originals is not in the range; the originals keep theirs", async () => {
    expect(
      await printRange("[1 4]", editor => {
        editor.removePage(2);
        editor.addPage();
      })
    ).toEqual({ DisplayDocTitle: true, PrintPageRange: [1, 3] });
    // Nothing moved: the range is written as it was.
    expect(
      await printRange("[1 4]", editor => {
        editor.addPage();
      })
    ).toEqual({ DisplayDocTitle: true, PrintPageRange: [1, 4] });
  });

  it("each part is what removing every other page writes", async () => {
    await withFixedClock(() => expectSplitMatchesReference(ranged("[1 2 4 4]")));
  });

  it("a part of splitPages prints its page when it was in the range", async () => {
    const parts = await Pdf.Editor.read(ranged("[1 2 4 4]")).splitPages();
    expect(
      parts.map(part => {
        const { doc, catalog } = catalogOf(part);
        return plain(doc, catalog.get("ViewerPreferences"));
      })
    ).toEqual([
      { DisplayDocTitle: true, PrintPageRange: [1, 1] },
      { DisplayDocTitle: true, PrintPageRange: [1, 1] },
      { DisplayDocTitle: true },
      { DisplayDocTitle: true, PrintPageRange: [1, 1] }
    ]);
  });
});

describe("PdfEditor full rewrite — page labels, open action and named destinations at the edges", () => {
  const labelsAfter = async (...removed: number[]): Promise<unknown> => {
    const editor = Pdf.Editor.read(fixture());
    for (const p of removed) {
      editor.removePage(p);
    }
    const { doc, catalog } = catalogOf(await editor.toBuffer());
    return plain(doc, catalog.get("PageLabels"));
  };

  it("page labels (§12.4.2): every page keeps its label, and index 0 always has a range", async () => {
    // i, ii, A-5, A-6 as read.
    expect(await labelsAfter(0)).toEqual({
      Nums: [0, { S: "/r", St: 2 }, 1, { S: "/D", P: "A-", St: 5 }]
    });
    expect(await labelsAfter(3)).toEqual({
      Nums: [0, { S: "/r" }, 2, { S: "/D", P: "A-", St: 5 }]
    });
    // Every page of the first range removed.
    expect(await labelsAfter(0, 1)).toEqual({ Nums: [0, { S: "/D", P: "A-", St: 5 }] });
    // The middle of a range removed: the range continues where its page left off.
    expect(await labelsAfter(2)).toEqual({
      Nums: [0, { S: "/r" }, 2, { S: "/D", P: "A-", St: 6 }]
    });
  });

  it("a copied page continues the numbering of the page before it", async () => {
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(0);
    editor.copyPagesFrom(fixture(), { pageIndices: [0] });
    const { doc, catalog } = catalogOf(await editor.toBuffer());
    // ii, A-5, A-6, then the copied page is A-7.
    expect(plain(doc, catalog.get("PageLabels"))).toEqual({
      Nums: [0, { S: "/r", St: 2 }, 1, { S: "/D", P: "A-", St: 5 }]
    });
  });

  it("open action (Table 28, §12.3.2, §12.6): follows its page, or is dropped with it", async () => {
    /** The open action written after removing `removed`: its kind and the page it goes to. */
    const open = async (action: string, removed: number[]) => {
      const editor = Pdf.Editor.read(fixture(action));
      for (const p of removed) {
        editor.removePage(p);
      }
      const { doc, catalog } = catalogOf(await editor.toBuffer());
      const value = doc.deref(catalog.get("OpenAction") ?? null);
      if (value === null) {
        return null;
      }
      if (value instanceof Uint8Array) {
        return { kind: "name", page: decodePdfStringBytes(value) };
      }
      const dest = isPdfArray(value) ? value : (value as PdfDictValue).get("D");
      const page = isPdfArray(dest)
        ? pageNums(doc).indexOf((dest[0] as { objNum: number }).objNum)
        : undefined;
      return { kind: isPdfArray(value) ? "dest" : (value as PdfDictValue).get("S"), page };
    };
    // A destination to the last page follows it when the first goes, and goes with it.
    expect(await open("[6 0 R /Fit]", [0])).toEqual({ kind: "dest", page: 2 });
    expect(await open("[6 0 R /Fit]", [3])).toBeNull();
    // A GoTo action the same.
    expect(await open("<< /S /GoTo /D [3 0 R /Fit] >>", [1])).toEqual({ kind: "GoTo", page: 0 });
    expect(await open("<< /S /GoTo /D [3 0 R /Fit] >>", [0])).toBeNull();
    // A named destination stays while its page does (the name is kept in /Names, renumbered).
    expect(await open("(sec12)", [0])).toEqual({ kind: "name", page: "sec12" });
    expect(await open("(sec12)", [2])).toBeNull();
    // An action that goes to no page is kept whatever is removed.
    expect(await open("<< /S /Named /N /NextPage >>", [0, 1])).toEqual({
      kind: "Named",
      page: undefined
    });
  });

  it("named destinations (§12.3.2.3): first and last pages removed, and a split part", async () => {
    const namesOf = (bytes: Uint8Array) => {
      const { doc, catalog } = catalogOf(bytes);
      const pages = pageNums(doc);
      const names = doc.derefDict(catalog.get("Names") ?? null);
      const tree = names ? doc.derefDict(names.get("Dests") ?? null) : null;
      const leaf = tree ? (doc.deref(tree.get("Names") ?? null) as PdfObject[]) : [];
      const out: Record<string, number> = {};
      for (let i = 0; i + 1 < leaf.length; i += 2) {
        const value = doc.deref(leaf[i + 1]);
        const dest = (value instanceof Map ? value.get("D") : value) as Array<{ objNum: number }>;
        out[decodePdfStringBytes(leaf[i] as Uint8Array)] = pages.indexOf(dest[0].objNum);
      }
      const dests = doc.derefDict(catalog.get("Dests") ?? null);
      for (const [key, dest] of dests ?? []) {
        out[`/${key}`] = pages.indexOf((dest as Array<{ objNum: number }>)[0].objNum);
      }
      return out;
    };
    const editor = Pdf.Editor.read(fixture());
    editor.removePage(0);
    // p0 → page 0 is gone; sec12 → page 2 is now 1; /old → page 1 is now 0; /gone → page 2 is now 1.
    expect(namesOf(await editor.toBuffer())).toEqual({ sec12: 1, "/old": 0, "/gone": 1 });
    const last = Pdf.Editor.read(fixture());
    last.removePage(3);
    expect(namesOf(await last.toBuffer())).toEqual({ p0: 0, sec12: 2, "/old": 1, "/gone": 2 });
    const parts = await Pdf.Editor.read(fixture()).splitPages();
    expect(parts.map(namesOf)).toEqual([{ p0: 0 }, { "/old": 0 }, { sec12: 0, "/gone": 0 }, {}]);
  });
});

describe("PdfEditor full rewrite — outline root /Count and action chains", () => {
  it("the outline root has /Count only while an item is open (§12.3.3, Table 152)", async () => {
    const editor = Pdf.Editor.read(fixture());
    // Chapter 1 loses both children, so nothing is open any more.
    editor.removePage(1).removePage(2);
    const { doc, catalog } = catalogOf(await editor.toBuffer());
    expect(doc.derefDict(catalog.get("Outlines"))!.has("Count")).toBe(false);
    // Chapter 2 is closed, so it still says how many it would show.
    expect(counts(doc, catalog).find(([title]) => title === "Chapter 2")).toEqual([
      "Chapter 2",
      -1
    ]);
  });

  /** The open action written after removing `removed`, as plain data. */
  const openAfter = async (action: string, removed: number[]): Promise<unknown> => {
    const editor = Pdf.Editor.read(fixture(action));
    for (const p of removed) {
      editor.removePage(p);
    }
    const { doc, catalog } = catalogOf(await editor.toBuffer());
    return plain(doc, catalog.get("OpenAction"));
  };

  it("an open GoTo whose page is gone is replaced by the rest of its chain", async () => {
    expect(
      await openAfter("<< /S /GoTo /D [5 0 R /Fit] /Next << /S /JavaScript /JS (a;) >> >>", [2])
    ).toEqual({ S: "/JavaScript", JS: "a;" });
    // A /Next array, itself starting with a GoTo to a removed page: what is left runs in order.
    expect(
      await openAfter(
        "<< /S /GoTo /D [5 0 R /Fit] /Next [<< /S /GoTo /D [5 0 R /Fit] /Next << /S /JavaScript /JS (b;) >> >> << /S /Named /N /NextPage >>] >>",
        [2]
      )
    ).toEqual({ S: "/JavaScript", JS: "b;", Next: { S: "/Named", N: "/NextPage" } });
    // Nothing left: dropped.
    expect(
      await openAfter("<< /S /GoTo /D [5 0 R /Fit] /Next << /S /GoTo /D [5 0 R /Fit] >> >>", [2])
    ).toBeNull();
    // A kept GoTo loses only the GoTo actions in its /Next that lead to a removed page.
    expect(
      await openAfter(
        "<< /S /GoTo /D [3 0 R /Fit] /Next [<< /S /GoTo /D [5 0 R /Fit] >> << /S /Named /N /NextPage >>] >>",
        [2]
      )
    ).toMatchObject({ S: "/GoTo", Next: { S: "/Named", N: "/NextPage" } });
  });

  it("a bookmark whose GoTo page is gone keeps the rest of its chain, and stays", async () => {
    const bodies = [
      "<< /Type /Catalog /Pages 2 0 R /Outlines 7 0 R >>",
      "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
      page(5),
      page(6),
      content("P0"),
      content("P1"),
      "<< /Type /Outlines /First 8 0 R /Last 8 0 R /Count 1 >>",
      "<< /Title (X) /Parent 7 0 R /A << /S /GoTo /D [4 0 R /Fit] /Next << /S /URI /URI (https://e.com/) >> >> >>",
      "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    ];
    const editor = Pdf.Editor.read(buildRawPdf(bodies));
    editor.removePage(1);
    const { doc, catalog } = catalogOf(await editor.toBuffer());
    const item = doc.derefDict(doc.derefDict(catalog.get("Outlines"))!.get("First"))!;
    expect(plain(doc, item.get("A"))).toEqual({ S: "/URI", URI: "https://e.com/" });
  });
});
