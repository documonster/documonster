/**
 * The `w:id` of a revision (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`) is an
 * annotation identifier, which ECMA-376 §17.13.1 requires to be unique within
 * the document. The model keeps one tracked run per run, so a revision that
 * wrapped several runs comes back as several tracked runs sharing one
 * `RevisionInfo`; written one wrapper per run they would all carry that id.
 * The writer therefore writes adjacent tracked runs of one revision as the one
 * wrapper they came from, and gives every other wrapper an id of its own —
 * across every part of the package — while keeping its author and date.
 */
import { unzip } from "@archive/read-archive";
import { createZip } from "@archive/zip/zip-bytes";
import { describe, it, expect } from "vitest";

import { Build, Document, Io, Package, Query, Streaming, Transform } from "../index";
import type { DocxDocument, Paragraph, RevisionInfo, StyleDef, Table } from "../types";

const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const enc = new TextEncoder();

/** A package with `body` as its document body. */
async function docx(body: string): Promise<Uint8Array> {
  return createZip([
    {
      name: "[Content_Types].xml",
      data: enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
      )
    },
    {
      name: "_rels/.rels",
      data: enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
      )
    },
    {
      name: "word/document.xml",
      data: enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`
      )
    }
  ]);
}

async function parts(bytes: Uint8Array): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for await (const entry of unzip(bytes).entries()) {
    if (entry.path.endsWith(".xml")) {
      out.set(entry.path, new TextDecoder().decode(await entry.bytes()));
    }
  }
  return out;
}

const WRAPPER = /<w:(ins|del|moveFrom|moveTo)\b([^>]*)>/g;

interface Wrapper {
  readonly tag: string;
  readonly id: string;
  readonly author: string | undefined;
  readonly date: string | undefined;
}

/** Every run-level revision wrapper in `xml`. */
function wrappers(xml: string): Wrapper[] {
  return [...xml.matchAll(WRAPPER)]
    .filter(m => !m[2]!.endsWith("/"))
    .map(m => ({
      tag: m[1]!,
      id: /\sw:id="([^"]*)"/.exec(m[2]!)![1]!,
      author: /\sw:author="([^"]*)"/.exec(m[2]!)?.[1],
      date: /\sw:date="([^"]*)"/.exec(m[2]!)?.[1]
    }));
}

/** Every revision wrapper of a written package, over all of its parts. */
async function packageWrappers(bytes: Uint8Array): Promise<Wrapper[]> {
  const out: Wrapper[] = [];
  for (const [path, xml] of await parts(bytes)) {
    if (path.startsWith("word/")) {
      out.push(...wrappers(xml));
    }
  }
  return out;
}

function expectUniqueIds(list: readonly Wrapper[]): void {
  const ids = list.map(w => w.id);
  expect(new Set(ids).size).toBe(ids.length);
}

const REV = `w:author="Ann" w:date="2024-01-01T00:00:00Z"`;
const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const delRun = (text: string) => `<w:r><w:delText>${text}</w:delText></w:r>`;

describe("revision wrappers: one wrapper per revision", () => {
  it.each(["ins", "del", "moveFrom", "moveTo"] as const)(
    "round-trips a w:%s of three runs as one wrapper",
    async tag => {
      const make = tag === "del" ? delRun : run;
      const bytes = await docx(
        `<w:p><w:${tag} w:id="7" ${REV}>${make("a")}${make("b")}${make("c")}</w:${tag}></w:p>`
      );
      const doc = await Io.read(bytes);
      const xml = (await parts(await Io.toBuffer(doc))).get("word/document.xml")!;
      const found = wrappers(xml);
      expect(found).toEqual([{ tag, id: "7", author: "Ann", date: "2024-01-01T00:00:00Z" }]);
      const inner = xml.slice(xml.indexOf(`<w:${tag} `), xml.indexOf(`</w:${tag}>`));
      expect(inner.match(/<w:r>/g)).toHaveLength(3);
    }
  );

  it("keeps a bookmark between the runs of one revision inside it", async () => {
    const bytes = await docx(
      `<w:p><w:ins w:id="2" ${REV}>${run("a")}<w:bookmarkStart w:id="0" w:name="m"/>${run("b")}<w:bookmarkEnd w:id="0"/>${run("c")}</w:ins></w:p>`
    );
    const xml = (await parts(await Io.toBuffer(await Io.read(bytes)))).get("word/document.xml")!;
    expect(wrappers(xml)).toHaveLength(1);
    const inner = xml.slice(xml.indexOf("<w:ins "), xml.indexOf("</w:ins>"));
    expect(inner).toContain('<w:bookmarkStart w:id="0" w:name="m"/>');
    expect(inner).toContain('<w:bookmarkEnd w:id="0"/>');
  });

  it("does not take markers that only follow the revision into it", async () => {
    const bytes = await docx(
      `<w:p><w:ins w:id="2" ${REV}>${run("a")}${run("b")}</w:ins><w:bookmarkStart w:id="0" w:name="m"/>${run("c")}<w:bookmarkEnd w:id="0"/></w:p>`
    );
    const xml = (await parts(await Io.toBuffer(await Io.read(bytes)))).get("word/document.xml")!;
    expect(wrappers(xml)).toHaveLength(1);
    expect(xml.indexOf("<w:bookmarkStart")).toBeGreaterThan(xml.indexOf("</w:ins>"));
  });

  it("does not join revisions of different kinds or different revision info", () => {
    const a: RevisionInfo = { id: 1, author: "Ann" };
    const b: RevisionInfo = { id: 1, author: "Bob" };
    const doc = Document.create();
    Document.addContent(
      doc,
      Build.paragraph([
        Build.insertedRun(Build.text("x"), a),
        Build.insertedRun(Build.text("y"), b),
        Build.deletedRun(Build.text("z"), b)
      ])
    );
    return Io.toBuffer(doc).then(async bytes => {
      const found = await packageWrappers(bytes);
      expect(found.map(w => [w.tag, w.author])).toEqual([
        ["ins", "Ann"],
        ["ins", "Bob"],
        ["del", "Bob"]
      ]);
      expectUniqueIds(found);
    });
  });
});

describe("revision wrappers: unique ids", () => {
  it("gives the pieces of a revision split by a comment reference ids of their own", async () => {
    const bytes = await docx(
      `<w:p><w:ins w:id="4" ${REV}><w:r><w:t>see</w:t><w:commentReference w:id="9"/><w:t>more</w:t></w:r></w:ins></w:p>`
    );
    const doc = await Io.read(bytes);
    const found = await packageWrappers(await Io.toBuffer(doc));
    expect(found).toHaveLength(2);
    expectUniqueIds(found);
    expect(found[0]!.id).toBe("4");
    for (const w of found) {
      expect(w.author).toBe("Ann");
      expect(w.date).toBe("2024-01-01T00:00:00Z");
    }
  });

  it("numbers apart the non-adjacent pieces of one revision, keeping author and date", async () => {
    const rev: RevisionInfo = { id: 3, author: "Ann", date: "2024-02-02T00:00:00Z" };
    const doc = Document.create();
    Document.addContent(
      doc,
      Build.paragraph([
        Build.insertedRun(Build.text("a"), rev),
        Build.text(" plain "),
        Build.insertedRun(Build.text("b"), rev)
      ])
    );
    Document.addContent(doc, Build.paragraph([Build.insertedRun(Build.text("c"), rev)]));
    const found = await packageWrappers(await Io.toBuffer(doc));
    expect(found).toHaveLength(3);
    expectUniqueIds(found);
    expect(found[0]!.id).toBe("3");
    expect(found.every(w => w.author === "Ann" && w.date === "2024-02-02T00:00:00Z")).toBe(true);
  });

  it("makes the revision ids of merged documents unique", async () => {
    const make = (text: string): DocxDocument => {
      const doc = Document.create();
      Document.addContent(
        doc,
        Build.paragraph([Build.insertedRun(Build.text(text), { id: 1, author: "Ann" })])
      );
      Document.addContent(
        doc,
        Build.paragraph([Build.deletedRun(Build.text(text), { id: 2, author: "Ann" })])
      );
      return doc;
    };
    const merged = Transform.merge([make("first"), make("second")]);
    const found = await packageWrappers(await Io.toBuffer(merged));
    expect(found).toHaveLength(4);
    expectUniqueIds(found);
  });

  it("is unique across the parts of a package", async () => {
    const rev: RevisionInfo = { id: 5, author: "Ann" };
    const header: Paragraph = Build.paragraph([Build.insertedRun(Build.text("head"), rev)]);
    const doc: DocxDocument = {
      ...Document.create(),
      body: [Build.paragraph([Build.insertedRun(Build.text("body"), rev)])],
      headers: new Map([["rId10", { content: { children: [header] }, rId: "rId10" }]]),
      sectionProperties: { headers: [{ type: "default", rId: "rId10" }] }
    };
    const found = await packageWrappers(await Io.toBuffer(doc));
    expect(found).toHaveLength(2);
    expectUniqueIds(found);
  });

  it("is unique in a streamed package", async () => {
    const rev: RevisionInfo = { id: 1, author: "Ann" };
    const writer = Streaming.createWriter();
    writer.add(Build.paragraph([Build.insertedRun(Build.text("a"), rev)]));
    writer.add(Build.paragraph([Build.insertedRun(Build.text("b"), rev)]));
    const found = await packageWrappers(await writer.finalize());
    expect(found).toHaveLength(2);
    expectUniqueIds(found);
  });

  it("keeps clear of the revision ids of parts Package.edit leaves untouched", async () => {
    const doc = Document.create();
    Document.addContent(doc, Build.paragraph([Build.text("body")]));
    const seeded = await Package.edit(await Io.toBuffer(doc), [
      {
        type: "replacePartText",
        path: "word/header1.xml",
        text: `<?xml version="1.0"?><w:hdr ${W}><w:p><w:ins w:id="6" ${REV}>${run("h")}</w:ins></w:p></w:hdr>`
      }
    ]);
    const edited = await Package.edit(seeded, [
      {
        type: "replaceBody",
        body: [Build.paragraph([Build.insertedRun(Build.text("b"), { id: 6, author: "Ann" })])]
      }
    ]);
    const found = await packageWrappers(edited);
    expect(found).toHaveLength(2);
    expectUniqueIds(found);
  });

  it("leaves accept and reject, which work on the model, unaffected", async () => {
    const bytes = await docx(
      `<w:p><w:ins w:id="4" ${REV}><w:r><w:t>see</w:t><w:commentReference w:id="9"/><w:t>more</w:t></w:r></w:ins>${run(" tail")}</w:p>`
    );
    const doc = await Io.read(await Io.toBuffer(await Io.read(bytes)));
    Query.rejectAllRevisions(doc);
    const xml = (await parts(await Io.toBuffer(doc))).get("word/document.xml")!;
    expect(wrappers(xml)).toHaveLength(0);
    expect(xml).not.toContain("see");
    expect(xml).not.toContain("more");
    expect(xml).toContain(" tail");
  });
});

/** Every revision element of a written package — wrappers, marks and property changes. */
const REVISION =
  /<w:(ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|pPrChange|rPrChange|tblPrChange|trPrChange|tcPrChange|sectPrChange)\b[^>]*?\sw:id="([^"]*)"/g;

async function revisionIds(bytes: Uint8Array): Promise<{ tag: string; id: string }[]> {
  const out: { tag: string; id: string }[] = [];
  for (const [path, xml] of await parts(bytes)) {
    if (path.startsWith("word/")) {
      for (const m of xml.matchAll(REVISION)) {
        out.push({ tag: m[1]!, id: m[2]! });
      }
    }
  }
  return out;
}

describe("revision ids: every kind of revision", () => {
  /** A document whose every revision — of every kind, in body and styles — carries id 1. */
  function mixed(): DocxDocument {
    const rev: RevisionInfo = { id: 1, author: "Ann", date: "2024-03-03T00:00:00Z" };
    const table: Table = Build.table(
      [
        Build.row(
          [
            Build.cell("cell", {
              inserted: { revision: rev },
              propertyChange: { revision: rev, previousProperties: {} }
            })
          ],
          { inserted: { revision: rev }, propertyChange: { revision: rev, previousProperties: {} } }
        )
      ],
      { propertyChange: { revision: rev, previousProperties: {} } }
    );
    const style: StyleDef = {
      type: "character",
      styleId: "Marked",
      name: "Marked",
      runProperties: { bold: true, propertyChange: { revision: rev, previousProperties: {} } }
    };
    const doc = Document.create();
    return {
      ...doc,
      styles: [...(doc.styles ?? []), style],
      body: [
        Build.paragraph(
          [
            Build.text("changed", {
              italic: true,
              propertyChange: { revision: rev, previousProperties: {} }
            }),
            Build.insertedRun(Build.text(" added"), rev),
            Build.deletedRun(Build.text(" removed"), rev)
          ],
          {
            alignment: "center",
            propertyChange: { revision: rev, previousProperties: {} },
            paragraphInsertion: rev
          }
        ),
        table,
        Build.paragraph([Build.text("after")])
      ],
      sectionProperties: {
        ...doc.sectionProperties,
        propertyChange: { revision: rev, previousProperties: {} }
      }
    };
  }

  it("gives every revision element of the package an id of its own", async () => {
    const found = await revisionIds(await Io.toBuffer(mixed()));
    expect(new Set(found.map(r => r.tag))).toEqual(
      new Set([
        "ins",
        "del",
        "pPrChange",
        "rPrChange",
        "tblPrChange",
        "trPrChange",
        "tcPrChange",
        "cellIns",
        "sectPrChange"
      ])
    );
    expect(new Set(found.map(r => r.id)).size).toBe(found.length);
  });

  it("does so in a streamed package too", async () => {
    const doc = mixed();
    const writer = Streaming.createWriter({
      styles: doc.styles,
      sectionProperties: doc.sectionProperties
    });
    for (const block of doc.body) {
      writer.add(block);
    }
    const found = await revisionIds(await writer.finalize());
    expect(found.length).toBeGreaterThan(9);
    expect(new Set(found.map(r => r.id)).size).toBe(found.length);
  });

  it("leaves accept and reject unaffected", async () => {
    // They work on the model, so a document read back accepts and rejects
    // exactly as the model it was written from does.
    const tags = async (doc: DocxDocument) =>
      (await revisionIds(await Io.toBuffer(doc))).map(r => r.tag);
    for (const process of [Query.acceptAllRevisions, Query.rejectAllRevisions]) {
      const original = mixed();
      const read = await Io.read(await Io.toBuffer(mixed()));
      expect(process(read)).toBe(process(original));
      expect(await tags(read)).toEqual(await tags(original));
    }
    const rejected = await Io.read(await Io.toBuffer(mixed()));
    Query.rejectAllRevisions(rejected);
    const xml = (await parts(await Io.toBuffer(rejected))).get("word/document.xml")!;
    expect(xml).toContain(" removed");
    expect(xml).not.toContain(" added");
  });
});

describe("revision ids: one id space per package", () => {
  const ids = async (bytes: Uint8Array) => (await packageWrappers(bytes)).map(w => w.id);
  const tracked = (text: string) =>
    Build.paragraph([Build.insertedRun(Build.text(text), { id: 1, author: "Ann" })]);

  it("keeps interleaved streamed packages apart", async () => {
    const a = Streaming.createWriter();
    const b = Streaming.createWriter();
    a.add(tracked("a1"));
    b.add(tracked("b1"));
    a.add(tracked("a2"));
    b.add(tracked("b2"));
    const [bytesA, bytesB] = await Promise.all([a.finalize(), b.finalize()]);
    expect(await ids(bytesA)).toEqual(["1", "2"]);
    expect(await ids(bytesB)).toEqual(["1", "2"]);
  });

  it("keeps concurrent package writes apart", async () => {
    const make = () => {
      const doc = Document.create();
      Document.addContent(doc, tracked("x"));
      Document.addContent(doc, tracked("y"));
      return doc;
    };
    const written = await Promise.all([make(), make(), make()].map(doc => Io.toBuffer(doc)));
    for (const bytes of written) {
      expect(await ids(bytes)).toEqual(["1", "2"]);
    }
  });
});
