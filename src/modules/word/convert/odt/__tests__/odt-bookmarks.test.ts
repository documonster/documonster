/**
 * ODT bookmarks: the reader numbers ranges uniquely and pairs ends by
 * `text:name`; the writer pairs ends with their starts even when the model
 * reuses an id.
 */
import { unzip } from "@archive/read-archive";
import { describe, it, expect } from "vitest";

import { Build, Convert, Document } from "../../../index";
import type { DocxDocument, Paragraph, ParagraphChild } from "../../../index";

async function contentXml(bytes: Uint8Array): Promise<string> {
  for await (const entry of unzip(bytes).entries()) {
    if (entry.path === "content.xml") {
      return new TextDecoder().decode(await entry.bytes());
    }
  }
  throw new Error("no content.xml");
}

function markers(xml: string): string[] {
  return [...xml.matchAll(/<text:bookmark(-start|-end)? text:name="([^"]*)"/g)].map(
    m => `${m[1] === "-start" ? "S" : m[1] === "-end" ? "E" : "P"}:${m[2]}`
  );
}

function bookmarkChildren(doc: DocxDocument): ParagraphChild[] {
  return doc.body.flatMap(b =>
    b.type === "paragraph"
      ? b.children.filter(
          c => "type" in c && (c.type === "bookmarkStart" || c.type === "bookmarkEnd")
        )
      : []
  );
}

describe("ODT bookmarks", () => {
  it("writes the right name on each end when the model reuses an id", async () => {
    const doc = Document.create();
    Document.addParagraph(
      doc,
      Build.paragraph([
        Build.bookmarkStart(0, "outer"),
        Build.text("a"),
        Build.bookmarkStart(0, "inner"),
        Build.text("b"),
        Build.bookmarkEnd(0),
        Build.bookmarkEnd(0)
      ])
    );
    Document.addParagraph(
      doc,
      Build.paragraph([
        Build.bookmarkStart(5, "x"),
        Build.bookmarkStart(5, "y"),
        Build.bookmarkEnd(5, { name: "x" }),
        Build.bookmarkEnd(5, { name: "y" }),
        // A stray end closes nothing and is not written.
        Build.bookmarkEnd(9)
      ])
    );
    expect(markers(await contentXml(await Convert.toOdt(doc)))).toEqual([
      "S:outer",
      "S:inner",
      "E:inner",
      "E:outer",
      "S:x",
      "S:y",
      "E:x",
      "E:y"
    ]);
  });

  it("reads unique ids and pairs ends by name, including overlapping ranges", async () => {
    const doc = Document.create();
    Document.addParagraph(
      doc,
      Build.paragraph([
        Build.bookmarkStart(0, "first"),
        Build.text("a"),
        Build.bookmarkStart(1, "second"),
        Build.text("b"),
        Build.bookmarkEnd(0, { name: "first" }),
        Build.text("c"),
        Build.bookmarkEnd(1, { name: "second" })
      ])
    );
    const read = await Convert.fromOdt(await Convert.toOdt(doc));
    const marks = bookmarkChildren(read) as { type: string; id: number; name?: string }[];
    expect(marks.map(m => m.type)).toEqual([
      "bookmarkStart",
      "bookmarkStart",
      "bookmarkEnd",
      "bookmarkEnd"
    ]);
    const [s1, s2, e1, e2] = marks;
    expect(s1!.id).not.toBe(s2!.id);
    expect(e1).toMatchObject({ id: s1!.id, name: "first" });
    expect(e2).toMatchObject({ id: s2!.id, name: "second" });
  });

  it("round-trips through DOCX with every range intact", async () => {
    const doc = Document.create();
    for (const name of ["a", "b", "c"]) {
      Document.addParagraph(
        doc,
        Build.paragraph([Build.bookmarkStart(0, name), Build.text(name), Build.bookmarkEnd(0)])
      );
    }
    const read = await Convert.fromOdt(await Convert.toOdt(doc));
    const starts = bookmarkChildren(read).filter(
      (c): c is Extract<ParagraphChild, { type: "bookmarkStart" }> =>
        "type" in c && c.type === "bookmarkStart"
    );
    expect(new Set(starts.map(s => s.id)).size).toBe(3);
    const xml = await contentXml(await Convert.toOdt(read));
    expect(markers(xml)).toEqual(["S:a", "E:a", "S:b", "E:b", "S:c", "E:c"]);
  });

  it("reads a point bookmark as a start and an end", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "x");
    const odt = await Convert.toOdt(doc);
    // Splice a point bookmark into the written content.
    const entries = new Map<string, Uint8Array>();
    for await (const entry of unzip(odt).entries()) {
      entries.set(entry.path, await entry.bytes());
    }
    const xml = new TextDecoder()
      .decode(entries.get("content.xml"))
      .replace(/(<text:p[^>]*>)/, '$1<text:bookmark text:name="here"/>');
    const { zip } = await import("@archive/create-archive");
    const archive = zip({ noSort: true });
    for (const [path, data] of entries) {
      archive.add(path, path === "content.xml" ? new TextEncoder().encode(xml) : data);
    }
    const read = await Convert.fromOdt(await archive.toBuffer());
    const para = read.body[0] as Paragraph;
    expect(para.children.slice(0, 2)).toEqual([
      { type: "bookmarkStart", id: 0, name: "here" },
      { type: "bookmarkEnd", id: 0, name: "here" }
    ]);
  });
});
