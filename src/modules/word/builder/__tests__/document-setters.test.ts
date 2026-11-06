/**
 * `Document.set*` on a document read from a file: what the caller does not
 * pass is not silently dropped, and a header or footer set by type is the one
 * the section shows.
 */
import { describe, it, expect } from "vitest";

import { Build, Document, DocxArgumentError, Io, Layout, Query, Security } from "../../index";
import type { DocxDocument } from "../../index";

async function readBack(doc: DocxDocument): Promise<DocxDocument> {
  return Io.read(await Io.toBuffer(doc));
}

async function readDocWithHeaderAndFooter(): Promise<DocxDocument> {
  const doc = Document.create();
  Document.addParagraph(doc, "body");
  Document.setHeader(doc, "default", { children: [Build.textParagraph("OLD HEADER")] });
  Document.setFooter(doc, "default", { children: [Build.textParagraph("OLD FOOTER")] });
  return readBack(doc);
}

function pageSvg(doc: DocxDocument, page = 1): string {
  return Layout.pageToSvg(doc, page);
}

describe("Document.setHeader / setFooter", () => {
  it("replaces the header the section shows when given its type on a read document", async () => {
    const doc = await readDocWithHeaderAndFooter();
    expect([...doc.headers!.keys()]).not.toContain("default");
    Document.setHeader(doc, "default", { children: [Build.textParagraph("NEW HEADER")] });
    Document.setFooter(doc, "default", { children: [Build.textParagraph("NEW FOOTER")] });
    expect(doc.headers!.size).toBe(1);
    expect(doc.footers!.size).toBe(1);
    const again = await readBack(doc);
    expect(pageSvg(again)).toContain("NEW HEADER");
    expect(pageSvg(again)).toContain("NEW FOOTER");
    expect(pageSvg(again)).not.toContain("OLD HEADER");
  });

  it("wires a header type the read section does not reference yet", async () => {
    const doc = await readDocWithHeaderAndFooter();
    Document.setHeader(doc, "first", { children: [Build.textParagraph("FIRST PAGE")] });
    const again = await readBack(doc);
    expect(again.sectionProperties?.titlePage).toBe(true);
    expect(pageSvg(again)).toContain("FIRST PAGE");
    expect(again.sectionProperties?.headers?.map(h => h.type).sort()).toEqual(["default", "first"]);
  });

  it("turns on the title page for a first-page header on a built document", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "body");
    Document.setHeader(doc, "first", { children: [Build.textParagraph("FIRST")] });
    expect(doc.sectionProperties?.titlePage).toBe(true);
    expect(pageSvg(await readBack(doc))).toContain("FIRST");
  });

  it("turns on even/odd headers for an even-page header", () => {
    const doc = Document.create();
    Document.setFooter(doc, "even", { children: [Build.textParagraph("EVEN")] });
    expect(doc.settings?.evenAndOddHeaders).toBe(true);
  });

  it("still replaces a header by the key Io.read gave it", async () => {
    const doc = await readDocWithHeaderAndFooter();
    const [key] = doc.headers!.keys();
    Document.setHeader(doc, key!, { children: [Build.textParagraph("BY KEY")] });
    expect(pageSvg(await readBack(doc))).toContain("BY KEY");
  });

  it("refuses a key that is neither a type nor a header the document has", async () => {
    const doc = await readDocWithHeaderAndFooter();
    expect(() =>
      Document.setHeader(doc, "bogus", { children: [Build.textParagraph("x")] })
    ).toThrow(DocxArgumentError);
    expect(() =>
      Document.setFooter(Document.create(), "rId99", { children: [Build.textParagraph("x")] })
    ).toThrow(DocxArgumentError);
  });
});

describe("Document.setSectionProperties", () => {
  it("keeps the header and footer references when none are given", async () => {
    const doc = await readDocWithHeaderAndFooter();
    Document.setSectionProperties(doc, { pageSize: { width: 11906, height: 16838 } });
    const again = await readBack(doc);
    expect(again.sectionProperties?.pageSize?.width).toBe(11906);
    expect(pageSvg(again)).toContain("OLD HEADER");
    expect(pageSvg(again)).toContain("OLD FOOTER");
  });

  it("uses the references it is given, including none", async () => {
    const doc = await readDocWithHeaderAndFooter();
    Document.setSectionProperties(doc, { headers: [], footers: [] });
    expect(doc.sectionProperties?.headers).toEqual([]);
  });
});

describe("Document.setSettings", () => {
  it("keeps what a settings object cannot say unless it says it", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "x");
    Document.setCompatibilityMode(doc, 14);
    const protectedDoc = await Security.protect(doc, {
      edit: "readOnly",
      password: "pw",
      spinCount: 1
    });
    const read = await readBack(protectedDoc);
    const preserved = read.settings?.preservedXml;

    Document.setSettings(read, { zoom: 120 });
    expect(read.settings?.zoom).toBe(120);
    expect(Query.getCompatibilityMode(read)).toBe(14);
    expect(Security.isProtected(read)).toBe(true);
    expect(read.settings?.preservedXml).toEqual(preserved);

    Document.setSettings(read, { zoom: 90, compatibilityMode: 15, documentProtection: undefined });
    expect(Query.getCompatibilityMode(read)).toBe(15);
    expect(Security.isProtected(read)).toBe(false);
  });

  it("replaces the fields a settings object can say", () => {
    const doc = Document.create();
    Document.setSettings(doc, { zoom: 150, trackRevisions: true });
    Document.setSettings(doc, { zoom: 90 });
    expect(doc.settings).toEqual({ zoom: 90 });
  });
});
