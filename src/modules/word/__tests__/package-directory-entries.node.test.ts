/**
 * ZIP directory entries are not OPC parts.
 *
 * A ZIP may carry an entry for a folder (`word/`, `docProps/`) — Word does not
 * write them, but repackers do. A part name must not end with "/" (ECMA-376
 * Part 2 §9.1.1.1), so treating such an entry as a part and declaring it in
 * `[Content_Types].xml` produces a package Word refuses as unreadable.
 */

import { readFileSync } from "node:fs";
import { join } from "node:path";

import { unzip } from "@archive/read-archive";
import { ZipArchive } from "@archive/zip/zip-archive";
import { DocxWriteError } from "@word/errors";
import { Build, Document, Io, Package } from "@word/index";
import { createContentTypes, typeUntypedParts } from "@word/writer/content-types";
import { describe, expect, it } from "vitest";

const FIXTURE = join(__dirname, "data", "word-native", "toc-word-zu.docx");

/** Repack a package with an explicit directory entry for every folder in it. */
async function withDirectoryEntries(bytes: Uint8Array): Promise<Uint8Array> {
  const files = new Map<string, Uint8Array>();
  for await (const entry of unzip(bytes).entries()) {
    files.set(entry.path, await entry.bytes());
  }
  const archive = new ZipArchive();
  const folders = new Set<string>();
  for (const path of files.keys()) {
    const segments = path.split("/");
    for (let i = 1; i < segments.length; i++) {
      folders.add(`${segments.slice(0, i).join("/")}/`);
    }
  }
  for (const folder of folders) {
    archive.addDirectory(folder);
  }
  for (const [path, data] of files) {
    // A folder the source already carries was added above; a ZIP holds one entry per name.
    if (!folders.has(path)) {
      archive.add(path, data);
    }
  }
  return archive.toBuffer();
}

async function inventory(bytes: Uint8Array): Promise<{ names: string[]; contentTypes: string }> {
  const names: string[] = [];
  let contentTypes = "";
  for await (const entry of unzip(bytes).entries()) {
    names.push(entry.path);
    if (entry.path === "[Content_Types].xml") {
      contentTypes = new TextDecoder().decode(await entry.bytes());
    }
  }
  return { names, contentTypes };
}

function expectNoDirectoryParts(result: { names: string[]; contentTypes: string }): void {
  expect(result.names.filter(n => n.endsWith("/"))).toEqual([]);
  expect(result.contentTypes).not.toMatch(/PartName="[^"]*\/"/);
  expect(result.contentTypes).toContain('PartName="/word/document.xml"');
}

describe("ZIP directory entries in a DOCX", () => {
  it("the fixture really carries directory entries", async () => {
    const input = await inventory(await withDirectoryEntries(readFileSync(FIXTURE)));
    expect(input.names).toEqual(expect.arrayContaining(["word/", "docProps/", "word/theme/"]));
  });

  it("Io.read → add paragraphs → Io.toBuffer writes no directory entry and declares none", async () => {
    const doc = await Io.read(await withDirectoryEntries(readFileSync(FIXTURE)));
    Document.addParagraph(doc, Build.paragraph([Build.text("added")]));
    expectNoDirectoryParts(await inventory(await Io.toBuffer(doc)));
  });

  it("Package.edit neither copies nor declares directory entries", async () => {
    const edited = await Package.edit(await withDirectoryEntries(readFileSync(FIXTURE)), [
      { type: "replacePartText", path: "customXml/new.xml", text: "<a/>" }
    ]);
    expectNoDirectoryParts(await inventory(edited));
  });

  it("Package.listParts lists parts, not folders", async () => {
    const parts = await Package.listParts(await withDirectoryEntries(readFileSync(FIXTURE)));
    expect(parts.filter(p => p.endsWith("/"))).toEqual([]);
    expect(parts).toContain("word/document.xml");
  });

  it("content-type completion refuses a name that is not a valid part name", () => {
    const state = createContentTypes();
    expect(() => typeUntypedParts(state, ["word/"])).toThrow(DocxWriteError);
    expect(() => typeUntypedParts(state, ["word//a.bin"])).toThrow(DocxWriteError);
    expect(state.overrides).toEqual([]);
  });
});
