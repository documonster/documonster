/**
 * `Transform.merge` keeps every appended document's own headers, footers and
 * picture bullets: each section of the result shows what it showed in the
 * document it came from, and every list keeps counting as it did there.
 */
import { unzip } from "@archive/read-archive";
import { describe, expect, it } from "vitest";

import { Build, Document, Io, Transform } from "../index";
import { WordMarkdown } from "../markdown";
import type {
  BodyContent,
  DocxDocument,
  DrawingShape,
  HeaderFooterRef,
  Paragraph,
  SectionProperties
} from "../types";

const decoder = new TextDecoder();

function png(letter: string): Uint8Array {
  return new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, letter.charCodeAt(0)]);
}

async function entries(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for await (const entry of unzip(bytes).entries()) {
    out.set(entry.path.replace(/^\/+/, ""), await entry.bytes());
  }
  return out;
}

function text(blocks: readonly unknown[]): string {
  return (
    JSON.stringify(blocks)
      .match(/"text":"([^"]*)"/g)
      ?.map(m => m.slice(8, -1))
      .join("") ?? ""
  );
}

/** Each section's sectPr, in document order: inner ones, then the final one. */
function sections(doc: DocxDocument): SectionProperties[] {
  const out: SectionProperties[] = [];
  for (const block of doc.body) {
    if (block.type === "paragraph" && block.properties?.sectionProperties) {
      out.push(block.properties.sectionProperties);
    }
  }
  out.push(doc.sectionProperties ?? {});
  return out;
}

/**
 * The text of the header (or footer) of `type` each section shows, applying
 * OOXML inheritance: a section without a reference of that type shows its
 * predecessor's.
 */
function shown(
  doc: DocxDocument,
  part: "headers" | "footers",
  type: HeaderFooterRef["type"]
): string[] {
  let current = "";
  return sections(doc).map(sect => {
    const ref = sect[part]?.find(r => r.type === type);
    if (ref) {
      const def = doc[part]?.get(ref.rId);
      current = def ? text(def.content.children) : "?";
    }
    return current;
  });
}

function letterhead(name: string, logo?: string): DocxDocument {
  const doc = Document.create();
  const children: Paragraph[] = [Build.textParagraph(`Header ${name}`)];
  if (logo) {
    const key = `rIdLogo${name}`;
    (doc as { images?: unknown }).images = [
      { data: png(logo), mediaType: "png", fileName: "logo.png", rId: key }
    ];
    children.push(
      Build.paragraph([{ content: [{ type: "image", rId: key, width: 9525, height: 9525 }] }])
    );
  }
  Document.setHeader(doc, "default", { children });
  Document.setFooter(doc, "default", { children: [Build.textParagraph(`Footer ${name}`)] });
  Document.addParagraph(doc, `Body ${name}`);
  return doc;
}

describe("Transform.merge keeps each document's headers and footers", () => {
  it("every section shows the header and footer of the document it came from", async () => {
    const merged = Transform.merge([letterhead("A"), letterhead("B"), letterhead("C")]);
    const read = await Io.read(await Io.toBuffer(merged));
    expect(shown(read, "headers", "default")).toEqual(["Header A", "Header B", "Header C"]);
    expect(shown(read, "footers", "default")).toEqual(["Footer A", "Footer B", "Footer C"]);
  });

  it("a document without headers does not inherit the previous one's", async () => {
    const plain = Document.create();
    Document.addParagraph(plain, "Body plain");
    const merged = Transform.merge([letterhead("A"), plain, letterhead("C")]);
    const read = await Io.read(await Io.toBuffer(merged));
    expect(shown(read, "headers", "default")).toEqual(["Header A", "", "Header C"]);
    expect(shown(read, "footers", "default")).toEqual(["Footer A", "", "Footer C"]);
  });

  it("sections within an appended document keep inheriting as they did there", async () => {
    // Standalone, B's first section names no header (so shows none) and its
    // final section gets the parts keyed by type.
    const b = letterhead("B");
    Document.addSectionBreak(b, {});
    Document.addParagraph(b, "Body B, second section");
    const standalone = await Io.read(await Io.toBuffer(b));
    expect(shown(standalone, "headers", "default")).toEqual(["", "Header B"]);

    const merged = Transform.merge([letterhead("A"), b]);
    const read = await Io.read(await Io.toBuffer(merged));
    expect(shown(read, "headers", "default")).toEqual(["Header A", "", "Header B"]);
  });

  it("each header keeps its own logo", async () => {
    const merged = Transform.merge([letterhead("A", "a"), letterhead("B", "b")]);
    const files = await entries(await Io.toBuffer(merged));
    const logos: string[] = [];
    for (const [path, data] of files) {
      const m = /^word\/(header\d+)\.xml$/.exec(path);
      if (!m) {
        continue;
      }
      const xml = decoder.decode(data);
      const rels = decoder.decode(files.get(`word/_rels/${m[1]}.xml.rels`) ?? new Uint8Array());
      for (const [, id] of xml.matchAll(/r:embed="([^"]+)"/g)) {
        const target = new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(rels)?.[1];
        const media = target ? files.get(`word/${target}`) : undefined;
        logos.push(media ? String.fromCharCode(media[media.length - 1]!) : "?");
      }
    }
    expect(logos.sort()).toEqual(["a", "b"]);
  });
});

describe("Transform.merge numbering", () => {
  function bulleted(letter: string): DocxDocument {
    const doc = Document.create();
    const key = `rIdBullet${letter}`;
    (doc as { images?: unknown }).images = [
      { data: png(letter), mediaType: "png", fileName: "bullet.png", rId: key }
    ];
    (doc as { numPicBullets?: unknown }).numPicBullets = [
      { id: 0, rId: key, width: 9525, height: 9525 }
    ];
    (doc as { abstractNumberings?: unknown }).abstractNumberings = [
      {
        abstractNumId: 0,
        levels: [{ level: 0, start: 1, format: "bullet", text: "\uF0B7", picBulletId: 0 }]
      }
    ];
    (doc as { numberingInstances?: unknown }).numberingInstances = [{ numId: 1, abstractNumId: 0 }];
    Document.addParagraph(
      doc,
      Build.textParagraph(`item ${letter}`, { numbering: { numId: 1, level: 0 } })
    );
    return doc;
  }

  it("picture bullets keep their own picture", async () => {
    const merged = Transform.merge([bulleted("x"), bulleted("y")]);
    expect(merged.numPicBullets?.map(b => b.id)).toEqual([0, 1]);
    const levels = merged.abstractNumberings!.map(a => a.levels[0]!.picBulletId);
    expect(levels).toEqual([0, 1]);

    const files = await entries(await Io.toBuffer(merged));
    const numbering = decoder.decode(files.get("word/numbering.xml")!);
    const rels = decoder.decode(files.get("word/_rels/numbering.xml.rels")!);
    const pictures = [
      ...numbering.matchAll(/<w:numPicBullet w:numPicBulletId="(\d+)">.*?r:id="([^"]+)"/g)
    ].map(([, bullet, id]) => {
      const target = new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(rels)?.[1];
      const media = target ? files.get(`word/${target}`) : undefined;
      return `${bullet}:${media ? String.fromCharCode(media[media.length - 1]!) : "?"}`;
    });
    expect(pictures).toEqual(["0:x", "1:y"]);
  });

  it("each document's lists keep counting as they did in it", async () => {
    const numbered = (prefix: string) => {
      const doc = Document.create();
      Document.addNumberedList(doc, [`${prefix} one`, `${prefix} two`]);
      Document.addParagraph(doc, "between");
      Document.addNumberedList(doc, [`${prefix} again`]);
      return doc;
    };
    const merged = Transform.merge([numbered("A"), numbered("B")]);
    const markdown = WordMarkdown.fromDocx(await Io.read(await Io.toBuffer(merged)));
    const items = markdown
      .split("\n")
      .filter(line => /^\d+\. /.test(line))
      .map(line => line.replace(/\s+/g, " "));
    expect(items).toEqual([
      "1. A one",
      "2. A two",
      "1. A again",
      "1. B one",
      "2. B two",
      "1. B again"
    ]);
  });
});

describe("the writer resolves header references by model key", () => {
  it("a key that equals a freshly assigned relationship id still names its own part", async () => {
    // The writer numbers header relationships afresh (rId5, rId6, …). A model
    // key that happens to equal one of those ids must still select the part it
    // keys, not the part that received that id.
    const doc = Document.create();
    Document.addParagraph(doc, "body");
    const headers = new Map([
      ["rId6", { content: { children: [Build.textParagraph("Header A")] } }],
      ["rId5", { content: { children: [Build.textParagraph("Header B")] } }]
    ]);
    const withRefs: DocxDocument = {
      ...doc,
      headers,
      sectionProperties: {
        ...doc.sectionProperties,
        titlePage: true,
        headers: [
          { type: "default", rId: "rId5" },
          { type: "first", rId: "rId6" }
        ]
      }
    };
    const read = await Io.read(await Io.toBuffer(withRefs));
    expect(shown(read, "headers", "default")).toEqual(["Header B"]);
    expect(shown(read, "headers", "first")).toEqual(["Header A"]);
  });
});

describe("Transform.merge renames ids inside markup kept verbatim", () => {
  function source(tag: string): DocxDocument {
    const doc = Document.create();
    Document.addNumberedList(doc, [`list ${tag}`]);
    Document.addComment(doc, tag, `comment ${tag}`);
    Document.addFootnote(doc, `footnote ${tag}`);
    Document.addEndnote(doc, `endnote ${tag}`);
    return doc;
  }

  /** A text box whose content is kept verbatim (it holds a table) and refers to the source's ids. */
  function textBox(numId: number, commentId: number, noteId: number): BodyContent {
    const xml =
      `<w:txbxContent><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>` +
      `<w:commentRangeStart w:id="${commentId}"/><w:r><w:t>Item</w:t></w:r><w:commentRangeEnd w:id="${commentId}"/>` +
      `<w:r><w:commentReference w:id="${commentId}"/></w:r>` +
      `<w:r><w:footnoteReference w:id="${noteId}"/></w:r><w:r><w:endnoteReference w:id="${noteId}"/></w:r>` +
      `<w:bookmarkStart w:id="0" w:name="box${numId}"/><w:permStart w:id="0" w:edGrp="everyone"/>` +
      `<w:r><w:t>marked</w:t></w:r><w:permEnd w:id="0"/><w:bookmarkEnd w:id="0"/></w:p>` +
      `<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:txbxContent>`;
    return {
      type: "drawingShape",
      shapeType: "rect",
      width: 914400,
      height: 914400,
      _preserved: {
        read: { shapeType: "rect", width: 914400, height: 914400 },
        txbxContentXml: xml
      }
    };
  }

  it("lists, comments and notes referred to from a text box follow their renamed definitions", async () => {
    const a = source("A");
    Document.addContent(a, textBox(1, 0, 1));
    const b = source("B");
    const listB = b.body.find(block => block.type === "paragraph")! as Paragraph;
    const numIdB = listB.properties!.numbering!.numId;
    Document.addContent(b, textBox(numIdB, b.comments![0]!.id, b.footnotes![0]!.id));

    const merged = Transform.merge([a, b]);
    const shape = merged.body.filter(block => block.type === "drawingShape")[1] as DrawingShape;
    const markup = shape._preserved!.txbxContentXml!;

    // B's ids, as the merge renamed them for B's modelled content.
    const mergedList = merged.body.filter(
      (block): block is Paragraph => block.type === "paragraph" && text([block]) === "list B"
    )[0]!;
    const numId = mergedList.properties!.numbering!.numId;
    const comment = merged.comments!.find(c => c.author === "B")!.id;
    const footnote = merged.footnotes!.find(f => text(f.content) === "footnote B")!.id;
    const endnote = merged.endnotes!.find(e => text(e.content) === "endnote B")!.id;
    expect([numId, comment, footnote, endnote]).not.toEqual([numIdB, 0, 1, 1]);

    expect(markup).toContain(`<w:numId w:val="${numId}"/>`);
    expect(markup).toContain(`<w:commentRangeStart w:id="${comment}"/>`);
    expect(markup).toContain(`<w:commentRangeEnd w:id="${comment}"/>`);
    expect(markup).toContain(`<w:commentReference w:id="${comment}"/>`);
    expect(markup).toContain(`<w:footnoteReference w:id="${footnote}"/>`);
    expect(markup).toContain(`<w:endnoteReference w:id="${endnote}"/>`);
    // The base document's own ids are not touched where they mean the base.
    expect(merged.comments!.find(c => c.author === "A")!.id).toBe(0);

    // Bookmark and permission ranges kept verbatim in both documents no longer share ids,
    // and each still closes the range it opened.
    const bookmark = /<w:bookmarkStart w:id="(\d+)"/.exec(markup)![1];
    expect(bookmark).not.toBe("0");
    expect(markup).toContain(`<w:bookmarkEnd w:id="${bookmark}"/>`);
    const perm = /<w:permStart w:id="(\d+)"/.exec(markup)![1];
    expect(perm).not.toBe("0");
    expect(markup).toContain(`<w:permEnd w:id="${perm}"/>`);

    // And the written package agrees.
    const written = decoder.decode(
      (await entries(await Io.toBuffer(merged))).get("word/document.xml")!
    );
    expect(written).toContain(`<w:footnoteReference w:id="${footnote}"/>`);
    const starts = [...written.matchAll(/<w:bookmarkStart w:id="(\d+)"/g)].map(m => m[1]);
    expect(new Set(starts).size).toBe(starts.length);
  });
});
