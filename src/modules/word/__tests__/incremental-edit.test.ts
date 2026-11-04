/**
 * DOCX Module - Incremental Edit Tests
 */

import { describe, it, expect } from "vitest";

import { DocxArgumentError, DocxParseError } from "../errors";
import { Build, Convert, Document, Io, Package } from "../index";
import type { Paragraph, Run } from "../types";

async function buildDocxBuffer(text: string): Promise<Uint8Array> {
  const doc = Document.create();
  Document.addContent(doc, Build.textParagraph(text));
  return Io.toBuffer(doc);
}

function paraText(p: Paragraph): string {
  let t = "";
  for (const child of p.children) {
    if ("content" in child && Array.isArray((child as Run).content)) {
      for (const c of (child as Run).content) {
        if (c.type === "text") {
          t += c.text;
        }
      }
    }
  }
  return t;
}

describe("listDocxParts", () => {
  it("lists all parts in a DOCX", async () => {
    const buffer = await buildDocxBuffer("test");
    const parts = await Package.listParts(buffer);

    expect(parts.length).toBeGreaterThan(0);
    expect(parts).toContain("[Content_Types].xml");
    expect(parts).toContain("word/document.xml");
    expect(parts).toContain("_rels/.rels");
  });
});

describe("readDocxPart", () => {
  it("reads document.xml content", async () => {
    const buffer = await buildDocxBuffer("hello world");
    const data = await Package.readPart(buffer, "word/document.xml");
    expect(data).toBeDefined();
    const xml = new TextDecoder().decode(data!);
    expect(xml).toContain("<?xml");
    expect(xml).toContain("hello world");
  });

  it("returns undefined for missing part", async () => {
    const buffer = await buildDocxBuffer("test");
    const data = await Package.readPart(buffer, "word/no-such-part.xml");
    expect(data).toBeUndefined();
  });

  it("reads [Content_Types].xml", async () => {
    const buffer = await buildDocxBuffer("test");
    const data = await Package.readPart(buffer, "[Content_Types].xml");
    expect(data).toBeDefined();
  });
});

describe("editDocxIncremental", () => {
  it("replaces a part with raw bytes", async () => {
    const buffer = await buildDocxBuffer("original");
    const newXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
    const edited = await Package.edit(buffer, [
      {
        type: "replacePart",
        path: "[Content_Types].xml",
        data: new TextEncoder().encode(newXml)
      }
    ]);

    expect(edited).toBeInstanceOf(Uint8Array);
    const ct = await Package.readPart(edited, "[Content_Types].xml");
    expect(new TextDecoder().decode(ct!)).toBe(newXml);
  });

  it("replaces a part with text", async () => {
    const buffer = await buildDocxBuffer("test");
    // Replace content types - use a minimal valid XML
    const newXml =
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';
    const edited = await Package.edit(buffer, [
      { type: "replacePartText", path: "[Content_Types].xml", text: newXml }
    ]);

    const ct = await Package.readPart(edited, "[Content_Types].xml");
    expect(new TextDecoder().decode(ct!)).toBe(newXml);
  });

  it("deletes a part", async () => {
    const buffer = await buildDocxBuffer("test");
    const partsBefore = await Package.listParts(buffer);

    // Add a part first by editing, then delete it
    const withExtra = await Package.edit(buffer, [
      { type: "replacePartText", path: "extra.txt", text: "hello" }
    ]);

    const withExtraParts = await Package.listParts(withExtra);
    expect(withExtraParts).toContain("extra.txt");

    const deleted = await Package.edit(withExtra, [{ type: "deletePart", path: "extra.txt" }]);

    const finalParts = await Package.listParts(deleted);
    expect(finalParts).not.toContain("extra.txt");
    expect(finalParts.length).toBe(partsBefore.length);
  });

  it("preserves unchanged parts", async () => {
    const buffer = await buildDocxBuffer("preserve me");
    const stylesXmlBefore = await Package.readPart(buffer, "word/styles.xml");
    const ctBefore = await Package.readPart(buffer, "[Content_Types].xml");

    // A part the package already types (`xml` has a Default), so the edit has nothing to declare;
    // an untyped one would gain a Default — see image-media-type.test.ts.
    const edited = await Package.edit(buffer, [
      { type: "replacePartText", path: "extra.xml", text: "<new/>" }
    ]);

    const stylesAfter = await Package.readPart(edited, "word/styles.xml");
    const ctAfter = await Package.readPart(edited, "[Content_Types].xml");

    if (stylesXmlBefore && stylesAfter) {
      expect(new TextDecoder().decode(stylesAfter)).toBe(new TextDecoder().decode(stylesXmlBefore));
    }
    if (ctBefore && ctAfter) {
      expect(new TextDecoder().decode(ctAfter)).toBe(new TextDecoder().decode(ctBefore));
    }
  });

  it("replaces document body and resulting DOCX is readable", async () => {
    const buffer = await buildDocxBuffer("original text");

    const newPara: Paragraph = {
      type: "paragraph",
      children: [{ content: [{ type: "text", text: "REPLACED CONTENT" }] } as Run]
    };

    const edited = await Package.edit(buffer, [{ type: "replaceBody", body: [newPara] }]);

    expect(edited).toBeInstanceOf(Uint8Array);
    // Verify by reading back
    const parsed = await Io.read(edited);
    let foundReplaced = false;
    let foundOriginal = false;
    for (const block of parsed.body) {
      if (block.type === "paragraph") {
        const text = paraText(block);
        if (text.includes("REPLACED")) {
          foundReplaced = true;
        }
        if (text.includes("original")) {
          foundOriginal = true;
        }
      }
    }
    expect(foundReplaced).toBe(true);
    expect(foundOriginal).toBe(false);
  });

  it("supports compression level option", async () => {
    const buffer = await buildDocxBuffer("compress me");
    const edited = await Package.edit(
      buffer,
      [{ type: "replacePartText", path: "extra.txt", text: "x" }],
      { compressionLevel: 9 }
    );
    expect(edited).toBeInstanceOf(Uint8Array);
  });

  it("applies multiple edits in order", async () => {
    const buffer = await buildDocxBuffer("test");
    const edited = await Package.edit(buffer, [
      { type: "replacePartText", path: "first.txt", text: "1" },
      { type: "replacePartText", path: "second.txt", text: "2" },
      { type: "replacePartText", path: "first.txt", text: "1-updated" }
    ]);

    const first = await Package.readPart(edited, "first.txt");
    const second = await Package.readPart(edited, "second.txt");

    expect(new TextDecoder().decode(first!)).toBe("1-updated");
    expect(new TextDecoder().decode(second!)).toBe("2");
  });
});

describe("editDocxIncremental — replaceHeader / replaceFooter", () => {
  it("replaces a header part with plain-text content", async () => {
    const buffer = await buildDocxBuffer("body");
    // Inject a placeholder header so we have a header part to replace.
    const seedHeader =
      '<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>';
    const seeded = await Package.edit(buffer, [
      { type: "replacePartText", path: "word/header1.xml", text: seedHeader }
    ]);

    const newPara: Paragraph = {
      type: "paragraph",
      children: [{ content: [{ type: "text", text: "Header replaced" }] } as Run]
    };
    const edited = await Package.edit(seeded, [
      { type: "replaceHeader", path: "word/header1.xml", children: [newPara] }
    ]);

    const headerXml = await Package.readPart(edited, "word/header1.xml");
    expect(headerXml).toBeDefined();
    const decoded = new TextDecoder().decode(headerXml!);
    expect(decoded).toContain("<w:hdr");
    expect(decoded).toContain("Header replaced");
  });

  it("replaces a footer part with plain-text content", async () => {
    const buffer = await buildDocxBuffer("body");
    const seedFooter =
      '<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:ftr>';
    const seeded = await Package.edit(buffer, [
      { type: "replacePartText", path: "word/footer1.xml", text: seedFooter }
    ]);

    const newPara: Paragraph = {
      type: "paragraph",
      children: [{ content: [{ type: "text", text: "Footer replaced" }] } as Run]
    };
    const edited = await Package.edit(seeded, [
      { type: "replaceFooter", path: "word/footer1.xml", children: [newPara] }
    ]);

    const footerXml = await Package.readPart(edited, "word/footer1.xml");
    expect(footerXml).toBeDefined();
    const decoded = new TextDecoder().decode(footerXml!);
    expect(decoded).toContain("<w:ftr");
    expect(decoded).toContain("Footer replaced");
  });

  it("rejects header content with an inline image (.rels not rewritten)", async () => {
    const buffer = await buildDocxBuffer("body");
    const seed =
      '<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>';
    const seeded = await Package.edit(buffer, [
      { type: "replacePartText", path: "word/header1.xml", text: seed }
    ]);

    const paraWithImage: Paragraph = {
      type: "paragraph",
      children: [
        {
          content: [
            {
              type: "image",
              rId: "rId99",
              width: 100000,
              height: 100000
            }
          ]
        } as Run
      ]
    };

    await expect(
      Package.edit(seeded, [
        { type: "replaceHeader", path: "word/header1.xml", children: [paraWithImage] }
      ])
    ).rejects.toThrow(/inline image/);
  });

  it("rejects footer content with a hyperlink (.rels not rewritten)", async () => {
    const buffer = await buildDocxBuffer("body");
    const seed =
      '<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:ftr>';
    const seeded = await Package.edit(buffer, [
      { type: "replacePartText", path: "word/footer1.xml", text: seed }
    ]);

    const paraWithLink: Paragraph = {
      type: "paragraph",
      children: [
        {
          type: "hyperlink",
          url: "https://example.com",
          children: [{ content: [{ type: "text", text: "click" }] } as Run]
        }
      ]
    };

    await expect(
      Package.edit(seeded, [
        { type: "replaceFooter", path: "word/footer1.xml", children: [paraWithLink] }
      ])
    ).rejects.toThrow(/hyperlink/);
  });

  it("rejects replaceBody with an inline image", async () => {
    const buffer = await buildDocxBuffer("body");
    const paraWithImage: Paragraph = {
      type: "paragraph",
      children: [
        {
          content: [
            {
              type: "image",
              rId: "rId42",
              width: 100000,
              height: 100000
            }
          ]
        } as Run
      ]
    };

    await expect(
      Package.edit(buffer, [{ type: "replaceBody", body: [paraWithImage] }])
    ).rejects.toThrow(/inline image/);
  });
});

// =============================================================================
// replaceBody must not be confused by literal "<w:body>" / "</w:body>"
// occurrences inside XML comments, CDATA sections, or processing
// instructions. These are exotic but legal in the input document.xml the
// caller supplied (e.g. an upstream tool that injected an authoring note),
// and the previous purely-regex scan would happily slice on them, leaving
// the output document corrupted.
// =============================================================================

describe("editDocxIncremental — replaceBody scanner robustness", () => {
  async function bodyXmlOf(buf: Uint8Array): Promise<string> {
    const part = await Package.readPart(buf, "word/document.xml");
    return new TextDecoder().decode(part!);
  }

  it("ignores </w:body> sitting inside an XML comment", async () => {
    const buffer = await buildDocxBuffer("hello");
    const original = await bodyXmlOf(buffer);
    // Inject a comment containing fake body tags BEFORE the real <w:body>.
    const stitched = original.replace("<w:body", "<!-- <w:body><w:p/></w:body> --><w:body");
    const seeded = await Package.edit(buffer, [
      { type: "replacePartText", path: "word/document.xml", text: stitched }
    ]);

    const newPara: Paragraph = {
      type: "paragraph",
      children: [{ content: [{ type: "text", text: "POST-COMMENT" }] } as Run]
    };
    const edited = await Package.edit(seeded, [{ type: "replaceBody", body: [newPara] }]);

    // The decoy comment must survive untouched (proof we didn't pick the
    // fake closing tag inside the comment) and the real body must contain
    // the new paragraph text.
    const after = await bodyXmlOf(edited);
    expect(after).toContain("<!-- <w:body><w:p/></w:body> -->");
    expect(after).toContain("POST-COMMENT");
    expect(after).not.toContain("hello"); // original body content was swapped
  });

  it("ignores </w:body> sitting inside a CDATA section", async () => {
    const buffer = await buildDocxBuffer("hello");
    const original = await bodyXmlOf(buffer);
    const stitched = original.replace("<w:body", "<![CDATA[ </w:body> ]]><w:body");
    const seeded = await Package.edit(buffer, [
      { type: "replacePartText", path: "word/document.xml", text: stitched }
    ]);

    const newPara: Paragraph = {
      type: "paragraph",
      children: [{ content: [{ type: "text", text: "POST-CDATA" }] } as Run]
    };
    const edited = await Package.edit(seeded, [{ type: "replaceBody", body: [newPara] }]);

    const after = await bodyXmlOf(edited);
    expect(after).toContain("<![CDATA[ </w:body> ]]>");
    expect(after).toContain("POST-CDATA");
    expect(after).not.toContain("hello");
  });

  it("ignores <w:body inside a processing instruction", async () => {
    const buffer = await buildDocxBuffer("hello");
    const original = await bodyXmlOf(buffer);
    const stitched = original.replace("<w:body", "<?fake <w:body><w:p/></w:body> ?><w:body");
    const seeded = await Package.edit(buffer, [
      { type: "replacePartText", path: "word/document.xml", text: stitched }
    ]);

    const newPara: Paragraph = {
      type: "paragraph",
      children: [{ content: [{ type: "text", text: "POST-PI" }] } as Run]
    };
    const edited = await Package.edit(seeded, [{ type: "replaceBody", body: [newPara] }]);

    const after = await bodyXmlOf(edited);
    expect(after).toContain("<?fake <w:body><w:p/></w:body> ?>");
    expect(after).toContain("POST-PI");
    expect(after).not.toContain("hello");
  });
});

describe("editDocxIncremental — comments the edit leaves without an anchor", () => {
  async function commentedDocx(): Promise<Uint8Array> {
    const doc = Document.create();
    const kept = Document.addComment(doc, "Ann", [
      { ...Build.textParagraph("kept"), paraId: "0000000A" }
    ]);
    const dropped = Document.addComment(doc, "Bob", [
      { ...Build.textParagraph("dropped"), paraId: "0000000B" }
    ]);
    for (const id of [kept, dropped]) {
      Document.addContent(
        doc,
        Build.paragraph([
          Build.commentRangeStart(id),
          Build.text(`anchor ${id}`),
          Build.commentRangeEnd(id),
          Build.commentReference(id)
        ])
      );
    }
    // Mark both resolved so word/commentsExtended.xml carries an entry each.
    return Io.toBuffer({ ...doc, comments: doc.comments!.map(c => ({ ...c, done: true })) });
  }

  async function partText(buffer: Uint8Array, path: string): Promise<string> {
    const data = await Package.readPart(buffer, path);
    return data ? new TextDecoder().decode(data) : "";
  }

  it("drops a comment replaceBody leaves without a reference, with its extended entry", async () => {
    const buffer = await commentedDocx();
    const before = await Io.read(buffer);
    const keptId = before.comments!.find(c => c.author === "Ann")!.id;
    const edited = await Package.edit(buffer, [
      {
        type: "replaceBody",
        body: [
          Build.paragraph([
            Build.commentRangeStart(keptId),
            Build.text("still anchored"),
            Build.commentRangeEnd(keptId),
            Build.commentReference(keptId)
          ])
        ]
      }
    ]);
    const comments = await partText(edited, "word/comments.xml");
    expect(comments).toContain("kept");
    expect(comments).not.toContain("dropped");
    const extended = await partText(edited, "word/commentsExtended.xml");
    expect(extended).toContain('w15:paraId="0000000A"');
    expect(extended).not.toContain('w15:paraId="0000000B"');

    const after = await Io.read(edited);
    expect(after.comments?.map(c => c.author)).toEqual(["Ann"]);
  });

  it("drops every comment, and its durable id, when the new body references none", async () => {
    const buffer = await Package.edit(await commentedDocx(), [
      {
        type: "replacePartText",
        path: "word/commentsIds.xml",
        text:
          '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">' +
          '<w16cid:commentId w16cid:paraId="0000000A" w16cid:durableId="1111"/>' +
          '<w16cid:commentId w16cid:paraId="0000000B" w16cid:durableId="2222"/>' +
          "</w16cid:commentsIds>"
      }
    ]);
    const edited = await Package.edit(buffer, [
      { type: "replaceBody", body: [Build.textParagraph("no comments here")] }
    ]);
    expect(await partText(edited, "word/comments.xml")).not.toMatch(/<w:comment\b/);
    expect(await partText(edited, "word/commentsExtended.xml")).not.toMatch(/<w15:commentEx\b/);
    expect(await partText(edited, "word/commentsIds.xml")).not.toMatch(/<w16cid:commentId\b/);
    expect((await Io.read(edited)).comments ?? []).toEqual([]);
  });

  it("drops only what the replaced body alone anchored", async () => {
    const doc = Document.create();
    const anchor = (id: number) =>
      Build.paragraph([
        Build.commentRangeStart(id),
        Build.text(`anchor ${id}`),
        Build.commentRangeEnd(id),
        Build.commentReference(id)
      ]);
    const bodyOnly = Document.addComment(doc, "A", "body only");
    const alsoHeader = Document.addComment(doc, "A", "body and header");
    const inFooter = Document.addComment(doc, "A", "in footer");
    const inNote = Document.addComment(doc, "A", "in note");
    Document.addComment(doc, "A", "loose before the edit");
    Document.addContent(doc, anchor(bodyOnly));
    Document.addContent(doc, anchor(alsoHeader));
    Document.setHeader(doc, "default", { children: [anchor(alsoHeader)] });
    Document.setFooter(doc, "default", { children: [anchor(inFooter)] });
    const note = Document.addFootnote(doc, [anchor(inNote)]);
    Document.addContent(
      doc,
      Build.paragraph([{ ...Build.text(""), content: [{ type: "footnoteRef", id: note }] }])
    );
    const buffer = await Io.toBuffer(doc);
    expect(await partText(buffer, "word/comments.xml")).toContain("loose before the edit");

    const edited = await Package.edit(buffer, [
      { type: "replaceBody", body: [Build.textParagraph("new body")] }
    ]);
    const comments = await partText(edited, "word/comments.xml");
    expect(comments).not.toContain("body only");
    for (const kept of ["body and header", "in footer", "in note", "loose before the edit"]) {
      expect(comments).toContain(kept);
    }
  });

  it("leaves comments alone when only raw parts are replaced", async () => {
    const buffer = await commentedDocx();
    const edited = await Package.edit(buffer, [
      { type: "replacePartText", path: "word/document.xml", text: "<w:document/>" }
    ]);
    expect(await partText(edited, "word/comments.xml")).toContain("dropped");
  });
});

describe("a missing or malformed buffer is the Word API's own error", () => {
  // These used to surface the archive module's `ArchiveError`; a value that
  // is not bytes is now an argument error, and bytes that are not a ZIP a
  // parse error — what `Io.read` reports for the same inputs.
  const calls: [string, (b: unknown) => Promise<unknown>][] = [
    ["listDocxParts", b => Io.listDocxParts(b as Uint8Array)],
    ["readDocxPart", b => Io.readDocxPart(b as Uint8Array, "word/document.xml")],
    ["editDocxIncremental", b => Io.editDocxIncremental(b as Uint8Array, [])],
    ["Convert.readOdt", b => Convert.readOdt(b as Uint8Array)]
  ];
  it.each(calls)("%s", async (_, call) => {
    for (const input of [null, undefined, {}]) {
      await expect(call(input)).rejects.toBeInstanceOf(DocxArgumentError);
    }
    await expect(call(new Uint8Array([1, 2, 3]))).rejects.toBeInstanceOf(DocxParseError);
  });
});
