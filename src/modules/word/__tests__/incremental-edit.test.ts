/**
 * DOCX Module - Incremental Edit Tests
 */

import { describe, it, expect } from "vitest";

import {
  Document,
  editDocxIncremental,
  listDocxParts,
  packageDocx,
  readDocx,
  readDocxPart,
  textParagraph
} from "../index";
import type { Paragraph, Run } from "../types";

async function buildDocxBuffer(text: string): Promise<Uint8Array> {
  const doc = Document.create();
  Document.addContent(doc, textParagraph(text));
  return packageDocx(Document.build(doc));
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
    const parts = await listDocxParts(buffer);

    expect(parts.length).toBeGreaterThan(0);
    expect(parts).toContain("[Content_Types].xml");
    expect(parts).toContain("word/document.xml");
    expect(parts).toContain("_rels/.rels");
  });
});

describe("readDocxPart", () => {
  it("reads document.xml content", async () => {
    const buffer = await buildDocxBuffer("hello world");
    const data = await readDocxPart(buffer, "word/document.xml");
    expect(data).toBeDefined();
    const xml = new TextDecoder().decode(data!);
    expect(xml).toContain("<?xml");
    expect(xml).toContain("hello world");
  });

  it("returns undefined for missing part", async () => {
    const buffer = await buildDocxBuffer("test");
    const data = await readDocxPart(buffer, "word/no-such-part.xml");
    expect(data).toBeUndefined();
  });

  it("reads [Content_Types].xml", async () => {
    const buffer = await buildDocxBuffer("test");
    const data = await readDocxPart(buffer, "[Content_Types].xml");
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
    const edited = await editDocxIncremental(buffer, [
      {
        type: "replacePart",
        path: "[Content_Types].xml",
        data: new TextEncoder().encode(newXml)
      }
    ]);

    expect(edited).toBeInstanceOf(Uint8Array);
    const ct = await readDocxPart(edited, "[Content_Types].xml");
    expect(new TextDecoder().decode(ct!)).toBe(newXml);
  });

  it("replaces a part with text", async () => {
    const buffer = await buildDocxBuffer("test");
    // Replace content types - use a minimal valid XML
    const newXml =
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';
    const edited = await editDocxIncremental(buffer, [
      { type: "replacePartText", path: "[Content_Types].xml", text: newXml }
    ]);

    const ct = await readDocxPart(edited, "[Content_Types].xml");
    expect(new TextDecoder().decode(ct!)).toBe(newXml);
  });

  it("deletes a part", async () => {
    const buffer = await buildDocxBuffer("test");
    const partsBefore = await listDocxParts(buffer);

    // Add a part first by editing, then delete it
    const withExtra = await editDocxIncremental(buffer, [
      { type: "replacePartText", path: "extra.txt", text: "hello" }
    ]);

    const withExtraParts = await listDocxParts(withExtra);
    expect(withExtraParts).toContain("extra.txt");

    const deleted = await editDocxIncremental(withExtra, [
      { type: "deletePart", path: "extra.txt" }
    ]);

    const finalParts = await listDocxParts(deleted);
    expect(finalParts).not.toContain("extra.txt");
    expect(finalParts.length).toBe(partsBefore.length);
  });

  it("preserves unchanged parts", async () => {
    const buffer = await buildDocxBuffer("preserve me");
    const stylesXmlBefore = await readDocxPart(buffer, "word/styles.xml");
    const ctBefore = await readDocxPart(buffer, "[Content_Types].xml");

    const edited = await editDocxIncremental(buffer, [
      { type: "replacePartText", path: "extra.txt", text: "new" }
    ]);

    const stylesAfter = await readDocxPart(edited, "word/styles.xml");
    const ctAfter = await readDocxPart(edited, "[Content_Types].xml");

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

    const edited = await editDocxIncremental(buffer, [{ type: "replaceBody", body: [newPara] }]);

    expect(edited).toBeInstanceOf(Uint8Array);
    // Verify by reading back
    const parsed = await readDocx(edited);
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
    const edited = await editDocxIncremental(
      buffer,
      [{ type: "replacePartText", path: "extra.txt", text: "x" }],
      { compressionLevel: 9 }
    );
    expect(edited).toBeInstanceOf(Uint8Array);
  });

  it("applies multiple edits in order", async () => {
    const buffer = await buildDocxBuffer("test");
    const edited = await editDocxIncremental(buffer, [
      { type: "replacePartText", path: "first.txt", text: "1" },
      { type: "replacePartText", path: "second.txt", text: "2" },
      { type: "replacePartText", path: "first.txt", text: "1-updated" }
    ]);

    const first = await readDocxPart(edited, "first.txt");
    const second = await readDocxPart(edited, "second.txt");

    expect(new TextDecoder().decode(first!)).toBe("1-updated");
    expect(new TextDecoder().decode(second!)).toBe("2");
  });
});
