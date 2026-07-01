/**
 * DOCX Module - Document IO (patchDocument/compileTemplate) Tests
 */

import { describe, it, expect } from "vitest";

import { Document, Build, Io } from "../index";
import { applyPatchesToDocument } from "../patcher";
import type {
  DocxDocument,
  DrawingShape,
  Hyperlink,
  ImageDef,
  Paragraph,
  Run,
  StructuredDocumentTag,
  TableOfContents,
  TextBox
} from "../types";

function makeRun(text: string): Run {
  return { content: [{ type: "text", text }] } as Run;
}
function makePara(text: string): Paragraph {
  return { type: "paragraph", children: [makeRun(text)] };
}

// Create a minimal DOCX buffer for testing
async function createTestDocx(content: string): Promise<Uint8Array> {
  const doc = Document.create();
  Document.addContent(doc, Build.textParagraph(content));
  return Io.package(Document.build(doc));
}

describe("patchDocument", () => {
  it("replaces text placeholder", async () => {
    const buffer = await createTestDocx("Hello {{name}}!");
    const result = await Io.patchDocument(buffer, [
      { placeholder: "{{name}}", content: { type: "text", text: "World" } }
    ]);

    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(0);

    // Verify by reading back
    const doc = await Io.read(result);
    let found = false;
    for (const block of doc.body) {
      if (block.type === "paragraph") {
        for (const child of block.children) {
          if ("content" in child && Array.isArray(child.content)) {
            for (const c of child.content) {
              if (c.type === "text" && c.text.includes("World")) {
                found = true;
              }
            }
          }
        }
      }
    }
    expect(found).toBe(true);
  });

  it("handles multiple patches in separate paragraphs", async () => {
    // patchDocument processes one placeholder per paragraph, so use separate paragraphs
    const doc = Document.create();
    Document.addContent(doc, Build.textParagraph("{{first}}"));
    Document.addContent(doc, Build.textParagraph("{{last}}"));
    const buffer = await Io.package(Document.build(doc));

    const result = await Io.patchDocument(buffer, [
      { placeholder: "{{first}}", content: { type: "text", text: "John" } },
      { placeholder: "{{last}}", content: { type: "text", text: "Doe" } }
    ]);

    const parsed = await Io.read(result);
    let text = "";
    for (const block of parsed.body) {
      if (block.type === "paragraph") {
        for (const child of block.children) {
          if ("content" in child && Array.isArray(child.content)) {
            for (const c of child.content) {
              if (c.type === "text") {
                text += c.text + " ";
              }
            }
          }
        }
      }
    }
    expect(text).toContain("John");
    expect(text).toContain("Doe");
  });

  it("handles multiple text patches in same paragraph", async () => {
    const buffer = await createTestDocx("{{first}} {{last}}");
    const result = await Io.patchDocument(buffer, [
      { placeholder: "{{first}}", content: { type: "text", text: "Jane" } },
      { placeholder: "{{last}}", content: { type: "text", text: "Smith" } }
    ]);

    const parsed = await Io.read(result);
    let text = "";
    for (const block of parsed.body) {
      if (block.type === "paragraph") {
        for (const child of block.children) {
          if ("content" in child && Array.isArray(child.content)) {
            for (const c of child.content) {
              if (c.type === "text") {
                text += c.text;
              }
            }
          }
        }
      }
    }
    expect(text).toContain("Jane");
    expect(text).toContain("Smith");
  });

  it("returns valid DOCX when no patches match", async () => {
    const buffer = await createTestDocx("No placeholders here");
    const result = await Io.patchDocument(buffer, [
      { placeholder: "{{missing}}", content: { type: "text", text: "nope" } }
    ]);

    expect(result).toBeInstanceOf(Uint8Array);
    const doc = await Io.read(result);
    expect(doc.body.length).toBeGreaterThan(0);
  });
});

describe("compileTemplate / patchTemplate", () => {
  it("compiles a template and patches it multiple times", async () => {
    const buffer = await createTestDocx("Dear {{name}},");
    const template = await Io.compileTemplate(buffer);

    const result1 = await Io.patchTemplate(template, [
      { placeholder: "{{name}}", content: { type: "text", text: "Alice" } }
    ]);
    const result2 = await Io.patchTemplate(template, [
      { placeholder: "{{name}}", content: { type: "text", text: "Bob" } }
    ]);

    expect(result1).toBeInstanceOf(Uint8Array);
    expect(result2).toBeInstanceOf(Uint8Array);

    // Both should produce valid DOCX
    const doc1 = await Io.read(result1);
    const doc2 = await Io.read(result2);
    expect(doc1.body.length).toBeGreaterThan(0);
    expect(doc2.body.length).toBeGreaterThan(0);
  });

  it("does not mutate the template between patches", async () => {
    const buffer = await createTestDocx("Value: {{val}}");
    const template = await Io.compileTemplate(buffer);

    await Io.patchTemplate(template, [
      { placeholder: "{{val}}", content: { type: "text", text: "first" } }
    ]);

    // Second patch should still find the placeholder
    const result = await Io.patchTemplate(template, [
      { placeholder: "{{val}}", content: { type: "text", text: "second" } }
    ]);

    const doc = await Io.read(result);
    let text = "";
    for (const block of doc.body) {
      if (block.type === "paragraph") {
        for (const child of block.children) {
          if ("content" in child && Array.isArray(child.content)) {
            for (const c of child.content) {
              if (c.type === "text") {
                text += c.text;
              }
            }
          }
        }
      }
    }
    expect(text).toContain("second");
    expect(text).not.toContain("first");
  });
});

describe("toBuffer / toBase64", () => {
  it("toBuffer produces valid bytes", async () => {
    const doc = Document.create();
    Document.addContent(doc, Build.textParagraph("Test"));
    const result = await Io.toBuffer(Document.build(doc));
    expect(result).toBeInstanceOf(Uint8Array);
    // ZIP magic number
    expect(result[0]).toBe(0x50);
    expect(result[1]).toBe(0x4b);
  });

  it("toBase64 produces valid base64 string", async () => {
    const doc = Document.create();
    Document.addContent(doc, Build.textParagraph("Test"));
    const result = await Io.toBase64(Document.build(doc));
    expect(typeof result).toBe("string");
    // Should be valid base64
    expect(() => atob(result)).not.toThrow();
  });
});

describe("fillTemplateFromBuffer", () => {
  it("fills template variables", async () => {
    const buffer = await createTestDocx("Name: {{name}}");
    const result = await Io.fillTemplateFromBuffer(buffer, { name: "Test User" });

    expect(result).toBeInstanceOf(Uint8Array);
    const doc = await Io.read(result);
    let text = "";
    for (const block of doc.body) {
      if (block.type === "paragraph") {
        for (const child of block.children) {
          if ("content" in child && Array.isArray(child.content)) {
            for (const c of child.content) {
              if (c.type === "text") {
                text += c.text;
              }
            }
          }
        }
      }
    }
    expect(text).toContain("Test User");
  });
});

// =============================================================================
// Container coverage regression
//
// patchDocument / applyPatchesToDocument used to walk only paragraphs and
// tables at the top level (and inside table cells), silently skipping
// placeholders that lived inside text boxes, drawing shapes, table-of-
// contents cached entries, or block-level structured-document tags. These
// tests pin down the expanded coverage so the bug can't regress.
// =============================================================================

describe("applyPatchesToDocument: container coverage", () => {
  it("replaces text inside a text box", () => {
    const tb: TextBox = {
      type: "textBox",
      content: [makePara("Hello {{name}}")]
    };
    const doc: DocxDocument = { body: [tb] };
    const result = applyPatchesToDocument(doc, [
      { placeholder: "{{name}}", content: { type: "text", text: "World" } }
    ]);
    const out = result.body[0] as TextBox;
    const run = out.content[0].children[0] as Run;
    const segs = run.content
      .filter((c): c is { type: "text"; text: string } => c.type === "text")
      .map(c => c.text)
      .join("");
    expect(segs).toBe("Hello World");
  });

  it("replaces text inside a drawing shape's text body", () => {
    const shape: DrawingShape = {
      type: "drawingShape",
      shapeType: "rect",
      width: 100,
      height: 100,
      textContent: [makePara("Title: {{title}}")]
    };
    const doc: DocxDocument = { body: [shape] };
    const result = applyPatchesToDocument(doc, [
      { placeholder: "{{title}}", content: { type: "text", text: "Hi" } }
    ]);
    const out = result.body[0] as DrawingShape;
    const run = out.textContent![0].children[0] as Run;
    const text = (run.content[0] as { type: "text"; text: string }).text;
    expect(text).toBe("Title: Hi");
  });

  it("replaces text inside a TOC's cached paragraphs", () => {
    const toc: TableOfContents = {
      type: "tableOfContents",
      cachedParagraphs: [makePara("Chapter — {{chapter}}")]
    };
    const doc: DocxDocument = { body: [toc] };
    const result = applyPatchesToDocument(doc, [
      { placeholder: "{{chapter}}", content: { type: "text", text: "1" } }
    ]);
    const out = result.body[0] as TableOfContents;
    const run = out.cachedParagraphs![0].children[0] as Run;
    const text = (run.content[0] as { type: "text"; text: string }).text;
    expect(text).toBe("Chapter — 1");
  });

  it("replaces text inside a block-level SDT", () => {
    const sdt: StructuredDocumentTag = {
      type: "sdt",
      content: [makePara("Field: {{value}}")]
    };
    const doc: DocxDocument = { body: [sdt] };
    const result = applyPatchesToDocument(doc, [
      { placeholder: "{{value}}", content: { type: "text", text: "OK" } }
    ]);
    const out = result.body[0] as StructuredDocumentTag;
    const para = out.content[0] as Paragraph;
    const run = para.children[0] as Run;
    const text = (run.content[0] as { type: "text"; text: string }).text;
    expect(text).toBe("Field: OK");
  });

  it("replaces text spanning multiple inline runs of an inline SDT", () => {
    const r1 = makeRun("foo");
    const r2 = makeRun("bar");
    const sdt: StructuredDocumentTag = {
      type: "sdt",
      content: [r1, r2]
    };
    const doc: DocxDocument = { body: [sdt] };
    const result = applyPatchesToDocument(doc, [
      { placeholder: "foobar", content: { type: "text", text: "QUUX" } }
    ]);
    const out = result.body[0] as StructuredDocumentTag;
    // Stitched: replacement lands in the first run, the second is emptied.
    const t0 = ((out.content[0] as Run).content[0] as { type: "text"; text: string }).text;
    const t1 = ((out.content[1] as Run).content[0] as { type: "text"; text: string }).text;
    expect(t0 + t1).toBe("QUUX");
  });

  it("replaces a paragraph-typed placeholder inside a block-level SDT", () => {
    // paragraph patches expand into the SDT's content stream.
    const sdt: StructuredDocumentTag = {
      type: "sdt",
      content: [makePara("{{block}}")]
    };
    const doc: DocxDocument = { body: [sdt] };
    const result = applyPatchesToDocument(doc, [
      {
        placeholder: "{{block}}",
        content: {
          type: "paragraph",
          children: [makePara("First"), makePara("Second")]
        }
      }
    ]);
    const out = result.body[0] as StructuredDocumentTag;
    expect(out.content.length).toBe(2);
    const first = (
      ((out.content[0] as Paragraph).children[0] as Run).content[0] as {
        type: "text";
        text: string;
      }
    ).text;
    const second = (
      ((out.content[1] as Paragraph).children[0] as Run).content[0] as {
        type: "text";
        text: string;
      }
    ).text;
    expect(first).toBe("First");
    expect(second).toBe("Second");
  });

  it("does not corrupt the SDT object when a paragraph-typed placeholder maps to a Table", () => {
    // Previous code did `Object.assign(paragraph, table)` here, producing a
    // hybrid object with `type: "table"` but stray paragraph fields. We now
    // keep tables that come out of structural patches as proper Table
    // entries in the SDT's content stream and never mutate the original
    // paragraph in place.
    const para = makePara("{{tbl}}");
    const sdt: StructuredDocumentTag = {
      type: "sdt",
      content: [para]
    };
    const doc: DocxDocument = { body: [sdt] };
    const result = applyPatchesToDocument(doc, [
      {
        placeholder: "{{tbl}}",
        content: {
          type: "table",
          table: {
            type: "table",
            rows: [{ cells: [{ content: [makePara("A")] }] }]
          }
        }
      }
    ]);
    const out = result.body[0] as StructuredDocumentTag;
    expect(out.content.length).toBe(1);
    expect((out.content[0] as { type: string }).type).toBe("table");
    // The original paragraph object must NOT have been overwritten with the
    // table's fields.
    expect(para.type).toBe("paragraph");
    expect("rows" in para).toBe(false);
  });
});

// =============================================================================
// Image rId consistency regression
//
// Two image patches sharing a fileName must end up with a single rId both
// in the body reference and in the registered images[] entry — otherwise
// the body has a dangling r:embed and Word renders blank images.
// =============================================================================

describe("applyPatchesToDocument: image rId normalization", () => {
  it("uses one rId for two patches that share a fileName", () => {
    const sharedImg: ImageDef = {
      data: new Uint8Array([1, 2, 3]),
      mediaType: "png",
      fileName: "shared.png"
      // no rId — patcher must invent a stable one and reuse it
    };
    const doc: DocxDocument = {
      body: [makePara("{{a}}"), makePara("{{b}}")]
    };
    const result = applyPatchesToDocument(doc, [
      {
        placeholder: "{{a}}",
        content: { type: "image", image: sharedImg, width: 100, height: 100 }
      },
      {
        placeholder: "{{b}}",
        content: { type: "image", image: sharedImg, width: 200, height: 200 }
      }
    ]);

    // Both patched paragraphs should reference the same rId.
    const p0 = result.body[0] as Paragraph;
    const p1 = result.body[1] as Paragraph;
    const rId0 = ((p0.children[0] as Run).content[0] as { rId: string }).rId;
    const rId1 = ((p1.children[0] as Run).content[0] as { rId: string }).rId;
    expect(rId0).toBe(rId1);

    // images[] should list the shared file exactly once with that same rId.
    expect(result.images?.length).toBe(1);
    expect(result.images![0].rId).toBe(rId0);
  });

  it("respects an existing rId from doc.images", () => {
    const sharedImg: ImageDef = {
      data: new Uint8Array([1, 2, 3]),
      mediaType: "png",
      fileName: "existing.png",
      rId: "rIdExisting"
    };
    const doc: DocxDocument = {
      body: [makePara("{{x}}")],
      images: [sharedImg]
    };
    const result = applyPatchesToDocument(doc, [
      {
        placeholder: "{{x}}",
        content: {
          type: "image",
          // Caller passes the same fileName but a different (stale) rId —
          // the normalizer must overwrite it with the existing one.
          image: { ...sharedImg, rId: "rIdStale" },
          width: 100,
          height: 100
        }
      }
    ]);
    const p = result.body[0] as Paragraph;
    const rId = ((p.children[0] as Run).content[0] as { rId: string }).rId;
    expect(rId).toBe("rIdExisting");
  });
});

// =============================================================================
// Visible-run coverage regression (hyperlinks & tracked-insert wrappers)
//
// Placeholders inside a hyperlink display text or inside `insertedRun` /
// `movedToRun` track-change wrappers used to be skipped — patcher only
// looked at top-level runs. Pending deletions (`deletedRun` /
// `movedFromRun`) must STILL be skipped, since by convention they don't
// contribute to the document's visible text.
// =============================================================================

describe("applyPatchesToDocument: visible-run coverage", () => {
  it("replaces text inside a hyperlink", () => {
    const link = {
      type: "hyperlink" as const,
      url: "https://example.com",
      children: [makeRun("Visit {{site}}")]
    };
    const para: Paragraph = { type: "paragraph", children: [link] };
    const doc: DocxDocument = { body: [para] };
    const result = applyPatchesToDocument(doc, [
      { placeholder: "{{site}}", content: { type: "text", text: "Acme" } }
    ]);
    const p = result.body[0] as Paragraph;
    const hl = p.children[0] as Hyperlink;
    const txt = (hl.children[0].content[0] as { type: "text"; text: string }).text;
    expect(txt).toBe("Visit Acme");
  });

  it("replaces text inside an insertedRun (tracked insert) wrapper", () => {
    const innerRun = makeRun("Hello {{name}}");
    const ins = {
      type: "insertedRun" as const,
      revision: { id: 1, author: "x" },
      run: innerRun
    };
    const para: Paragraph = { type: "paragraph", children: [ins] as never };
    const doc: DocxDocument = { body: [para] };
    const result = applyPatchesToDocument(doc, [
      { placeholder: "{{name}}", content: { type: "text", text: "World" } }
    ]);
    const p = result.body[0] as Paragraph;
    const wrapped = p.children[0] as { run: Run };
    const txt = (wrapped.run.content[0] as { type: "text"; text: string }).text;
    expect(txt).toBe("Hello World");
  });

  it("does NOT replace text inside a deletedRun (pending deletion is invisible)", () => {
    const innerRun = makeRun("delete {{x}}");
    const del = {
      type: "deletedRun" as const,
      revision: { id: 1, author: "x" },
      run: innerRun
    };
    const para: Paragraph = { type: "paragraph", children: [del] as never };
    const doc: DocxDocument = { body: [para] };
    applyPatchesToDocument(doc, [
      { placeholder: "{{x}}", content: { type: "text", text: "WAS HERE" } }
    ]);
    // Inner run text must still contain the placeholder — pending
    // deletions don't contribute to the document's visible text and
    // therefore must not receive replacements either.
    const txt = (innerRun.content[0] as { type: "text"; text: string }).text;
    expect(txt).toBe("delete {{x}}");
  });

  it("stitches a placeholder split across a top-level run and a hyperlink run", () => {
    // The placeholder "{{name}}" is split: "{{na" lives in a normal run
    // and "me}}" lives inside a hyperlink. Cross-run replacement must
    // still find and replace the full token.
    const r1 = makeRun("Hi {{na");
    const link = {
      type: "hyperlink" as const,
      url: "#",
      children: [makeRun("me}}!")]
    };
    const para: Paragraph = { type: "paragraph", children: [r1, link] };
    const doc: DocxDocument = { body: [para] };
    const result = applyPatchesToDocument(doc, [
      { placeholder: "{{name}}", content: { type: "text", text: "Friend" } }
    ]);
    const p = result.body[0] as Paragraph;
    const t0 = ((p.children[0] as Run).content[0] as { type: "text"; text: string }).text;
    const t1 = (
      (p.children[1] as Hyperlink).children[0].content[0] as {
        type: "text";
        text: string;
      }
    ).text;
    // The full replaced text "Hi Friend!" lands in the first text node;
    // the hyperlink's text node is cleared. (The cross-run fallback
    // collapses everything into the first segment by design.)
    expect(t0 + t1).toBe("Hi Friend!");
  });
});
