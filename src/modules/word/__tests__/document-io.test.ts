/**
 * DOCX Module - Document IO (patchDocument/compileTemplate) Tests
 */

import { describe, it, expect } from "vitest";

import {
  patchDocument,
  compileTemplate,
  patchTemplate,
  toBuffer,
  toBase64,
  fillTemplateFromBuffer,
  textParagraph,
  packageDocx,
  readDocx,
  Document
} from "../index";

// Create a minimal DOCX buffer for testing
async function createTestDocx(content: string): Promise<Uint8Array> {
  const doc = Document.create();
  Document.addContent(doc, textParagraph(content));
  return packageDocx(Document.build(doc));
}

describe("patchDocument", () => {
  it("replaces text placeholder", async () => {
    const buffer = await createTestDocx("Hello {{name}}!");
    const result = await patchDocument(buffer, [
      { placeholder: "{{name}}", content: { type: "text", text: "World" } }
    ]);

    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(0);

    // Verify by reading back
    const doc = await readDocx(result);
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
    Document.addContent(doc, textParagraph("{{first}}"));
    Document.addContent(doc, textParagraph("{{last}}"));
    const buffer = await packageDocx(Document.build(doc));

    const result = await patchDocument(buffer, [
      { placeholder: "{{first}}", content: { type: "text", text: "John" } },
      { placeholder: "{{last}}", content: { type: "text", text: "Doe" } }
    ]);

    const parsed = await readDocx(result);
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
    const result = await patchDocument(buffer, [
      { placeholder: "{{first}}", content: { type: "text", text: "Jane" } },
      { placeholder: "{{last}}", content: { type: "text", text: "Smith" } }
    ]);

    const parsed = await readDocx(result);
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
    const result = await patchDocument(buffer, [
      { placeholder: "{{missing}}", content: { type: "text", text: "nope" } }
    ]);

    expect(result).toBeInstanceOf(Uint8Array);
    const doc = await readDocx(result);
    expect(doc.body.length).toBeGreaterThan(0);
  });
});

describe("compileTemplate / patchTemplate", () => {
  it("compiles a template and patches it multiple times", async () => {
    const buffer = await createTestDocx("Dear {{name}},");
    const template = await compileTemplate(buffer);

    const result1 = await patchTemplate(template, [
      { placeholder: "{{name}}", content: { type: "text", text: "Alice" } }
    ]);
    const result2 = await patchTemplate(template, [
      { placeholder: "{{name}}", content: { type: "text", text: "Bob" } }
    ]);

    expect(result1).toBeInstanceOf(Uint8Array);
    expect(result2).toBeInstanceOf(Uint8Array);

    // Both should produce valid DOCX
    const doc1 = await readDocx(result1);
    const doc2 = await readDocx(result2);
    expect(doc1.body.length).toBeGreaterThan(0);
    expect(doc2.body.length).toBeGreaterThan(0);
  });

  it("does not mutate the template between patches", async () => {
    const buffer = await createTestDocx("Value: {{val}}");
    const template = await compileTemplate(buffer);

    await patchTemplate(template, [
      { placeholder: "{{val}}", content: { type: "text", text: "first" } }
    ]);

    // Second patch should still find the placeholder
    const result = await patchTemplate(template, [
      { placeholder: "{{val}}", content: { type: "text", text: "second" } }
    ]);

    const doc = await readDocx(result);
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
    Document.addContent(doc, textParagraph("Test"));
    const result = await toBuffer(Document.build(doc));
    expect(result).toBeInstanceOf(Uint8Array);
    // ZIP magic number
    expect(result[0]).toBe(0x50);
    expect(result[1]).toBe(0x4b);
  });

  it("toBase64 produces valid base64 string", async () => {
    const doc = Document.create();
    Document.addContent(doc, textParagraph("Test"));
    const result = await toBase64(Document.build(doc));
    expect(typeof result).toBe("string");
    // Should be valid base64
    expect(() => atob(result)).not.toThrow();
  });
});

describe("fillTemplateFromBuffer", () => {
  it("fills template variables", async () => {
    const buffer = await createTestDocx("Name: {{name}}");
    const result = await fillTemplateFromBuffer(buffer, { name: "Test User" });

    expect(result).toBeInstanceOf(Uint8Array);
    const doc = await readDocx(result);
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
