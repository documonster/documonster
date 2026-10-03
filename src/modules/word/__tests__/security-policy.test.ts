/**
 * DOCX Module - SecurityPolicy Tests
 *
 * Verifies the parts of WordSecurityPolicy that the reader actually
 * enforces. Adding a new field to the policy interface without enforcing
 * it should fail the corresponding test below — the goal is to keep the
 * public policy API honest about what is and isn't implemented.
 */

import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { DocxLimitExceededError } from "@word/errors";
import { describe, it, expect } from "vitest";

import { Document, Io, Security } from "../index";
import type { Hyperlink, Paragraph, Run } from "../types";

describe("WordSecurityPolicy: allowExternalTargets", () => {
  it("default policy keeps external hyperlink URLs", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          type: "hyperlink",
          url: "https://example.com",
          children: [{ content: [{ type: "text", text: "click" }] }]
        }
      ]
    } as Paragraph);
    const bytes = await Io.package(Document.build(h));

    const parsed = await Io.read(bytes);
    const link = (parsed.body[0] as Paragraph).children.find(
      (c): c is Hyperlink => "type" in c && c.type === "hyperlink"
    )!;
    expect(link.url).toBe("https://example.com");
  });

  it("strict policy strips external hyperlink URLs but keeps inner text", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          type: "hyperlink",
          url: "https://example.com",
          children: [{ content: [{ type: "text", text: "click" }] }]
        }
      ]
    } as Paragraph);
    const bytes = await Io.package(Document.build(h));

    const parsed = await Io.read(bytes, { securityPolicy: Security.STRICT_SECURITY_POLICY });
    const link = (parsed.body[0] as Paragraph).children.find(
      (c): c is Hyperlink => "type" in c && c.type === "hyperlink"
    );
    // The hyperlink wrapper survives, but its URL is dropped — the inner
    // run text is still visible to downstream consumers.
    expect(link).toBeDefined();
    expect(link!.url).toBeUndefined();
    const inner = link!.children[0] as Run;
    expect((inner.content[0] as { text: string }).text).toBe("click");
  });
});

describe("WordSecurityPolicy: preserveVbaProject", () => {
  it("strict policy drops vbaProject binary on .docm round-trip", async () => {
    // Synthesize a minimal docm: a regular doc with a fake VBA blob
    // attached. The packager wires up the rels + content type for us.
    const h = Document.create();
    Document.addParagraph(h, "Body");
    const doc = Document.build(h);
    const docm = {
      ...doc,
      docType: "macroEnabledDocument" as const,
      vbaProject: new Uint8Array([0x01, 0x02, 0x03])
    };
    const bytes = await Io.package(docm);

    const strict = await Io.read(bytes, { securityPolicy: Security.STRICT_SECURITY_POLICY });
    expect(strict.vbaProject).toBeUndefined();

    const lenient = await Io.read(bytes);
    expect(lenient.vbaProject).toBeDefined();
  });
});

describe("WordSecurityPolicy: maxXmlDepth", () => {
  /** Rebuild a minimal package with `part` replaced by `xml`. */
  async function withPart(part: string, xml: string): Promise<Uint8Array> {
    const h = Document.create();
    Document.addParagraph(h, "hello");
    const files = await extractAll(await Io.package(Document.build(h)));
    const entries = [...files].map(([name, entry]) => ({ name, data: entry.data }));
    const data = new TextEncoder().encode(xml);
    const existing = entries.find(e => e.name === part);
    if (existing) {
      existing.data = data;
    } else {
      entries.push({ name: part, data });
    }
    return createZip(entries);
  }

  function nestedBody(depth: number): string {
    // w:document > w:body > w:sdt > w:sdtContent > … > w:p
    const open = "<w:sdt><w:sdtContent>".repeat(depth);
    const close = "</w:sdtContent></w:sdt>".repeat(depth);
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `<w:body>${open}<w:p><w:r><w:t>x</w:t></w:r></w:p>${close}</w:body></w:document>`
    );
  }

  it("raises DocxLimitExceededError when document.xml nests beyond the policy", async () => {
    const bytes = await withPart("word/document.xml", nestedBody(30));
    const error = await Io.read(bytes, { securityPolicy: { maxXmlDepth: 40 } }).catch(
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf(DocxLimitExceededError);
    expect((error as DocxLimitExceededError).limit).toBe("xmlDepth");
    expect((error as DocxLimitExceededError).maximum).toBe(40);
  });

  it("reads the same document under a policy that allows its depth", async () => {
    const bytes = await withPart("word/document.xml", nestedBody(30));
    await expect(Io.read(bytes, { securityPolicy: { maxXmlDepth: 200 } })).resolves.toBeDefined();
  });

  it("strict preset is tighter than the default", async () => {
    const bytes = await withPart("word/document.xml", nestedBody(80));
    await expect(Io.read(bytes)).resolves.toBeDefined();
    await expect(
      Io.read(bytes, { securityPolicy: Security.STRICT_SECURITY_POLICY })
    ).rejects.toBeInstanceOf(DocxLimitExceededError);
  });

  it("is not swallowed by best-effort parsing of auxiliary parts", async () => {
    const deep =
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      "<w:x>".repeat(50) +
      "</w:x>".repeat(50) +
      "</w:styles>";
    const bytes = await withPart("word/styles.xml", deep);
    await expect(Io.read(bytes, { securityPolicy: { maxXmlDepth: 20 } })).rejects.toBeInstanceOf(
      DocxLimitExceededError
    );
  });
});
