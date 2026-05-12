/**
 * DOCX Module - SecurityPolicy Tests
 *
 * Verifies the parts of WordSecurityPolicy that the reader actually
 * enforces. Adding a new field to the policy interface without enforcing
 * it should fail the corresponding test below — the goal is to keep the
 * public policy API honest about what is and isn't implemented.
 */

import { describe, it, expect } from "vitest";

import { Document, packageDocx, readDocx, STRICT_SECURITY_POLICY } from "../index";
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
    const bytes = await packageDocx(Document.build(h));

    const parsed = await readDocx(bytes);
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
    const bytes = await packageDocx(Document.build(h));

    const parsed = await readDocx(bytes, { securityPolicy: STRICT_SECURITY_POLICY });
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
    const bytes = await packageDocx(docm);

    const strict = await readDocx(bytes, { securityPolicy: STRICT_SECURITY_POLICY });
    expect(strict.vbaProject).toBeUndefined();

    const lenient = await readDocx(bytes);
    expect(lenient.vbaProject).toBeDefined();
  });
});
