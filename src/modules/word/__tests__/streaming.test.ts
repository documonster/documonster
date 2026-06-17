/**
 * DOCX Module - Streaming Writer Tests
 */

import { extractAll } from "@archive/unzip/extract";
import { describe, it, expect } from "vitest";

import { Build, Streaming } from "../index";
import type { BodyContent, HeaderDef, ImageDef, CommentDef } from "../types";

const MINI_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 6, 0, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 10, 0x49, 0x44, 0x41, 0x54, 0x78,
  0xda, 0x62, 0, 0, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82
]);

async function extract(bytes: Uint8Array): Promise<Map<string, string>> {
  const files = await extractAll(bytes);
  const out = new Map<string, string>();
  for (const [path, entry] of files) {
    out.set(path, new TextDecoder().decode(entry.data));
  }
  return out;
}

describe("StreamingDocxWriter", () => {
  it("basic usage: add 3 paragraphs, finalize, verify output", async () => {
    const stream = Streaming.createDocxStream();
    stream.add(Build.textParagraph("First"));
    stream.add(Build.textParagraph("Second"));
    stream.add(Build.textParagraph("Third"));

    const output = await stream.finalize();
    expect(output).toBeInstanceOf(Uint8Array);
    expect(output.length).toBeGreaterThan(0);
    // ZIP magic number
    expect(output[0]).toBe(0x50);
    expect(output[1]).toBe(0x4b);
  });

  it("addMany with 100 elements", async () => {
    const stream = Streaming.createDocxStream();
    const elements: BodyContent[] = [];
    for (let i = 0; i < 100; i++) {
      elements.push(Build.textParagraph(`Paragraph ${i}`));
    }
    stream.addMany(elements);
    expect(stream.elementCount).toBe(100);

    const output = await stream.finalize();
    expect(output).toBeInstanceOf(Uint8Array);
    expect(output.length).toBeGreaterThan(0);
  });

  it("progress callback reports progress", async () => {
    const stream = Streaming.createDocxStream({ chunkSize: 5 });
    const reports: Array<{ elementsWritten: number; phase: string }> = [];
    stream.onProgress(info => {
      reports.push({ ...info });
    });

    for (let i = 0; i < 15; i++) {
      stream.add(Build.textParagraph(`Para ${i}`));
    }
    await stream.finalize();

    // Should have reported at least at 5, 10, 15 elements
    expect(reports.some(r => r.elementsWritten === 5)).toBe(true);
    expect(reports.some(r => r.elementsWritten === 10)).toBe(true);
    expect(reports.some(r => r.elementsWritten === 15)).toBe(true);
    // Should have a "finalizing" phase report
    expect(reports.some(r => r.phase === "finalizing")).toBe(true);
  });

  it("throws on add after finalize", async () => {
    const stream = Streaming.createDocxStream();
    stream.add(Build.textParagraph("test"));
    await stream.finalize();

    expect(() => stream.add(Build.textParagraph("too late"))).toThrow();
  });

  it("reset for reuse", async () => {
    const stream = Streaming.createDocxStream();
    stream.add(Build.textParagraph("first run"));
    await stream.finalize();

    stream.reset();
    stream.add(Build.textParagraph("second run"));
    const output = await stream.finalize();
    expect(output).toBeInstanceOf(Uint8Array);
    expect(output.length).toBeGreaterThan(0);
    expect(stream.elementCount).toBe(1);
  });

  it("throws when an image-bearing element references an rId not provided in options.images", () => {
    const stream = Streaming.createDocxStream();
    const para: BodyContent = {
      type: "paragraph",
      children: [
        {
          content: [
            {
              type: "image",
              rId: "missingRId",
              width: 100000,
              height: 100000,
              fileName: "ghost.png"
            }
          ]
        }
      ]
    } as unknown as BodyContent;

    expect(() => stream.add(para)).toThrow(/missingRId/);
  });

  it("missingImagePolicy=warn skips silently with a console.warn", async () => {
    const warnings: string[] = [];
    const originalWarn = console.warn;
    console.warn = (...args: unknown[]) => {
      warnings.push(String(args[0]));
    };
    try {
      const stream = Streaming.createDocxStream({ missingImagePolicy: "warn" });
      const para: BodyContent = {
        type: "paragraph",
        children: [
          {
            content: [
              {
                type: "image",
                rId: "missingRId",
                width: 100000,
                height: 100000,
                fileName: "ghost.png"
              }
            ]
          }
        ]
      } as unknown as BodyContent;

      stream.add(para);
      const output = await stream.finalize();
      expect(output.length).toBeGreaterThan(0);
      expect(warnings.some(w => w.includes("missingRId"))).toBe(true);
    } finally {
      console.warn = originalWarn;
    }
  });

  it("writes header.xml.rels with hyperlink and image relationships referenced inside the header", async () => {
    const headerDef: HeaderDef = {
      content: {
        children: [
          {
            type: "paragraph",
            children: [
              {
                type: "hyperlink",
                url: "https://example.com",
                children: [{ content: [{ type: "text", text: "link" }] }]
              },
              {
                content: [
                  {
                    type: "image",
                    rId: "rIdHdrImg",
                    width: 100000,
                    height: 100000,
                    fileName: "hdr.png"
                  }
                ]
              }
            ]
          } as never
        ]
      }
    };
    const headers = new Map<string, HeaderDef>([["default", headerDef]]);
    const images: ImageDef[] = [
      { rId: "rIdHdrImg", fileName: "hdr.png", data: MINI_PNG, mediaType: "png" }
    ];
    const stream = Streaming.createDocxStream({ headers, images });
    stream.add(Build.textParagraph("body"));
    const bytes = await stream.finalize();
    const files = await extract(bytes);

    const headerRels = files.get("word/_rels/header1.xml.rels");
    expect(headerRels).toBeDefined();
    expect(headerRels!).toContain("https://example.com");
    expect(headerRels!).toContain('Id="rIdHdrImg"');
    // The hyperlink relationship must NOT also live in document.xml.rels —
    // header rIds must resolve against the header's own .rels file.
    const docRels = files.get("word/_rels/document.xml.rels");
    expect(docRels).toBeDefined();
    expect(docRels!).not.toContain("https://example.com");
  });

  it("writes footnote/endnote/comment .rels for hyperlinks living inside those parts", async () => {
    const stream = Streaming.createDocxStream({
      footnotes: [
        {
          id: 2,
          content: [
            {
              type: "paragraph",
              children: [
                {
                  type: "hyperlink",
                  url: "https://footnote.example",
                  children: [{ content: [{ type: "text", text: "fn" }] }]
                }
              ]
            } as unknown as BodyContent as never
          ]
        }
      ] as never,
      endnotes: [
        {
          id: 2,
          content: [
            {
              type: "paragraph",
              children: [
                {
                  type: "hyperlink",
                  url: "https://endnote.example",
                  children: [{ content: [{ type: "text", text: "en" }] }]
                }
              ]
            } as unknown as BodyContent as never
          ]
        }
      ] as never,
      comments: [
        {
          id: 1,
          author: "tester",
          content: [
            {
              type: "paragraph",
              children: [
                {
                  type: "hyperlink",
                  url: "https://comment.example",
                  children: [{ content: [{ type: "text", text: "cm" }] }]
                }
              ]
            } as unknown as BodyContent as never
          ]
        } as CommentDef
      ]
    });
    stream.add(Build.textParagraph("body"));
    const bytes = await stream.finalize();
    const files = await extract(bytes);

    expect(files.get("word/_rels/footnotes.xml.rels")).toContain("footnote.example");
    expect(files.get("word/_rels/endnotes.xml.rels")).toContain("endnote.example");
    expect(files.get("word/_rels/comments.xml.rels")).toContain("comment.example");

    // Document-level rels must NOT carry these external URLs.
    const docRels = files.get("word/_rels/document.xml.rels");
    expect(docRels).toBeDefined();
    expect(docRels!).not.toContain("footnote.example");
    expect(docRels!).not.toContain("endnote.example");
    expect(docRels!).not.toContain("comment.example");
  });

  it("rewires sectionProperties header/footer rIds to match document.xml.rels", async () => {
    // sectionProperties references "default" (a logical type) — the
    // streaming writer must allocate matching rIds and rewrite the
    // section-property references so they resolve in document.xml.rels.
    const headers = new Map<string, HeaderDef>([
      ["default", { content: { children: [{ type: "paragraph", children: [] } as never] } }]
    ]);
    const stream = Streaming.createDocxStream({
      headers,
      sectionProperties: { headers: [{ type: "default", rId: "" }] }
    });
    stream.add(Build.textParagraph("body"));
    const bytes = await stream.finalize();
    const files = await extract(bytes);

    const documentXml = files.get("word/document.xml")!;
    const documentRels = files.get("word/_rels/document.xml.rels")!;
    // Pull the rId emitted by headerReference and assert it resolves in rels.
    const m = documentXml.match(/<w:headerReference[^>]*r:id="([^"]+)"/);
    expect(m).toBeTruthy();
    const rId = m![1]!;
    expect(rId).not.toBe("");
    expect(documentRels).toContain(`Id="${rId}"`);
  });

  it("synthesises section-property header refs when caller did not author any", async () => {
    const headers = new Map<string, HeaderDef>([
      ["default", { content: { children: [{ type: "paragraph", children: [] } as never] } }]
    ]);
    const stream = Streaming.createDocxStream({ headers });
    stream.add(Build.textParagraph("body"));
    const bytes = await stream.finalize();
    const files = await extract(bytes);

    const documentXml = files.get("word/document.xml")!;
    expect(documentXml).toMatch(/<w:headerReference[^>]*r:id="[^"]+"/);
  });

  it("emits a watermark even when explicit headers are also provided", async () => {
    // Regression: previously the watermark branch was guarded with
    // `!headers`, so callers that provided both lost the watermark
    // silently.
    const headers = new Map<string, HeaderDef>([
      ["default", { content: { children: [{ type: "paragraph", children: [] } as never] } }]
    ]);
    const stream = Streaming.createDocxStream({
      headers,
      watermark: { type: "text", text: "DRAFT" }
    });
    stream.add(Build.textParagraph("body"));
    const bytes = await stream.finalize();
    const files = await extract(bytes);

    // Two header XML parts: the explicit one + the watermark.
    const headerKeys = [...files.keys()].filter(p => /^word\/header\d+\.xml$/.test(p));
    expect(headerKeys.length).toBe(2);
  });

  // ===========================================================================
  // OOXML compliance for empty / minimal streams (regression for Word
  // refusing to open packages that lacked the bare-minimum body / sectPr /
  // Content_Types entries).
  // ===========================================================================

  it("finalising an empty stream produces a body with at least one <w:p>", async () => {
    // Word rejects <w:body/> with no <w:p>. Make sure the writer fills
    // the gap when the caller streamed zero elements.
    const stream = Streaming.createDocxStream();
    const bytes = await stream.finalize();
    const files = await extract(bytes);
    const docXml = files.get("word/document.xml")!;
    expect(docXml).toBeDefined();
    expect(/<w:p\s*\/>/.test(docXml) || /<w:p>\s*<\/w:p>/.test(docXml)).toBe(true);
  });

  it("always emits a final <w:sectPr> even when sectionProperties is omitted", async () => {
    // ECMA-376 CT_Body requires a final sectPr; without it Word does
    // not know the page geometry and falls back to error-prone defaults.
    const stream = Streaming.createDocxStream();
    stream.add(Build.textParagraph("body"));
    const bytes = await stream.finalize();
    const docXml = (await extract(bytes)).get("word/document.xml")!;
    expect(docXml).toContain("<w:sectPr");
    expect(docXml).toContain("<w:pgSz");
  });

  it("[Content_Types].xml declares core/app properties Overrides", async () => {
    // _rels/.rels references docProps/core.xml + docProps/app.xml on
    // every package. If the corresponding Override entries are missing,
    // Word/LibreOffice refuse to open the file at the OPC layer.
    const stream = Streaming.createDocxStream();
    stream.add(Build.textParagraph("body"));
    const bytes = await stream.finalize();
    const ct = (await extract(bytes)).get("[Content_Types].xml")!;
    expect(ct).toContain('PartName="/docProps/core.xml"');
    expect(ct).toContain("application/vnd.openxmlformats-package.core-properties+xml");
    expect(ct).toContain('PartName="/docProps/app.xml"');
    expect(ct).toContain("application/vnd.openxmlformats-officedocument.extended-properties+xml");
  });
});
