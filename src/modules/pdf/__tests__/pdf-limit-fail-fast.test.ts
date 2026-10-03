/**
 * Once a resource limit is hit, a document refuses every further operation,
 * so readers that tolerate malformed objects stop at their next step instead
 * of decoding the rest of a hostile file.
 *
 * Asserted through the document's own contract rather than by mocking the
 * decoder: the node suite runs with `isolate: false`, where `vi.mock` of an
 * already-loaded module is not reliable.
 */

import { PdfLimitExceededError } from "@pdf/errors";
import { PdfDocument } from "@pdf/reader/pdf-document";
import type { PdfStream } from "@pdf/reader/pdf-parser";
import { describe, it, expect } from "vitest";

import { Pdf } from "../index";

/** A classic-xref PDF whose objects are numbered from 1. */
function rawPdf(objects: string[]): Uint8Array {
  let pdf = "%PDF-1.4\n";
  const rows: string[] = [];
  objects.forEach((body, i) => {
    rows.push(`${String(pdf.length).padStart(10, "0")} 00000 n \n`);
    pdf += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${rows.join("")}`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

/** `pageCount` pages, each with its own content stream that decodes to 5000 bytes. */
function manyBombPages(pageCount: number): Uint8Array {
  const bomb = "00".repeat(5000) + ">";
  const objects = ["<< /Type /Catalog /Pages 2 0 R >>", ""];
  const kids: string[] = [];
  for (let i = 0; i < pageCount; i++) {
    const pageNum = objects.length + 1;
    kids.push(`${pageNum} 0 R`);
    objects.push(
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents ${pageNum + 1} 0 R >>`
    );
    objects.push(
      `<< /Length ${bomb.length} /Filter /ASCIIHexDecode >>\nstream\n${bomb}\nendstream`
    );
  }
  objects[1] = `<< /Type /Pages /Kids [${kids.join(" ")}] /Count ${pageCount} >>`;
  return rawPdf(objects);
}

describe("limit errors fail fast", () => {
  /** Page `i`'s content stream object number in `manyBombPages`. */
  const contentObj = (i: number) => 4 + 2 * i;

  it("poisons the document: every later operation throws the same error", () => {
    const doc = new PdfDocument(manyBombPages(20), 100);
    const stream = (i: number) => doc.resolve(contentObj(i)) as PdfStream;
    const first = stream(0);
    let hit: unknown;
    try {
      doc.getStreamData(first, contentObj(0));
    } catch (err) {
      hit = err;
    }
    expect(hit).toBeInstanceOf(PdfLimitExceededError);
    expect(doc.limitError).toBe(hit);
    // Nothing further is decoded or resolved — not even a fresh object.
    expect(() => doc.resolve(contentObj(5))).toThrow(hit as Error);
    expect(() => doc.deref({ objNum: 1, gen: 0 } as never)).toThrow(hit as Error);
    expect(() => doc.getStreamData(first, contentObj(0))).toThrow(hit as Error);
  });

  it("Pdf.read rejects instead of returning truncated pages", async () => {
    await expect(Pdf.read(manyBombPages(20), { maxDecodedBytes: 100 })).rejects.toThrow(
      PdfLimitExceededError
    );
  });

  it.each([Number.NaN, -1])("rejects maxDecodedBytes %s as a RangeError", async bad => {
    await expect(Pdf.read(manyBombPages(1), { maxDecodedBytes: bad })).rejects.toThrow(RangeError);
  });

  it("accepts Infinity as an unbounded maxDecodedBytes", async () => {
    const result = await Pdf.read(manyBombPages(1), { maxDecodedBytes: Infinity });
    expect(result.pages).toHaveLength(1);
  });

  it("still reads every page when no limit is hit", async () => {
    const result = await Pdf.read(manyBombPages(3));
    expect(result.pages).toHaveLength(3);
  });
});
