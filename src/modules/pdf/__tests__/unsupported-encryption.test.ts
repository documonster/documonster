/**
 * An encryption the reader recognises but has no handler for is a `PdfUnsupportedFeatureError` —
 * still a `PdfStructureError`, the class it was reported as before.
 */
import { Pdf, PdfStructureError, PdfUnsupportedFeatureError } from "@pdf/index";
import { describe, expect, it } from "vitest";

function pdfWithEncryptFilter(filter: string): Uint8Array {
  const objects = [
    "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
    "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n",
    "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>\nendobj\n",
    `4 0 obj\n<< /Filter /${filter} /V 1 /R 2 /O <00> /U <00> /P -4 >>\nendobj\n`
  ];
  let body = "%PDF-1.4\n";
  const offsets: number[] = [];
  for (const object of objects) {
    offsets.push(body.length);
    body += object;
  }
  const xref = body.length;
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets) {
    body += `${String(offset).padStart(10, "0")} 00000 n \n`;
  }
  body +=
    `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R /Encrypt 4 0 R /ID [<01> <01>] >>\n` +
    `startxref\n${xref}\n%%EOF\n`;
  return new TextEncoder().encode(body);
}

describe("unsupported PDF encryption", () => {
  it("is a PdfUnsupportedFeatureError, and still a PdfStructureError", async () => {
    const error = await Promise.resolve()
      .then(() => Pdf.extract(pdfWithEncryptFilter("Adobe.PubSec")))
      .then(
        () => undefined,
        (e: unknown) => e
      );
    expect(error).toBeInstanceOf(PdfUnsupportedFeatureError);
    expect(error).toBeInstanceOf(PdfStructureError);
    expect((error as Error).message).toContain("Unsupported encryption filter: Adobe.PubSec");
  });
});
