/**
 * A date the reader cannot read is absent (`undefined`), on every path — the lenient scan of an XMP packet that is not
 * namespaced XML used to hand out an `Invalid Date`.
 */
import { Pdf } from "@pdf/index";
import { describe, expect, it } from "vitest";

function pdfWithXmp(xmp: string): Uint8Array {
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R /Metadata 4 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>",
    `<< /Type /Metadata /Subtype /XML /Length ${xmp.length} >>\nstream\n${xmp}\nendstream`
  ];
  let text = "%PDF-1.4\n";
  const offsets: number[] = [];
  objects.forEach((body, i) => {
    offsets.push(text.length);
    text += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = text.length;
  text +=
    `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n` +
    offsets.map(offset => `${String(offset).padStart(10, "0")} 00000 n \n`).join("") +
    `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new TextEncoder().encode(text);
}

describe("PdfMetadata dates", () => {
  it("are absent when the lenient XMP scan finds text that is not a date", async () => {
    const { metadata } = await Pdf.extract(
      pdfWithXmp("<xmp:CreateDate>not-a-date</xmp:CreateDate><xmp:ModifyDate>x</xmp:ModifyDate>")
    );
    expect(metadata.creationDate).toBeUndefined();
    expect(metadata.modDate).toBeUndefined();
  });

  it("are read when the lenient XMP scan finds one", async () => {
    const { metadata } = await Pdf.extract(
      pdfWithXmp("<xmp:CreateDate>2024-05-06T07:08:09Z</xmp:CreateDate>")
    );
    expect(metadata.creationDate?.toISOString()).toBe("2024-05-06T07:08:09.000Z");
  });
});
