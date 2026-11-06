/**
 * Hand-assembled PDFs whose only cross-reference data is a `/Type /XRef` stream — no
 * `xref` table, no `trailer` keyword. The builder writes classic tables, so these are
 * the fixtures for the other format.
 */

import { Pdf } from "documonster/pdf";

import { generateSelfSignedCertificate } from "./self-signed-certificate.js";

/** `bodies[i]` is the body of object `i + 1`; object 1 is the catalog. */
export function buildXrefStreamPdf(bodies: readonly string[]): Uint8Array {
  let text = "%PDF-1.7\n";
  const offsets: number[] = [];
  bodies.forEach((body, i) => {
    offsets.push(text.length);
    text += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xrefNum = bodies.length + 1;
  const xrefAt = text.length;
  const rows: Array<[number, number, number]> = [
    [0, 0, 0xffff],
    ...offsets.map((offset): [number, number, number] => [1, offset, 0]),
    [1, xrefAt, 0]
  ];
  const data = Buffer.alloc(rows.length * 7);
  rows.forEach(([type, offset, gen], i) => {
    data.writeUInt8(type, i * 7);
    data.writeUInt32BE(offset, i * 7 + 1);
    data.writeUInt16BE(gen, i * 7 + 5);
  });
  return new Uint8Array(
    Buffer.concat([
      Buffer.from(
        text +
          `${xrefNum} 0 obj\n<< /Type /XRef /Size ${xrefNum + 1} /W [1 4 2] /Root 1 0 R /Length ${data.length} >>\nstream\n`,
        "latin1"
      ),
      data,
      Buffer.from(`\nendstream\nendobj\nstartxref\n${xrefAt}\n%%EOF\n`, "latin1")
    ])
  );
}

const PAGE_CONTENT = "<< /Length 0 >>\nstream\n\nendstream";

/** One page, one text field `fullName`. */
export function xrefStreamForm(): Uint8Array {
  return buildXrefStreamPdf([
    "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] >> >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents 5 0 R /Annots [4 0 R] >>",
    "<< /Type /Annot /Subtype /Widget /FT /Tx /T (fullName) /Rect [60 690 300 712] /P 3 0 R /DA (/Helv 12 Tf 0 g) >>",
    PAGE_CONTENT
  ]);
}

/** `pages` blank pages, unsigned. */
export function xrefStreamPages(pages: number): Uint8Array {
  const kids = Array.from({ length: pages }, (_, i) => `${3 + i * 2} 0 R`).join(" ");
  const bodies = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    `<< /Type /Pages /Kids [${kids}] /Count ${pages} >>`
  ];
  for (let i = 0; i < pages; i++) {
    bodies.push(
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents ${4 + i * 2} 0 R >>`,
      PAGE_CONTENT
    );
  }
  return buildXrefStreamPdf(bodies);
}

/**
 * Two pages and a signature over the whole file, written into an xref-stream PDF by
 * `Pdf.Signature.fillPlaceholder` — which fills a placeholder in place, so the file keeps its format.
 */
export async function signedXrefStreamPdf(): Promise<Uint8Array> {
  const { dictString } = Pdf.Signature.buildDictPlaceholder({ name: "Fixture Signer" });
  const unsigned = buildXrefStreamPdf([
    "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] /SigFlags 3 >> >>",
    "<< /Type /Pages /Kids [3 0 R 7 0 R] /Count 2 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents 4 0 R /Annots [5 0 R] >>",
    PAGE_CONTENT,
    "<< /Type /Annot /Subtype /Widget /FT /Sig /T (Signature1) /Rect [0 0 0 0] /F 4 /P 3 0 R /V 6 0 R >>",
    dictString,
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents 8 0 R >>",
    PAGE_CONTENT
  ]);
  const { certificate, privateKey } = await generateSelfSignedCertificate("Fixture Signer");
  return Pdf.Signature.fillPlaceholder(unsigned, { certificate, privateKey });
}
