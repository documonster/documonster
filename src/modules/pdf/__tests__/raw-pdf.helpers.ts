/**
 * Assemble a PDF from hand-written object bodies, with either a classic cross-reference
 * table or a cross-reference stream. For fixtures the builder cannot produce: a field
 * tree spanning pages, a field with no widget, a file whose only xref is a stream.
 */

export interface RawPdfOptions {
  /** Write a `/Type /XRef` stream instead of an `xref` table and `trailer`. */
  readonly xrefStream?: boolean;
  /** Object number of the catalog. Defaults to 1. */
  readonly root?: number;
}

/** `bodies[i]` is the body of object `i + 1` (without `obj`/`endobj`). */
export function buildRawPdf(bodies: readonly string[], options: RawPdfOptions = {}): Uint8Array {
  const root = options.root ?? 1;
  let out = "%PDF-1.7\n%\u00e2\u00e3\u00cf\u00d3\n";
  const offsets: number[] = [];
  bodies.forEach((body, i) => {
    offsets.push(out.length);
    out += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });

  if (!options.xrefStream) {
    const xrefAt = out.length;
    out += `xref\n0 ${bodies.length + 1}\n0000000000 65535 f \n`;
    for (const offset of offsets) {
      out += `${String(offset).padStart(10, "0")} 00000 n \n`;
    }
    out += `trailer\n<< /Size ${bodies.length + 1} /Root ${root} 0 R >>\nstartxref\n${xrefAt}\n%%EOF\n`;
    return latin1Bytes(out);
  }

  // The stream lists itself, so it is object n + 1 and /Size is n + 2.
  const xrefNum = bodies.length + 1;
  const xrefAt = out.length;
  const rows: number[][] = [[0, 0, 0xffff]];
  for (const offset of offsets) {
    rows.push([1, offset, 0]);
  }
  rows.push([1, xrefAt, 0]);
  const data = new Uint8Array(rows.length * 7);
  rows.forEach(([type, field2, field3], i) => {
    const at = i * 7;
    data[at] = type;
    data[at + 1] = (field2 >>> 24) & 0xff;
    data[at + 2] = (field2 >>> 16) & 0xff;
    data[at + 3] = (field2 >>> 8) & 0xff;
    data[at + 4] = field2 & 0xff;
    data[at + 5] = (field3 >>> 8) & 0xff;
    data[at + 6] = field3 & 0xff;
  });
  const head = latin1Bytes(
    out +
      `${xrefNum} 0 obj\n<< /Type /XRef /Size ${xrefNum + 1} /W [1 4 2] /Root ${root} 0 R /Length ${data.length} >>\nstream\n`
  );
  const tail = latin1Bytes(`\nendstream\nendobj\nstartxref\n${xrefAt}\n%%EOF\n`);
  const bytes = new Uint8Array(head.length + data.length + tail.length);
  bytes.set(head);
  bytes.set(data, head.length);
  bytes.set(tail, head.length + data.length);
  return bytes;
}

function latin1Bytes(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) {
    bytes[i] = text.charCodeAt(i) & 0xff;
  }
  return bytes;
}

/**
 * A two-page form:
 *
 * - `name` — one widget on page 1;
 * - `group` — a parent field with one widget on each page;
 * - `solo` — one widget on page 2;
 * - `hidden` — a field with a value and no widget on any page.
 */
export function twoPageFormBodies(): string[] {
  const page = (annots: string, contents: number): string =>
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents ${contents} 0 R /Annots [${annots}] >>`;
  const widget = (extra: string, pageObj: number, y: number): string =>
    `<< /Type /Annot /Subtype /Widget /Rect [60 ${y} 300 ${y + 22}] /P ${pageObj} 0 R /DA (/Helv 12 Tf 0 g) ${extra} >>`;
  const content = "<< /Length 0 >>\nstream\n\nendstream";
  return [
    /* 1 */ "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R 8 0 R 9 0 R 10 0 R] >> >>",
    /* 2 */ "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
    /* 3 */ page("6 0 R 8 0 R", 11),
    /* 4 */ page("7 0 R 9 0 R", 12),
    /* 5 */ "<< /FT /Tx /T (group) /Kids [6 0 R 7 0 R] >>",
    /* 6 */ widget("/Parent 5 0 R", 3, 650),
    /* 7 */ widget("/Parent 5 0 R", 4, 650),
    /* 8 */ widget("/FT /Tx /T (name)", 3, 700),
    /* 9 */ widget("/FT /Tx /T (solo)", 4, 700),
    /* 10 */ "<< /FT /Tx /T (hidden) /V (kept) >>",
    /* 11 */ content,
    /* 12 */ content
  ];
}
