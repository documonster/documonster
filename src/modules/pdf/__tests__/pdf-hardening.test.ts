import { zlibSync } from "@archive/compression/compress";
import {
  PdfLimitExceededError,
  PdfSignatureInvalidationError,
  PdfStructureError
} from "@pdf/errors";
import { PdfDocument } from "@pdf/reader/pdf-document";
import { parseObject } from "@pdf/reader/pdf-parser";
import { PdfTokenizer } from "@pdf/reader/pdf-tokenizer";
import { decodeStreamFilters } from "@pdf/reader/stream-filters";
import { describe, it, expect } from "vitest";

import { generateSelfSignedCertificate } from "../examples/utils/self-signed-certificate";
import { Pdf } from "../index";

const enc = new TextEncoder();
const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);

/** A classic-xref PDF whose objects are numbered from 1, with optional generations. */
function rawPdf(objects: Array<string | { gen: number; body: string }>): string {
  let pdf = "%PDF-1.4\n";
  const rows: string[] = [];
  objects.forEach((obj, i) => {
    const { gen, body } = typeof obj === "string" ? { gen: 0, body: obj } : obj;
    rows.push(`${String(pdf.length).padStart(10, "0")} ${String(gen).padStart(5, "0")} n \n`);
    pdf += `${i + 1} ${gen} obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${rows.join("")}`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return pdf;
}

function parse(text: string): unknown {
  return parseObject(new PdfTokenizer(enc.encode(text)));
}

describe("parser nesting depth", () => {
  it("rejects 100k nested arrays with PdfStructureError instead of overflowing", () => {
    const deep = "[".repeat(100_000) + "]".repeat(100_000);
    expect(() => parse(deep)).toThrow(PdfStructureError);
  });

  it("rejects deeply nested dictionaries", () => {
    const deep = "<< /A ".repeat(100_000) + "1" + " >>".repeat(100_000);
    expect(() => parse(deep)).toThrow(/nested deeper/);
  });

  it("accepts nesting up to the limit", () => {
    const n = 512; // MAX_PDF_NESTING_DEPTH
    expect(() => parse("[".repeat(n) + "]".repeat(n))).not.toThrow();
    expect(() => parse("[".repeat(n + 1) + "]".repeat(n + 1))).toThrow(PdfStructureError);
  });
});

describe("decoded stream limits", () => {
  const filterDict = (filter: unknown): Map<string, unknown> => new Map([["Filter", filter]]);

  it("caps FlateDecode output", () => {
    const bomb = zlibSync(new Uint8Array(1_000_000));
    expect(() => decodeStreamFilters(bomb, filterDict("FlateDecode") as never, 1000)).toThrow(
      PdfStructureError
    );
    expect(decodeStreamFilters(bomb, filterDict("FlateDecode") as never).length).toBe(1_000_000);
  });

  it("aborts RunLengthDecode early", () => {
    // Each 2-byte run (0x81, b) expands to 128 bytes
    const runs = new Uint8Array(20_000);
    for (let i = 0; i < runs.length; i += 2) {
      runs[i] = 0x81;
    }
    expect(() => decodeStreamFilters(runs, filterDict("RunLengthDecode") as never, 4096)).toThrow(
      /RunLengthDecode/
    );
  });

  it("aborts LZWDecode early", () => {
    // 9-bit code 0 repeated: each emits one byte, then grows the table
    const data = new Uint8Array(10_000);
    expect(() => decodeStreamFilters(data, filterDict("LZWDecode") as never, 100)).toThrow(
      /LZWDecode/
    );
  });

  it("caps ASCII85Decode, including 'z' expansion", () => {
    // Each 'z' is one input byte expanding to four zero bytes
    const zs = enc.encode("z".repeat(10_000) + "~>");
    expect(decodeStreamFilters(zs, filterDict("ASCII85Decode") as never).length).toBe(40_000);
    expect(() => decodeStreamFilters(zs, filterDict("ASCII85Decode") as never, 1000)).toThrow(
      PdfLimitExceededError
    );
    expect(() =>
      decodeStreamFilters(enc.encode('87cURD]i,"Ebo80~>'), filterDict("A85") as never, 5)
    ).toThrow(/ASCII85Decode/);
    const ok = decodeStreamFilters(enc.encode('87cURD]i,"Ebo80~>'), filterDict("A85") as never, 12);
    expect(new TextDecoder().decode(ok)).toBe("Hello World!");
  });

  it("caps ASCIIHexDecode", () => {
    const hex = enc.encode("41 42 43 4>");
    expect(latin1(decodeStreamFilters(hex, filterDict("AHx") as never))).toBe("ABC@");
    expect(latin1(decodeStreamFilters(hex, filterDict("AHx") as never, 4))).toBe("ABC@");
    expect(() => decodeStreamFilters(hex, filterDict("ASCIIHexDecode") as never, 3)).toThrow(
      PdfLimitExceededError
    );
    expect(() =>
      decodeStreamFilters(enc.encode("00".repeat(5000)), filterDict("AHx") as never, 100)
    ).toThrow(/ASCIIHexDecode/);
  });

  it("rejects an overlong filter chain", () => {
    const chain = Array.from({ length: 17 }, () => "ASCIIHexDecode");
    expect(() => decodeStreamFilters(enc.encode("00>"), filterDict(chain) as never)).toThrow(
      /filter chain/
    );
  });

  it("is plumbed through Pdf.read", async () => {
    const doc = new Pdf.Builder();
    doc.addPage().drawText("x".repeat(2000), { x: 10, y: 700, fontSize: 4 });
    const bytes = await doc.build();
    // A limit hit is a hard failure, never a page warning with a truncated result
    await expect(Pdf.read(bytes, { maxDecodedBytes: 16 })).rejects.toThrow(PdfLimitExceededError);
    await expect(Pdf.read(bytes, { maxDecodedBytes: 16 })).rejects.toThrow(/maxDecodedBytes/);
    const full = await Pdf.read(bytes);
    expect(full.pages[0].warnings).toEqual([]);
  });

  it("is not swallowed by the editor's tolerant reads", async () => {
    // A page whose content stream carries an overlong filter chain: load
    // succeeds (content is not decoded), and save must refuse rather than
    // write the page with its content silently dropped.
    const chain = Array.from({ length: 17 }, () => "/AHx").join(" ");
    const pdf = rawPdf([
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
      `<< /Length 3 /Filter [${chain}] >>\nstream\n00>\nendstream`
    ]);

    const editor = Pdf.Editor.load(enc.encode(pdf));
    await expect(editor.save()).rejects.toThrow(PdfLimitExceededError);
    editor.getPage(0).drawText("x", { x: 10, y: 10 });
    await expect(editor.saveIncremental()).rejects.toThrow(PdfLimitExceededError);
  });
});

describe("incremental update generations", () => {
  it("rewrites an existing object under its own generation", async () => {
    const pdf = rawPdf([
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 2 R] /Count 1 >>",
      { gen: 2, body: "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>" }
    ]);
    const editor = Pdf.Editor.load(enc.encode(pdf));
    editor.getPage(0).drawText("Updated", { x: 10, y: 100, fontSize: 12 });
    const out = await editor.saveIncremental();
    const appended = latin1(out.subarray(pdf.length));
    expect(appended).toContain("3 2 obj");
    expect(appended).toMatch(/ 00002 n \n/);
    const read = await Pdf.read(out);
    expect(read.text).toContain("Updated");
  });
});

describe("incremental form fill", () => {
  it("numbers a text field's new appearance stream in the original file's space", async () => {
    const pdf = rawPdf([
      "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] >> >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Annots [4 0 R] >>",
      "<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /V (old) /Rect [10 10 150 30] /P 3 0 R >>"
    ]);
    const editor = Pdf.Editor.load(enc.encode(pdf));
    editor.setFormField("name", "new value");
    const out = await editor.saveIncremental();
    const appended = latin1(out.subarray(pdf.length));
    const apRef = /\/AP << \/N (\d+) 0 R >>/.exec(appended);
    expect(apRef).not.toBeNull();
    const apNum = Number(apRef![1]);
    // A fresh number past the original /Size (5), and the object is appended.
    expect(apNum).toBeGreaterThanOrEqual(5);
    expect(appended).toContain(`${apNum} 0 obj`);
    const doc = new PdfDocument(out);
    const ap = doc.resolve(apNum);
    expect(ap && typeof ap === "object" && "dict" in ap && ap.dict.get("Subtype")).toBe("Form");
  });
});

async function buildSigned(): Promise<Uint8Array> {
  const { certificate, privateKey } = await generateSelfSignedCertificate("Hardening");
  const doc = new Pdf.Builder();
  doc.addPage().drawText("Signed", { x: 72, y: 750, fontSize: 14 });
  const unsigned = await doc.build();
  return Pdf.Editor.load(unsigned).sign({ certificate, privateKey });
}

describe("signature preservation", () => {
  it("detects signatures", async () => {
    const signed = await buildSigned();
    expect(Pdf.Editor.load(signed).hasSignatures).toBe(true);
    const builder = new Pdf.Builder();
    builder.addPage();
    const plain = await builder.build();
    expect(Pdf.Editor.load(plain).hasSignatures).toBe(false);
  });

  it("refuses a full save of a signed document unless opted in", async () => {
    const signed = await buildSigned();
    const editor = Pdf.Editor.load(signed);
    await expect(editor.save()).rejects.toThrow(PdfSignatureInvalidationError);
    await expect(editor.save({ invalidateSignatures: true })).resolves.toBeInstanceOf(Uint8Array);
  });

  it("refuses to re-sign or fall back from saveIncremental", async () => {
    const signed = await buildSigned();
    const { certificate, privateKey } = await generateSelfSignedCertificate("Second");
    await expect(Pdf.Editor.load(signed).sign({ certificate, privateKey })).rejects.toThrow(
      PdfSignatureInvalidationError
    );
    const editor = Pdf.Editor.load(signed);
    editor.addPage();
    await expect(editor.saveIncremental()).rejects.toThrow(PdfSignatureInvalidationError);
  });

  it("saveIncremental keeps the signed bytes as a prefix and the signature valid", async () => {
    const signed = await buildSigned();
    const editor = Pdf.Editor.load(signed);
    editor.getPage(0).drawText("Annotated", { x: 72, y: 700, fontSize: 12 });
    const updated = await editor.saveIncremental();
    expect(updated.length).toBeGreaterThan(signed.length);
    expect(updated.subarray(0, signed.length)).toEqual(signed);

    const text = latin1(signed);
    const contents = text.match(/\/Contents\s*<([0-9a-fA-F]+)>/)![1];
    const br = text.match(/\/ByteRange\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]/)!;
    const range = [1, 2, 3, 4].map(i => parseInt(br[i], 10)) as [number, number, number, number];
    const result = await Pdf.verifySignature(updated, contents, range);
    expect(result.valid).toBe(true);
    expect(result.coversWholeFile).toBe(false);
  });
});

const PLAIN_OBJECTS = [
  "<< /Type /Catalog /Pages 2 0 R >>",
  "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
  "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>"
];

const FILE_ID = "<0123456789abcdef0123456789abcdef>";

/**
 * A PDF whose cross-reference data is an xref stream. `objs` are numbered
 * from 1; the xref stream takes the next number.
 */
function buildXrefStreamPdf(objs: string[] = PLAIN_OBJECTS, trailerExtra = ""): Uint8Array {
  const parts: string[] = ["%PDF-1.5\n"];
  const offsets: number[] = [];
  let len = parts[0].length;
  objs.forEach((body, i) => {
    offsets.push(len);
    const s = `${i + 1} 0 obj\n${body}\nendobj\n`;
    parts.push(s);
    len += s.length;
  });
  const xrefOffset = len;
  const size = objs.length + 2;
  const rows = new Uint8Array(size * 6);
  const entries: Array<[number, number, number]> = [
    [0, 0, 255],
    ...offsets.map(o => [1, o, 0] as [number, number, number]),
    [1, xrefOffset, 0]
  ];
  entries.forEach(([t, o, g], i) => {
    rows.set([t, (o >>> 24) & 255, (o >>> 16) & 255, (o >>> 8) & 255, o & 255, g], i * 6);
  });
  const head = enc.encode(
    parts.join("") +
      `${size - 1} 0 obj\n<< /Type /XRef /Size ${size} /W [1 4 1] /Root 1 0 R${trailerExtra} ` +
      `/Length ${rows.length} >>\nstream\n`
  );
  const tail = enc.encode(`\nendstream\nendobj\nstartxref\n${xrefOffset}\n%%EOF\n`);
  const out = new Uint8Array(head.length + rows.length + tail.length);
  out.set(head);
  out.set(rows, head.length);
  out.set(tail, head.length + rows.length);
  return out;
}

const FORM_OBJECTS = [
  "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] >> >>",
  "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
  "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Annots [4 0 R] >>",
  "<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /V (old) /Rect [10 10 150 30] /P 3 0 R >>"
];

/** Last `startxref` value and the dictionary of the section it points at. */
function lastXrefSection(bytes: Uint8Array): { offset: number; dict: string } {
  const text = latin1(bytes);
  const offset = parseInt(text.slice(text.lastIndexOf("startxref") + 9).trim(), 10);
  const dict = text.slice(offset, text.indexOf("stream", offset));
  return { offset, dict };
}

describe("incremental update of an xref-stream file", () => {
  it("appends an xref stream section instead of rewriting", async () => {
    const original = buildXrefStreamPdf();
    const editor = Pdf.Editor.load(original);
    editor.getPage(0).drawText("Hi", { x: 10, y: 100, fontSize: 12 });
    const updated = await editor.saveIncremental();
    expect(updated.subarray(0, original.length)).toEqual(original);
    const appended = latin1(updated.subarray(original.length));
    expect(appended).toContain("/Type /XRef");
    expect(appended).toContain("/Prev ");
    expect(appended).not.toContain("trailer");

    const read = await Pdf.read(updated);
    expect(read.pages[0].text).toContain("Hi");
  });
});

describe("incremental update that only modifies existing objects (xref stream)", () => {
  it("numbers the xref stream past the original /Size and keeps every object", async () => {
    const original = buildXrefStreamPdf(FORM_OBJECTS, ` /ID [${FILE_ID} ${FILE_ID}]`);
    const originalSize = FORM_OBJECTS.length + 2; // objects 1..4, xref stream 5, plus 0
    const editor = Pdf.Editor.load(original);
    editor.setFormField("name", "new value");
    const updated = await editor.saveIncremental();
    expect(updated.subarray(0, original.length)).toEqual(original);

    const { offset, dict } = lastXrefSection(updated);
    expect(offset).toBeGreaterThan(original.length);
    const header = latin1(updated)
      .slice(offset)
      .match(/^(\d+) 0 obj/)!;
    const xrefObjNum = parseInt(header[1], 10);
    expect(xrefObjNum).toBeGreaterThanOrEqual(originalSize);
    expect(dict).toContain(`/Size ${xrefObjNum + 1}`);
    // ID[0] (and here ID[1]) carried over unchanged
    expect(dict).toContain(`/ID [${FILE_ID} ${FILE_ID}]`);
    // Only modified objects (page 3, widget 4), the widget's new appearance
    // stream and the xref stream itself are listed
    const index = dict
      .match(/\/Index \[([^\]]*)\]/)![1]
      .trim()
      .split(/\s+/)
      .map(Number);
    const listed: number[] = [];
    for (let i = 0; i < index.length; i += 2) {
      for (let n = 0; n < index[i + 1]; n++) {
        listed.push(index[i] + n);
      }
    }
    expect(listed).toEqual([3, 4, originalSize, xrefObjNum]);
    expect(xrefObjNum).toBe(originalSize + 1);

    const doc = new PdfDocument(updated);
    expect(doc.trailer.get("Size")).toBe(xrefObjNum + 1);
    for (let n = 1; n <= FORM_OBJECTS.length; n++) {
      expect(doc.resolve(n)).toBeInstanceOf(Map);
    }
    // The original xref stream object is still reachable, not shadowed
    expect(doc.resolve(originalSize - 1)).toMatchObject({ type: "stream" });

    const read = await Pdf.read(updated);
    expect(read.formFields.find(f => f.name === "name")?.value).toBe("new value");
  });
});

describe("encrypted documents", () => {
  async function buildEncrypted(): Promise<Uint8Array> {
    const doc = new Pdf.Builder();
    doc.setEncryption({ userPassword: "user", ownerPassword: "owner" });
    doc.addPage().drawText("Secret body", { x: 72, y: 700, fontSize: 12 });
    return doc.build();
  }

  it("saveIncremental encrypts the appended objects with the document's own handler", async () => {
    const encrypted = await buildEncrypted();
    const editor = Pdf.Editor.load(encrypted, { password: "user" });
    editor.getPage(0).drawText("Overlay", { x: 72, y: 600, fontSize: 12 });
    const updated = await editor.saveIncremental();
    // The original bytes are untouched and the update claims no new handler.
    expect(updated.subarray(0, encrypted.length)).toEqual(encrypted);
    const appended = latin1(updated.subarray(encrypted.length));
    expect(appended).not.toContain("Overlay");
    const read = await Pdf.read(updated, { password: "user" });
    expect(read.pages[0].text).toContain("Secret body");
    expect(read.pages[0].text).toContain("Overlay");
  });

  it("save() writes a decrypted copy that does not claim /Encrypt", async () => {
    const encrypted = await buildEncrypted();
    const editor = Pdf.Editor.load(encrypted, { password: "user" });
    editor.getPage(0).drawText("Overlay", { x: 72, y: 600, fontSize: 12 });
    const saved = await editor.save();
    expect(latin1(saved)).not.toContain("/Encrypt");
    const read = await Pdf.read(saved);
    expect(read.pages[0].text).toContain("Secret body");
    expect(read.pages[0].text).toContain("Overlay");
  });
});

describe("signature detection", () => {
  const signedValue = "<< /Type /Sig /ByteRange [0 10 20 10] /Contents <00> >>";

  function withFields(acroForm: string, extra: string[]): Uint8Array {
    return buildXrefStreamPdf([
      `<< /Type /Catalog /Pages 2 0 R /AcroForm ${acroForm} >>`,
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
      ...extra
    ]);
  }

  it("ignores /SigFlags without a signed field", () => {
    const pdf = withFields("<< /Fields [] /SigFlags 3 >>", []);
    expect(Pdf.Editor.load(pdf).hasSignatures).toBe(false);
  });

  it("ignores an unsigned signature field", () => {
    const pdf = withFields("<< /Fields [4 0 R] /SigFlags 3 >>", ["<< /FT /Sig /T (s) >>"]);
    expect(Pdf.Editor.load(pdf).hasSignatures).toBe(false);
  });

  it("requires both /ByteRange and /Contents", () => {
    const pdf = withFields("<< /Fields [4 0 R] >>", [
      "<< /FT /Sig /T (s) /V << /ByteRange [0 1 2 3] >> >>"
    ]);
    expect(Pdf.Editor.load(pdf).hasSignatures).toBe(false);
  });

  it("finds a signed kid whose /FT is inherited from its parent", () => {
    const pdf = withFields("<< /Fields [4 0 R] >>", [
      "<< /FT /Sig /T (parent) /Kids [5 0 R] >>",
      `<< /Parent 4 0 R /T (kid) /V ${signedValue} >>`
    ]);
    expect(Pdf.Editor.load(pdf).hasSignatures).toBe(true);
  });

  it("guards the inline-page fallback inside saveIncremental", async () => {
    const pdf = buildXrefStreamPdf([
      "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [3 0 R] >> >>",
      "<< /Type /Pages /Kids [<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>] /Count 1 >>",
      `<< /FT /Sig /T (s) /V ${signedValue} >>`
    ]);
    const editor = Pdf.Editor.load(pdf);
    expect(editor.hasSignatures).toBe(true);
    editor.getPage(0).drawText("x", { x: 10, y: 10, fontSize: 12 });
    await expect(editor.saveIncremental()).rejects.toThrow(PdfSignatureInvalidationError);
    const rewritten = await editor.saveIncremental({ invalidateSignatures: true });
    const read = await Pdf.read(rewritten);
    expect(read.pages[0].text).toContain("x");
  });
});

describe("creation date", () => {
  it("writes a supplied creation date, making builds reproducible", async () => {
    const creationDate = new Date(Date.UTC(2020, 0, 2, 3, 4, 5));
    const build = async (): Promise<Uint8Array> => {
      const doc = new Pdf.Builder();
      doc.setMetadata({ title: "T", creationDate, modDate: creationDate });
      doc.addPage().drawText("same", { x: 10, y: 700, fontSize: 12 });
      return doc.build();
    };
    const a = await build();
    expect(latin1(a)).toContain("/CreationDate (D:20200102030405Z)");
    expect(latin1(a)).toContain("/ModDate (D:20200102030405Z)");
    await new Promise(r => setTimeout(r, 1100));
    expect(await build()).toEqual(a);
  });
});
