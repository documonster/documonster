import { describe, expect, it } from "vitest";

import { buildCmsSignedData, certificateCommonName } from "../core/digital-signature";
import { buildIncremental } from "../core/pdf-writer";
import { findSignatureContents } from "../core/signature-dict";
import { generateSelfSignedCertificate } from "../examples/utils/self-signed-certificate";
import type { Asn1Node } from "../index";
import { Pdf } from "../index";
import { initDecryption } from "../reader/pdf-decrypt";
import { PdfDocument } from "../reader/pdf-document";
import { buildRawPdf } from "./raw-pdf.helpers";

const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);

async function identity(name = "Verifier Test") {
  return generateSelfSignedCertificate(name);
}

describe("Pdf.verifySignatures", () => {
  it("reports a signature written by Builder#sign", async () => {
    const { certificate, privateKey } = await identity();
    const doc = new Pdf.Builder();
    doc.addPage().drawText("Signed", { x: 72, y: 750, fontSize: 14 });
    doc.sign({
      certificate,
      privateKey,
      name: "Alice",
      reason: "Approval",
      location: "Berlin",
      contactInfo: "alice@example.com"
    });
    const signed = await doc.toBuffer();

    const reports = await Pdf.verifySignatures(signed);
    expect(reports).toHaveLength(1);
    const [report] = reports;
    expect(report.intact).toBe(true);
    expect(report.coversWholeFile).toBe(true);
    expect(report.byteRangeValid).toBe(true);
    expect(report.failure).toBeUndefined();
    expect(report.fieldName).toBe("Signature1");
    expect(report.claimedSigner).toBe("Verifier Test");
    expect(report.name).toBe("Alice");
    expect(report.reason).toBe("Approval");
    expect(report.location).toBe("Berlin");
    expect(report.contactInfo).toBe("alice@example.com");
    expect(report.subFilter).toBe("adbe.pkcs7.detached");
    expect(report.signingTime).toBeInstanceOf(Date);
    expect(report.byteRange[0]).toBe(0);
    expect(report.byteRange[2] + report.byteRange[3]).toBe(signed.length);
  });

  it("reports a signature written by Editor#sign", async () => {
    const { certificate, privateKey } = await identity("Editor Signer");
    const doc = new Pdf.Builder();
    doc.addPage().drawText("Invoice", { x: 72, y: 750, fontSize: 14 });
    const signed = await Pdf.Editor.read(await doc.toBuffer()).sign({
      certificate,
      privateKey,
      name: "Bob"
    });

    const [report] = await Pdf.verifySignatures(signed);
    expect(report.intact).toBe(true);
    expect(report.coversWholeFile).toBe(true);
    expect(report.claimedSigner).toBe("Editor Signer");
    expect(report.name).toBe("Bob");
  });

  it("returns an empty list for an unsigned document", async () => {
    const doc = new Pdf.Builder();
    doc.addPage();
    expect(await Pdf.verifySignatures(await doc.toBuffer())).toEqual([]);
  });

  it("fails when a signed byte is changed", async () => {
    const { certificate, privateKey } = await identity();
    const doc = new Pdf.Builder();
    doc.addPage().drawText("Tamper", { x: 72, y: 750, fontSize: 14 });
    doc.sign({ certificate, privateKey });
    const signed = await doc.toBuffer();

    // Change a character of the page text, which sits inside the first signed range.
    const tampered = new Uint8Array(signed);
    const at = latin1(signed).indexOf("%PDF") + 5;
    tampered[at] = tampered[at] === 0x31 ? 0x32 : 0x31;

    const [report] = await Pdf.verifySignatures(tampered);
    expect(report.intact).toBe(false);
    expect(report.failure).toMatch(/digest mismatch/i);
  });

  it("fails when the signature value itself is changed", async () => {
    const { certificate, privateKey } = await identity();
    const doc = new Pdf.Builder();
    doc.addPage();
    doc.sign({ certificate, privateKey });
    const signed = await doc.toBuffer();
    const [{ byteRange }] = await Pdf.verifySignatures(signed);

    const tampered = new Uint8Array(signed);
    // A hex digit well inside the CMS structure.
    const at = byteRange[0] + byteRange[1] + 1 + 400;
    tampered[at] = tampered[at] === 0x61 ? 0x62 : 0x61;

    const [report] = await Pdf.verifySignatures(tampered);
    expect(report.intact).toBe(false);
  });

  it("rejects a /ByteRange that leaves bytes unsigned", async () => {
    const { certificate, privateKey } = await identity();
    const doc = new Pdf.Builder();
    doc.addPage();
    doc.sign({ certificate, privateKey });
    const signed = await doc.toBuffer();
    const [{ byteRange }] = await Pdf.verifySignatures(signed);

    // Same width, so no offset moves: shrink the first range by one byte, which leaves
    // the byte before the `<` outside both ranges.
    const text = latin1(signed);
    const open = text.lastIndexOf("/ByteRange [");
    const close = text.indexOf("]", open);
    const body = text.slice(open + "/ByteRange [".length, close);
    const tokens = [...body.matchAll(/\d+/g)];
    const second = tokens[1];
    const replacement = String(byteRange[1] - 1).padStart(second[0].length, "0");
    expect(replacement.length).toBe(second[0].length);
    const tampered = new Uint8Array(signed);
    const start = open + "/ByteRange [".length + second.index!;
    for (let i = 0; i < replacement.length; i++) {
      tampered[start + i] = replacement.charCodeAt(i);
    }

    const [report] = await Pdf.verifySignatures(tampered);
    expect(report.byteRangeValid).toBe(false);
    expect(report.intact).toBe(false);
  });

  it("keeps a signature valid across an incremental update, but not whole-file", async () => {
    const { certificate, privateKey } = await identity();
    const doc = new Pdf.Builder();
    doc.addPage().drawText("Signed", { x: 72, y: 750, fontSize: 14 });
    doc.sign({ certificate, privateKey });
    const signed = await doc.toBuffer();

    const editor = Pdf.Editor.read(signed);
    editor.getPage(0).drawText("Later", { x: 72, y: 700, fontSize: 12 });
    const updated = await editor.toBuffer({ incremental: true });

    const [report] = await Pdf.verifySignatures(updated);
    expect(report.intact).toBe(true);
    expect(report.byteRangeValid).toBe(true);
    expect(report.coversWholeFile).toBe(false);
  });

  it("agrees with Pdf.Signature.verify on the same signature", async () => {
    const { certificate, privateKey } = await identity();
    const doc = new Pdf.Builder();
    doc.addPage();
    doc.sign({ certificate, privateKey });
    const signed = await doc.toBuffer();
    const [report] = await Pdf.verifySignatures(signed);
    const { byteRange } = report;
    const hex = latin1(signed.subarray(byteRange[0] + byteRange[1] + 1, byteRange[2] - 1));
    const single = await Pdf.Signature.verify(signed, hex, byteRange);
    expect(single.intact).toBe(report.intact);
    expect(single.coversWholeFile).toBe(report.coversWholeFile);
    expect(single.digestAlgorithm).toBe(report.digestAlgorithm);
  });
});

// --- DER helpers for rebuilding a CMS structure -----------------------------------

function derLength(length: number): number[] {
  if (length < 0x80) {
    return [length];
  }
  const bytes: number[] = [];
  for (let l = length; l > 0; l >>= 8) {
    bytes.unshift(l & 0xff);
  }
  return [0x80 | bytes.length, ...bytes];
}

function der(tag: number, body: Uint8Array): Uint8Array {
  return Uint8Array.from([tag, ...derLength(body.length), ...body]);
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

function encodeNode(node: Asn1Node): Uint8Array {
  return der(node.tag, node.bytes);
}

const toHex = (bytes: Uint8Array): string =>
  Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");

function fromHex(hex: string): Uint8Array {
  const out = new Uint8Array(hex.length / 2);
  for (let i = 0; i < out.length; i++) {
    out[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  }
  return out;
}

/** Replace the signature gap's hex digits, keeping its width. */
function patchGap(signed: Uint8Array, byteRange: number[], hex: string): Uint8Array {
  const start = byteRange[0] + byteRange[1] + 1;
  const width = byteRange[2] - 1 - start;
  expect(hex.length).toBeLessThanOrEqual(width);
  const out = new Uint8Array(signed);
  const padded = hex.padEnd(width, "0");
  for (let i = 0; i < width; i++) {
    out[start + i] = padded.charCodeAt(i);
  }
  return out;
}

async function signedSample(name = "Real Signer") {
  const { certificate, privateKey } = await identity(name);
  const doc = new Pdf.Builder();
  doc.addPage().drawText("Signed", { x: 72, y: 750, fontSize: 14 });
  doc.sign({ certificate, privateKey });
  const signed = await doc.toBuffer();
  const [report] = await Pdf.verifySignatures(signed);
  return { signed, report };
}

describe("Pdf.verifySignatures — what a report does and does not claim", () => {
  it("names the certificate the SignerInfo identifies, not the first one embedded", async () => {
    const { signed, report } = await signedSample();
    const decoy = (await identity("Decoy Authority")).certificate;
    const gap = latin1(
      signed.subarray(report.byteRange[0] + report.byteRange[1] + 1, report.byteRange[2] - 1)
    );
    const root = Pdf.Signature.parseAsn1(fromHex(gap));
    const signedData = root.children[1].children[0];
    const rebuilt = signedData.children.map(child =>
      child.tag === 0xa0
        ? der(0xa0, concat([decoy, ...child.children.map(encodeNode)]))
        : encodeNode(child)
    );
    const cms = der(
      0x30,
      concat([encodeNode(root.children[0]), der(0xa0, der(0x30, concat(rebuilt)))])
    );
    const patched = patchGap(signed, report.byteRange, toHex(cms));

    const [after] = await Pdf.verifySignatures(patched);
    expect(after.intact).toBe(true);
    expect(after.claimedSigner).toBe("Real Signer");
    expect(Pdf.Signature.parseAsn1(after.certificate!).children.length).toBeGreaterThan(0);
  });

  it("reports the certificate and its validity period, and no trust verdict", async () => {
    const { report } = await signedSample();
    expect(report.certificate).toBeInstanceOf(Uint8Array);
    expect(report.certificateValidity!.notBefore.getTime()).toBeLessThan(
      report.certificateValidity!.notAfter.getTime()
    );
    expect(report).not.toHaveProperty("valid");
    expect(report).not.toHaveProperty("signer");
  });

  it("accepts white space inside the /Contents hex string", async () => {
    const { signed, report } = await signedSample();
    const start = report.byteRange[0] + report.byteRange[1] + 1;
    const end = report.byteRange[2] - 1;
    const patched = new Uint8Array(signed);
    // Trailing zero padding becomes white space: same signature, spelt differently.
    patched[end - 1] = 0x20;
    patched[end - 2] = 0x0a;
    patched[end - 3] = 0x20;
    patched[end - 4] = 0x0d;
    expect(latin1(patched.subarray(start, start + 4))).toMatch(/^[0-9a-f]{4}$/i);
    const [after] = await Pdf.verifySignatures(patched);
    expect(after.byteRangeValid).toBe(true);
    expect(after.intact).toBe(true);
  });

  it("rejects a /ByteRange with absurd numbers without throwing", async () => {
    const { signed, report } = await signedSample();
    const text = latin1(signed);
    const open = text.lastIndexOf("/ByteRange [");
    const close = text.indexOf("]", open);
    const width = close - (open + "/ByteRange [".length);
    const body = `0 99999999999999999999 1 ${report.byteRange[3]}`.padEnd(width, " ");
    expect(body.length).toBe(width);
    const patched = new Uint8Array(signed);
    for (let i = 0; i < width; i++) {
      patched[open + "/ByteRange [".length + i] = body.charCodeAt(i);
    }
    const [after] = await Pdf.verifySignatures(patched);
    expect(after.byteRangeValid).toBe(false);
    expect(after.intact).toBe(false);
  });

  it("reports every signature of a file signed twice", async () => {
    const first = await identity("First");
    const second = await identity("Second");
    const doc = new Pdf.Builder();
    doc.addPage().drawText("Twice", { x: 72, y: 750, fontSize: 14 });
    doc.sign({ ...first, name: "one" });
    const twice = await appendSecondSignature(await doc.toBuffer(), second);

    const reports = await Pdf.verifySignatures(twice);
    expect(reports.map(r => r.claimedSigner)).toEqual(["First", "Second"]);
    expect(reports.map(r => r.name)).toEqual(["one", "two"]);
    expect(reports.map(r => r.intact)).toEqual([true, true]);
    expect(reports.map(r => r.coversWholeFile)).toEqual([false, true]);
  });

  it("checks /Contents against the signed gap in an encrypted file too", async () => {
    const { certificate, privateKey } = await identity("Encrypted");
    const doc = new Pdf.Builder();
    doc.addPage().drawText("Secret", { x: 72, y: 750, fontSize: 14 });
    doc.setEncryption({ ownerPassword: "owner", userPassword: "" });
    doc.sign({ certificate, privateKey, name: "Enc", reason: "Approval" });
    const signed = await doc.toBuffer();

    const [report] = await Pdf.verifySignatures(signed);
    expect(report.intact).toBe(true);
    expect(report.byteRangeValid).toBe(true);
    expect(report.name).toBe("Enc");
    // The other strings of the signature dictionary are encrypted and decrypt normally.
    expect(report.reason).toBe("Approval");
    expect(report.signingTime).toBeInstanceOf(Date);
    expect(Math.abs(report.signingTime!.getTime() - Date.now())).toBeLessThan(3_600_000);

    // Point the gap at another hex string in the file — the trailer's /ID.
    const text = latin1(signed);
    const idOpen = text.indexOf("<", text.lastIndexOf("/ID"));
    const idClose = text.indexOf(">", idOpen) + 1;
    const open = text.lastIndexOf("/ByteRange [");
    const close = text.indexOf("]", open);
    const width = close - (open + "/ByteRange [".length);
    const body = `0 ${idOpen} ${idClose} ${signed.length - idClose}`.padEnd(width, " ");
    expect(body.length).toBe(width);
    const patched = new Uint8Array(signed);
    for (let i = 0; i < width; i++) {
      patched[open + "/ByteRange [".length + i] = body.charCodeAt(i);
    }
    const [after] = await Pdf.verifySignatures(patched);
    expect(after.byteRangeValid).toBe(false);
    expect(after.intact).toBe(false);
  });

  it("walks a very deep field tree without exhausting the stack", async () => {
    const depth = 20000;
    const bodies = [
      "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] >> >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>"
    ];
    for (let i = 0; i < depth; i++) {
      const num = 4 + i;
      bodies.push(
        i === depth - 1 ? "<< /T (leaf) /FT /Sig >>" : `<< /T (n) /Kids [${num + 1} 0 R] >>`
      );
    }
    expect(await Pdf.verifySignatures(buildRawPdf(bodies))).toEqual([]);
  });
});

describe("an encrypted file's signature dictionary written again", () => {
  it("an incremental update that re-serialises it keeps /Contents in the clear", async () => {
    const { certificate, privateKey } = await identity("Rewritten");
    const doc = new Pdf.Builder();
    doc.addPage().drawText("Secret", { x: 72, y: 750, fontSize: 14 });
    doc.setEncryption({ ownerPassword: "owner", userPassword: "" });
    doc.sign({ certificate, privateKey, name: "Enc" });
    const signed = await doc.toBuffer();

    // Re-write the signature object as a new revision, as any path that serialises the
    // dictionary into an update would: its /ByteRange and /Contents unchanged.
    const text = latin1(signed);
    const brAt = text.lastIndexOf("/ByteRange [");
    const objStart = text.lastIndexOf(" 0 obj", brAt);
    const objNum = Number(/(\d+)$/.exec(text.slice(objStart - 12, objStart))![1]);
    const byteRange = text.slice(brAt, text.indexOf("]", brAt) + 1);
    const contentsAt = text.indexOf("/Contents <", objStart);
    const contents = text.slice(contentsAt, text.indexOf(">", contentsAt) + 1);
    const body = `<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached ${byteRange} ${contents} >>`;

    const source = new PdfDocument(signed);
    expect(initDecryption(source, "")).toBe(true);
    const trailer = new Map<string, string>([["Size", String(source.objectNumberBound)]]);
    for (const key of ["Root", "Encrypt", "ID"]) {
      const at = text.lastIndexOf(`/${key} `);
      trailer.set(key, /^\/\w+ (\d+ \d+ R|\[[^\]]*\])/.exec(text.slice(at))![1]);
    }
    const updated = buildIncremental(signed, new Map([[objNum, { gen: 0, body }]]), trailer, {
      xrefStream: false,
      prevXrefOffset: source.startxrefOffset!,
      encrypt: source.encryptFn ?? undefined
    });

    expect(latin1(updated.subarray(signed.length))).toContain(contents);
    const [report] = await Pdf.verifySignatures(updated);
    expect(report.intact).toBe(true);
    expect(report.coversWholeFile).toBe(false);
  });

  it("indirect /Contents and /ByteRange: an ordinary dictionary on both sides", async () => {
    // ISO 32000-2 §12.8.1 requires both to be direct objects, so a dictionary that references
    // them is not a signature the writer exempts — and the reader does not exempt it either.
    // What matters is that the two agree: the referenced string is encrypted on write and
    // decrypted on read, so it reads back as written.
    const doc = new Pdf.Builder();
    doc.addPage().drawText("Secret", { x: 72, y: 750, fontSize: 14 });
    doc.setEncryption({ ownerPassword: "owner", userPassword: "" });
    const original = await doc.toBuffer();
    const text = latin1(original);
    const source = new PdfDocument(original);
    expect(initDecryption(source, "")).toBe(true);
    const size = source.objectNumberBound;
    const [sig, range, contents] = [size, size + 1, size + 2];
    const trailer = new Map<string, string>([["Size", String(size + 3)]]);
    for (const key of ["Root", "Encrypt", "ID"]) {
      const at = text.lastIndexOf(`/${key} `);
      trailer.set(key, /^\/\w+ (\d+ \d+ R|\[[^\]]*\])/.exec(text.slice(at))![1]);
    }
    const updated = buildIncremental(
      original,
      new Map([
        [
          sig,
          { gen: 0, body: `<< /Type /Sig /ByteRange ${range} 0 R /Contents ${contents} 0 R >>` }
        ],
        [range, { gen: 0, body: "[0 1 2 3]" }],
        [contents, { gen: 0, body: "<00ff7a>" }]
      ]),
      trailer,
      { xrefStream: false, prevXrefOffset: source.startxrefOffset!, encrypt: source.encryptFn! }
    );
    expect(latin1(updated.subarray(original.length))).not.toContain("<00ff7a>");

    const back = new PdfDocument(updated);
    expect(initDecryption(back, "")).toBe(true);
    const dict = back.derefDict(back.resolve(sig))!;
    expect(back.deref(dict.get("Contents"))).toEqual(new Uint8Array([0x00, 0xff, 0x7a]));
    expect(back.deref(dict.get("ByteRange"))).toEqual([0, 1, 2, 3]);
  });

  it("findSignatureContents applies the reader's rule to text", () => {
    const at = (text: string): number[] => [...findSignatureContents(text)];
    const sig = "<< /Type /Sig /ByteRange [0 10 20 30] /Contents <00ff> >>";
    expect(at(sig)).toEqual([sig.indexOf("<00ff>")]);
    // /Contents before /ByteRange, a literal string, nested in an annotation's /V.
    const nested = "<< /V << /Contents (x\\)y) /ByteRange [0 1 2 3] >> /T (sig) >>";
    expect(at(nested)).toEqual([nested.indexOf("(x")]);
    // Not a signature: a /ByteRange that is not four numbers, or a reference.
    expect(at("<< /ByteRange [0 1 2] /Contents <00> >>")).toEqual([]);
    expect(at("<< /ByteRange [0 1 2 3 R] /Contents <00> >>")).toEqual([]);
    expect(at("<< /ByteRange 5 0 R /Contents <00> >>")).toEqual([]);
    // /Contents of a sibling dictionary is not exempt.
    expect(at("<< /A << /ByteRange [0 1 2 3] >> /Contents <00> >>")).toEqual([]);
  });
});

describe("Pdf.Signature.fillPlaceholder", () => {
  it("explains what is missing when the file has no placeholder", async () => {
    const { certificate, privateKey } = await identity();
    const doc = new Pdf.Builder();
    doc.addPage();
    await expect(
      Pdf.Signature.fillPlaceholder(await doc.toBuffer(), { certificate, privateKey })
    ).rejects.toThrow(/no signature placeholder/i);
  });
});

/**
 * Append an incremental update that adds a second signature field and signs it — which
 * the library itself does not do, so the fixture is assembled by hand.
 */
async function appendSecondSignature(
  signed: Uint8Array,
  signer: { certificate: Uint8Array; privateKey: Uint8Array }
): Promise<Uint8Array> {
  const text = latin1(signed);
  const size = Number(/\/Size (\d+)/.exec(text.slice(text.lastIndexOf("trailer")))![1]);
  const root = Number(/\/Root (\d+) 0 R/.exec(text.slice(text.lastIndexOf("trailer")))![1]);
  const prev = Number(/startxref\s+(\d+)/.exec(text.slice(text.lastIndexOf("startxref")))![1]);
  const catalog = new RegExp(`\\n${root} 0 obj([\\s\\S]*?)endobj`).exec(text)![1];
  const oldFields = /\/Fields \[([^\]]*)\]/.exec(catalog)![1];
  const pages = /\/Pages (\d+ 0 R)/.exec(catalog)![1];
  const [sig, widget, cat] = [size, size + 1, size + 2];
  const { dictString } = Pdf.Signature.buildDictPlaceholder({ name: "two" });

  let update = "\n";
  const offsets: number[] = [];
  const add = (num: number, body: string) => {
    offsets[num] = signed.length + update.length;
    update += `${num} 0 obj\n${body}\nendobj\n`;
  };
  add(sig, dictString);
  add(
    widget,
    `<< /Type /Annot /Subtype /Widget /FT /Sig /T (Second) /V ${sig} 0 R /Rect [0 0 0 0] >>`
  );
  add(
    cat,
    `<< /Type /Catalog /Pages ${pages} /AcroForm << /Fields [${oldFields} ${widget} 0 R] /SigFlags 3 >> >>`
  );
  const xrefAt = signed.length + update.length;
  update += `xref\n${sig} 3\n`;
  for (const num of [sig, widget, cat]) {
    update += `${String(offsets[num]).padStart(10, "0")} 00000 n \n`;
  }
  update += `trailer\n<< /Size ${size + 3} /Root ${cat} 0 R /Prev ${prev} >>\nstartxref\n${xrefAt}\n%%EOF\n`;

  const out = new Uint8Array(signed.length + update.length);
  out.set(signed);
  for (let i = 0; i < update.length; i++) {
    out[signed.length + i] = update.charCodeAt(i) & 0xff;
  }

  // Fill in the new placeholder: its ByteRange, then the CMS over the two ranges.
  const all = latin1(out);
  const brOpen = all.lastIndexOf("/ByteRange [") + "/ByteRange [".length;
  const brWidth = all.indexOf("]", brOpen) - brOpen;
  const gapStart = all.lastIndexOf("/Contents <") + "/Contents ".length;
  const gapEnd = all.indexOf(">", gapStart) + 1;
  const range = `0 ${gapStart} ${gapEnd} ${out.length - gapEnd}`.padEnd(brWidth, " ");
  for (let i = 0; i < brWidth; i++) {
    out[brOpen + i] = range.charCodeAt(i);
  }
  const data = concat([out.subarray(0, gapStart), out.subarray(gapEnd)]);
  const cms = await buildCmsSignedData({ ...signer, data });
  return patchGap(out, [0, gapStart, gapEnd, out.length - gapEnd], toHex(cms));
}

describe("certificateCommonName — string types", () => {
  // A TBSCertificate skeleton: only the subject is read.
  const certWithCn = (tag: number, value: Uint8Array): Uint8Array => {
    const empty = der(0x30, new Uint8Array());
    const cn = der(0x30, concat([der(0x06, Uint8Array.from([0x55, 0x04, 0x03])), der(tag, value)]));
    const subject = der(0x30, der(0x31, cn));
    const tbs = der(
      0x30,
      concat([der(0x02, Uint8Array.from([1])), empty, empty, empty, subject, empty])
    );
    return der(0x30, tbs);
  };

  it("decodes UTF8String, BMPString, UniversalString and T61String", () => {
    expect(certificateCommonName(certWithCn(0x0c, new TextEncoder().encode("Zoë 张")))).toBe(
      "Zoë 张"
    );
    expect(certificateCommonName(certWithCn(0x1e, Uint8Array.from([0x5f, 0x20, 0x00, 0x41])))).toBe(
      "张A"
    );
    expect(
      certificateCommonName(certWithCn(0x1c, Uint8Array.from([0, 1, 0xf6, 0x00, 0, 0, 0, 0x42])))
    ).toBe("\u{1f600}B");
    expect(certificateCommonName(certWithCn(0x14, Uint8Array.from([0x5a, 0x6f, 0xeb])))).toBe(
      "Zoë"
    );
  });
});
