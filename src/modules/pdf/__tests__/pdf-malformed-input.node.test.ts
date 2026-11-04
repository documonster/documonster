/**
 * Corrupted files found by fuzzing, as files: each must end in a result or a typed
 * `PdfError`. `data/malformed/` holds the mutated inputs that first exposed each root cause;
 * the platform-free cases are in `pdf-malformed-input.test.ts`.
 */
import { readFileSync } from "node:fs";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { Pdf, PdfEncryptionError, PdfError, PdfStructureError } from "../index";

// Byte-for-byte: `TextDecoder("latin1")` is windows-1252 and would remap 0x80–0x9F.
const latin1 = (bytes: Uint8Array): string => Buffer.from(bytes).toString("latin1");
const bytesOf = (text: string): Uint8Array => new Uint8Array(Buffer.from(text, "latin1"));
const load = (...path: string[]): Uint8Array =>
  new Uint8Array(readFileSync(join(__dirname, "data", ...path)));

/** Run a read and report how it ended: `"ok"` or the error's class name. */
async function outcome(read: () => unknown): Promise<string> {
  try {
    await read();
    return "ok";
  } catch (error) {
    expect(error).toBeInstanceOf(PdfError);
    return (error as Error).constructor.name;
  }
}

/** Append a revision redefining `objects` (number → body) with a classic xref section. */
function appendRevision(original: Uint8Array, objects: Record<number, string>, trailer: string) {
  let text = latin1(original);
  const prev = Number(/startxref\s+(\d+)/.exec(text.slice(text.lastIndexOf("startxref")))![1]);
  if (!text.endsWith("\n")) {
    text += "\n";
  }
  const rows: string[] = [];
  for (const [num, body] of Object.entries(objects)) {
    rows.push(`${num} 1\n${String(text.length).padStart(10, "0")} 00000 n \n`);
    text += `${num} 0 obj\n${body}\nendobj\n`;
  }
  const xref = text.length;
  text += `xref\n${rows.join("")}trailer\n<< ${trailer} /Prev ${prev} >>\nstartxref\n${xref}\n%%EOF\n`;
  return bytesOf(text);
}

describe("V=5 encryption with a truncated /UE", () => {
  it("is a typed error, not the AES decryptor's own", async () => {
    const text = latin1(load("encrypted", "aes-256-r6.pdf"));
    // Shorten /UE in place, padding with spaces so every xref offset stays right.
    const truncated = text.replace(/\/UE <([0-9a-f]{64})>/, (match, hex: string) =>
      `/UE <${hex.slice(0, 20)}>`.padEnd(match.length, " ")
    );
    expect(truncated).not.toBe(text);
    const bytes = bytesOf(truncated);
    await expect(Pdf.extract(bytes, { password: "user" })).rejects.toThrow(PdfStructureError);
    expect(() => Pdf.Editor.read(bytes, { password: "user" })).toThrow(PdfStructureError);
  });
});

describe("V=5 encryption with a truncated /OE", () => {
  it("still opens with the user password, which never reads /OE", async () => {
    const text = latin1(load("encrypted", "aes-256-r6.pdf"));
    const truncated = text.replace(/\/OE <([0-9a-f]{64})>/, (match, hex: string) =>
      `/OE <${hex.slice(0, 20)}>`.padEnd(match.length, " ")
    );
    expect(truncated).not.toBe(text);
    const bytes = bytesOf(truncated);
    expect((await Pdf.extract(bytes, { password: "user" })).pages[0].text).toContain("Base text");
    expect(Pdf.Editor.read(bytes, { password: "user" }).pageCount).toBe(1);
    // The owner password is the one that needs /OE: that is still a typed error.
    expect(() => Pdf.Editor.read(bytes, { password: "owner" })).toThrow(PdfStructureError);
  });
});

describe("an /Encrypt dictionary a rewrite cannot copy", () => {
  // A full rewrite copies /Encrypt with every reference resolved in place. With no cycle
  // check that copy recursed to its depth cap down every branch — 2^32 visits for a
  // dictionary referring to itself twice — and the write never finished.
  const original = load("encrypted", "rc4-40.pdf");
  const encrypt = /9 0 obj\s*(<<[^]*?>>)\s*endobj/.exec(latin1(original))![1];
  const trailer =
    "/Root 1 0 R /Size 10 /ID [<f918ad5ab487441c676451974b980eba><f918ad5ab487441c676451974b980eba>] /Encrypt 9 0 R";
  const withEntry = (entry: string, more: Record<number, string> = {}): Uint8Array =>
    appendRevision(original, { 9: encrypt.replace(/>>$/, `${entry} >>`), ...more }, trailer);

  it("a dictionary that refers back to itself", async () => {
    const bytes = withEntry("/X [9 0 R 9 0 R]");
    expect((await Pdf.extract(bytes, { password: "user" })).pages[0].text).toContain("Base text");
    const editor = Pdf.Editor.read(bytes, { password: "user" });
    await expect(editor.toBuffer()).rejects.toThrow(PdfEncryptionError);
    await expect(editor.toBuffer()).rejects.toThrow(/refers back to itself/);
    const decrypted = await editor.toBuffer({ encryption: "remove" });
    expect((await Pdf.extract(decrypted)).pages[0].text).toContain("Base text");
  });

  it("a value shared down many levels", async () => {
    const levels: Record<number, string> = {};
    for (let i = 0; i < 40; i++) {
      levels[10 + i] = `[${11 + i} 0 R ${11 + i} 0 R]`;
    }
    levels[50] = "(leaf)";
    const bytes = withEntry("/X 10 0 R", levels);
    const editor = Pdf.Editor.read(bytes, { password: "user" });
    await expect(editor.toBuffer()).rejects.toThrow(/too large to copy/);
  });

  it("an ordinary dictionary is still copied", async () => {
    const editor = Pdf.Editor.read(withEntry("/X [1 2 3]"), { password: "user" });
    const out = await editor.toBuffer();
    expect(latin1(out)).toContain("/X [1 2 3]");
    expect((await Pdf.extract(out, { password: "user" })).pages[0].text).toContain("Base text");
  });
});

describe("fuzz fixtures", () => {
  it("a stray ')' in an object stream reached from xref reconstruction", async () => {
    const bytes = load("malformed", "stray-delimiter-objstm.pdf");
    expect(await outcome(() => Pdf.extract(bytes))).toBeTypeOf("string");
    expect(await outcome(() => Pdf.Editor.read(bytes).toBuffer())).toBeTypeOf("string");
  });
});
