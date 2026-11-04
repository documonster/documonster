/**
 * Names are byte sequences (ISO 32000-1 §7.3.5): text is written as its UTF-8 bytes, and
 * every byte outside `!`–`~`, plus `#` and the delimiters `( ) < > [ ] { } / %`, as `#xx`.
 * `pdfName` used to escape UTF-16 code units, so `选` became `#9009` — which a reader takes
 * as the byte `0x90` followed by the characters `09`.
 */
import { pdfName, pdfNameBytes } from "@pdf/core/pdf-object";
import { Pdf, PdfArgumentError } from "@pdf/index";
import { PdfTokenizer, TokenType } from "@pdf/reader/pdf-tokenizer";
import { describe, expect, it } from "vitest";

function utf8Hex(text: string): string {
  return Array.from(
    new TextEncoder().encode(text),
    b => `#${b.toString(16).padStart(2, "0")}`
  ).join("");
}

/** The bytes a reader decodes a written name to, as a string of char codes 0–255. */
function readBack(written: string): string {
  const token = new PdfTokenizer(new TextEncoder().encode(`${written} `)).next();
  expect(token.type).toBe(TokenType.Name);
  return token.strValue!;
}

const latin1 = (bytes: Uint8Array) => String.fromCharCode(...bytes);

describe("pdfName", () => {
  it("writes non-ASCII text as escaped UTF-8 bytes", () => {
    expect(pdfName("Ünïcødé 选项")).toBe(
      `/${utf8Hex("Ü")}n${utf8Hex("ï")}c${utf8Hex("ø")}d${utf8Hex("é")}#20${utf8Hex("选项")}`
    );
    expect(readBack(pdfName("Ünïcødé 选项"))).toBe(
      latin1(new TextEncoder().encode("Ünïcødé 选项"))
    );
    // An astral character is one code point, four UTF-8 bytes — not two surrogates.
    expect(pdfName("😀")).toBe("/#f0#9f#98#80");
  });

  it("escapes every delimiter, including % which would start a comment", () => {
    expect(pdfName("a%b(c)<d>[e]{f}/g#h")).toBe("/a#25b#28c#29#3cd#3e#5be#5d#7bf#7d#2fg#23h");
  });

  it("leaves regular ASCII names as they were", () => {
    expect(pdfName("Type")).toBe("/Type");
    expect(pdfName("F1-Subset_x.y")).toBe("/F1-Subset_x.y");
  });

  it("pdfNameBytes writes a name the reader decoded back to the same bytes", () => {
    // The reader decodes a name to one character per byte; re-encoding that as UTF-8 would
    // double-encode every byte from 0x80 up, so editor round-trips use the byte form.
    for (const bytes of [new TextEncoder().encode("选项"), Uint8Array.of(0xe9, 0x41, 0xff)]) {
      const name = latin1(bytes);
      expect(readBack(pdfNameBytes(name))).toBe(name);
    }
  });
});

describe("radio export values", () => {
  it("are written as UTF-8 names and read back as the same text", async () => {
    const value = "Ünïcødé 选项";
    const doc = new Pdf.Builder();
    doc.addPage().addFormField({
      type: "radio",
      name: "choice",
      buttons: [
        { rect: [10, 10, 30, 30], value },
        { rect: [40, 10, 60, 30], value: "plain" }
      ],
      selected: value
    });
    const text = latin1(await doc.toBuffer());
    const name = `/${utf8Hex("Ü")}n${utf8Hex("ï")}c${utf8Hex("ø")}d${utf8Hex("é")}#20${utf8Hex("选项")}`;
    expect(text).toContain(`/V ${name}`);
    expect(text).toContain(`/AS ${name}`);
  });

  it("are written the same way by the editor's overlay pages", async () => {
    const source = new Pdf.Builder();
    source.addPage();
    const editor = Pdf.Editor.read(await source.toBuffer());
    editor.getPage(0).addFormField({
      type: "radio",
      name: "choice",
      buttons: [{ rect: [10, 10, 30, 30], value: "选项" }],
      selected: "选项"
    });
    expect(latin1(await editor.toBuffer())).toContain(`/V /${utf8Hex("选项")}`);
  });

  it("refuses a value holding NUL, which no name can carry", () => {
    expect(() =>
      new Pdf.Builder().addPage().addFormField({
        type: "radio",
        name: "r",
        buttons: [{ rect: [0, 0, 1, 1], value: "a\0b" }]
      })
    ).toThrow(PdfArgumentError);
  });
});
