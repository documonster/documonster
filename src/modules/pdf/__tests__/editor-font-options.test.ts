/**
 * `Pdf.Editor.read` takes the font options of `Pdf.Builder` (R4P-G1): text drawn on an
 * editor page is embedded, and reported, as the same text on a builder page is. The editor
 * used to have no font discovery and no `onWarning`, so Chinese overlay text came out as
 * `.notdef` boxes while `Pdf.extract` still read it back, and nothing said so.
 */
import { Pdf, PdfArgumentError } from "@pdf/index";
import { describe, expect, it } from "vitest";

import { buildTtfWithCmap, buildMinimalTtf } from "./ttf-test-utils";

const ZHONG = 0x4e2d;

/** A face covering `中`, `A` and `B`. */
function cjkFace(): Uint8Array {
  return buildTtfWithCmap(
    [
      { start: 0x41, end: 0x42, delta: 1 - 0x41 },
      { start: ZHONG, end: ZHONG, delta: 3 - ZHONG }
    ],
    4,
    { familyName: "TestHan" }
  );
}

async function blank(): Promise<Uint8Array> {
  const b = new Pdf.Builder();
  b.addPage();
  return b.toBuffer();
}

function latin1(bytes: Uint8Array): string {
  return new TextDecoder("latin1").decode(bytes);
}

function embedsTrueType(bytes: Uint8Array): boolean {
  return /\/FontFile2/.test(latin1(bytes));
}

describe("Pdf.Editor.read font options", () => {
  it("embeds the configured faces for overlay text, full rewrite and incremental", async () => {
    const source = await blank();
    for (const incremental of [false, true]) {
      const editor = Pdf.Editor.read(source, { fonts: { default: { regular: cjkFace() } } });
      editor.getPage(0).drawText("中AB", { x: 10, y: 10 });
      const out = await editor.toBuffer({ incremental });
      expect(embedsTrueType(out)).toBe(true);
      expect((await Pdf.extract(out)).text).toContain("中AB");
    }
  });

  it("does what the builder does with the same options", async () => {
    const fonts = { default: { regular: cjkFace() } };
    const b = new Pdf.Builder({ fonts });
    b.addPage().drawText("中", { x: 10, y: 10 });
    const built = latin1(await b.toBuffer());
    const editor = Pdf.Editor.read(await blank(), { fonts });
    editor.getPage(0).drawText("中", { x: 10, y: 10 });
    const edited = latin1(await editor.toBuffer());
    const fontName = (text: string) => /\/FontName\s*\/([^\s/>]+)/.exec(text)?.[1];
    expect(fontName(edited)).toBe(fontName(built));
  });

  it("reports text no face can draw through onWarning, and uses no host font when told not to", async () => {
    const warnings: string[] = [];
    const editor = Pdf.Editor.read(await blank(), {
      disableFontAutoDiscovery: true,
      onWarning: message => warnings.push(message)
    });
    editor.getPage(0).drawText("中", { x: 10, y: 10 });
    const out = await editor.toBuffer();
    expect(embedsTrueType(out)).toBe(false);
    expect(warnings.some(message => message.includes(".notdef"))).toBe(true);
    // An added page draws through the same fonts and the same report.
    const added: string[] = [];
    const other = Pdf.Editor.read(await blank(), {
      disableFontAutoDiscovery: true,
      onWarning: message => added.push(message)
    });
    other.addPage().drawText("中", { x: 10, y: 10 });
    await other.toBuffer();
    expect(added.some(message => message.includes(".notdef"))).toBe(true);
  });

  it("checks the font options as the builder does, before reading the file", () => {
    for (const [options, field] of [
      [{ fonts: 42 }, "options.fonts"],
      [{ fonts: { default: { regular: "x" } } }, "options.fonts.default.regular"],
      [{ preferSystemFonts: "Arial" }, "options.preferSystemFonts"],
      [{ textLanguage: "fr" }, "options.textLanguage"],
      [{ disableFontAutoDiscovery: "yes" }, "options.disableFontAutoDiscovery"],
      [{ onWarning: 1 }, "options.onWarning"],
      [{ password: 1 }, "options.password"]
    ] as const) {
      let error: unknown;
      try {
        Pdf.Editor.read(new Uint8Array([0x25, 0x50, 0x44, 0x46]), options as never);
      } catch (err) {
        error = err;
      }
      expect(error).toBeInstanceOf(PdfArgumentError);
      expect((error as PdfArgumentError).field).toBe(field);
      expect((error as Error).message.startsWith("Pdf.Editor.read: ")).toBe(true);
    }
  });

  it("refuses embedFont when the document was read with fonts", async () => {
    const editor = Pdf.Editor.read(await blank(), { fonts: { default: { regular: cjkFace() } } });
    expect(() => editor.embedFont(buildMinimalTtf())).toThrow(PdfArgumentError);
    // Without `fonts`, embedFont is the way to name the face, as it always was.
    const plain = Pdf.Editor.read(await blank());
    plain.embedFont(cjkFace());
    plain.getPage(0).drawText("中", { x: 10, y: 10 });
    expect(embedsTrueType(await plain.toBuffer())).toBe(true);
  });
});
