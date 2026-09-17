import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { collectFontWarnings, pdfFontOptions } from "../tools/pdf-fonts.js";

describe("collectFontWarnings", () => {
  it("says nothing when the writer raised nothing", () => {
    expect(collectFontWarnings().notes()).toEqual([]);
  });

  it("marks a character that will draw as a box more strongly than a note", () => {
    // The two conditions are not equally serious and must not read alike. An
    // auto-embedded face means the page is fine but the bytes depend on the host; a
    // missing glyph means the page is visibly wrong. Issue #218 was reported as boxes
    // in a document, so this is the line that has to stand out.
    const fonts = collectFontWarnings();
    fonts.onWarning("Auto-embedded system font 'Songti SC' to render 61 character(s).");
    fonts.onWarning(
      "2 character(s) have no glyph in any available font and will render as .notdef boxes: CJK Unified Ideographs (2)."
    );

    const notes = fonts.notes();

    expect(notes[0]).toContain("**font coverage**");
    expect(notes[0]).toContain("no glyph in any available font");
    expect(notes[1]).toBe(
      "- font: Auto-embedded system font 'Songti SC' to render 61 character(s)."
    );
  });

  it("keeps every message it was given", () => {
    const fonts = collectFontWarnings();
    fonts.onWarning("first");
    fonts.onWarning("second");

    expect(fonts.notes()).toEqual(["- font: first", "- font: second"]);
  });
});

describe("pdfFontOptions", () => {
  it("passes no font when none is configured", () => {
    const { options } = pdfFontOptions({});
    expect(options.fonts).toBeUndefined();
    expect(typeof options.onWarning).toBe("function");
  });

  it("supplies the configured font as the default regular face", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "documonster-mcp-font-"));
    const file = path.join(dir, "face.ttf");
    const bytes = Buffer.from([0x00, 0x01, 0x00, 0x00, 0x11, 0x22]);
    await writeFile(file, bytes);

    const { options } = pdfFontOptions({ pdfFont: { path: file } });

    expect(options.fonts?.default.regular).toEqual(new Uint8Array(bytes));
  });

  it("reads the font once, however many conversions ask for it", async () => {
    // A CJK face is tens of megabytes; a conversion-heavy session must not
    // re-read it per call.
    const dir = await mkdtemp(path.join(tmpdir(), "documonster-mcp-font-"));
    const file = path.join(dir, "cached.ttf");
    await writeFile(file, Buffer.from([0x00, 0x01, 0x00, 0x00]));

    expect(pdfFontOptions({ pdfFont: { path: file } }).options.fonts).toBe(
      pdfFontOptions({ pdfFont: { path: file } }).options.fonts
    );
  });

  it("turns the fallbacks into a real chain the library will follow", async () => {
    // `{ default: { regular } }` alone cannot express "try this, then that", which is why one
    // configured face meant a mixed-script document lost a script to `.notdef` boxes.
    const dir = await mkdtemp(path.join(tmpdir(), "documonster-mcp-font-"));
    const primary = path.join(dir, "cjk.ttf");
    const arabic = path.join(dir, "arabic.ttf");
    await writeFile(primary, Buffer.from([0x00, 0x01, 0x00, 0x00, 0x01]));
    await writeFile(arabic, Buffer.from([0x00, 0x01, 0x00, 0x00, 0x02]));

    const { options } = pdfFontOptions({
      pdfFont: { path: primary },
      pdfFontFallbacks: [{ path: arabic }]
    });

    expect(options.fonts?.fallbackFamilies).toEqual(["fallback-1"]);
    expect(options.fonts?.families?.[0]?.faces.regular).toEqual(
      new Uint8Array([0x00, 0x01, 0x00, 0x00, 0x02])
    );
  });

  it("keeps collectionIndex attached to the bytes", async () => {
    // Handing over just the bytes embeds face 0, which for `Songti.ttc` is weight 900.
    const dir = await mkdtemp(path.join(tmpdir(), "documonster-mcp-font-"));
    const file = path.join(dir, "faces.ttc");
    await writeFile(file, Buffer.from([0x74, 0x74, 0x63, 0x66]));

    const { options } = pdfFontOptions({ pdfFont: { path: file, collectionIndex: 3 } });

    expect(options.fonts?.default.regular).toEqual({
      data: new Uint8Array([0x74, 0x74, 0x63, 0x66]),
      collectionIndex: 3
    });
  });

  it("caches per face, not per file", async () => {
    // Two faces of one collection must not collapse to one cache entry.
    const dir = await mkdtemp(path.join(tmpdir(), "documonster-mcp-font-"));
    const file = path.join(dir, "two.ttc");
    await writeFile(file, Buffer.from([0x74, 0x74, 0x63, 0x66]));

    const first = pdfFontOptions({ pdfFont: { path: file, collectionIndex: 0 } }).options.fonts;
    const second = pdfFontOptions({ pdfFont: { path: file, collectionIndex: 3 } }).options.fonts;
    expect(first).not.toBe(second);
    expect(pdfFontOptions({ pdfFont: { path: file, collectionIndex: 3 } }).options.fonts).toBe(
      second
    );
  });

  it("reports a font deleted after startup as itself", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "documonster-mcp-font-"));
    expect(() => pdfFontOptions({ pdfFont: { path: path.join(dir, "gone.ttf") } })).toThrow(
      /--pdf-font/
    );
  });
});
