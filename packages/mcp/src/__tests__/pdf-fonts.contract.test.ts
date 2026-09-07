import { mkdtemp, readFile, realpath, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { Cell, Workbook } from "documonster/excel";
import { Pdf } from "documonster/pdf";
import { describe, expect, it } from "vitest";

import { resolveConfig, type ServerConfig } from "../config.js";
import { docConvertTool } from "../tools/doc-convert.js";
import { collectFontWarnings } from "../tools/pdf-fonts.js";

interface Fixture {
  readonly config: ServerConfig;
  readonly root: string;
}

async function fixture(): Promise<Fixture> {
  const root = await realpath(await mkdtemp(path.join(tmpdir(), "documonster-mcp-fonts-")));
  const base = resolveConfig([], { cwd: root });
  return { config: { ...base, outputRoot: root, allowInPlace: true }, root };
}

async function run(fx: Fixture, args: Record<string, unknown>): Promise<string> {
  const result = await docConvertTool.handler(args, { config: fx.config });
  return result.content
    .filter((block): block is { type: "text"; text: string } => block.type === "text")
    .map(block => block.text)
    .join("\n");
}

/**
 * The wording the fail-closed check depends on, pinned against the library itself.
 *
 * `pdf-fonts.ts` decides "this PDF has boxes in it" by looking for `.notdef` in the
 * writer's warning. That is a string coupling across a package boundary, and the honest
 * way to hold it is to run a real export down each path the library takes and assert the
 * phrase is still there. A reword in `font-manager.ts` then fails here rather than
 * silently turning the check off and letting boxed PDFs through again.
 */
describe("missing-glyph warnings, as the library actually words them", () => {
  /**
   * A character no font draws, on any host: an unassigned code point.
   *
   * Not a private-use one, which was the first attempt — `U+E000` is mapped by plenty of
   * real fonts, and on this machine `.Geeza Pro` was auto-embedded to cover it, so the test
   * asserted the opposite of what it meant to on a machine with fewer fonts. An unassigned
   * code point is in no `cmap` by definition and in no glyph table here, which makes the
   * outcome a property of the code rather than of the host.
   */
  const UNDRAWABLE = "\u0378";

  it("says .notdef when nothing available covers a character", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", `x ${UNDRAWABLE}`);

    const warnings: string[] = [];
    await Pdf.fromExcel(wb, { onWarning: message => warnings.push(message) });

    expect(warnings.join(" ")).toContain(".notdef");
  });

  it("says .notdef with the host scan turned off, which is the container case", async () => {
    // The same condition reached the other way: nothing to discover, so nothing covers the
    // character. `--pdf-font` narrows the gap rather than closing it, and the wording for a
    // configured face that does not reach far enough also carries `.notdef` — asserted here
    // through the path a test can drive without shipping a font fixture.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", `x ${UNDRAWABLE}`);

    const warnings: string[] = [];
    // A font that exists but is narrow: the standard-14 metrics only, via the legacy
    // single-font option being absent and `fonts` naming a face without the code point.
    // Reusing the same undrawable character keeps the assertion about the wording.
    await Pdf.fromExcel(wb, {
      disableFontAutoDiscovery: true,
      onWarning: message => warnings.push(message)
    });

    expect(warnings.join(" ")).toContain(".notdef");
  });

  it("does not say .notdef for a character the built-in glyphs draw", async () => {
    // The discriminator has to be narrow: an arrow or a Greek letter is drawn by the
    // Type3 fallback, and refusing to write that PDF would make the check useless.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", "→ Δ Кириллица ł ₂");

    const warnings: string[] = [];
    await Pdf.fromExcel(wb, {
      disableFontAutoDiscovery: true,
      onWarning: message => warnings.push(message)
    });

    expect(warnings.join(" ")).not.toContain(".notdef");
  });
});

describe("assertDrawable", () => {
  it("passes when nothing was reported", () => {
    expect(() => collectFontWarnings().assertDrawable(false)).not.toThrow();
  });

  it("passes for a warning that is only a note", () => {
    const fonts = collectFontWarnings();
    fonts.onWarning("Auto-embedded system font 'Songti SC' to render 61 character(s).");
    fonts.onWarning("2 character(s) will be drawn with built-in Type3 glyphs: Arrows (2).");

    expect(() => fonts.assertDrawable(false)).not.toThrow();
    expect(fonts.missingGlyphs()).toEqual([]);
  });

  it("refuses a PDF that will have boxes in it", () => {
    const fonts = collectFontWarnings();
    fonts.onWarning(
      "3 character(s) have no glyph in any available font and will render as .notdef boxes: Han (3)."
    );

    expect(() => fonts.assertDrawable(false)).toThrow(/was not written/);
    // The way forward is in the hint, which is where this server puts remedies.
    try {
      fonts.assertDrawable(false);
      expect.unreachable("assertDrawable should have thrown");
    } catch (error) {
      expect((error as { hint?: string }).hint).toContain("--pdf-font");
      expect((error as { hint?: string }).hint).toContain("allowMissingGlyphs");
    }
    expect(fonts.missingGlyphs()).toHaveLength(1);
  });

  it("writes it anyway when the caller has decided", () => {
    const fonts = collectFontWarnings();
    fonts.onWarning("1 character(s) will render with the .notdef glyph (e.g. U+1F389).");

    expect(() => fonts.assertDrawable(true)).not.toThrow();
  });
});

describe("doc_convert refuses rather than writing boxes", () => {
  it("fails the conversion, and leaves no file behind", async () => {
    const fx = await fixture();
    await writeFile(path.join(fx.root, "in.md"), "# Report \u0378\n", "utf8");

    await expect(
      docConvertTool.handler({ from: "in.md", to: "out.pdf" }, { config: fx.config })
    ).rejects.toThrow(/was not written/);

    // Nothing promoted to the destination: the check runs before the bytes are written.
    await expect(readFile(path.join(fx.root, "out.pdf"))).rejects.toThrow();
  });

  it("writes it when told to", async () => {
    const fx = await fixture();
    await writeFile(path.join(fx.root, "in.md"), "# Report \u0378\n", "utf8");

    const report = await run(fx, {
      from: "in.md",
      to: "out.pdf",
      allowMissingGlyphs: true
    });

    expect(report).toContain("font coverage");
    expect((await readFile(path.join(fx.root, "out.pdf"))).subarray(0, 5).toString("latin1")).toBe(
      "%PDF-"
    );
  });
});
