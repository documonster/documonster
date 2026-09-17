/**
 * Every place this server draws text of its own must be able to reach the operator's font.
 *
 * `--pdf-font` existed but only two call sites used it, and the gaps were invisible: a
 * Chinese watermark stamped by `pdf_edit` simply did not appear, a Chinese label inside a
 * diagram rendered to PNG came out blank, and both calls reported success. The PDF writer at
 * least draws a `.notdef` box and has a fail-closed check for it; a rasteriser has no notdef,
 * so nothing was drawn and nothing was said.
 *
 * These tests assert *reachability* — that the configuration arrives — rather than pixels.
 * Whether the glyph is correct is the library's business and is tested there; whether this
 * server bothered to pass the font is this package's, and that is what regressed.
 */

import { mkdtemp, realpath, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { resolveConfig, type ServerConfig } from "../config.js";
import { docWriteTool } from "../tools/doc-write.js";
import { diagramFontOptions, overlayFaces, textLanguageOption } from "../tools/pdf-fonts.js";

async function root(): Promise<string> {
  return realpath(await mkdtemp(path.join(tmpdir(), "documonster-mcp-reach-")));
}

/** A file with a valid sfnt header, which is all the startup check reads. */
async function fakeTtf(dir: string, name: string): Promise<string> {
  const file = path.join(dir, name);
  await writeFile(file, Buffer.from([0x00, 0x01, 0x00, 0x00, 0x11, 0x22]));
  return file;
}

async function configured(names: readonly string[]): Promise<ServerConfig> {
  const cwd = await root();
  for (const name of names) {
    await fakeTtf(cwd, name);
  }
  const [primary, ...rest] = names;
  const argv = ["--pdf-font", primary as string];
  for (const name of rest) {
    argv.push("--pdf-font-fallback", name);
  }
  return resolveConfig(argv, { cwd });
}

describe("the operator's font reaches every backend", () => {
  it("is offered to the rasteriser, not only to the PDF writer", async () => {
    // The gap that made a diagram's Chinese labels blank while the prose around them was
    // fine: the raster backend takes bytes and was never handed any.
    const config = await configured(["cjk.ttf"]);
    const options = diagramFontOptions(config);

    expect(options.pdf).toBeDefined();
    expect(options.raster).toHaveLength(1);
  });

  it("offers every configured face to the rasteriser, primary and fallbacks alike", async () => {
    // The rasteriser resolves each character against a chain itself, so there is no "default"
    // face for it to be primary of — handing it only the primary would drop the others.
    const config = await configured(["cjk.ttf", "arabic.ttf", "symbols.ttf"]);

    expect(diagramFontOptions(config).raster).toHaveLength(3);
  });

  it("switches the host off, so the render depends only on the named faces", async () => {
    // Not cosmetic, and it is what makes the next test possible. `--pdf-font` exists to take
    // the machine out of the answer; consulting installed fonts for the raster backend while
    // the PDF backend ignores them would leave half the problem in place.
    expect(diagramFontOptions(await configured(["cjk.ttf"])).useSystemFonts).toBe(false);
    expect(diagramFontOptions(resolveConfig([], { cwd: await root() })).useSystemFonts).toBe(
      undefined
    );
  });

  it("yields nothing when no font is configured", async () => {
    // Must stay empty rather than `{ raster: [] }`: an empty array passed as `fonts` with
    // `useSystemFonts` untouched is a different request from saying nothing.
    expect(diagramFontOptions(resolveConfig([], { cwd: await root() }))).toEqual({});
  });

  it("passes textLanguage through only when the caller stated it", async () => {
    // Han characters are shared between the three languages and drawn differently, so this
    // decides the shape of the glyphs. Defaulting it here would override the library's own
    // inference from the document's text, which is better than a guess made in this package.
    expect(textLanguageOption({ textLanguage: "ja" })).toEqual({ textLanguage: "ja" });
    expect(textLanguageOption({})).toEqual({});
  });
});

describe("overlay text gets the face, including a face inside a collection", () => {
  it("carries collectionIndex through to the editor", async () => {
    // The bug this replaced was worse than the one before it. Handing `embedFont` only the
    // bytes meant a `.ttc` face selection was dropped — and the code did that by testing
    // `instanceof Uint8Array` and skipping anything else, so configuring `Songti.ttc#6` left
    // overlays with no face at all while `Songti.ttf` worked. Silent, and inconsistent.
    const cwd = await root();
    await fakeTtf(cwd, "faces.ttc");
    await writeFile(
      path.join(cwd, "faces.ttc"),
      Buffer.from("ttcf\u0000\u0001\u0000\u0000", "latin1")
    );
    const config = resolveConfig(["--pdf-font", "faces.ttc#6"], { cwd });

    const faces = overlayFaces(config);
    expect(faces).toHaveLength(1);
    expect(faces[0]?.collectionIndex).toBe(6);
  });

  it("omits collectionIndex for a single-face file rather than sending 0", async () => {
    // `embedFont(bytes, 0)` and `embedFont(bytes)` mean the same thing today, but stating an
    // index that was never asked for turns a default into a decision.
    const config = await configured(["face.ttf"]);
    expect(overlayFaces(config)[0]).not.toHaveProperty("collectionIndex");
  });

  it("is empty with no font configured, so nothing is embedded", async () => {
    expect(overlayFaces(resolveConfig([], { cwd: await root() }))).toEqual([]);
  });
});

describe("an embedded diagram reports what it could not draw", () => {
  it("names the uncovered code points in the tool result", async () => {
    // The most hidden version of the silent failure: the reader opens a .docx, sees a diagram
    // with missing labels, and has nothing pointing at fonts. Rendering a standalone PNG
    // reported this; embedding one did not, and the embedding path is the one a document
    // conversion takes.
    const cwd = await root();
    const fontPath = await fakeTtf(cwd, "unusable.ttf");
    const base = resolveConfig(["--pdf-font", "unusable.ttf"], { cwd });
    const config: ServerConfig = { ...base, outputRoot: cwd, allowInPlace: true };

    const result = await docWriteTool.handler(
      {
        path: "out.docx",
        markdown: "# T\n\n```mermaid\nflowchart LR\n  A[\u4e2d\u6587] --> B[ok]\n```\n",
        overwrite: true
      },
      { config }
    );
    const text = result.content
      .filter((block): block is { type: "text"; text: string } => block.type === "text")
      .map(block => block.text)
      .join("\n");

    expect(fontPath).toContain("unusable.ttf");
    expect(text).toContain("U+4E2D");
    expect(text.toLowerCase()).toContain("blank");
  });
});
