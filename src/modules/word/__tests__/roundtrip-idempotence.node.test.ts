/**
 * Every `.docx` checked into the repository survives a round trip intact and
 * stably: read → write → read → write gives the same bytes twice (core
 * properties' timestamps aside), and the written package is structurally
 * valid ({@link packageIssues}). Reads fixtures off disk, hence `.node`.
 *
 * `data/example-outputs/` holds packages the Word examples write — watermarks,
 * a TOC, an SVG with its fallback — each of which once grew on every pass.
 */

import { readdirSync, readFileSync } from "node:fs";
import { join, relative, resolve } from "node:path";

import { describe, expect, it } from "vitest";

import { Io } from "../index";
import {
  differingParts,
  packageIssues,
  packageParts,
  partText,
  tocBlockInstructions
} from "./roundtrip-integrity.helpers";

const ROOT = resolve(__dirname, "../../../..");
// `office-encrypted` holds password-protected files: compound files, not packages to round-trip.
const SKIP = new Set(["node_modules", "dist", "tmp", "coverage", ".git", "office-encrypted"]);

function findDocx(dir: string, out: string[] = []): string[] {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      if (!SKIP.has(entry.name)) {
        findDocx(join(dir, entry.name), out);
      }
    } else if (entry.name.endsWith(".docx")) {
      out.push(join(dir, entry.name));
    }
  }
  return out;
}

const fixtures = findDocx(ROOT).sort();

describe("round trip of every .docx fixture", () => {
  it("finds the fixtures", () => {
    expect(fixtures.length).toBeGreaterThan(0);
  });

  it.each(fixtures.map(f => [relative(ROOT, f), f]))("%s", async (_name, file) => {
    const once = await Io.toBuffer(await Io.read(readFileSync(file)));
    const twice = await Io.toBuffer(await Io.read(once));
    const first = await packageParts(once);
    const second = await packageParts(twice);
    expect(packageIssues(first)).toEqual([]);
    expect(differingParts(first, second)).toEqual([]);
  });
});

/** The visible text of every TOC block in `xml`: its `w:t` text and tabs, not field instructions. */
function tocBlockTexts(xml: string): string[] {
  return [...xml.matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>/g)]
    .filter(m => /<w:docPartGallery w:val="Table of Contents"\/>/.test(m[1]!))
    .map(m =>
      [
        ...m[1]!
          .replace(/<w:sdtPr>[\s\S]*?<\/w:sdtPr>/, "")
          .matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<\/w:p>/g)
      ]
        .map(([token, text]) => text ?? (token === "<w:tab/>" ? "\t" : "\u00b6"))
        .join("")
    );
}

describe("TOC blocks of every .docx fixture", () => {
  it.each(fixtures.map(f => [relative(ROOT, f), f]))("%s", async (_name, file) => {
    const source = readFileSync(file);
    const before = partText(await packageParts(source), "word/document.xml");
    const after = partText(
      await packageParts(await Io.toBuffer(await Io.read(source))),
      "word/document.xml"
    );
    expect(tocBlockInstructions(after)).toEqual(tocBlockInstructions(before));
    expect(tocBlockTexts(after)).toEqual(tocBlockTexts(before));
  });
});
