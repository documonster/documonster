/**
 * A table of contents Word wrote — its field opening in the first entry's
 * paragraph, closing in a paragraph after the last — is written back with the
 * same paragraphs: no paragraph added for the field to open in, none for it to
 * close in, and the field characters where they were. The fixtures are
 * Microsoft Word's own output (Word 16 updating a TOC field and saving).
 * Reads fixtures off disk, hence `.node`.
 */

import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { describe, expect, it } from "vitest";

import { Io } from "../index";
import type { TableOfContents } from "../types";
import { packageParts, partText, tocBlockInstructions } from "./roundtrip-integrity.helpers";

const DATA = resolve(__dirname, "data/word-native");

/**
 * Per paragraph of every TOC block: its style, then its content as tokens —
 * field characters, instruction text (joined, trimmed), text, tabs — with a
 * hyperlink's content inside `link(anchor)[…]`.
 */
function tocParagraphs(xml: string): string[] {
  const out: string[] = [];
  for (const [, sdt] of xml.matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>/g)) {
    if (!sdt!.includes('w:docPartGallery w:val="Table of Contents"')) {
      continue;
    }
    for (const [p] of sdt!.matchAll(/<w:p\b[^>]*?(?:\/>|>[\s\S]*?<\/w:p>)/g)) {
      const style = /<w:pStyle w:val="([^"]+)"/.exec(p)?.[1] ?? "-";
      const body = p.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, "");
      const tokens = (s: string): string[] => {
        const list: string[] = [];
        for (const m of s.matchAll(
          /<w:fldChar w:fldCharType="(\w+)"|<w:instrText[^>]*>([^<]*)<\/w:instrText>|<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>/g
        )) {
          if (m[1]) {
            list.push(m[1]);
          } else if (m[2] !== undefined) {
            const text = m[2].replace(/&quot;/g, '"');
            if (list.at(-1)?.startsWith("instr:")) {
              list[list.length - 1] += text;
            } else {
              list.push(`instr:${text}`);
            }
          } else if (m[3] !== undefined) {
            list.push(`t:${m[3]}`);
          } else {
            list.push("tab");
          }
        }
        return list.map(t => (t.startsWith("instr:") ? `instr:${t.slice(6).trim()}` : t));
      };
      const parts: string[] = [];
      let rest = body;
      for (const link of body.matchAll(
        /<w:hyperlink\b[^>]*w:anchor="([^"]*)"[^>]*>([\s\S]*?)<\/w:hyperlink>/g
      )) {
        const [before] = rest.split(link[0]);
        parts.push(...tokens(before!), `link(${link[1]})[${tokens(link[2]!).join(" ")}]`);
        rest = rest.slice(before!.length + link[0].length);
      }
      parts.push(...tokens(rest));
      out.push(`${style}: ${parts.join(" ")}`);
    }
  }
  return out;
}

describe.each(["toc-word.docx", "toc-word-zu.docx"])("Word's own TOC: %s", file => {
  const source = readFileSync(resolve(DATA, file));

  it("keeps the paragraphs of its TOC, and where the field characters are", async () => {
    const before = partText(await packageParts(source), "word/document.xml");
    const once = await Io.toBuffer(await Io.read(source));
    const after = partText(await packageParts(once), "word/document.xml");
    expect(tocParagraphs(after)).toEqual(tocParagraphs(before));
    expect(tocBlockInstructions(after)).toEqual(tocBlockInstructions(before));
    const twice = partText(
      await packageParts(await Io.toBuffer(await Io.read(once))),
      "word/document.xml"
    );
    expect(twice).toBe(after);
  });

  it("opens its field in the first entry's paragraph", async () => {
    const before = tocParagraphs(partText(await packageParts(source), "word/document.xml"));
    expect(before[0]).toMatch(/^TOC1: begin instr:TOC .* separate link\(/);
    expect(before.at(-1)).toBe("-: end");
    const doc = await Io.read(source);
    const toc = doc.body.find((b): b is TableOfContents => b.type === "tableOfContents")!;
    expect(toc.field).toBe("entries");
    expect(toc.headingStyleRange).toBe("1-3");
  });
});
