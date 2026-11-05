/**
 * The streaming writer writes a section break inside the body as it streams
 * past, so the break names its headers and footers by ids allocated before
 * anything is streamed. They were allocated only at the end, and a break
 * named the caller's map key — a part the package did not have.
 */

import { describe, expect, it } from "vitest";

import { Build, Document, Io, Stream } from "../index";
import type { HeaderDef, SectionProperties } from "../index";
import { packageIssues, packageParts, partText } from "./roundtrip-integrity.helpers";

const PAGE = { pageSize: { width: 12240, height: 15840 } } as const;
const part = (text: string): HeaderDef => ({ content: { children: [Build.textParagraph(text)] } });

const headers = new Map([
  ["cover", part("Cover header")],
  ["default", part("Body header")]
]);
const footers = new Map([
  ["coverFoot", part("Cover footer")],
  ["default", part("Body footer")]
]);
const inner: SectionProperties = {
  ...PAGE,
  headers: [{ type: "default", rId: "cover" }],
  footers: [{ type: "default", rId: "coverFoot" }]
};
const final: SectionProperties = {
  ...PAGE,
  headers: [{ type: "default", rId: "default" }],
  footers: [{ type: "default", rId: "default" }]
};

/** Per section, each reference as `kind type:text`. */
function sections(parts: Map<string, Uint8Array>): string[][] {
  const doc = partText(parts, "word/document.xml");
  const rels = partText(parts, "word/_rels/document.xml.rels");
  return [...doc.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map(([sect]) =>
    [...sect.matchAll(/<w:(header|footer)Reference w:type="(\w+)" r:id="([^"]+)"/g)].map(
      ([, kind, type, rId]) => {
        const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1];
        const xml = target ? partText(parts, `word/${target}`) : "";
        const text = [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join("");
        return `${kind} ${type}:${target ? text : "MISSING"}`;
      }
    )
  );
}

describe("streaming writer section breaks", () => {
  it("name the headers and footers the package has, as the buffered writer does", async () => {
    const stream = new Stream.DocumentWriter({ headers, footers, sectionProperties: final });
    stream.add(Build.paragraph([Build.text("Cover")], { sectionProperties: inner }));
    stream.add(Build.textParagraph("Body"));
    const fromStream = await packageParts(await stream.finalize());

    const d = Document.create();
    Document.addParagraph(d, "Cover");
    Document.addSectionBreak(d, inner);
    Document.addParagraph(d, "Body");
    Document.setSectionProperties(d, final);
    for (const [key, def] of headers) {
      Document.setHeader(d, key, def.content);
    }
    for (const [key, def] of footers) {
      Document.setFooter(d, key, def.content);
    }
    const fromBuffer = await packageParts(await Io.toBuffer(d));

    expect(packageIssues(fromStream)).toEqual([]);
    expect(sections(fromStream)).toEqual([
      ["header default:Cover header", "footer default:Cover footer"],
      ["header default:Body header", "footer default:Body footer"]
    ]);
    expect(sections(fromStream)).toEqual(sections(fromBuffer));
  });
});
