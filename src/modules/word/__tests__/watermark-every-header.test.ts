/**
 * `DocxDocument.watermark` is the document's watermark, so it is drawn on
 * every page, as Word's Design › Watermark does: in every header a section
 * shows — default, first page (`titlePage`), even pages
 * (`evenAndOddHeaders`) — in every section. A section showing no header of a
 * type is given one holding only the watermark. Reading applies the converse:
 * a watermark is the document's only when every header shown draws it, and
 * is otherwise kept as header content. The streaming writer agrees.
 */

import { zip } from "@archive/create-archive";
import { describe, expect, it } from "vitest";

import { utf8Encoder } from "../core/internal-utils";
import { Build, Document, Io, Stream } from "../index";
import type { BodyContent, DocxDocument, HeaderDef, SectionProperties, Watermark } from "../index";
import {
  differingParts,
  packageIssues,
  packageParts,
  partText
} from "./roundtrip-integrity.helpers";

const DRAFT: Watermark = { type: "text", text: "DRAFT" };
const PAGE = { pageSize: { width: 12240, height: 15840 } } as const;

const header = (text: string): HeaderDef => ({
  content: { children: [Build.textParagraph(text)] }
});

interface Case {
  readonly name: string;
  readonly headers?: ReadonlyMap<string, HeaderDef>;
  readonly footers?: ReadonlyMap<string, HeaderDef>;
  readonly final?: SectionProperties;
  /** Section breaks, each after a paragraph of its own. */
  readonly inner?: readonly SectionProperties[];
  readonly evenAndOddHeaders?: boolean;
  /** Per section: each reference's type, text, and watermark count. */
  readonly expected: readonly (readonly string[])[];
}

const CASES: Case[] = [
  {
    name: "a single section with a default header",
    headers: new Map([["default", header("Head")]]),
    expected: [["default:Head:1"]]
  },
  {
    name: "a single section without a header",
    expected: [["default::1"]]
  },
  {
    name: "a distinct first page with both headers",
    headers: new Map([
      ["default", header("Head")],
      ["first", header("Cover")]
    ]),
    final: { ...PAGE, titlePage: true },
    expected: [["default:Head:1", "first:Cover:1"]]
  },
  {
    name: "a distinct first page without a first-page header",
    headers: new Map([["default", header("Head")]]),
    final: { ...PAGE, titlePage: true },
    expected: [["default:Head:1", "first::1"]]
  },
  {
    name: "even and odd headers without an even-page header",
    headers: new Map([["default", header("Head")]]),
    evenAndOddHeaders: true,
    expected: [["default:Head:1", "even::1"]]
  },
  {
    name: "a section break before a section with a header",
    headers: new Map([["default", header("Head")]]),
    inner: [PAGE],
    expected: [["default::1"], ["default:Head:1"]]
  },
  {
    name: "a section break, neither section naming a header",
    inner: [PAGE],
    // The first section names the watermark's header; the last inherits it.
    expected: [["default::1"], []]
  },
  {
    name: "sections naming headers and footers of their own",
    headers: new Map([
      ["inner", header("Inner")],
      ["default", header("Head")]
    ]),
    footers: new Map([
      ["innerFoot", header("Inner foot")],
      ["default", header("Foot")]
    ]),
    inner: [
      {
        ...PAGE,
        headers: [{ type: "default", rId: "inner" }],
        footers: [{ type: "default", rId: "innerFoot" }]
      }
    ],
    final: {
      ...PAGE,
      headers: [{ type: "default", rId: "default" }],
      footers: [{ type: "default", rId: "default" }]
    },
    expected: [
      ["default:Inner:1", "footer default:Inner foot:0"],
      ["default:Head:1", "footer default:Foot:0"]
    ]
  },
  {
    name: "a section break inheriting the first section's header",
    headers: new Map([["inner", header("Inner")]]),
    inner: [{ ...PAGE, headers: [{ type: "default", rId: "inner" }] }],
    final: { ...PAGE, titlePage: true },
    expected: [["default:Inner:1"], ["first::1"]]
  }
];

function buffered(c: Case): DocxDocument {
  const d = Document.create();
  for (const [i, sect] of (c.inner ?? []).entries()) {
    Document.addParagraph(d, `Section ${i + 1}`);
    Document.addSectionBreak(d, sect);
  }
  Document.addParagraph(d, "Last section");
  if (c.final) {
    Document.setSectionProperties(d, c.final);
  }
  if (c.evenAndOddHeaders) {
    Document.setSettings(d, { ...d.settings, evenAndOddHeaders: true });
  }
  for (const [key, def] of c.headers ?? []) {
    Document.setHeader(d, key, def.content);
  }
  for (const [key, def] of c.footers ?? []) {
    Document.setFooter(d, key, def.content);
  }
  Document.setWatermark(d, DRAFT);
  return d;
}

async function streamed(c: Case): Promise<Uint8Array> {
  const stream = new Stream.DocumentWriter({
    watermark: DRAFT,
    headers: c.headers,
    footers: c.footers,
    sectionProperties: c.final ?? PAGE,
    settings: c.evenAndOddHeaders ? { evenAndOddHeaders: true } : undefined
  });
  for (const [i, sect] of (c.inner ?? []).entries()) {
    stream.add(Build.paragraph([Build.text(`Section ${i + 1}`)], { sectionProperties: sect }));
  }
  stream.add(Build.textParagraph("Last section"));
  return stream.finalize();
}

/**
 * Per section, each header reference as `type:text:watermarks`, then each
 * footer reference as `footer type:text:watermarks`.
 */
function sectionSummary(parts: Map<string, Uint8Array>): string[][] {
  const doc = partText(parts, "word/document.xml");
  const rels = partText(parts, "word/_rels/document.xml.rels");
  return [...doc.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map(([sect]) =>
    [...sect.matchAll(/<w:(header|footer)Reference w:type="(\w+)" r:id="([^"]+)"/g)]
      .sort((a, b) => (a[1] === b[1] ? 0 : a[1] === "header" ? -1 : 1))
      .map(([, kind, type, rId]) => {
        const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1];
        const xml = target ? partText(parts, `word/${target}`) : "";
        const text = [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join("");
        const summary = `${type}:${target ? text : "MISSING"}:${xml.match(/PowerPlusWaterMarkObject/g)?.length ?? 0}`;
        return kind === "footer" ? `footer ${summary}` : summary;
      })
  );
}

describe("the document watermark", () => {
  it.each(CASES.map(c => [c.name, c] as const))("is in every header shown: %s", async (_n, c) => {
    const parts = await packageParts(await Io.toBuffer(buffered(c)));
    expect(packageIssues(parts)).toEqual([]);
    expect(sectionSummary(parts)).toEqual(c.expected);
  });

  it.each(CASES.map(c => [c.name, c] as const))(
    "is read back once and written back the same: %s",
    async (_n, c) => {
      const written = await Io.toBuffer(buffered(c));
      const read = await Io.read(written);
      expect(read.watermark).toMatchObject(DRAFT);
      // Taken out of every header that drew it: no header content carries it.
      for (const def of read.headers?.values() ?? []) {
        expect(JSON.stringify(def.content)).not.toContain("WaterMark");
      }
      const again = await packageParts(await Io.toBuffer(read));
      expect(sectionSummary(again)).toEqual(c.expected);
      const thrice = await packageParts(await Io.toBuffer(await Io.read(await Io.toBuffer(read))));
      expect(differingParts(again, thrice)).toEqual([]);
    }
  );

  it.each(CASES.map(c => [c.name, c] as const))(
    "is placed the same by the streaming writer: %s",
    async (_n, c) => {
      const fromBuffer = await packageParts(await Io.toBuffer(buffered(c)));
      const fromStream = await packageParts(await streamed(c));
      expect(packageIssues(fromStream)).toEqual([]);
      expect(sectionSummary(fromStream)).toEqual(sectionSummary(fromBuffer));
    }
  );

  it.each(CASES.map(c => [c.name, c] as const))(
    "leaves the streaming reader the section references Io.read leaves: %s",
    async (_n, c) => {
      const written = await Io.toBuffer(buffered(c));
      const read = await Io.read(written);
      const reader = new Stream.DocumentReader(written);
      const body: BodyContent[] = [];
      for await (const block of reader) {
        body.push(block);
      }
      expect(body).toEqual(read.body);
      expect(reader.sectionProperties).toEqual(read.sectionProperties);
    }
  );

  it("is header content when a header shown does not draw it", async () => {
    const c = CASES[2]!; // default and first-page headers
    const parts = await packageParts(await Io.toBuffer(buffered(c)));
    // Take the watermark out of the first-page header only.
    const rels = partText(parts, "word/_rels/document.xml.rels");
    const doc = partText(parts, "word/document.xml");
    const firstRId = /<w:headerReference w:type="first" r:id="([^"]+)"/.exec(doc)![1]!;
    const firstPart = `word/${new RegExp(`Id="${firstRId}"[^>]*Target="([^"]+)"`).exec(rels)![1]!}`;
    const archive = zip();
    for (const [path, bytes] of parts) {
      archive.add(
        path,
        path === firstPart
          ? utf8Encoder.encode(
              partText(parts, path).replace(
                /<w:r>(?:(?!<w:r>)[\s\S])*?PowerPlusWaterMarkObject[\s\S]*?<\/w:r>/,
                ""
              )
            )
          : bytes
      );
    }
    const edited = await archive.toBuffer();
    expect(sectionSummary(await packageParts(edited))).toEqual([
      ["default:Head:1", "first:Cover:0"]
    ]);

    const read = await Io.read(edited);
    expect(read.watermark).toBeUndefined();
    const out = await packageParts(await Io.toBuffer(read));
    // Kept where it was: in the default header, as its content.
    expect(sectionSummary(out)).toEqual([["default:Head:1", "first:Cover:0"]]);
  });
});
