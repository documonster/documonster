/**
 * Range editing permissions (`w:permStart` / `w:permEnd`): the editable
 * regions of a protected document. They sit inside paragraphs, between blocks,
 * between table rows and cells, after the last block and inside SDTs, and must
 * survive a round-trip with every attribute and with start/end still paired.
 */

import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { describe, expect, it } from "vitest";

import { utf8Decoder, utf8Encoder } from "../core/internal-utils";
import { Build, Document, Io, Query, Security, Stream } from "../index";
import type { BodyContent, DocxDocument, Paragraph, RangeMarker } from "../types";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

async function entriesOf(buffer: Uint8Array): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for await (const entry of unzip(buffer).entries()) {
    out.set(entry.path.replace(/^\/+/, ""), await entry.bytes());
  }
  return out;
}

/** A real package, protected read-only, whose body is `bodyXml`. */
async function protectedDocxWithBody(bodyXml: string): Promise<Uint8Array> {
  const doc = Document.create();
  Document.addContent(doc, Build.textParagraph("placeholder"));
  const entries = await entriesOf(
    await Io.toBuffer(await Security.protect(doc, { edit: "readOnly" }))
  );
  entries.set(
    "word/document.xml",
    utf8Encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:document xmlns:w="${W_NS}"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`
    )
  );
  const archive = zip();
  for (const [p, d] of entries) {
    archive.add(p, d);
  }
  return archive.toBuffer();
}

async function documentXml(bytes: Uint8Array): Promise<string> {
  return utf8Decoder.decode((await entriesOf(bytes)).get("word/document.xml")!);
}

/** Every range marker in document order, as reported by the walker. */
function markers(doc: DocxDocument): RangeMarker[] {
  const out: RangeMarker[] = [];
  Query.each(doc, { visitRangeMarker: m => void out.push(m) });
  return out;
}

/** The `w:permStart` / `w:permEnd` elements of a part, in order. */
function permElements(xml: string): string[] {
  return xml.match(/<w:perm(?:Start|End)\b[^>]*\/>/g) ?? [];
}

function p(text: string, before = "", after = ""): string {
  return `<w:p>${before}<w:r><w:t>${text}</w:t></w:r>${after}</w:p>`;
}

function tc(text: string): string {
  return `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
}

const GRID = `<w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>`;

/** What Word writes for a read-only document with editable regions. */
const PROTECTED_BODY = [
  p("Read only"),
  p("Anyone may edit this", `<w:permStart w:id="1" w:edGrp="everyone"/>`, `<w:permEnd w:id="1"/>`),
  `<w:permStart w:id="2" w:ed="DOMAIN\\jane"/>`,
  `<w:tbl>${GRID}`,
  `<w:tr>${tc("A1")}${tc("B1")}</w:tr>`,
  `<w:tr><w:permStart w:id="3" w:edGrp="editors" w:colFirst="0" w:colLast="1" w:displacedByCustomXml="next"/>`,
  `${tc("A2")}${tc("B2")}<w:permEnd w:id="3" w:displacedByCustomXml="prev"/></w:tr>`,
  `</w:tbl>`,
  p("Jane's region ends here", "", `<w:permEnd w:id="2"/>`),
  `<w:permStart w:id="4" w:ed="jane@example.com"/><w:permEnd w:id="4"/>`
].join("");

const EXPECTED: RangeMarker[] = [
  { type: "permStart", id: 1, edGrp: "everyone" },
  { type: "permEnd", id: 1 },
  { type: "permStart", id: 2, ed: "DOMAIN\\jane" },
  {
    type: "permStart",
    id: 3,
    edGrp: "editors",
    colFirst: 0,
    colLast: 1,
    displacedByCustomXml: "next"
  },
  { type: "permEnd", id: 3, displacedByCustomXml: "prev" },
  { type: "permEnd", id: 2 },
  { type: "permStart", id: 4, ed: "jane@example.com" },
  { type: "permEnd", id: 4 }
];

describe("range permissions (w:permStart / w:permEnd)", () => {
  it("reads every permission range, with all attributes, where it sits", async () => {
    const doc = await Io.read(await protectedDocxWithBody(PROTECTED_BODY));
    expect(Security.isProtected(doc)).toBe(true);
    expect(markers(doc)).toEqual(EXPECTED);
    // Paragraph-level, block-level, table-level and trailing positions.
    const editable = doc.body[1]!;
    expect(editable.type === "paragraph" && editable.children[0]!.rangeMarkersBefore).toEqual([
      EXPECTED[0]
    ]);
    expect(editable.type === "paragraph" && editable.rangeMarkersAfter).toEqual([EXPECTED[1]]);
    expect(doc.body[2]!.rangeMarkersBefore).toEqual([EXPECTED[2]]);
    const table = doc.body[2]!;
    expect(table.type === "table" && table.rows[1]!.cells[0]!.rangeMarkersBefore).toEqual([
      EXPECTED[3]
    ]);
    expect(table.type === "table" && table.rows[1]!.rangeMarkersAfter).toEqual([EXPECTED[4]]);
    expect(doc.bodyRangeMarkersAfter).toEqual([EXPECTED[6], EXPECTED[7]]);
    expect(doc.body.every(b => b.type === "paragraph" || b.type === "table")).toBe(true);
  });

  it("never puts a permission marker among a paragraph's children", async () => {
    const bytes = await protectedDocxWithBody(
      PROTECTED_BODY +
        p("Only a permission", `<w:permStart w:id="6" w:edGrp="everyone"/>`, "") +
        `<w:p><w:permStart w:id="7" w:edGrp="everyone"/><w:permEnd w:id="7"/></w:p>` +
        `<w:p><w:smartTag w:element="x"><w:permStart w:id="8" w:edGrp="everyone"/><w:r><w:t>tagged</w:t></w:r><w:permEnd w:id="8"/></w:smartTag></w:p>`
    );
    // Permission 6 has no end: it is read, and not written back — a range
    // that cannot be paired in a part is removed from it.
    for (const [doc, expected] of [
      [await Io.read(bytes), [1, 2, 3, 4, 6, 7, 8]],
      [await Io.read(await Io.toBuffer(await Io.read(bytes))), [1, 2, 3, 4, 7, 8]]
    ] as const) {
      const seen: string[] = [];
      Query.each(doc, {
        enterParagraph(para) {
          for (const child of para.children) {
            seen.push("type" in child ? child.type : "run");
          }
        }
      });
      expect(seen.length).toBeGreaterThan(0);
      expect(seen.filter(t => t === "permStart" || t === "permEnd")).toEqual([]);
      const ids = markers(doc)
        .filter(m => m.type === "permStart")
        .map(m => m.id);
      expect(ids).toEqual(expected);
    }
  });

  it("round-trips the ranges unchanged, byte for byte in document.xml", async () => {
    const bytes = await protectedDocxWithBody(PROTECTED_BODY);
    const first = await Io.read(bytes);
    const written = await Io.toBuffer(first);
    expect(permElements(await documentXml(written))).toEqual(
      permElements(await documentXml(bytes))
    );
    const second = await Io.read(written);
    expect(markers(second)).toEqual(EXPECTED);
    expect(Security.isProtected(second)).toBe(true);
  });

  it("keeps written permission ids unique and each end on its start", async () => {
    // Two documents each using perm id 1 — as merging produces.
    const doc = Document.create();
    const editable = (text: string, start: RangeMarker): Paragraph => ({
      ...Build.paragraph([{ ...Build.text(text), rangeMarkersBefore: [start] }]),
      rangeMarkersAfter: [{ type: "permEnd", id: 1 }]
    });
    Document.addContent(doc, editable("first", { type: "permStart", id: 1, edGrp: "everyone" }));
    Document.addContent(
      doc,
      editable("second", { type: "permStart", id: 1, ed: "jane@example.com" })
    );
    const xml = await documentXml(await Io.toBuffer(doc));
    const elements = permElements(xml);
    const ids = elements.map(e => /w:id="(\d+)"/.exec(e)![1]);
    expect(ids[0]).toBe(ids[1]);
    expect(ids[2]).toBe(ids[3]);
    expect(ids[0]).not.toBe(ids[2]);
    expect(elements[2]).toContain(`w:ed="jane@example.com"`);
  });

  it("drops a permEnd that closes nothing, as for bookmarks", async () => {
    const doc = Document.create();
    Document.addContent(doc, {
      ...Build.paragraph([Build.text("x")]),
      rangeMarkersAfter: [{ type: "permEnd", id: 9 }]
    });
    expect(permElements(await documentXml(await Io.toBuffer(doc)))).toEqual([]);
  });

  it("the streaming reader and writer carry them too", async () => {
    const bytes = await protectedDocxWithBody(PROTECTED_BODY);
    const reader = new Stream.DocumentReader(bytes);
    const body: BodyContent[] = [];
    for await (const item of reader) {
      body.push(item);
    }
    const streamed: DocxDocument = { body, bodyRangeMarkersAfter: reader.bodyRangeMarkersAfter };
    expect(markers(streamed)).toEqual(EXPECTED);

    const stream = new Stream.DocumentWriter({
      bodyRangeMarkersAfter: streamed.bodyRangeMarkersAfter
    });
    stream.addMany(body);
    const out = await stream.finalize();
    expect(permElements(await documentXml(out))).toEqual(permElements(await documentXml(bytes)));
    expect(markers(await Io.read(out))).toEqual(EXPECTED);
  });

  it("keeps markers in an SDT made of runs, where no paragraph can hold them", async () => {
    const xml = [
      p("Before"),
      `<w:sdt><w:sdtPr><w:alias w:val="Name"/></w:sdtPr><w:sdtContent>`,
      `<w:permStart w:id="5" w:edGrp="everyone"/><w:bookmarkStart w:id="0" w:name="name"/>`,
      `<w:r><w:t>Jane</w:t></w:r>`,
      `<w:bookmarkEnd w:id="0"/><w:permEnd w:id="5"/>`,
      `</w:sdtContent></w:sdt>`,
      p("After")
    ].join("");
    const bytes = await protectedDocxWithBody(xml);
    const doc = await Io.read(bytes);
    const sdt = doc.body[1]!;
    expect(sdt.type).toBe("sdt");
    if (sdt.type !== "sdt") {
      return;
    }
    expect(sdt.content[0]).toMatchObject({
      rangeMarkersBefore: [
        { type: "permStart", id: 5, edGrp: "everyone" },
        { type: "bookmarkStart", id: 0, name: "name" }
      ]
    });
    expect(sdt.contentRangeMarkersAfter).toEqual([
      { type: "bookmarkEnd", id: 0 },
      { type: "permEnd", id: 5 }
    ]);
    const written = await documentXml(await Io.toBuffer(doc));
    expect(written).toContain(
      `<w:sdtContent><w:permStart w:id="5" w:edGrp="everyone"/><w:bookmarkStart w:id="0" w:name="name"/><w:r>`
    );
    expect(written).toContain(
      `</w:r><w:bookmarkEnd w:id="0"/><w:permEnd w:id="5"/></w:sdtContent>`
    );
    expect(markers(await Io.read(await Io.toBuffer(doc)))).toEqual(markers(doc));
  });
});
