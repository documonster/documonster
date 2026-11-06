/**
 * Round trips that lost content or wrote an invalid package. Each case reads a
 * hand-written package, writes it, and checks the result is structurally
 * valid ({@link packageIssues}), keeps what the source had, and is stable — a
 * second read → write yields the same bytes.
 */

import { describe, expect, it } from "vitest";

import { keepClonedRangeMarkers } from "../core/range-markers";
import { WordHtml } from "../html";
import { Build, Document, Font, Io, Package, Query, Stream, Transform } from "../index";
import type { BodyContent, Paragraph } from "../types";
import {
  buildDocx,
  differingParts,
  packageIssues,
  packageParts,
  partText
} from "./roundtrip-integrity.helpers";

const P = (inner: string, pPr = ""): string => `<w:p>${pPr}${inner}</w:p>`;
const R = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const BS = (id: number, name: string): string => `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`;
const BE = (id: number): string => `<w:bookmarkEnd w:id="${id}"/>`;
const REV = `w:author="A" w:date="2024-01-01T00:00:00Z"`;

/** Write `source` read back, then once more; both packages and the body of the first. */
async function roundTrip(source: Uint8Array): Promise<{
  first: Map<string, Uint8Array>;
  second: Map<string, Uint8Array>;
  body: string;
}> {
  const once = await Io.toBuffer(await Io.read(source));
  const twice = await Io.toBuffer(await Io.read(once));
  const first = await packageParts(once);
  const second = await packageParts(twice);
  const xml = partText(first, "word/document.xml");
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(xml)![1]!;
  return { first, second, body };
}

/** The bookmark names, in document order, of a part. */
function bookmarkNames(xml: string): string[] {
  return [...xml.matchAll(/<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"/g)].map(m => m[1]!);
}

/** `roundTrip` plus the checks every case shares. */
async function expectSoundRoundTrip(source: Uint8Array): Promise<string> {
  const { first, second, body } = await roundTrip(source);
  expect(packageIssues(first)).toEqual([]);
  expect(differingParts(first, second)).toEqual([]);
  return body;
}

describe("block-level w:customXml", () => {
  it("keeps the content of a custom XML block, nested blocks and the markers inside", async () => {
    const source = await buildDocx({
      body:
        `<w:customXml w:uri="urn:x" w:element="section"><w:customXmlPr><w:attr w:name="a" w:val="1"/></w:customXmlPr>` +
        P(BS(1, "inside") + R("first") + BE(1)) +
        `<w:customXml w:element="inner">${P(R("nested"))}` +
        `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${P(R("cell"))}</w:tc></w:tr></w:tbl>` +
        `</w:customXml>` +
        BS(2, "across") +
        P(R("last")) +
        `</w:customXml>` +
        P(R("after") + BE(2))
    });
    const doc = await Io.read(source);
    expect(Query.getText(doc)).toContain("first");
    expect(Query.getText(doc)).toContain("nested");
    expect(Query.getText(doc)).toContain("cell");
    expect(Query.getText(doc)).toContain("last");
    const body = await expectSoundRoundTrip(source);
    expect(bookmarkNames(body)).toEqual(["inside", "across"]);
    expect(body).toMatch(/first[\s\S]*nested[\s\S]*cell[\s\S]*last[\s\S]*after/);
  });

  it("keeps custom XML around rows, cells and runs, and in a header", async () => {
    const source = await buildDocx({
      body:
        `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:customXml w:element="row"><w:tr><w:tc>${P(R("r1c1"))}</w:tc>` +
        `<w:customXml w:element="cell"><w:tc>${P(R("r1c2"))}</w:tc></w:customXml></w:tr></w:customXml>` +
        `</w:tbl>` +
        P(`<w:customXml w:element="run">${BS(1, "run")}${R("inline")}${BE(1)}</w:customXml>`),
      header: `<w:customXml w:element="h">${P(R("head"))}</w:customXml>`
    });
    const body = await expectSoundRoundTrip(source);
    expect(body).toMatch(/r1c1[\s\S]*r1c2[\s\S]*inline/);
    expect(bookmarkNames(body)).toEqual(["run"]);
    const { first } = await roundTrip(source);
    const headerPath = [...first.keys()].find(p => /^word\/header\d*\.xml$/.test(p))!;
    expect(partText(first, headerPath)).toContain("head");
  });

  it("is read the same by the streaming reader", async () => {
    const source = await buildDocx({
      body:
        `<w:customXml w:element="s">${BS(1, "s")}${P(R("streamed"))}${BE(1)}</w:customXml>` +
        P(R("tail"))
    });
    const reader = new Stream.DocumentReader(source);
    const blocks: BodyContent[] = [];
    for await (const block of reader) {
      blocks.push(block);
    }
    const batch = await Io.read(source);
    expect(blocks).toEqual(batch.body);
    expect(Query.getText(batch)).toContain("streamed");
  });
});

describe("range markers inside paragraph-level wrappers", () => {
  it("keeps bookmarks inside a hyperlink where they were", async () => {
    const body = await expectSoundRoundTrip(
      await buildDocx({
        body:
          P(
            `<w:hyperlink w:anchor="tgt">${BS(1, "lead")}${R("a")}${BS(2, "mid")}${R("b")}${BE(1)}${BE(2)}</w:hyperlink>`
          ) +
          P(
            BS(3, "spans") +
              R("pre ") +
              `<w:hyperlink w:anchor="tgt">${R("link")}${BE(3)}</w:hyperlink>`
          ) +
          P(BS(4, "tgt") + R("target") + BE(4))
      })
    );
    expect(bookmarkNames(body)).toEqual(["lead", "mid", "spans", "tgt"]);
    const links = [...body.matchAll(/<w:hyperlink\b[^>]*>([\s\S]*?)<\/w:hyperlink>/g)].map(m =>
      m[1]!.replace(/<w:r>[\s\S]*?<\/w:r>/g, "R").replace(/ w:name="[^"]*"/g, "")
    );
    expect(links).toEqual([
      `<w:bookmarkStart w:id="1"/>R<w:bookmarkStart w:id="2"/>R<w:bookmarkEnd w:id="1"/><w:bookmarkEnd w:id="2"/>`,
      `R<w:bookmarkEnd w:id="3"/>`
    ]);
  });

  it("keeps bookmarks inside tracked insertions, deletions and moves", async () => {
    const body = await expectSoundRoundTrip(
      await buildDocx({
        body:
          P(`<w:ins w:id="10" ${REV}>${BS(1, "ins")}${R("i")}${BE(1)}</w:ins>`) +
          P(
            `<w:del w:id="11" ${REV}><w:r><w:delText>d</w:delText></w:r>${BS(2, "del")}<w:r><w:delText>e</w:delText></w:r>${BE(2)}</w:del>`
          ) +
          P(`<w:moveFrom w:id="12" ${REV}>${BS(3, "from")}${R("m")}${BE(3)}</w:moveFrom>`) +
          P(`<w:moveTo w:id="13" ${REV}>${BS(4, "to")}${R("m")}${BE(4)}</w:moveTo>`) +
          P(`<w:ins w:id="14" ${REV}>${BS(5, "open")}${R("x")}</w:ins>`) +
          P(R("y") + BE(5))
      })
    );
    expect(bookmarkNames(body)).toEqual(["ins", "del", "from", "to", "open"]);
    // The deletion's two runs keep the bookmark between them.
    expect(body).toMatch(
      /d<\/w:delText>[\s\S]*<w:bookmarkStart w:id="\d+" w:name="del"\/>[\s\S]*e<\/w:delText>/
    );
  });

  it("keeps bookmarks inside a simple field, a smart tag and a run-level content control", async () => {
    const body = await expectSoundRoundTrip(
      await buildDocx({
        body:
          P(`<w:fldSimple w:instr=" DATE ">${BS(1, "field")}${R("today")}${BE(1)}</w:fldSimple>`) +
          P(
            `<w:smartTag w:uri="u" w:element="e">${BS(2, "smart")}${R("tag")}${BE(2)}</w:smartTag>`
          ) +
          P(
            R("x") +
              `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${BS(3, "sdt")}${R("in")}${BE(3)}</w:sdtContent></w:sdt>`
          ) +
          P(BS(4, "after") + R("a") + BE(4))
      })
    );
    expect(bookmarkNames(body)).toEqual(["field", "smart", "sdt", "after"]);
  });

  it("pairs a model bookmark with its end inside preserved markup when ids are reused", async () => {
    // Sequential reuse of an id is common in merged documents; the second
    // bookmark's end sits inside an element the reader keeps verbatim.
    const body = await expectSoundRoundTrip(
      await buildDocx({
        body:
          P(BS(1, "first") + R("a") + BE(1)) +
          P(BS(1, "second") + R("b") + `<w:bdo w:val="rtl">${R("c")}${BE(1)}</w:bdo>`) +
          P(`<w:bdo w:val="rtl">${BS(2, "inRaw")}${R("c")}</w:bdo>` + R("d") + BE(2)) +
          P(BS(2, "dupe") + R("e") + BE(2))
      })
    );
    expect(bookmarkNames(body)).toEqual(["first", "second", "inRaw", "dupe"]);
  });
});

describe("paragraph properties", () => {
  it("writes w:pPr once and remaps a section break's header and footer references", async () => {
    const source = await buildDocx({
      body:
        P(R("centered"), `<w:pPr><w:jc w:val="center"/></w:pPr>`) +
        P(
          R("section one"),
          `<w:pPr><w:jc w:val="right"/><w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>`
        ) +
        P(R("deleted mark"), `<w:pPr><w:rPr><w:del w:id="3" ${REV}/></w:rPr></w:pPr>`) +
        P(R("inserted mark"), `<w:pPr><w:rPr><w:ins w:id="4" ${REV}/></w:rPr></w:pPr>`),
      header: P(R("H"))
    });
    const body = await expectSoundRoundTrip(source);
    expect(body.match(/<w:pPr>/g)).toHaveLength(4);
    expect(body).toMatch(
      /<w:pPr><w:jc w:val="right"\/><w:sectPr><w:headerReference w:type="default" r:id="rId\d+"\/>/
    );
    const doc = await Io.read(source);
    for (const block of doc.body) {
      if (block.type === "paragraph") {
        expect(
          block.children.filter(c => "type" in c && c.type === "opaqueParagraphChild")
        ).toEqual([]);
      }
    }
  });
});

describe("font table", () => {
  const fontNames = (parts: Map<string, Uint8Array>): string[] =>
    [...partText(parts, "word/fontTable.xml").matchAll(/<w:font w:name="([^"]*)"/g)].map(
      m => m[1]!
    );

  it("declares each default font once, so the first write is already stable", async () => {
    const doc = Document.create();
    Document.addContent(doc, Build.textParagraph("defaults"));
    const source = await Io.toBuffer(doc);
    const names = fontNames(await packageParts(source));
    expect(new Set(names).size).toBe(names.length);
    const { first, second } = await roundTrip(source);
    expect(differingParts(await packageParts(source), first)).toEqual([]);
    expect(differingParts(first, second)).toEqual([]);
  });

  it("keeps embedded fonts and their relationships stable across round trips", async () => {
    const doc = Document.create();
    Document.addContent(doc, Build.textParagraph("fonts"));
    const ttf = new Uint8Array(64);
    ttf.set([0, 1, 0, 0]);
    const embedded = Font.addEmbedded(doc, [
      Font.embed({ name: "Embedded Face", data: ttf, style: "regular" })
    ]);
    const source = await Io.toBuffer(embedded);
    const { first, second } = await roundTrip(source);
    expect(differingParts(first, second)).toEqual([]);
    expect(packageIssues(first)).toEqual([]);
    const fontTable = partText(first, "word/fontTable.xml");
    const names = fontNames(first);
    expect(new Set(names).size).toBe(names.length);
    expect(fontTable).toMatch(
      /<w:font w:name="Embedded Face">[\s\S]*?<w:embedRegular r:id="([^"]+)"/
    );
    expect(partText(first, "word/_rels/fontTable.xml.rels")).toContain("fonts/");
    // The very first write is already stable.
    expect(differingParts(await packageParts(source), first)).toEqual([]);
  });
});

describe("drawing ids", () => {
  const NSX =
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ` +
    `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
    `xmlns:v="urn:schemas-microsoft-com:vml"`;
  /** A text box the way Word writes one: DrawingML in `mc:Choice`, VML in `mc:Fallback`. */
  const textBox = (docPrId: number, inner: string): string =>
    `<w:r><mc:AlternateContent ${NSX}><mc:Choice Requires="wps"><w:drawing>` +
    `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="914400" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
    `<wp:docPr id="${docPrId}" name="Text Box ${docPrId}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
    `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
    `<mc:Fallback><w:pict><v:rect style="width:72pt;height:72pt"><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback>` +
    `</mc:AlternateContent></w:r>`;

  it("writes a body text box once, with an id no other drawing has", async () => {
    const source = await buildDocx({
      body: P(BS(1, "body") + R("body") + BE(1)) + P(textBox(1, P(R("in box")))) + P(R("after")),
      header: P(textBox(1, P(R("header box"))))
    });
    const body = await expectSoundRoundTrip(source);
    expect(body.match(/<wp:anchor\b/g)).toHaveLength(1);
    expect(body.match(/in box/g)?.length).toBeGreaterThanOrEqual(1);
    const doc = await Io.read(source);
    expect(doc.body.filter(b => b.type === "drawingShape")).toHaveLength(1);
  });

  it("numbers model drawings around the ids preserved markup already uses", async () => {
    // The header is not searched for floating drawings, so its text box is
    // kept verbatim with its id; the body's are written from the model.
    const source = await buildDocx({
      body: P(R("a")) + P(textBox(2, P(R("one")))) + P(textBox(1, P(R("two")))),
      header: P(textBox(1, P(R("kept verbatim"))))
    });
    await expectSoundRoundTrip(source);
  });
});

/** `xml` with the noise a writer may legitimately change removed, for comparing paragraph content. */
function contentShape(xml: string): string {
  return xml
    .replace(/<w:(pPr|rPr)\b[\s\S]*?<\/w:\1>/g, "")
    .replace(/<w:t\b[^>]*>/g, "<w:t>")
    .replace(/<w:delText\b[^>]*>/g, "<w:delText>")
    .replace(/<w:instrText\b[^>]*>/g, "<w:instrText>")
    .replace(/<\/w:t><\/w:r><w:r><w:t>/g, "");
}

/** The first `w:p` (outermost) of a body. */
function firstParagraph(body: string): string {
  const start = body.indexOf("<w:p>");
  let depth = 0;
  for (const m of body.slice(start).matchAll(/<w:p>|<\/w:p>/g)) {
    depth += m[0] === "<w:p>" ? 1 : -1;
    if (depth === 0) {
      return body.slice(start, start + m.index! + m[0].length);
    }
  }
  return body.slice(start);
}

describe("bidirectional and smart tag wrappers", () => {
  const source = P(
    R("a") +
      `<w:dir w:val="rtl">${BS(1, "inDir")}${R("b")}<w:bdo w:val="ltr">${R("c")}</w:bdo>${BE(1)}</w:dir>` +
      `<w:smartTag w:uri="urn:x" w:element="place"><w:smartTagPr><w:attr w:name="k" w:val="v"/></w:smartTagPr>${R("d")}</w:smartTag>` +
      R("e")
  );

  it("writes w:dir, w:bdo and w:smartTag back as they were", async () => {
    const body = await expectSoundRoundTrip(await buildDocx({ body: source }));
    expect(contentShape(firstParagraph(body))).toBe(contentShape(source));
  });

  it("still exposes the text inside them", async () => {
    const doc = await Io.read(await buildDocx({ body: source }));
    expect(Query.getText(doc)).toContain("abcde");
  });

  it("reads the wrappers as markers riding on ordinary children, never as children", async () => {
    const doc = await Io.read(
      await buildDocx({
        body:
          source +
          P(
            `<w:hyperlink w:anchor="t"><w:dir w:val="rtl">${R("x")}</w:dir>` +
              `<w:smartTag w:element="e">${R("y")}</w:smartTag></w:hyperlink>`
          )
      })
    );
    const spanTypes = new Set(["bidiStart", "bidiEnd", "smartTagStart", "smartTagEnd"]);
    const childTypes: string[] = [];
    const markerTypes: string[] = [];
    Query.each(doc, {
      enterParagraph(para) {
        for (const child of para.children) {
          childTypes.push("type" in child ? child.type : "run");
        }
      },
      visitRangeMarker(marker) {
        markerTypes.push(marker.type);
      }
    });
    for (const block of doc.body) {
      if (block.type !== "paragraph") {
        continue;
      }
      for (const child of block.children) {
        if ("type" in child && child.type === "hyperlink") {
          expect(child.children.every(run => !("type" in run))).toBe(true);
          for (const n of child.nestedChildren ?? []) {
            expect(spanTypes.has(n.child.type)).toBe(false);
          }
        }
      }
    }
    expect(childTypes.filter(t => spanTypes.has(t))).toEqual([]);
    // The smart tag's content is unwrapped into runs, as it always was.
    const first = doc.body[0] as Paragraph;
    expect(first.children.map(c => ("type" in c ? c.type : "run"))).toEqual([
      "run",
      "bookmarkStart",
      "run",
      "run",
      "bookmarkEnd",
      "run",
      "run"
    ]);
    expect(markerTypes).toEqual(
      expect.arrayContaining(["bidiStart", "bidiEnd", "smartTagStart", "smartTagEnd"])
    );
  });

  it("renders the direction in HTML", async () => {
    const doc = await Io.read(await buildDocx({ body: source }));
    const { html } = WordHtml.fromDocx(doc);
    expect(html).toMatch(
      /<span dir="rtl">[\s\S]*b[\s\S]*<bdo dir="ltr">[\s\S]*c[\s\S]*<\/bdo>[\s\S]*<\/span>/
    );
  });
});

describe("hyperlinks with nested content", () => {
  it("keeps fields, revisions, smart tags, content controls and bidi spans inside a link", async () => {
    const link =
      `<w:hyperlink w:anchor="tgt">${R("a")}` +
      `<w:fldSimple w:instr=" PAGE ">${R("1")}</w:fldSimple>` +
      `<w:ins w:id="20" ${REV}>${BS(1, "inIns")}${R("b")}${BE(1)}</w:ins>` +
      `<w:del w:id="21" ${REV}><w:r><w:delText>c</w:delText></w:r></w:del>` +
      `<w:smartTag w:uri="urn:x" w:element="t">${R("d")}</w:smartTag>` +
      `<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent>${R("e")}</w:sdtContent></w:sdt>` +
      `<w:dir w:val="rtl">${R("f")}</w:dir>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">PAGEREF tgt</w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${R("7")}<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `${R("g")}</w:hyperlink>`;
    const body = await expectSoundRoundTrip(
      await buildDocx({ body: P(link) + P(BS(2, "tgt") + R("target") + BE(2)) })
    );
    const written = /<w:hyperlink\b[\s\S]*<\/w:hyperlink>/.exec(firstParagraph(body))![0];
    // fldSimple is read as a field, which the writer spells as a complex field.
    // A tracked change is one revision element per run, with the markers
    // that were inside it beside it — the same text positions.
    const expected = contentShape(link)
      .replace(
        `<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>`,
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r>` +
          `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
      )
      .replace(
        `<w:ins w:id="20" ${REV}>${BS(1, "inIns")}<w:r><w:t>b</w:t></w:r>${BE(1)}</w:ins>`,
        `${BS(1, "inIns")}<w:ins w:id="20" ${REV}><w:r><w:t>b</w:t></w:r></w:ins>${BE(1)}`
      );
    expect(contentShape(written)).toBe(expected);
    const doc = await Io.read(await buildDocx({ body: P(link) }));
    expect(Query.getText(doc)).toContain("a");
    expect(Query.getText(doc)).toMatch(/b[\s\S]*f[\s\S]*g/);
  });
});

describe("hyperlinks with nested content, edited", () => {
  const link =
    `<w:hyperlink w:anchor="t">${R("a")}<w:ins w:id="30" ${REV}>${R("new")}</w:ins>` +
    `<w:del w:id="31" ${REV}><w:r><w:delText>old</w:delText></w:r></w:del>${R("z")}</w:hyperlink>`;

  it("accepts and rejects the tracked changes inside a link", async () => {
    const accepted = await Io.read(await buildDocx({ body: P(link) }));
    expect(Query.acceptAllRevisions(accepted)).toBe(2);
    const acceptedBody = /<w:hyperlink\b[\s\S]*?<\/w:hyperlink>/.exec(
      partText(await packageParts(await Io.toBuffer(accepted)), "word/document.xml")
    )![0];
    // `contentShape` joins adjacent text runs.
    expect(contentShape(acceptedBody)).toBe(
      `<w:hyperlink w:anchor="t"><w:r><w:t>anewz</w:t></w:r></w:hyperlink>`
    );
    const rejected = await Io.read(await buildDocx({ body: P(link) }));
    expect(Query.rejectAllRevisions(rejected)).toBe(2);
    expect(Query.getText(rejected)).toContain("aoldz");
  });

  it("keeps the nested content through Transform.map", async () => {
    const doc = await Io.read(await buildDocx({ body: P(link) }));
    const mapped = Transform.map(doc, { transformRun: run => run });
    const body = /<w:hyperlink\b[\s\S]*?<\/w:hyperlink>/.exec(
      partText(await packageParts(await Io.toBuffer(mapped)), "word/document.xml")
    )![0];
    expect(contentShape(body)).toBe(contentShape(link));
  });
});

describe("range markers inside complex fields", () => {
  const begin = `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`;
  const sep = `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
  const end = `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const instr = (text: string): string =>
    `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;

  /** The sequence of field parts, text and bookmark names in a paragraph. */
  function tokens(xml: string): string[] {
    const out: string[] = [];
    for (const m of xml.matchAll(
      /<w:fldChar w:fldCharType="(\w+)"\/>|<w:instrText[^>]*>([^<]*)<\/w:instrText>|<w:t\b[^>]*>([^<]*)<\/w:t>|<w:bookmarkStart\b[^>]*w:name="([^"]*)"\/>|<w:bookmarkEnd\b[^>]*\/>/g
    )) {
      out.push(
        m[1] !== undefined
          ? m[1]
          : m[2] !== undefined
            ? `i:${m[2]}`
            : m[3] !== undefined
              ? `t:${m[3]}`
              : m[4] !== undefined
                ? `[${m[4]}`
                : "]"
      );
    }
    // Adjacent instruction or text pieces are one string, however split.
    return out.reduce<string[]>((acc, t) => {
      const last = acc[acc.length - 1];
      if (
        last &&
        t.slice(0, 2) === last.slice(0, 2) &&
        (t.startsWith("i:") || t.startsWith("t:"))
      ) {
        acc[acc.length - 1] = last + t.slice(2);
      } else {
        acc.push(t);
      }
      return acc;
    }, []);
  }

  it("keeps markers between begin, the instruction, separate, the result and end where they were", async () => {
    const para = P(
      BS(1, "around") +
        begin +
        BS(2, "lead") +
        instr("REF ") +
        BS(3, "instr") +
        instr("tgt") +
        BE(3) +
        BE(2) +
        sep +
        BS(4, "result") +
        R("re") +
        BS(5, "mid") +
        R("sult") +
        BE(5) +
        BE(4) +
        end +
        BE(1)
    );
    const source = await buildDocx({ body: para + P(BS(6, "tgt") + R("x") + BE(6)) });
    const body = await expectSoundRoundTrip(source);
    expect(tokens(firstParagraph(body))).toEqual(tokens(para));
    const doc = await Io.read(source);
    const names: string[] = [];
    Query.each(doc, {
      visitBookmarkStart(bm) {
        names.push(bm.name);
      }
    });
    expect(names).toEqual(
      expect.arrayContaining(["around", "lead", "instr", "result", "mid", "tgt"])
    );
  });

  it("keeps a marker in place in a paragraph the field only starts in", async () => {
    const source = await buildDocx({
      body:
        P(R("x") + begin + instr(" TOC ") + sep + BS(1, "first") + R("one")) +
        P(R("two") + BE(1) + end)
    });
    const body = await expectSoundRoundTrip(source);
    expect(bookmarkNames(body)).toEqual(["first"]);
  });
});

describe("Package.edit", () => {
  it("numbers new drawings around the ids the parts it leaves alone use", async () => {
    const NSX =
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
    const drawing = (id: number): string =>
      `<w:r><w:drawing><wp:inline ${NSX}><wp:extent cx="10" cy="10"/><wp:docPr id="${id}" name="d${id}"/>` +
      `<a:graphic><a:graphicData uri="urn:other"/></a:graphic></wp:inline></w:drawing></w:r>`;
    const source = await buildDocx({ body: P(R("old")), header: P(drawing(1)) + P(drawing(2)) });
    const edited = await Package.edit(source, [
      {
        type: "replaceBody",
        body: [
          { type: "drawingShape", shapeType: "rect", width: 914400, height: 914400 },
          { type: "drawingShape", shapeType: "rect", width: 914400, height: 914400 }
        ]
      }
    ]);
    expect(packageIssues(await packageParts(edited))).toEqual([]);
  });
});

describe("span markers in repeated content", () => {
  it("keep their span on every copy, while a bookmark stays on the first", () => {
    const make = (): Paragraph => ({
      type: "paragraph",
      children: [
        {
          content: [{ type: "text", text: "x" }],
          rangeMarkersBefore: [{ type: "bidiStart", kind: "embedding", direction: "rtl" }]
        },
        { type: "bookmarkStart", id: 1, name: "b" },
        { type: "bookmarkEnd", id: 1, rangeMarkersBefore: [{ type: "bidiEnd" }] }
      ]
    });
    const copies = [0, 1, 2].map(i => keepClonedRangeMarkers(make(), i, 3));
    for (const [i, copy] of copies.entries()) {
      const types = copy.children.map(c => ("type" in c ? c.type : "run"));
      const markers = [
        ...copy.children.flatMap(c => c.rangeMarkersBefore ?? []),
        ...(copy.rangeMarkersAfter ?? [])
      ].map(m => m.type);
      expect(markers).toEqual(["bidiStart", "bidiEnd"]);
      expect(types.includes("bookmarkStart")).toBe(i === 0);
    }
  });
});
