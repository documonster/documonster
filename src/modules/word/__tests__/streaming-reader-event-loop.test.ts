/**
 * The streaming reader lets the event loop run while it reads, however the package stores it.
 *
 * It parsed each piece synchronously and handed elements out through promises alone, so the event loop ran only when
 * something happened to wait — inflating a part. A stored (uncompressed) `document.xml` never waits, the source being
 * in memory, and was read whole in one task; a Flat OPC document and a large style part were each parsed in one call
 * before the first element came out.
 */

import { Archive } from "@archive/index";
import { runStepsPaced, scheduleMacrotask } from "@utils/yield";
import { bodyChildSteps, parseBodyChild } from "@word/reader/docx-reader";
import { createReaderContext } from "@word/reader/reader-context";
import { parseXml } from "@xml/dom";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { Document, Stream } from "../index";

// The budget is measured with `performance.now`, so on a fast enough machine a whole read could fit inside it and
// never yield — a test that passes or fails by the speed of the machine. A clock that advances a millisecond per
// reading makes the budget run out by the amount of work done instead: what is checked is that the reader asks
// between its pieces.
beforeEach(() => {
  let clock = 0;
  vi.spyOn(performance, "now").mockImplementation(() => (clock += 1));
});
afterEach(() => {
  vi.restoreAllMocks();
});

const PARAGRAPHS = 20_000;

function paragraphs(count: number): ReturnType<typeof Document.create> {
  const doc = Document.create();
  for (let i = 0; i < count; i++) {
    Document.addParagraph(doc, `Paragraph ${i} with a reasonable amount of body text in it.`);
  }
  return doc;
}

/** `doc` packed at `level`, its style part given `extraStyles` more styles first. */
async function packed(
  doc: ReturnType<typeof Document.create>,
  level: number,
  extraStyles = 0
): Promise<Uint8Array> {
  const parts = await Archive.unpackZip(await Document.toBuffer(doc));
  const entries = new Map<string, Uint8Array>();
  for (const [name, part] of parts) {
    entries.set(name, part.data);
  }
  if (extraStyles > 0) {
    let styles = "";
    for (let i = 0; i < extraStyles; i++) {
      styles +=
        `<w:style w:type="paragraph" w:customStyle="1" w:styleId="S${i}"><w:name w:val="Style ${i}"/>` +
        `<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="${i % 240}" w:after="120"/></w:pPr>` +
        `<w:rPr><w:b/><w:sz w:val="${20 + (i % 10)}"/></w:rPr></w:style>`;
    }
    const xml = new TextDecoder().decode(entries.get("word/styles.xml"));
    entries.set(
      "word/styles.xml",
      new TextEncoder().encode(xml.replace("</w:styles>", `${styles}</w:styles>`))
    );
  }
  return Archive.packZip(entries, { level });
}

/** Read every element, recording how many turns the event loop had taken when the first and the last arrived. */
async function readCountingTurns(
  bytes: Uint8Array
): Promise<{ items: number; turnsAtFirst: number; turnsAtLast: number }> {
  // A callback that reschedules itself runs once per turn of the event loop.
  let turns = 0;
  let running = true;
  const tick = (): void => {
    if (running) {
      turns++;
      scheduleMacrotask(tick);
    }
  };
  scheduleMacrotask(tick);
  let items = 0;
  let turnsAtFirst: number | undefined;
  let turnsAtLast = 0;
  try {
    for await (const _item of new Stream.DocumentReader(bytes)) {
      items++;
      turnsAtFirst ??= turns;
      turnsAtLast = turns;
    }
  } finally {
    running = false;
  }
  return { items, turnsAtFirst: turnsAtFirst ?? 0, turnsAtLast };
}

describe("Stream.DocumentReader and the event loop", () => {
  it.each([
    ["a deflated body", 6],
    ["a stored body", 0]
  ])("lets the event loop run while it reads %s", async (_name, level) => {
    const result = await readCountingTurns(await packed(paragraphs(PARAGRAPHS), level));
    expect(result.items).toBe(PARAGRAPHS);
    expect(result.turnsAtLast).toBeGreaterThan(result.turnsAtFirst);
  });

  // Both are read whole before the first element comes out, so the turns are counted up to it.
  it("lets the event loop run while it reads a large style part", async () => {
    const result = await readCountingTurns(await packed(paragraphs(10), 0, 5_000));
    expect(result.items).toBe(10);
    expect(result.turnsAtFirst).toBeGreaterThan(0);
  });

  it("lets the event loop run while it reads a Flat OPC document", async () => {
    const flat = new TextEncoder().encode(await Document.toFlatOpc(paragraphs(PARAGRAPHS)));
    const result = await readCountingTurns(flat);
    expect(result.items).toBe(PARAGRAPHS);
    expect(result.turnsAtFirst).toBeGreaterThan(0);
  });

  // One body element however large — the unit the reader hands out, not the unit of work. Read through the reader's
  // own step runner, so what is counted is the element's steps and not the inflating or parsing around it; and read
  // the way `Document.read` reads it too, which must give the same content.
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const paragraph = (i: number): string => `<w:p><w:r><w:t>Paragraph ${i}</w:t></w:r></w:p>`;
  const many = (count: number, item: (i: number) => string): string =>
    Array.from({ length: count }, (_, i) => item(i)).join("");
  it.each([
    [
      "a table of many rows",
      `<w:tbl ${W}>${many(2_000, i => `<w:tr><w:tc>${paragraph(i)}</w:tc></w:tr>`)}</w:tbl>`
    ],
    [
      "a table cell of many paragraphs",
      `<w:tbl ${W}><w:tr><w:tc>${many(2_000, paragraph)}</w:tc></w:tr></w:tbl>`
    ],
    [
      "a paragraph of many runs",
      `<w:p ${W}>${many(2_000, i => `<w:r><w:t>run ${i} </w:t></w:r>`)}</w:p>`
    ],
    [
      "a hyperlink of many runs",
      `<w:p ${W}><w:hyperlink w:anchor="a">${many(2_000, i => `<w:r><w:t>run ${i} </w:t></w:r>`)}</w:hyperlink></w:p>`
    ],
    [
      "a hyperlink of many fields, one left open",
      `<w:p ${W}><w:hyperlink w:anchor="a">${many(
        600,
        i =>
          `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGEREF _T${i} </w:instrText></w:r>` +
          `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${i}</w:t></w:r>` +
          `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
      )}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGEREF x </w:instrText></w:r>` +
        `<w:r><w:t>never ended</w:t></w:r></w:hyperlink><w:r><w:t>after</w:t></w:r></w:p>`
    ],
    [
      "a field across many runs around a hyperlink",
      `<w:p ${W}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF x </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${many(1_000, i => `<w:r><w:t>${i}</w:t></w:r>`)}` +
        `<w:hyperlink w:anchor="a">${many(1_000, i => `<w:r><w:t>l${i}</w:t></w:r>`)}</w:hyperlink>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
    ],
    ...(["ins", "del", "moveFrom", "moveTo"] as const).map(kind => [
      `a tracked ${kind} of many runs and markers`,
      `<w:p ${W}><w:${kind} w:id="1" w:author="a">${many(
        1_000,
        i =>
          `<w:bookmarkStart w:id="${i}" w:name="b${i}"/><w:r><w:${kind === "del" ? "delText" : "t"}>run ${i} ` +
          `</w:${kind === "del" ? "delText" : "t"}></w:r><w:bookmarkEnd w:id="${i}"/>`
      )}</w:${kind}></w:p>`
    ]),
    [
      "nested wrappers of many runs",
      `<w:p ${W}><w:smartTag w:element="e"><w:smartTagPr><w:attr w:name="n" w:val="v"/></w:smartTagPr>` +
        `<w:dir w:val="rtl">${many(800, i => `<w:r><w:t>r${i}</w:t></w:r>`)}` +
        `<w:hyperlink w:anchor="a"><w:ins w:id="2" w:author="a">${many(800, i => `<w:r><w:t>i${i}</w:t></w:r>`)}` +
        `</w:ins></w:hyperlink></w:dir></w:smartTag><w:bdo w:val="ltr">${many(400, i => `<w:r><w:t>o${i}</w:t></w:r>`)}</w:bdo></w:p>`
    ],
    [
      "a simple field of many runs",
      `<w:p ${W}><w:fldSimple w:instr=" PAGE ">${many(2_000, i => `<w:r><w:t>${i}</w:t></w:r>`)}</w:fldSimple></w:p>`
    ],
    [
      "one run of many lines",
      `<w:p ${W}><w:r><w:rPr><w:b/></w:rPr>${many(10_000, i => `<w:t>line ${i}</w:t><w:br/>`)}</w:r></w:p>`
    ],
    [
      "an inline content control of many runs",
      `<w:p ${W}><w:sdt><w:sdtPr><w:tag w:val="t"/></w:sdtPr><w:sdtContent>${many(5_000, i => `<w:r><w:t>run ${i}</w:t></w:r>`)}</w:sdtContent></w:sdt></w:p>`
    ],
    [
      "a paragraph of many fields read in pieces",
      `<w:p ${W}>${many(
        1_000,
        i =>
          `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>` +
          `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:tab/><w:t>${i}</w:t></w:r>` +
          `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
      )}</w:p>`
    ],
    [
      "a text box of many paragraphs",
      `<w:pict ${W} xmlns:v="urn:schemas-microsoft-com:vml"><v:shape style="width:100pt"><v:textbox>` +
        `<w:txbxContent>${many(2_000, paragraph)}</w:txbxContent></v:textbox></v:shape></w:pict>`
    ],
    [
      "a drawing shape of many paragraphs",
      `<w:p ${W} xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
        `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
        `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:r><w:drawing>` +
        `<wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
        `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
        `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
        `<wp:extent cx="100" cy="100"/><wp:wrapNone/><wp:docPr id="1" name="s"/><a:graphic><a:graphicData ` +
        `uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:txbx>` +
        `<w:txbxContent>${many(2_000, paragraph)}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
        `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
    ],
    [
      "a content control of many paragraphs",
      `<w:sdt ${W}><w:sdtPr><w:tag w:val="t"/></w:sdtPr><w:sdtContent>${many(2_000, paragraph)}</w:sdtContent></w:sdt>`
    ],
    // Runs longer than a step of their own, in each place that reads a run.
    [
      "one long run in a field's result and one in its instruction",
      `<w:p ${W}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r>${many(30_000, i => `<w:instrText>I${i} </w:instrText>`)}</w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r>${many(30_000, i => `<w:t>r${i}</w:t><w:br/>`)}</w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r>${many(30_000, i => `<w:t>after ${i}</w:t>`)}</w:r></w:p>`
    ],
    [
      "one long run holding comment marks",
      `<w:p ${W}><w:r><w:rPr><w:i/></w:rPr>${many(30_000, i => `<w:t>a${i}</w:t>${i % 9_000 === 0 ? `<w:commentReference w:id="${i}"/>` : ""}`)}</w:r></w:p>`
    ],
    ...(["ins", "del"] as const).map(kind => [
      `one long run in a tracked ${kind}`,
      `<w:p ${W}><w:${kind} w:id="1" w:author="a"><w:r>${many(30_000, i => `<w:${kind === "del" ? "delText" : "t"}>x${i}</w:${kind === "del" ? "delText" : "t"}><w:tab/>`)}` +
        `<w:commentReference w:id="3"/></w:r></w:${kind}></w:p>`
    ]),
    [
      "a ruby of many runs",
      `<w:p ${W}><w:r><w:ruby><w:rubyPr><w:hps w:val="10"/></w:rubyPr><w:rt>${many(2_000, i => `<w:r><w:t>t${i}</w:t></w:r>`)}</w:rt>` +
        `<w:rubyBase>${many(2_000, i => `<w:r><w:t>b${i}</w:t></w:r>`)}</w:rubyBase></w:ruby><w:t>after</w:t></w:r></w:p>`
    ],
    [
      "a content control of many runs",
      `<w:sdt ${W}><w:sdtContent>${many(2_000, i => `<w:r><w:t>r${i}</w:t></w:r>`)}${paragraph(0)}${many(2_000, i => `<w:r><w:t>s${i}</w:t></w:r>`)}</w:sdtContent></w:sdt>`
    ],
    [
      "a table cell's content control of many runs",
      `<w:tbl ${W}><w:tr><w:tc><w:sdt><w:sdtPr><w:tag w:val="c"/></w:sdtPr><w:sdtContent>${many(2_000, i => `<w:r><w:t>r${i}</w:t></w:r>`)}` +
        `${paragraph(0)}${many(2_000, i => `<w:r><w:t>s${i}</w:t></w:r>`)}</w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>`
    ],
    [
      "a table of contents of many entries",
      `<w:sdt ${W}><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>` +
        `<w:sdtContent>${many(2_000, paragraph)}</w:sdtContent></w:sdt>`
    ]
  ])("reads %s in steps, to what a whole read gives", async (_name, xml) => {
    const element = parseXml(xml).root;
    let asked = 0;
    const paced = await runStepsPaced(
      bodyChildSteps(element, { ...createReaderContext(), paced: true }),
      () => {
        asked++;
        return undefined;
      }
    );
    expect(asked).toBeGreaterThanOrEqual(7);
    expect(paced).toEqual(parseBodyChild(parseXml(xml).root, createReaderContext()));
  });
});
