/**
 * A complex field inside a wrapper — a hyperlink (each entry of Word's TOC is
 * one, with a PAGEREF inside), a revision (`w:ins`, `w:del`, a move), a smart
 * tag, an inline content control — or crossing a wrapper's boundary keeps
 * every character and every piece of content where it was. A field in a
 * revision was read with the revision's runs alone and lost its characters
 * and instruction; one crossing a boundary was assembled from the half on one
 * side, inventing a field with an empty instruction there.
 */

import { describe, expect, it } from "vitest";

import { Io } from "../index";
import { buildDocx, packageIssues, packageParts, partText } from "./roundtrip-integrity.helpers";

const R = (t: string): string => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const TAB = `<w:r><w:tab/></w:r>`;
const C = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const I = (t: string): string => `<w:r><w:instrText xml:space="preserve">${t}</w:instrText></w:r>`;
const DT = (t: string): string => `<w:r><w:delText xml:space="preserve">${t}</w:delText></w:r>`;
const DI = (t: string): string =>
  `<w:r><w:delInstrText xml:space="preserve">${t}</w:delInstrText></w:r>`;
const REV = 'w:id="9" w:author="A" w:date="2020-01-01T00:00:00Z"';
const F = (instruction = " PAGE ", result = "1"): string =>
  C("begin") + I(instruction) + C("separate") + R(result) + C("end");
const P = (inner: string): string => `<w:p>${inner}</w:p>`;
const LINK = (anchor: string, inner: string): string =>
  `<w:hyperlink w:anchor="${anchor}">${inner}</w:hyperlink>`;

const CASES: Record<string, string> = {
  "a hyperlink holding a whole field (Word's TOC entry)": P(
    LINK("_T1", R("Intro") + TAB + F(" PAGEREF _T1 \\h ", "3"))
  ),
  "a field beginning in a hyperlink and ending after it": P(
    LINK("_T1", C("begin") + I(" PAGEREF _T1 ") + C("separate") + R("3")) + C("end") + R(" tail")
  ),
  "a field beginning before a hyperlink and ending in it": P(
    C("begin") + I(" PAGEREF _T1 ") + C("separate") + LINK("_T1", R("3") + C("end")) + R(" tail")
  ),
  "a field in an insertion": P(R("a ") + `<w:ins ${REV}>${F()}</w:ins>` + R(" b")),
  "a field beginning in an insertion and ending after it": P(
    `<w:ins ${REV}>${C("begin")}${I(" PAGE ")}</w:ins>` + C("separate") + R("1") + C("end")
  ),
  "a field in a deletion": P(
    R("a ") +
      `<w:del ${REV}>${C("begin")}${DI(" PAGE ")}${C("separate")}${DT("1")}${C("end")}</w:del>` +
      R(" b")
  ),
  "a field beginning in a deletion and ending after it": P(
    `<w:del ${REV}>${C("begin")}${DI(" PAGE ")}${C("separate")}${DT("1")}</w:del>` + C("end")
  ),
  "a field in a move": P(
    `<w:moveTo ${REV}>${F()}</w:moveTo>` + `<w:moveFrom ${REV}>${F()}</w:moveFrom>`
  ),
  "a field in a smart tag": P(`<w:smartTag w:uri="u" w:element="e">${F()}</w:smartTag>` + R(" b")),
  "a field in an inline content control": P(
    `<w:sdt><w:sdtPr><w:alias w:val="x"/></w:sdtPr><w:sdtContent>${F()}</w:sdtContent></w:sdt>` +
      R(" b")
  ),
  "a field beginning in an inline content control and ending after it": P(
    `<w:sdt><w:sdtPr/><w:sdtContent>${C("begin")}${I(" PAGE ")}${C("separate")}</w:sdtContent></w:sdt>` +
      R("1") +
      C("end")
  ),
  "a TOC beginning in an insertion, its entries hyperlinks with a PAGEREF each":
    P(
      `<w:ins ${REV}>${C("begin")}${I(" TOC \\o ")}${C("separate")}</w:ins>` +
        LINK("_T1", R("Intro") + TAB + F(" PAGEREF _T1 \\h ", "3"))
    ) +
    P(LINK("_T2", R("Two") + TAB + F(" PAGEREF _T2 \\h ", "4"))) +
    P(C("end")),
  "a field spanning paragraphs, ending in a hyperlink":
    P(C("begin") + I(" TOC ") + C("separate") + R("x")) + P(LINK("a", R("y") + C("end")))
};

const WRAPPERS = "hyperlink|ins|del|smartTag|sdtContent|moveTo|moveFrom|p";

/** The body as tokens: wrappers opening and closing, field characters, instruction text, text, tabs. */
function structure(xml: string): string[] {
  const body = /<w:body>([\s\S]*?)(?:<w:sectPr\b[\s\S]*)?<\/w:body>/.exec(xml)![1]!;
  return [
    ...body.matchAll(
      new RegExp(
        `<w:(${WRAPPERS})\\b[^>]*?>|</w:(${WRAPPERS})>|fldCharType="(\\w+)"|<w:(?:instrText|delInstrText)[^>]*>([^<]*)<|<w:(?:t|delText)\\b[^>]*>([^<]*)<|<w:tab/>`,
        "g"
      )
    )
  ].map(m =>
    m[1]
      ? `<${m[1]}>`
      : m[2]
        ? `</${m[2]}>`
        : (m[3] ??
          (m[4] !== undefined ? `instr:${m[4].trim()}` : m[5] !== undefined ? `t:${m[5]}` : "tab"))
  );
}

describe("a field in or across a wrapper", () => {
  it.each(Object.entries(CASES))("keeps its characters and content: %s", async (_name, body) => {
    const once = await Io.toBuffer(await Io.read(await buildDocx({ body })));
    const twice = await Io.toBuffer(await Io.read(once));
    const parts = await packageParts(once);
    expect(packageIssues(parts)).toEqual([]);
    const xml = partText(parts, "word/document.xml");
    expect(structure(xml)).toEqual(structure(`<w:body>${body}</w:body>`));
    expect(partText(await packageParts(twice), "word/document.xml")).toBe(xml);
    // Deleted text stays deleted text.
    for (const [, deleted] of xml.matchAll(/<w:del\b[^>]*>([\s\S]*?)<\/w:del>/g)) {
      expect(deleted).not.toMatch(/<w:instrText|<w:t[ >]/);
    }
  });
});
