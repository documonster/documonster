/**
 * Word → PDF sets a line break or a tab inside a run's text as one — as Word
 * sets a tab kept in `<w:t>` and the DOCX writer writes a line break — not as a
 * character, which no face has a glyph for and which used to come out as `a�b`
 * on one line.
 */
import { Build, Document } from "@word/index";
import { describe, expect, it } from "vitest";

import { Pdf } from "../index";

describe("Pdf.fromDocx: control characters in run text", () => {
  it("breaks the line at a newline and tabs at a tab", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, Build.paragraph([Build.text("line1\nline2\tTAB\r\nline3")]));
    const lines = (await Pdf.extract(await Pdf.fromDocx(doc))).pages[0]!.textLines;
    expect(lines.map(l => l.text)).toEqual(["line1", "line2 TAB", "line3"]);
    expect(lines[0]!.y).toBeGreaterThan(lines[1]!.y);
    expect(lines[1]!.y).toBeGreaterThan(lines[2]!.y);
    const tab = (await Pdf.extract(await Pdf.fromDocx(doc))).pages[0]!.textFragments.find(
      f => f.text === "TAB"
    );
    expect(tab!.x).toBeGreaterThanOrEqual(72 + 36);
  });
});
