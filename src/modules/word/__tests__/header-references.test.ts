/**
 * One encoding of "this section shows that header": `Document.setHeader` /
 * `setFooter` write an explicit reference into the section, so the layout
 * (and everything rendering through it, such as `Pdf.fromDocx`) shows what
 * the written file shows. A model built by hand without references gets the
 * same implied references in the layout as in the package.
 */

import { describe, it, expect } from "vitest";

import { Build, Document, Io, Layout } from "../index";
import type { DocxDocument } from "../types";

function twoPages(): DocxDocument {
  const d = Document.create();
  Document.addParagraph(d, "body");
  Document.addPageBreak(d);
  Document.addParagraph(d, "body2");
  return d;
}

function shown(d: DocxDocument): [boolean, boolean][] {
  return Layout.toSvg(d).map(svg => [svg.includes("HDRMARK"), svg.includes("FTRMARK")]);
}

describe("header and footer references", () => {
  it.each(["default", "first", "even"] as const)(
    "lays out a %s header set on a built document as the written file shows it",
    async key => {
      const d = twoPages();
      Document.setHeader(d, key, { children: [Build.textParagraph("HDRMARK")] });
      Document.setFooter(d, key, { children: [Build.textParagraph("FTRMARK")] });
      expect(d.sectionProperties?.headers).toEqual([{ type: key, rId: key }]);
      expect(d.sectionProperties?.footers).toEqual([{ type: key, rId: key }]);
      const reread = await Io.read(await Io.toBuffer(d));
      expect(shown(d)).toEqual(shown(reread));
      expect(shown(d).some(([h, f]) => h && f)).toBe(true);
    }
  );

  it("keeps a hand-built model's implied type-keyed headers when one is set", async () => {
    const d = twoPages();
    // Built by hand: a header stored under its type, no reference.
    (d as { headers?: Map<string, unknown> }).headers = new Map([
      ["even", { content: { children: [Build.textParagraph("EVENMARK")] } }]
    ]);
    Document.setSettings(d, { evenAndOddHeaders: true });
    Document.setHeader(d, "default", { children: [Build.textParagraph("HDRMARK")] });
    expect(d.sectionProperties?.headers).toEqual([
      { type: "even", rId: "even" },
      { type: "default", rId: "default" }
    ]);
    const svgs = Layout.toSvg(d);
    expect(svgs[0]).toContain("HDRMARK");
    expect(svgs[1]).toContain("EVENMARK");
  });

  it("lays out a hand-built model without references as the writer wires it", async () => {
    const d = twoPages();
    (d as { headers?: Map<string, unknown> }).headers = new Map([
      ["default", { content: { children: [Build.textParagraph("HDRMARK")] } }]
    ]);
    const reread = await Io.read(await Io.toBuffer(d));
    expect(shown(d)).toEqual(shown(reread));
    expect(shown(d)[0]![0]).toBe(true);
  });

  it("keeps the title page and even/odd setting other setters do not mention", () => {
    const d = twoPages();
    Document.setHeader(d, "first", { children: [Build.textParagraph("HDRMARK")] });
    Document.setHeader(d, "even", { children: [Build.textParagraph("EVENMARK")] });
    Document.setSectionProperties(d, { pageSize: { width: 11906, height: 16838 } });
    Document.setSettings(d, { zoom: 120 });
    expect(d.sectionProperties?.titlePage).toBe(true);
    expect(d.settings?.evenAndOddHeaders).toBe(true);
    expect(Layout.toSvg(d)[0]).toContain("HDRMARK");
  });
});
