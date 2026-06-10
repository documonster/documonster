/**
 * Sub / superscript: ECMA-376 `<w:vertAlign w:val="superscript"/>`
 * (and `subscript`) shrinks the run's effective font size and
 * shifts its baseline. Word renders these at roughly 65 % of the
 * surrounding text's size; the layout engine bakes the size scale
 * into `PositionedRun.fontSize` and surfaces the alignment via
 * `PositionedRun.verticalAlign` so renderers can apply the y-shift.
 */

import { describe, it, expect } from "vitest";

import { layoutDocumentFull } from "../layout/layout-full";
import { renderPageFromLayout } from "../layout/render-page";
import type { DocxDocument, Paragraph, Run } from "../types";

function plainRun(text: string, props?: Run["properties"]): Run {
  return { properties: props, content: [{ type: "text", text }] };
}

const baseDoc = (overrides: Partial<DocxDocument> = {}): DocxDocument => ({
  body: [],
  styles: [],
  abstractNumberings: [],
  numberingInstances: [],
  headers: new Map(),
  footers: new Map(),
  footnotes: [],
  endnotes: [],
  comments: [],
  images: [],
  fonts: [],
  embeddedFonts: [],
  customXmlParts: [],
  customProperties: [],
  opaqueParts: [],
  ...overrides
});

describe("sub / superscript layout", () => {
  it("shrinks the font size of a superscript run to ~65% of the source size", () => {
    // 12pt body (size = 24 half-points); superscript should land at
    // ~7.8pt. Allow a small tolerance.
    const para: Paragraph = {
      type: "paragraph",
      children: [
        plainRun("body ", { size: 24 }),
        plainRun("²", { size: 24, vertAlign: "superscript" })
      ]
    };
    const layout = layoutDocumentFull(baseDoc({ body: [para] }));
    const paraItem = layout.pages[0].content[0];
    expect(paraItem.type).toBe("paragraph");
    if (paraItem.type !== "paragraph") {
      return;
    }
    const items = paraItem.lines.flatMap(l => l.runs);
    const sup = items.find(i => i.type !== "image" && i.text === "²");
    const body = items.find(i => i.type !== "image" && i.text.startsWith("body"));
    expect(sup).toBeDefined();
    expect(body).toBeDefined();
    if (!sup || sup.type === "image" || !body || body.type === "image") {
      return;
    }
    expect(sup.fontSize).toBeLessThan(body.fontSize);
    expect(sup.fontSize).toBeCloseTo(12 * 0.65, 0);
    expect(sup.verticalAlign).toBe("superscript");
  });

  it("shrinks subscript runs identically and tags them subscript", () => {
    const para: Paragraph = {
      type: "paragraph",
      children: [
        plainRun("H", { size: 24 }),
        plainRun("2", { size: 24, vertAlign: "subscript" }),
        plainRun("O", { size: 24 })
      ]
    };
    const layout = layoutDocumentFull(baseDoc({ body: [para] }));
    const paraItem = layout.pages[0].content[0];
    if (paraItem.type !== "paragraph") {
      return;
    }
    const items = paraItem.lines.flatMap(l => l.runs);
    const sub = items.find(i => i.type !== "image" && i.text === "2");
    expect(sub).toBeDefined();
    if (!sub || sub.type === "image") {
      return;
    }
    expect(sub.fontSize).toBeCloseTo(12 * 0.65, 0);
    expect(sub.verticalAlign).toBe("subscript");
  });

  it("leaves regular runs at their full size with no verticalAlign", () => {
    const para: Paragraph = {
      type: "paragraph",
      children: [plainRun("plain", { size: 24 })]
    };
    const layout = layoutDocumentFull(baseDoc({ body: [para] }));
    const paraItem = layout.pages[0].content[0];
    if (paraItem.type !== "paragraph") {
      return;
    }
    const items = paraItem.lines.flatMap(l => l.runs);
    const run = items[0];
    if (!run || run.type === "image") {
      return;
    }
    expect(run.fontSize).toBe(12);
    expect(run.verticalAlign).toBeUndefined();
  });

  it("renders SVG with a baseline shifted upward for superscript", () => {
    const para: Paragraph = {
      type: "paragraph",
      children: [plainRun("x", { size: 24 }), plainRun("2", { size: 24, vertAlign: "superscript" })]
    };
    const layout = layoutDocumentFull(baseDoc({ body: [para] }));
    const svg = renderPageFromLayout(layout, 1);
    // Two <text> elements; the superscript one's y is smaller than
    // the base run's y (SVG y grows downward).
    const matches = [...svg.matchAll(/<text[^>]*y="([\d.]+)"[^>]*>([^<]*)<\/text>/g)];
    const xRow = matches.find(m => m[2] === "x");
    const supRow = matches.find(m => m[2] === "2");
    expect(xRow).toBeDefined();
    expect(supRow).toBeDefined();
    if (xRow && supRow) {
      const xY = parseFloat(xRow[1]);
      const supY = parseFloat(supRow[1]);
      // Superscript should sit higher on the page (smaller SVG y).
      expect(supY).toBeLessThan(xY);
    }
  });
});

describe("paragraph style resolution in full layout", () => {
  const headingStyle = {
    type: "paragraph" as const,
    styleId: "Heading1",
    name: "heading 1",
    runProperties: { size: 32, color: "2F5496", font: "Calibri Light" }
  };

  it("resolves a heading run's size, colour and font from the styles table", () => {
    const para: Paragraph = {
      type: "paragraph",
      properties: { style: "Heading1" },
      children: [plainRun("Title")]
    };
    const layout = layoutDocumentFull(baseDoc({ body: [para], styles: [headingStyle] }));
    const paraItem = layout.pages[0].content[0];
    expect(paraItem.type).toBe("paragraph");
    if (paraItem.type !== "paragraph") {
      return;
    }
    const run = paraItem.lines.flatMap(l => l.runs).find(r => r.type !== "image");
    expect(run).toBeDefined();
    if (!run) {
      return;
    }
    // sz=32 half-points → 16pt, NOT the hardcoded 2.0× heading scale (24pt).
    expect(run.fontSize).toBe(16);
    expect(run.color).toBe("2F5496");
    expect(run.font).toBe("Calibri Light");
  });

  it("falls back to the heuristic heading scale without a matching style", () => {
    const heading: Paragraph = {
      type: "paragraph",
      properties: { style: "Heading1" },
      children: [plainRun("Title")]
    };
    const normal: Paragraph = { type: "paragraph", children: [plainRun("Body")] };
    const layout = layoutDocumentFull(baseDoc({ body: [heading, normal] }));
    const runs = layout.pages[0].content
      .filter(c => c.type === "paragraph")
      .map(c => (c.type === "paragraph" ? c.lines.flatMap(l => l.runs)[0] : undefined));
    const headingRun = runs[0];
    const bodyRun = runs[1];
    if (!headingRun || headingRun.type === "image" || !bodyRun || bodyRun.type === "image") {
      throw new Error("expected text runs");
    }
    expect(headingRun.fontSize).toBeGreaterThan(bodyRun.fontSize);
  });
});
