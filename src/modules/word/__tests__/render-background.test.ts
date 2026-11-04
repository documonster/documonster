/**
 * Word's page renderers take `background` as Draw, Mermaid and the chart
 * renderers take it — a CSS colour string or an `{ r, g, b, a }` colour with
 * channels from 0 to 1 — checked the same way. The deprecated
 * `backgroundColor` still works; `background` wins over it.
 */

import { describe, it, expect } from "vitest";

import { Document, DocxArgumentError, Layout } from "../index";

function doc() {
  const d = Document.create();
  Document.addParagraph(d, "Hello");
  return d;
}

/** The page's background rectangle, or `undefined` when none is drawn. */
function backgroundRect(svg: string): string | undefined {
  return /<rect width="100%" height="100%"[^>]*\/>/.exec(svg)?.[0];
}

describe("Layout render options: background", () => {
  it("defaults to white", () => {
    expect(backgroundRect(Layout.pageToSvg(doc(), 1))).toContain('fill="white"');
  });

  it("takes a CSS colour", () => {
    expect(backgroundRect(Layout.pageToSvg(doc(), 1, { background: "#EEEEEE" }))).toContain(
      'fill="#EEEEEE"'
    );
    expect(backgroundRect(Layout.toSvg(doc(), { background: "rgb(1, 2, 3)" })[0]!)).toContain(
      'fill="rgb(1, 2, 3)"'
    );
  });

  it("takes an { r, g, b, a } colour, its alpha as fill-opacity", () => {
    const rect = backgroundRect(
      Layout.pageToSvg(doc(), 1, { background: { r: 1, g: 0.5, b: 0, a: 0.25 } })
    );
    expect(rect).toContain('fill="#ff8000"');
    expect(rect).toContain('fill-opacity="0.25"');
    expect(
      backgroundRect(Layout.pageToSvg(doc(), 1, { background: { r: 0, g: 0, b: 0, a: 1 } }))
    ).not.toContain("fill-opacity");
  });

  it("draws no background for a fully transparent colour or none", () => {
    expect(
      backgroundRect(Layout.pageToSvg(doc(), 1, { background: { r: 1, g: 1, b: 1, a: 0 } }))
    ).toBeUndefined();
    expect(backgroundRect(Layout.pageToSvg(doc(), 1, { background: "none" }))).toBeUndefined();
    expect(
      backgroundRect(Layout.pageToSvg(doc(), 1, { background: "transparent" }))
    ).toBeUndefined();
  });

  it("wins over the deprecated backgroundColor, which still works alone", () => {
    expect(
      backgroundRect(
        Layout.pageToSvg(doc(), 1, { background: "#000000", backgroundColor: "lightblue" })
      )
    ).toContain('fill="#000000"');
    expect(backgroundRect(Layout.pageToSvg(doc(), 1, { backgroundColor: "lightblue" }))).toContain(
      'fill="lightblue"'
    );
  });

  it("is checked as Draw checks it, naming the field", () => {
    const cases: [unknown, string][] = [
      ["not a colour", "options.background"],
      [42, "options.background"],
      [{ r: 2, g: 0, b: 0, a: 1 }, "options.background.r"],
      [{ r: 0, g: 0, b: 0 }, "options.background.a"]
    ];
    for (const [background, field] of cases) {
      let error: unknown;
      try {
        Layout.pageToSvg(doc(), 1, { background: background as never });
      } catch (e) {
        error = e;
      }
      expect(error, field).toBeInstanceOf(DocxArgumentError);
      expect((error as DocxArgumentError).field, field).toBe(field);
      expect((error as Error).message, field).toMatch(/^Layout\.pageToSvg: /);
    }
  });
});
