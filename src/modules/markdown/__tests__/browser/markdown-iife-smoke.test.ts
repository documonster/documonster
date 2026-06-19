import { loadIife } from "@test/browser/load-iife";
import { describe, it, expect, beforeAll } from "vitest";

/**
 * Smoke test for the shipped `documonster.markdown.iife.min.js` bundle: asserts
 * the `Documonster.Markdown` surface and a parse → format round-trip of a GFM
 * table run in a browser.
 */
describe("Documonster.Markdown IIFE bundle", () => {
  let Markdown: any;

  beforeAll(async () => {
    ({ Markdown } = await loadIife<{ Markdown: any }>("markdown", "Markdown"));
  }, 60000);

  it("exposes the Markdown namespace with parse/format", () => {
    expect(typeof Markdown.parse).toBe("function");
    expect(typeof Markdown.format).toBe("function");
  });

  it("parses a GFM table and formats it back (round-trip)", () => {
    const src = "| a | b |\n| --- | --- |\n| 1 | 2 |";
    const table = Markdown.parse(src);
    expect(table.headers).toEqual(["a", "b"]);
    expect(table.rows).toEqual([["1", "2"]]);

    // `format(headers, rows)` must produce text that re-parses equivalently.
    const out = Markdown.format(table.headers, table.rows);
    expect(typeof out).toBe("string");
    const reparsed = Markdown.parse(out);
    expect(reparsed.headers).toEqual(table.headers);
    expect(reparsed.rows).toEqual(table.rows);
  });
});
