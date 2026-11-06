/**
 * The documented way to put converted HTML or Markdown into a document that
 * already has lists of its own: merge the converted document, which
 * renumbers, rather than copy a fragment's fixed ids.
 */

import { describe, it, expect } from "vitest";

import { WordHtml } from "../html";
import { Document, Io, Transform, Validation } from "../index";
import { WordMarkdown } from "../markdown";

function host() {
  const d = Document.create();
  Document.addNumberedList(d, ["one", "two"]);
  return d;
}

describe("splicing converted content into a document", () => {
  it("Transform.merge renumbers what HTML and Markdown bring", async () => {
    const html = WordHtml.toDocx("<ul><li>a</li></ul><ol><li>b</li></ol>");
    const md = await WordMarkdown.toDocx("- a\n\n1. b\n");
    for (const converted of [html, md]) {
      const merged = Transform.merge([host(), converted]);
      const ids = (merged.numberingInstances ?? []).map(n => n.numId);
      expect(new Set(ids).size).toBe(ids.length);
      expect(Validation.validate(merged).issues.filter(i => i.severity === "error")).toEqual([]);
      await expect(Io.toBuffer(merged)).resolves.toBeInstanceOf(Uint8Array);
    }
  });
});
