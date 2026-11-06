/**
 * The bridge entries `documonster/word/html`, `/word/markdown` and
 * `/word/excel` export the errors their functions throw and the guard, as the
 * same class objects as `documonster/word` — as `documonster/word/crypto`
 * does — so a caller of one bridge alone can tell its errors apart.
 */

import { describe, it, expect } from "vitest";

import * as ExcelEntry from "../excel";
import * as HtmlEntry from "../html";
import * as Main from "../index";
import * as MarkdownEntry from "../markdown";

const entries = {
  "word/html": HtmlEntry,
  "word/markdown": MarkdownEntry,
  "word/excel": ExcelEntry
};

describe("bridge entries export their errors", () => {
  it("as the same class objects as documonster/word", () => {
    for (const [name, entry] of Object.entries(entries)) {
      const e = entry as Record<string, unknown>;
      expect(e.DocxError, name).toBe(Main.DocxError);
      expect(e.DocxArgumentError, name).toBe(Main.DocxArgumentError);
      expect(e.isDocxError, name).toBe(Main.isDocxError);
    }
  });

  it("and what they throw is one of them", async () => {
    const html = (() => {
      try {
        HtmlEntry.WordHtml.toDocx(42 as never);
      } catch (e) {
        return e;
      }
    })();
    expect(html).toBeInstanceOf(HtmlEntry.DocxArgumentError);
    expect(HtmlEntry.isDocxError(html)).toBe(true);
    const md = await MarkdownEntry.WordMarkdown.toDocx(42 as never).catch((e: unknown) => e);
    expect(md).toBeInstanceOf(MarkdownEntry.DocxArgumentError);
    const excel = (() => {
      try {
        ExcelEntry.WordExcel.toDocx(42 as never);
      } catch (e) {
        return e;
      }
    })();
    expect(excel).toBeInstanceOf(ExcelEntry.DocxArgumentError);
  });
});
