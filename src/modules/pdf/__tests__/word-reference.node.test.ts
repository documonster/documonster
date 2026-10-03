/**
 * Word → PDF against Microsoft Word's own output.
 *
 * Each `.docx` in `data/word-reference/` was exported to PDF by Microsoft Word
 * (16.x, macOS); `expected.json` records what that PDF shows per page: its
 * characters (whitespace removed, sorted, so line breaking and drawing order
 * do not matter), its paper size and, for the toggle-property fixtures, which
 * `K…` probe words are set in a bold face. The documents isolate behaviour this
 * library has to reproduce exactly — section breaks and blank pages, displayed
 * page numbers, even/odd footers, list numbering across shared definitions,
 * merged documents, literal `w:pgSz`, and toggle properties across style
 * levels — so a change that drifts from Word fails here, without Word in CI.
 *
 * To add a case: build the .docx, export it from Word, and append to
 * expected.json with the same normalisation (see the fields above).
 */

import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";

import { Pdf } from "@pdf/index";
import { Io } from "@word/index";
import { describe, expect, it } from "vitest";

interface ExpectedPage {
  readonly text: string;
  readonly width: number;
  readonly height: number;
  readonly bold?: Readonly<Record<string, boolean>>;
}

const DIR = join(__dirname, "data", "word-reference");
const expected = JSON.parse(readFileSync(join(DIR, "expected.json"), "utf8")) as Record<
  string,
  ExpectedPage[]
>;

describe("Word → PDF matches Microsoft Word", () => {
  it("has a fixture for every expectation and vice versa", () => {
    const docs = readdirSync(DIR)
      .filter(f => f.endsWith(".docx"))
      .map(f => f.slice(0, -5))
      .sort();
    expect(docs).toEqual(Object.keys(expected).sort());
  });

  for (const [name, pages] of Object.entries(expected)) {
    it(name, async () => {
      const doc = await Io.read(readFileSync(join(DIR, `${name}.docx`)));
      const read = await Pdf.read(await Pdf.fromDocx(doc));
      const actual = read.pages.map(page => {
        const bold: Record<string, boolean> = {};
        for (const fragment of page.textFragments) {
          const word = fragment.text.trim();
          if (/^K\w+$/.test(word)) {
            bold[word] = /bold/i.test(fragment.fontName);
          }
        }
        return {
          text: [...page.text.replace(/\s+/g, "")].sort().join(""),
          width: Math.round(page.width),
          height: Math.round(page.height),
          ...(Object.keys(bold).length > 0 ? { bold } : {})
        };
      });
      expect(actual).toEqual(pages);
    });
  }
});
