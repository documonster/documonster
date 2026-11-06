/**
 * Every module describes a refused value the same way in its argument errors —
 * through `@utils/describe-value` — so "must be <expected>, got <description>"
 * never contradicts itself: a `Map` is "a Map", not "object" or "an object";
 * a `Date` holding no time is "an invalid Date"; a valid `Date` passed where it
 * does not belong is "a Date", not "an invalid Date".
 */

import { Archive } from "@archive/index";
import { Csv } from "@csv/index";
import { Draw } from "@draw/index";
import { Workbook } from "@excel/index";
import { Formula } from "@formula/index";
import { Markdown } from "@markdown/index";
import { Mermaid } from "@mermaid/index";
import { Pdf } from "@pdf/index";
import { Xml } from "@xml/index";
import { describe, expect, it } from "vitest";

async function message(call: () => unknown): Promise<string> {
  try {
    await call();
  } catch (e) {
    return (e as Error).message;
  }
  return "(no error)";
}

const MAP = new Map([["k", 1]]);

describe("argument errors describe a refused value the same way in every module", () => {
  const cases: [string, () => unknown, string][] = [
    ["Excel", () => Workbook.protect(Workbook.create(), { password: MAP as never }), "got a Map"],
    ["Excel stream", () => Workbook.writeStream(Workbook.create(), MAP as never), "got a Map"],
    ["PDF", () => new Pdf.Builder().addPage({ width: MAP as never }), "got a Map"],
    [
      "Archive modTime",
      () => Archive.packZip(new Map([["a.txt", "x"]]), { modTime: MAP as never }),
      "got a Map"
    ],
    [
      "Archive invalid Date",
      () => Archive.packZip(new Map([["a.txt", "x"]]), { modTime: new Date(Number.NaN) }),
      "got an invalid Date"
    ],
    ["Archive TAR source", () => Archive.parseTar(MAP as never), "got a Map"],
    ["CSV input", () => Csv.parse(MAP as never), "got a Map"],
    ["CSV option", () => Csv.parse("a", { delimiter: MAP as never }), "got a Map"],
    ["Markdown", () => Markdown.parse(MAP as never), "got a Map"],
    ["XML", () => Xml.parse(MAP as never), "got a Map"],
    [
      "Mermaid option",
      () => Mermaid.toSvg(Mermaid.parse("graph TD\nA-->B"), { padding: MAP as never }),
      "got a Map"
    ],
    [
      "Draw",
      () => Draw.toSvg({ width: 10, height: 10, children: [] } as never, { width: MAP as never }),
      "got a Map"
    ],
    ["Formula tokenize", () => Formula.tokenize(MAP as never), "got a Map"],
    ["Formula parse", () => Formula.parse(MAP as never), "got a Map"]
  ];
  it.each(cases)("%s", async (_label, call, expected) => {
    const text = await message(call);
    expect(text).toContain(expected);
    expect(text).not.toMatch(/got (an )?object\b/);
  });
});
