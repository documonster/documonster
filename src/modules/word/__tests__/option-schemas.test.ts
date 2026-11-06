import { describe, expect, it } from "vitest";

import { WordHtml } from "../html";
import { Convert, Document, DocxArgumentError, Font, Query, Transform, Validation } from "../index";
import { WordMarkdown } from "../markdown";

function refusal(call: () => unknown): DocxArgumentError {
  try {
    call();
  } catch (error) {
    expect(error).toBeInstanceOf(DocxArgumentError);
    return error as DocxArgumentError;
  }
  throw new Error("expected a DocxArgumentError");
}

/**
 * Options that were accepted unchecked — a value of the wrong type was read as truthy or falsy,
 * or failed inside the renderer with a `TypeError` — and are now refused by the option's schema.
 */
describe("options checked by their schema", () => {
  const doc = () => Document.create();
  const cases: [string, string, string, () => unknown][] = [
    [
      "Convert.toSemantic",
      "options.includeFootnotes",
      "a boolean",
      () => Convert.toSemantic(doc(), { includeFootnotes: "yes" as never })
    ],
    [
      "Transform.map",
      "options.includeHeaders",
      "a boolean",
      () => Transform.map(doc(), {}, { includeHeaders: 1 as never })
    ],
    [
      "Query.each",
      "options.includeComments",
      "a boolean",
      () => Query.each(doc(), {}, { includeComments: null as never })
    ],
    [
      "Query.listTables",
      "options.topLevelOnly",
      "a boolean",
      () => Query.listTables(doc(), { topLevelOnly: "no" as never })
    ],
    [
      "WordHtml.fromDocx",
      "options.title",
      "a string",
      () => WordHtml.fromDocx(doc(), { title: 42 as never })
    ],
    [
      "WordHtml.fromDocx",
      "options.classPrefix",
      "a string",
      () => WordHtml.fromDocx(doc(), { classPrefix: {} as never })
    ],
    [
      "WordHtml.fromDocx",
      "options.imageMode",
      'one of "dataUrl", "filename", "none"',
      () => WordHtml.fromDocx(doc(), { imageMode: "inline" as never })
    ],
    [
      "WordHtml.fromDocx",
      "options.styleMap",
      "a Map or an object",
      () => WordHtml.fromDocx(doc(), { styleMap: "p => h1" as never })
    ],
    [
      "WordHtml.fromDocx",
      "options.chartRenderer",
      "a function",
      () => WordHtml.fromDocx(doc(), { chartRenderer: "svg" as never })
    ],
    [
      "WordMarkdown.fromDocx",
      "options.headingStyle",
      'one of "atx", "setext"',
      () => WordMarkdown.fromDocx(doc(), { headingStyle: "hash" as never })
    ],
    [
      "Validation.validate",
      "options.maxErrors",
      "a number ≥ 0 (Infinity for no limit)",
      () => Validation.validate(doc(), { maxErrors: -1 })
    ],
    [
      "Validation.validate",
      "options.maxSeverity",
      'one of "error", "warning", "info"',
      () => Validation.validate(doc(), { maxSeverity: "fatal" as never })
    ],
    [
      "Validation.validate",
      "options.compatibilityMode",
      'one of "word2007", "word2010", "word2013", "word2016", "word2019"',
      () => Validation.validate(doc(), { compatibilityMode: "word2003" as never })
    ],
    [
      "Validation.validate",
      "options.strict",
      "a boolean",
      () => Validation.validate(doc(), { strict: "true" as never })
    ],
    [
      "Font.createHyphenator",
      "options.hyphenChar",
      "a string",
      () => Font.createHyphenator({ language: "en", patterns: [] }, { hyphenChar: 45 as never })
    ]
  ];

  it.each(cases)("%s refuses %s", (where, field, expected, call) => {
    const err = refusal(call);
    expect(err.field).toBe(field);
    expect(err.message).toBe(
      `${where}: ${field} must be ${expected}, got ${err.message.split(", got ")[1]}`
    );
  });

  it("names an entry of options.classStyles as the other maps do", () => {
    const err = refusal(() =>
      WordHtml.toDocx("<p>x</p>", { classStyles: { "my-class": 1 as never } })
    );
    expect(err.field).toBe('options.classStyles["my-class"]');
  });

  it("ignores modTime in Document.toFlatOpc, which writes no timestamp", async () => {
    const xml = await Document.toFlatOpc(doc(), { modTime: "now" } as never);
    expect(typeof xml).toBe("string");
  });
});
