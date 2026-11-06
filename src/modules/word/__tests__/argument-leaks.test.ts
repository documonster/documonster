/**
 * Public functions that let a native `TypeError` escape on a wrong argument —
 * found by calling every namespace member with `[]`, `[null]`, `[{}]`, `[42]`,
 * … — refuse it with `DocxArgumentError`, naming the parameter and the call.
 */

import { describe, it, expect } from "vitest";

import { WordCrypto } from "../crypto";
import { DocxArgumentError } from "../errors";
import { WordExcel } from "../excel";
import { Document, Font, Package, Template } from "../index";

function refusal(run: () => unknown): DocxArgumentError {
  let error: unknown;
  try {
    run();
  } catch (e) {
    error = e;
  }
  expect(error).toBeInstanceOf(DocxArgumentError);
  return error as DocxArgumentError;
}

function expectRefused(run: () => unknown, field: string, where: string): void {
  const error = refusal(run);
  expect(error.field).toBe(field);
  expect(error.message.startsWith(`${where}: `)).toBe(true);
}

async function expectRejected(run: () => Promise<unknown>, field: string, where: string) {
  const error = await run().then(
    () => undefined,
    (e: unknown) => e
  );
  expect(error).toBeInstanceOf(DocxArgumentError);
  expect((error as DocxArgumentError).field).toBe(field);
  expect((error as DocxArgumentError).message.startsWith(`${where}: `)).toBe(true);
}

describe("Font.createHyphenator", () => {
  const where = "Font.createHyphenator";

  it("checks the pattern set's lists, not only that it is an object", () => {
    expectRefused(() => Font.createHyphenator({} as never), "patterns.patterns", where);
    expectRefused(
      () => Font.createHyphenator({ language: "x", patterns: "a1b" } as never),
      "patterns.patterns",
      where
    );
    expectRefused(
      () => Font.createHyphenator({ language: "x", patterns: ["a1b", "c1d", "e1f", 4] } as never),
      "patterns.patterns[3]",
      where
    );
    expectRefused(
      () => Font.createHyphenator({ language: "x", patterns: [], exceptions: "as-so" } as never),
      "patterns.exceptions",
      where
    );
    expectRefused(
      () => Font.createHyphenator({ language: "x", patterns: [], exceptions: [null] } as never),
      "patterns.exceptions[0]",
      where
    );
    expectRefused(
      () => Font.createHyphenator({ language: 7, patterns: [] } as never),
      "patterns.language",
      where
    );
  });

  it("checks minLeft and minRight", () => {
    const patterns = Font.ENGLISH_US_PATTERNS;
    expectRefused(
      () => Font.createHyphenator(patterns, { minLeft: "2" } as never),
      "options.minLeft",
      where
    );
    expectRefused(() => Font.createHyphenator(patterns, { minLeft: -1 }), "options.minLeft", where);
    expectRefused(
      () => Font.createHyphenator(patterns, { minRight: Number.NaN }),
      "options.minRight",
      where
    );
    expectRefused(() => Font.createHyphenator(patterns, 42 as never), "options", where);
  });

  it("still accepts a pattern set without exceptions", () => {
    const h = Font.createHyphenator({ language: "x", patterns: ["a1b"] });
    expect(h("aabb")).toEqual([2]);
  });
});

describe("Font.hyphenateWord / hyphenateText", () => {
  const h = Font.createHyphenator(Font.ENGLISH_US_PATTERNS);

  it("take the hyphen character as { hyphenChar }", () => {
    expect(Font.hyphenateWord("associate", h, { hyphenChar: "-" })).toBe("as-so-ciate");
    expect(Font.hyphenateWord("computer", h, {})).toBe("com\u00ADput\u00ADer");
    expect(Font.hyphenateText("Information technology is everywhere", h, { hyphenChar: "-" })).toBe(
      "In-for-ma-tion tech-nol-o-gy is every-where"
    );
  });

  it("keep the deprecated positional form, as v0.21.1 answered it", () => {
    // Expected values are v0.21.1's output for the same calls.
    expect(Font.hyphenateWord("associate", h, "-")).toBe("as-so-ciate");
    expect(Font.hyphenateWord("hyphenation", h, "|")).toBe("hy|phen|a|tion");
    expect(Font.hyphenateText("Information technology is everywhere", h, "-")).toBe(
      "In-for-ma-tion tech-nol-o-gy is every-where"
    );
    expect(Font.hyphenateWord("computer", h)).toBe("com\u00ADput\u00ADer");
    expect(Font.hyphenateWord("hyphenation", h, "|")).toBe(
      Font.hyphenateWord("hyphenation", h, { hyphenChar: "|" })
    );
  });

  it("refuse a hyphen character that is not a string", () => {
    expectRefused(
      () => Font.hyphenateWord("associate", h, { hyphenChar: 1 } as never),
      "options.hyphenChar",
      "Font.hyphenateWord"
    );
    expectRefused(
      () => Font.hyphenateText("associate", h, { hyphenChar: null } as never),
      "options.hyphenChar",
      "Font.hyphenateText"
    );
    expectRefused(
      () => Font.hyphenateWord("associate", h, 42 as never),
      "hyphenChar",
      "Font.hyphenateWord"
    );
  });
});

describe("Font text shaping", () => {
  it("refuses text that is not a string", () => {
    expectRefused(() => (Font.detectScript as () => unknown)(), "text", "Font.detectScript");
    expectRefused(() => Font.detectDirection(null as never), "text", "Font.detectDirection");
    expectRefused(() => Font.shapeText(new Uint8Array(2) as never), "text", "Font.shapeText");
  });

  it("checks shapeText's options", () => {
    const where = "Font.shapeText";
    expectRefused(() => Font.shapeText("abc", 42 as never), "options", where);
    expectRefused(
      () => Font.shapeText("abc", { direction: "up" } as never),
      "options.direction",
      where
    );
    expectRefused(
      () => Font.shapeText("abc", { arabicJoining: "no" } as never),
      "options.arabicJoining",
      where
    );
    expectRefused(
      () => Font.shapeText("abc", { bidiReorder: 0 } as never),
      "options.bidiReorder",
      where
    );
    expect(Font.shapeText("abc", { direction: "rtl" }).length).toBeGreaterThan(0);
  });
});

describe("Template.fill", () => {
  it("refuses something that is not a document, before reading data", () => {
    expectRefused(() => Template.fill({} as never, {}), "doc", "Template.fill");
    expectRefused(() => Template.fill("x" as never, {}), "doc", "Template.fill");
  });
});

describe("Package.patchTemplate", () => {
  it("refuses something that is not a compiled template", async () => {
    await expectRejected(
      () => Package.patchTemplate(Document.create() as never, []),
      "template",
      "Package.patchTemplate"
    );
    await expectRejected(
      () => Package.patchTemplate(null as never, []),
      "template",
      "Package.patchTemplate"
    );
  });
});

describe("WordExcel.toDocx", () => {
  it("refuses something that is not a workbook in the Word API's own error", () => {
    expectRefused(() => WordExcel.toDocx({} as never), "workbook", "WordExcel.toDocx");
    expectRefused(
      () => WordExcel.toDocx(Document.create() as never),
      "workbook",
      "WordExcel.toDocx"
    );
  });
});

describe("WordCrypto.Agile.parseEncryptionInfo", () => {
  it("refuses XML that is not a string", () => {
    const where = "WordCrypto.Agile.parseEncryptionInfo";
    expectRefused(() => (WordCrypto.Agile.parseEncryptionInfo as () => unknown)(), "xml", where);
    expectRefused(() => WordCrypto.Agile.parseEncryptionInfo(null as never), "xml", where);
  });
});
