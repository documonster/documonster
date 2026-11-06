import { describe, expect, it } from "vitest";

import { Build, DocxArgumentError } from "../index";

const cached = (run: ReturnType<typeof Build.field>) =>
  (run.content[0] as { cachedValue?: string }).cachedValue;

describe("every field builder takes its cached value as { cachedValue }", () => {
  it.each([
    ["field", () => Build.field(" PAGE ", { cachedValue: "x" }), () => Build.field(" PAGE ", "x")],
    [
      "pageNumberField",
      () => Build.pageNumberField({ cachedValue: "x" }),
      () => Build.pageNumberField("x")
    ],
    [
      "totalPagesField",
      () => Build.totalPagesField({ cachedValue: "x" }),
      () => Build.totalPagesField("x")
    ],
    [
      "sectionPagesField",
      () => Build.sectionPagesField({ cachedValue: "x" }),
      () => Build.sectionPagesField("x")
    ],
    ["sectionField", () => Build.sectionField({ cachedValue: "x" }), () => Build.sectionField("x")],
    ["authorField", () => Build.authorField({ cachedValue: "x" }), () => Build.authorField("x")],
    ["titleField", () => Build.titleField({ cachedValue: "x" }), () => Build.titleField("x")],
    ["subjectField", () => Build.subjectField({ cachedValue: "x" }), () => Build.subjectField("x")],
    [
      "keywordsField",
      () => Build.keywordsField({ cachedValue: "x" }),
      () => Build.keywordsField("x")
    ],
    [
      "fileSizeField",
      () => Build.fileSizeField({ cachedValue: "x" }),
      () => Build.fileSizeField("x")
    ],
    [
      "quoteField",
      () => Build.quoteField("q", { cachedValue: "x" }),
      () => Build.quoteField("q", "x")
    ],
    [
      "ifField",
      () => Build.ifField("1 = 1", "a", "b", { cachedValue: "x" }),
      () => Build.ifField("1 = 1", "a", "b", "x")
    ],
    [
      "includePictureField",
      () => Build.includePictureField("p.png", { cachedValue: "x" }),
      () => Build.includePictureField("p.png", "x")
    ]
  ])("%s, as the deprecated positional form did", (_name, named, positional) => {
    expect(named()).toEqual(positional());
    expect(cached(named())).toBe("x");
  });

  it("keeps the defaults, and checks the value", () => {
    expect(cached(Build.pageNumberField())).toBe("1");
    expect(cached(Build.pageNumberField({}))).toBe("1");
    expect(cached(Build.quoteField("q", {}))).toBe("q");
    expect(() => Build.authorField({ cachedValue: 3 as never })).toThrow(DocxArgumentError);
  });
});
