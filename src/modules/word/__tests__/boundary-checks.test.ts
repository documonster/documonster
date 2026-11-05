/**
 * Values that used to reach the package unchecked — written as a file Word
 * refuses, failing only at `Io.toBuffer` with another module's error, or
 * silently ignored — are refused where they are passed, with
 * `DocxArgumentError` naming the field, and nothing changes.
 */

import { describe, it, expect } from "vitest";

import { WordCrypto } from "../crypto";
import { DocxArgumentError, DocxUnsupportedFeatureError } from "../errors";
import {
  Build,
  Convert,
  Diff,
  Document,
  Font,
  Glossary,
  Io,
  Package,
  Query,
  Security,
  Stream,
  Template,
  Transform
} from "../index";
import type { DocxDocument } from "../types";

const NOT_A_DATE = "not a date";

function snapshot(value: unknown): string {
  return JSON.stringify(value);
}

function doc(): DocxDocument {
  const d = Document.create();
  Document.addParagraph(d, "Hello world");
  return d;
}

function refused(label: string, run: () => unknown, field?: string): void {
  let error: unknown;
  try {
    run();
  } catch (e) {
    error = e;
  }
  expect(error, label).toBeInstanceOf(DocxArgumentError);
  if (field) {
    expect((error as DocxArgumentError).field, label).toBe(field);
  }
}

async function refusedAsync(
  label: string,
  run: () => Promise<unknown>,
  field?: string
): Promise<void> {
  let error: unknown;
  try {
    await run();
  } catch (e) {
    error = e;
  }
  expect(error, label).toBeInstanceOf(DocxArgumentError);
  if (field) {
    expect((error as DocxArgumentError).field, label).toBe(field);
  }
}

describe("dates are xsd:dateTime", () => {
  it("refuses a comment or revision date that is not one", () => {
    refused(
      "addComment",
      () => Document.addComment(doc(), "A", "x", { date: NOT_A_DATE }),
      "options.date"
    );
    refused("addComment month 13", () =>
      Document.addComment(doc(), "A", "x", { date: "2024-13-01T00:00:00Z" })
    );
    for (const build of [
      Build.insertedRun,
      Build.deletedRun,
      Build.movedFromRun,
      Build.movedToRun
    ]) {
      refused(
        build.name,
        () => build(Build.text("x"), { id: 1, author: "A", date: NOT_A_DATE }),
        "revision.date"
      );
    }
    refused(
      "moveFromRangeStart",
      () => Build.moveFromRangeStart(1, "A", { date: NOT_A_DATE }),
      "options.date"
    );
    refused(
      "moveToRangeStart",
      () => Build.moveToRangeStart(1, "A", { date: NOT_A_DATE }),
      "options.date"
    );
  });

  it("accepts the forms Word writes", () => {
    for (const date of [
      "2024-01-02T03:04:05Z",
      "2024-01-02T03:04:05.123Z",
      "2024-01-02T03:04:05+08:00",
      "2024-01-02T03:04:05"
    ]) {
      expect(Document.addComment(doc(), "A", "x", { date })).toBeTypeOf("number");
    }
  });

  it("refuses core property dates that are not valid Dates", () => {
    for (const created of [NOT_A_DATE, new Date(Number.NaN)]) {
      refused(
        String(created),
        () => Document.setCoreProperties(doc(), { created: created as Date }),
        "props.created"
      );
    }
    refused(
      "modified",
      () => Document.setCoreProperties(doc(), { modified: new Date(Number.NaN) }),
      "props.modified"
    );
    refused(
      "title",
      () => Document.setCoreProperties(doc(), { title: 42 as never }),
      "props.title"
    );
  });
});

describe("charts", () => {
  const series = [{ name: "s", categories: ["a", "b"], values: [1, 2] }];

  it("Build.chart refuses values that cannot be written", () => {
    refused(
      "NaN",
      () => Build.chart({ type: "bar", series: [{ ...series[0]!, values: [Number.NaN, 2] }] }),
      "options.series[0].values[0]"
    );
    refused("type", () => Build.chart({ type: "bogus" as never, series }), "options.type");
    refused(
      "categories",
      () => Build.chart({ type: "bar", series: [{ ...series[0]!, categories: [1] as never }] }),
      "options.series[0].categories[0]"
    );
    refused(
      "name",
      () => Build.chart({ type: "bar", series: [{ ...series[0]!, name: 1 as never }] }),
      "options.series[0].name"
    );
  });

  it("Template.bindChartData refuses bad values and a chart it cannot find", () => {
    const d = doc();
    Document.addContent(d, Build.chart({ type: "bar", series, title: "T", name: "Sales" }));
    const before = snapshot(d);
    refused(
      "NaN",
      () =>
        Template.bindChartData(d, [
          { chartRef: 0, series: [{ name: "n", values: [Number.NaN, 1] }] }
        ]),
      "chartBindings[0].series[0].values[0]"
    );
    for (const [label, chartRef] of [
      ["negative index", -1],
      ["fractional index", 0.5],
      ["NaN index", Number.NaN],
      ["empty name", ""],
      ["object ref", {}]
    ] as const) {
      refused(
        label,
        () => Template.bindChartData(d, [{ chartRef: chartRef as never, series: [] }]),
        "chartBindings[0].chartRef"
      );
    }
    refused("length", () =>
      Template.bindChartData(d, [
        { chartRef: 0, series: [{ name: "n", values: [1, 2, 3] }], categories: ["x"] }
      ])
    );
    refused("not an array", () => Template.bindChartData(d, null as never), "chartBindings");
    expect(snapshot(d)).toBe(before);
    const bound = Template.bindChartData(d, [
      { chartRef: "Sales", series: [{ name: "n", values: [5, 6] }] }
    ]);
    expect(bound).not.toBe(d);
    // A binding naming no chart is skipped, as it always was: one set of
    // bindings may be applied to documents that lack some of the charts.
    const skipped = Template.bindChartData(d, [
      { chartRef: 9, series: [{ name: "n", values: [1, 2] }] },
      { chartRef: "Nope", series: [] }
    ]);
    expect(snapshot(skipped.body)).toBe(snapshot(d.body));
  });
});

describe("Document", () => {
  it("setCompatibilityMode takes a mode Word has", () => {
    for (const mode of [99, "x", 15.5]) {
      refused(String(mode), () => Document.setCompatibilityMode(doc(), mode as never), "mode");
    }
    const d = doc();
    Document.setCompatibilityMode(d, 12);
    expect(Query.getCompatibilityMode(d)).toBe(12);
  });

  it("addContent refuses a node that is not a block", () => {
    const d = doc();
    refused("bogus", () => Document.addContent(d, { type: "bogus" } as never), "content.type");
    refused(
      "insert bogus",
      () => Document.insertContent(d, 0, { type: "run" } as never),
      "content.type"
    );
    expect(Document.getContentCount(d)).toBe(1);
  });

  it("addStyle replaces a style with the same id rather than writing two", async () => {
    const d = doc();
    Document.addStyle(d, { styleId: "Dup", type: "paragraph", name: "First" });
    Document.addStyle(d, { styleId: "Dup", type: "paragraph", name: "Second" });
    expect(d.styles?.filter(s => s.styleId === "Dup")).toEqual([
      { styleId: "Dup", type: "paragraph", name: "Second" }
    ]);
    const read = await Io.read(await Io.toBuffer(d));
    expect(read.styles?.filter(s => s.styleId === "Dup")).toHaveLength(1);
  });

  it("replaceText, mailMerge and revisions check what they are given", () => {
    const d = doc();
    refused(
      "replacement null",
      () => Document.replaceText(d, "Hello", null as never),
      "replacement"
    );
    refused(
      "replacement number",
      () => Document.replaceText(d, "Hello", 42 as never),
      "replacement"
    );
    refused("search null", () => Document.replaceText(d, null as never, "x"), "search");
    expect(Query.getText(d)).toBe("Hello world");
    refused("mailMerge null", () => Document.mailMerge(d, null as never), "data");
    refused("mailMerge value", () => Document.mailMerge(d, { First: 5 as never }), 'data["First"]');
    refused("acceptRevision", () => Document.acceptRevision(d, "x" as never), "revisionId");
    refused("rejectRevision", () => Document.rejectRevision(d, Number.NaN), "revisionId");
  });
});

describe("Package.edit", () => {
  it("refuses an edit it does not know or one missing what it needs", async () => {
    const bytes = await Io.toBuffer(doc());
    await refusedAsync(
      "unknown type",
      () => Package.edit(bytes, [{ bogus: 1 } as never]),
      "edits[0].type"
    );
    await refusedAsync(
      "text",
      () =>
        Package.edit(bytes, [{ type: "replacePartText", path: "word/a.xml", text: 42 as never }]),
      "edits[0].text"
    );
    await refusedAsync(
      "data",
      () => Package.edit(bytes, [{ type: "replacePart", path: "word/a.bin", data: "x" as never }]),
      "edits[0].data"
    );
    await refusedAsync(
      "path",
      () => Package.edit(bytes, [{ type: "deletePart", path: 1 as never }]),
      "edits[0].path"
    );
    await refusedAsync(
      "body",
      () => Package.edit(bytes, [{ type: "replaceBody", body: null as never }]),
      "edits[0].body"
    );
    await refusedAsync(
      "children",
      () =>
        Package.edit(bytes, [
          { type: "replaceHeader", path: "word/header1.xml", children: null as never }
        ]),
      "edits[0].children"
    );
    await refusedAsync("edits", () => Package.edit(bytes, null as never), "edits");
    await refusedAsync(
      "level",
      () => Package.edit(bytes, [], { compressionLevel: 99 }),
      "options.compressionLevel"
    );
  });

  it("finds a part by the name it is written under, with or without the leading slash", async () => {
    const bytes = await Io.toBuffer(doc());
    expect(await Package.readPart(bytes, "/word/document.xml")).toEqual(
      await Package.readPart(bytes, "word/document.xml")
    );
  });
});

describe("security policy", () => {
  it("refuses a limit that would switch the limit off, and an unknown raw-XML policy", async () => {
    for (const key of ["maxPackageSize", "maxPartSize", "maxPartCount", "maxXmlDepth"] as const) {
      for (const bad of [Number.NaN, -1, "10"]) {
        refused(
          `${key}=${String(bad)}`,
          () => Security.resolvePolicy({ [key]: bad as never }),
          `policy.${key}`
        );
      }
    }
    refused(
      "rawXmlPolicy",
      () => Security.resolvePolicy({ rawXmlPolicy: "bogus" as never }),
      "policy.rawXmlPolicy"
    );
    refused(
      "boolean",
      () => Security.resolvePolicy({ preserveVbaProject: "yes" as never }),
      "policy.preserveVbaProject"
    );
    expect(
      Security.resolvePolicy({ maxPackageSize: Number.POSITIVE_INFINITY }).maxPackageSize
    ).toBe(Number.POSITIVE_INFINITY);
    const bytes = await Io.toBuffer(doc());
    await refusedAsync("Io.read", () =>
      Io.read(bytes, { securityPolicy: { maxPackageSize: Number.NaN } })
    );
  });

  it("does not let a caller change the defaults for everyone", () => {
    expect(Object.isFrozen(Security.DEFAULT_POLICY)).toBe(true);
    expect(Object.isFrozen(Security.STRICT_POLICY)).toBe(true);
  });

  it("keeps the deprecated policy names as the same objects", () => {
    expect(Security.DEFAULT_SECURITY_POLICY).toBe(Security.DEFAULT_POLICY);
    expect(Security.STRICT_SECURITY_POLICY).toBe(Security.STRICT_POLICY);
  });
});

describe("Glossary.createBlock", () => {
  it("refuses an empty name, an unknown gallery and content that is not blocks", () => {
    refused("name", () => Glossary.createBlock("", "autoText", []), "name");
    refused("gallery", () => Glossary.createBlock("x", "bogus" as never, []), "gallery");
    refused("content", () => Glossary.createBlock("x", "autoText", null as never), "content");
    refused("blocks", () => Glossary.createDocument(null as never), "blocks");
  });
});

describe("fonts", () => {
  it("refuses bytes that are not a font", () => {
    refused("subset null", () => Font.subset(null as never, "a"), "fontData");
    refused("subset junk", () => Font.subset(new Uint8Array([1, 2, 3, 4, 5]), "a"), "fontData");
    refused(
      "embed junk",
      () => Font.embed({ name: "F", data: new Uint8Array([1, 2, 3, 4, 5]) }),
      "options.data"
    );
    refused(
      "family junk",
      () => Font.embedFamily("F", { regular: new Uint8Array([1, 2, 3, 4, 5]) }),
      "variants.regular"
    );
  });
});

describe("output options", () => {
  it("refuses a compression level or stamp the package cannot use, in the Word API's own error", async () => {
    const d = doc();
    await refusedAsync(
      "toBuffer level",
      () => Io.toBuffer(d, { compressionLevel: 99 }),
      "options.compressionLevel"
    );
    await refusedAsync(
      "toBuffer NaN",
      () => Io.toBuffer(d, { compressionLevel: Number.NaN }),
      "options.compressionLevel"
    );
    await refusedAsync(
      "toBase64",
      () => Io.toBase64(d, { compressionLevel: 10 }),
      "options.compressionLevel"
    );
    await refusedAsync(
      "toFlatOpc",
      () => Io.toFlatOpc(d, { compressionLevel: 99 }),
      "options.compressionLevel"
    );
    await refusedAsync(
      "modTime",
      () => Io.toBuffer(d, { modTime: new Date(Number.NaN) }),
      "options.modTime"
    );
    const bytes = await Io.toBuffer(d);
    await refusedAsync(
      "fillTemplate",
      () => Package.fillTemplate(bytes, {}, { compressionLevel: 99 }),
      "options.compressionLevel"
    );
    await refusedAsync(
      "patch",
      () => Package.patch(bytes, [], { compressionLevel: 99 }),
      "options.compressionLevel"
    );
    refused(
      "stream",
      () => new Stream.DocumentWriter({ compressionLevel: 99 }),
      "options.compressionLevel"
    );
  });

  it("refuses a streamed element that is not a block, before anything is written", () => {
    const writer = new Stream.DocumentWriter();
    refused("null", () => writer.add(null as never), "element");
    refused("bogus", () => writer.add({ type: "bogus" } as never), "element.type");
    refused("addText", () => writer.addText(42 as never), "content");
  });
});

describe("WordCrypto.encrypt", () => {
  it("refuses bytes that are not a ZIP package", async () => {
    await refusedAsync(
      "junk",
      () => WordCrypto.encrypt(new Uint8Array([1, 2, 3]), "pw"),
      "zipBytes"
    );
  });
});

describe("unsupported", () => {
  it("is still its own error", () => {
    expect(new DocxUnsupportedFeatureError("x")).toBeInstanceOf(Error);
  });
});

describe("arguments other than the document", () => {
  it("are checked, not dereferenced", async () => {
    const d = doc();
    const cases: [string, () => unknown, string][] = [
      ["Query.searchText", () => Query.searchText(d, null as never), "query"],
      ["Query.searchByFormat", () => Query.searchByFormat(d, null as never), "criteria"],
      ["Query.getCountByFormat", () => Query.getCountByFormat(d, null as never), "criteria"],
      ["Query.getParagraphStyle", () => Query.getParagraphStyle(d, null as never), "para"],
      ["Query.getRunStyle", () => Query.getRunStyle(d, null as never), "run"],
      ["Query.getTableStyle", () => Query.getTableStyle(d, null as never), "tableStyleId"],
      ["Query.getNumberingLevel", () => Query.getNumberingLevel(d, null as never), "para"],
      ["Query.getBookmark", () => Query.getBookmark(d, null as never), "name"],
      ["Query.getComment", () => Query.getComment(d, "x" as never), "id"],
      ["Query.each", () => Query.each(d, null as never), "visitor"],
      ["Query.eachBlock", () => Query.eachBlock(null as never, {}), "blocks"],
      ["Transform.map", () => Transform.map(d, null as never), "transformer"],
      [
        "Transform.resolveDataBindings",
        () => Transform.resolveDataBindings(d, 42 as never),
        "data"
      ],
      ["Font.addEmbedded", () => Font.addEmbedded(d, null as never), "results"],
      ["Font.createHyphenator", () => Font.createHyphenator(null as never), "patterns"],
      ["Font.hyphenateWord", () => Font.hyphenateWord("word", null as never), "hyphenator"],
      ["Font.hyphenateText", () => Font.hyphenateText(null as never, () => []), "text"],
      ["Convert.isFlatOpc", () => Convert.isFlatOpc(null as never), "content"],
      ["Convert.flatOpcToParts", () => Convert.flatOpcToParts(null as never), "xmlContent"],
      ["Convert.partsToFlatOpc", () => Convert.partsToFlatOpc(null as never), "entries"],
      ["Convert.mathMlToOmml", () => Convert.mathMlToOmml(null as never), "mathml"],
      ["Convert.ommlToMathMl", () => Convert.ommlToMathMl(null as never), "content"],
      ["Diff.compare", () => Diff.compare(null as never, d), "oldDoc"],
      ["WordCrypto.isEncrypted", () => WordCrypto.isEncrypted(null as never), "buffer"],
      ["WordCrypto.Cfb.write", () => WordCrypto.Cfb.write(null as never), "entries"]
    ];
    for (const [label, run, field] of cases) {
      refused(label, run, field);
    }
  });
});

describe("Document.addCustomProperty", () => {
  it("replaces a property of the same name rather than writing two", async () => {
    const d = doc();
    Document.addCustomProperty(d, "Client", { type: "string", value: "A" });
    Document.addCustomProperty(d, "Other", { type: "number", value: 1 });
    Document.addCustomProperty(d, "Client", { type: "string", value: "B" });
    expect(d.customProperties).toEqual([
      { name: "Client", value: { type: "string", value: "B" } },
      { name: "Other", value: { type: "number", value: 1 } }
    ]);
    const read = await Io.read(await Io.toBuffer(d));
    expect(read.customProperties?.filter(p => p.name === "Client")).toHaveLength(1);
  });
});

describe("Document.setSettings / setAppProperties", () => {
  it("refuse the values their sibling entry points refuse", () => {
    const cases: [string, () => unknown, string][] = [
      ["zoom -5", () => Document.setSettings(doc(), { zoom: -5 }), "settings.zoom"],
      ["zoom x", () => Document.setSettings(doc(), { zoom: "x" as never }), "settings.zoom"],
      ["zoom NaN", () => Document.setSettings(doc(), { zoom: Number.NaN }), "settings.zoom"],
      [
        "zoomType",
        () => Document.setSettings(doc(), { zoomType: "bogus" as never }),
        "settings.zoomType"
      ],
      [
        "tab -5",
        () => Document.setSettings(doc(), { defaultTabStop: -5 }),
        "settings.defaultTabStop"
      ],
      [
        "tab 1.5",
        () => Document.setSettings(doc(), { defaultTabStop: 1.5 }),
        "settings.defaultTabStop"
      ],
      [
        "compat 99",
        () => Document.setSettings(doc(), { compatibilityMode: 99 }),
        "settings.compatibilityMode"
      ],
      [
        "protection edit",
        () => Document.setSettings(doc(), { documentProtection: { edit: "bogus" } }),
        "settings.documentProtection.edit"
      ],
      [
        "protection hash",
        () => Document.setSettings(doc(), { documentProtection: { hashAlgorithm: "MD5" } }),
        "settings.documentProtection.hashAlgorithm"
      ],
      [
        "flag",
        () => Document.setSettings(doc(), { trackRevisions: "yes" as never }),
        "settings.trackRevisions"
      ],
      [
        "app pages x",
        () => Document.setAppProperties(doc(), { pages: "x" as never }),
        "props.pages"
      ],
      [
        "app pages NaN",
        () => Document.setAppProperties(doc(), { pages: Number.NaN }),
        "props.pages"
      ],
      ["app pages -1", () => Document.setAppProperties(doc(), { pages: -1 }), "props.pages"],
      [
        "app company",
        () => Document.setAppProperties(doc(), { company: 1 as never }),
        "props.company"
      ]
    ];
    for (const [label, run, field] of cases) {
      refused(label, run, field);
    }
  });

  it("take back what Io.read gives", async () => {
    const d = doc();
    Document.setSettings(d, { zoom: 120, defaultTabStop: 720, compatibilityMode: 15 });
    Document.setAppProperties(d, { pages: 1, company: "C" });
    const read = await Io.read(await Io.toBuffer(d));
    Document.setSettings(read, { ...read.settings, zoom: 90 });
    Document.setAppProperties(read, { ...read.appProperties });
    expect(read.settings?.zoom).toBe(90);
  });
});

describe("Document.setWatermark sizes", () => {
  it("refuses a watermark that would draw nothing", () => {
    const text = { type: "text", text: "DRAFT" } as const;
    const cases: [string, object, string][] = [
      ["fontSize 0", { ...text, fontSize: 0 }, "watermark.fontSize"],
      ["fontSize -4", { ...text, fontSize: -4 }, "watermark.fontSize"],
      ["widthPt 0", { ...text, widthPt: 0 }, "watermark.widthPt"],
      ["heightPt 0", { ...text, heightPt: 0 }, "watermark.heightPt"],
      ["empty text", { ...text, text: "" }, "watermark.text"],
      ["image widthPt 0", { type: "image", rId: "r", widthPt: 0 }, "watermark.widthPt"]
    ];
    for (const [label, watermark, field] of cases) {
      refused(label, () => Document.setWatermark(doc(), watermark as never), field);
    }
    const d = doc();
    Document.setWatermark(d, { ...text, fontSize: 2, widthPt: 1, heightPt: 1 });
    expect(d.watermark).toBeDefined();
  });
});
