import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { Workbook } from "@excel/index";
import * as SharedUnits from "@utils/units";
import { describe, expect, it } from "vitest";

import { WordExcel } from "../excel";
import { WordHtml } from "../html";
import {
  Build,
  Document,
  DocxError,
  DocxArgumentError,
  Font,
  Io,
  Layout,
  Ole,
  Security,
  Styles,
  Template,
  Transform,
  Units,
  Validation,
  type BodyContent,
  type DiscardUnrecognizedValuesOptions,
  type DocxDocument,
  type Paragraph,
  type Run,
  type Table
} from "../index";
import { WordMarkdown } from "../markdown";

async function partText(bytes: Uint8Array, path: string): Promise<string> {
  for await (const entry of unzip(bytes).entries()) {
    if (entry.path === path) {
      return new TextDecoder().decode(await entry.bytes());
    }
  }
  throw new Error(`no ${path}`);
}

function failure(fn: () => unknown): DocxArgumentError {
  try {
    fn();
  } catch (err) {
    expect(err).toBeInstanceOf(DocxArgumentError);
    expect(err).toBeInstanceOf(DocxError);
    return err as DocxArgumentError;
  }
  throw new Error("expected a DocxArgumentError");
}

describe("Word arguments are checked where they are given", () => {
  it("Document.addHeading rejects a level outside 1–9", () => {
    const doc = Document.create();
    const err = failure(() => Document.addHeading(doc, "x", 12 as never));
    expect(err.field).toBe("level");
    expect(err.message).toBe("Document.addHeading: level must be an integer 1–9, got 12");
    expect(doc.body).toEqual([]);
  });

  it("Document.addParagraph rejects null text with a typed error", () => {
    const doc = Document.create();
    const err = failure(() => Document.addParagraph(doc, null as never));
    expect(err.field).toBe("content");
    expect(err.message).toContain("Document.addParagraph");
    expect(doc.body).toEqual([]);
  });

  it("Document.addParagraph rejects a NaN size at the call, not at Io.toBuffer", () => {
    const doc = Document.create();
    const err = failure(() => Document.addParagraph(doc, "x", { run: { size: NaN } }));
    expect(err.field).toBe("properties.run.size");
    expect(err.message).toContain("Document.addParagraph");
    expect(doc.body).toEqual([]);
  });

  it("Document.insertContent rejects an index past the end instead of appending", () => {
    const doc = Document.create();
    Document.addParagraph(doc, "a");
    const err = failure(() => Document.insertContent(doc, 99, Build.paragraph([])));
    expect(err.field).toBe("index");
    expect(err.message).toBe("Document.insertContent: index must be an integer 0–1, got 99");
    expect(doc.body).toHaveLength(1);
  });

  it("Document.insertContent still accepts 0 and the length", async () => {
    const doc = Document.create();
    Document.insertContent(doc, 0, Build.textParagraph("b"));
    Document.insertContent(doc, 0, Build.textParagraph("a"));
    Document.insertContent(doc, 2, Build.textParagraph("c"));
    expect(doc.body).toHaveLength(3);
    await Io.toBuffer(doc);
  });
});

type Doc = ReturnType<typeof Document.create>;
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const x = undefined as never;

/** [call as written, the field it must name, the call] */
const INVALID: [string, string, (doc: Doc) => unknown][] = [
  // Document.* — body
  ["Document.addContent(null)", "content", d => Document.addContent(d, null as never)],
  ["Document.addContent({})", "content.type", d => Document.addContent(d, {} as never)],
  ["Document.insertContent(-1)", "index", d => Document.insertContent(d, -1, Build.paragraph([]))],
  [
    "Document.insertContent(0.5)",
    "index",
    d => Document.insertContent(d, 0.5, Build.paragraph([]))
  ],
  [
    "Document.insertContent(NaN)",
    "index",
    d => Document.insertContent(d, NaN, Build.paragraph([]))
  ],
  ["Document.insertContent(0, null)", "content", d => Document.insertContent(d, 0, x)],
  ["Document.removeContent(NaN)", "index", d => Document.removeContent(d, NaN)],
  ["Document.removeContent(0.5)", "index", d => Document.removeContent(d, 0.5)],
  ["Document.getContent('0')", "index", d => Document.getContent(d, "0" as never)],
  // paragraphs and headings
  ["Document.addParagraph(undefined)", "content", d => Document.addParagraph(d, x)],
  ["Document.addParagraph(42)", "content", d => Document.addParagraph(d, 42 as never)],
  [
    "Document.addParagraph(table node)",
    "para",
    d => Document.addParagraph(d, Build.simpleTable([["a"]]) as never)
  ],
  [
    "addParagraph size Infinity",
    "properties.run.size",
    d => Document.addParagraph(d, "x", { run: { size: Infinity } })
  ],
  [
    "addParagraph size fractional",
    "properties.run.size",
    d => Document.addParagraph(d, "x", { run: { size: 21.5 } })
  ],
  [
    "addParagraph size negative",
    "properties.run.size",
    d => Document.addParagraph(d, "x", { run: { size: -2 } })
  ],
  [
    "addParagraph color with #",
    "properties.run.color",
    d => Document.addParagraph(d, "x", { run: { color: "#FF0000" } })
  ],
  [
    "addParagraph color spec",
    "properties.run.color.val",
    d => Document.addParagraph(d, "x", { run: { color: { val: "red" } } })
  ],
  [
    "addParagraph underline",
    "properties.run.underline",
    d => Document.addParagraph(d, "x", { run: { underline: "squiggle" as never } })
  ],
  [
    "addParagraph underline spec",
    "properties.run.underline.style",
    d => Document.addParagraph(d, "x", { run: { underline: { style: "zz" as never } } })
  ],
  [
    "addParagraph highlight",
    "properties.run.highlight",
    d => Document.addParagraph(d, "x", { run: { highlight: "orange" as never } })
  ],
  [
    "addParagraph vertAlign",
    "properties.run.vertAlign",
    d => Document.addParagraph(d, "x", { run: { vertAlign: "up" as never } })
  ],
  [
    "addParagraph scale",
    "properties.run.scale",
    d => Document.addParagraph(d, "x", { run: { scale: 0 } })
  ],
  [
    "addParagraph font",
    "properties.run.font.ascii",
    d => Document.addParagraph(d, "x", { run: { font: { ascii: 1 as never } } })
  ],
  [
    "addParagraph shading fill",
    "properties.shading.fill",
    d => Document.addParagraph(d, "x", { shading: { fill: x } })
  ],
  [
    "addParagraph alignment",
    "properties.alignment",
    d => Document.addParagraph(d, "x", { alignment: "middle" as never })
  ],
  [
    "addParagraph spacing.after NaN",
    "properties.spacing.after",
    d => Document.addParagraph(d, "x", { spacing: { after: NaN } })
  ],
  [
    "addParagraph spacing.lineRule",
    "properties.spacing.lineRule",
    d => Document.addParagraph(d, "x", { spacing: { lineRule: "double" as never } })
  ],
  [
    "addParagraph indent.hanging negative",
    "properties.indent.hanging",
    d => Document.addParagraph(d, "x", { indent: { hanging: -1 } })
  ],
  [
    "addParagraph numbering.level",
    "properties.numbering.level",
    d => Document.addParagraph(d, "x", { numbering: { numId: 1, level: 9 } })
  ],
  [
    "addParagraph border style",
    "properties.borders.top.style",
    d => Document.addParagraph(d, "x", { borders: { top: { style: "solid" as never } } })
  ],
  [
    "addParagraph border size",
    "properties.borders.bottom.size",
    d => Document.addParagraph(d, "x", { borders: { bottom: { style: "single", size: NaN } } })
  ],
  [
    "addParagraph tab position",
    "properties.tabs[0].position",
    d => Document.addParagraph(d, "x", { tabs: [{ type: "left", position: NaN }] })
  ],
  [
    "addParagraph outlineLevel",
    "properties.outlineLevel",
    d => Document.addParagraph(d, "x", { outlineLevel: 10 })
  ],
  [
    "addParagraph markRunProperties",
    "properties.markRunProperties.size",
    d => Document.addParagraph(d, "x", { markRunProperties: { size: NaN } })
  ],
  ["Document.addHeading(0)", "level", d => Document.addHeading(d, "x", 0 as never)],
  ["Document.addHeading(1.5)", "level", d => Document.addHeading(d, "x", 1.5 as never)],
  ["Document.addHeading('2')", "level", d => Document.addHeading(d, "x", "2" as never)],
  ["Document.addHeading(null)", "content", d => Document.addHeading(d, null as never)],
  ["Document.addHeading(undefined)", "content", d => Document.addHeading(d, x)],
  // tables
  ["Document.addTable(null)", "data", d => Document.addTable(d, null as never)],
  ["Document.addTable([['a', 1]])", "data[0][1]", d => Document.addTable(d, [["a", 1 as never]])],
  ["Document.addTable(['a'])", "data[0]", d => Document.addTable(d, ["a"] as never)],
  [
    "Document.addTable width",
    "options.width.value",
    d => Document.addTable(d, [["a"]], { width: { type: "pct", value: NaN } })
  ],
  [
    "Document.addTable columnWidths",
    "options.columnWidths[1]",
    d => Document.addTable(d, [["a"]], { columnWidths: [100, -1] })
  ],
  [
    "Document.addTable(paragraph node)",
    "table",
    d => Document.addTable(d, Build.paragraph([]) as never)
  ],
  // images
  ["Document.addImage(data: null)", "data", d => Document.addImage(d, null as never, "png", 1, 1)],
  [
    "Document.addImage(altText: 1)",
    "options.altText",
    d => Document.addImage(d, PNG, "png", 1, 1, { altText: 1 as never })
  ],
  ["Document.addImage(width NaN)", "width", d => Document.addImage(d, PNG, "png", NaN, 1)],
  [
    "Document.addFloatingImage(data: string)",
    "data",
    d => Document.addFloatingImage(d, "png" as never, "png", 1, 1)
  ],
  // notes, comments, lists
  ["Document.addFootnote(null)", "content", d => Document.addFootnote(d, null as never)],
  ["Document.addFootnote([null])", "content[0]", d => Document.addFootnote(d, [null as never])],
  ["Document.addEndnote(5)", "content", d => Document.addEndnote(d, 5 as never)],
  ["Document.addComment(author null)", "author", d => Document.addComment(d, null as never, "c")],
  [
    "Document.addComment(date 5)",
    "options.date",
    d => Document.addComment(d, "a", "c", { date: 5 as never })
  ],
  ["Document.addTextBox(null)", "content", d => Document.addTextBox(d, null as never)],
  ["Document.addMath(null)", "content", d => Document.addMath(d, null as never)],
  ["Document.addBulletList(null)", "items", d => Document.addBulletList(d, null as never)],
  ["Document.addBulletList([5])", "items[0]", d => Document.addBulletList(d, [5 as never])],
  ["Document.addBulletList(level 9)", "level", d => Document.addBulletList(d, ["a"], 9)],
  ["Document.addNumberedList(level -1)", "level", d => Document.addNumberedList(d, ["a"], -1)],
  [
    "Document.addListDefinition(kind)",
    "kind",
    d => Document.addListDefinition(d, "roman" as never)
  ],
  [
    "Document.addListDefinition(numId)",
    "options.numId",
    d => Document.addListDefinition(d, "bullet", { numId: 1.5 })
  ],
  [
    "Document.addTableOfContents(instruction)",
    "options.instruction",
    d => Document.addTableOfContents(d, { instruction: 5 as never })
  ],
  // document-level settings
  [
    "Document.setSectionProperties(null)",
    "props",
    d => Document.setSectionProperties(d, null as never)
  ],
  [
    "Document.setSectionProperties(pageSize)",
    "props.pageSize.width",
    d => Document.setSectionProperties(d, { pageSize: { width: NaN, height: 1 } })
  ],
  [
    "Document.setSectionProperties(margins)",
    "props.margins.left",
    d => Document.setSectionProperties(d, { margins: { top: 1, right: 1, bottom: 1, left: -1 } })
  ],
  [
    "Document.setSectionProperties(columns)",
    "props.columns.count",
    d => Document.setSectionProperties(d, { columns: { count: 0 } })
  ],
  [
    "Document.setSectionProperties(breakType)",
    "props.breakType",
    d => Document.setSectionProperties(d, { breakType: "page" as never })
  ],
  [
    "Document.addSectionBreak(orientation)",
    "props.pageSize.orientation",
    d =>
      Document.addSectionBreak(d, {
        pageSize: { width: 1, height: 1, orientation: "sideways" as never }
      })
  ],
  [
    "Document.setDocDefaults(run size)",
    "defaults.runProperties.size",
    d => Document.setDocDefaults(d, { runProperties: { size: NaN } })
  ],
  [
    "Document.setDocDefaults(paragraph spacing)",
    "defaults.paragraphProperties.spacing.before",
    d => Document.setDocDefaults(d, { paragraphProperties: { spacing: { before: -5 } } })
  ],
  ["Document.addStyle(null)", "style", d => Document.addStyle(d, null as never)],
  [
    "Document.addStyle(type)",
    "style.type",
    d => Document.addStyle(d, { type: "list" as never, styleId: "S", name: "S" })
  ],
  [
    "Document.addStyle(styleId)",
    "style.styleId",
    d => Document.addStyle(d, { type: "paragraph", styleId: "", name: "S" })
  ],
  [
    "Document.addStyle(runProperties)",
    "style.runProperties.color",
    d =>
      Document.addStyle(d, {
        type: "character",
        styleId: "S",
        name: "S",
        runProperties: { color: "blue" }
      })
  ],
  ["Document.setHeader(key)", "key", d => Document.setHeader(d, 1 as never, { children: [] })],
  [
    "Document.setFooter(children)",
    "content.children",
    d => Document.setFooter(d, "default", { children: null as never })
  ],
  [
    "Document.setWatermark(rotation)",
    "watermark.rotation",
    d => Document.setWatermark(d, { type: "text", text: "DRAFT", rotation: NaN })
  ],
  [
    "Document.setWatermark(color)",
    "watermark.color",
    d => Document.setWatermark(d, { type: "text", text: "DRAFT", color: "auto" })
  ],
  [
    "Document.setWatermark(type)",
    "watermark.type",
    d => Document.setWatermark(d, { type: "video" } as never)
  ],
  [
    "Document.setBackground(color)",
    "background.color",
    d => Document.setBackground(d, { color: "#FFFFFF" })
  ],
  ["Document.setSettings(null)", "settings", d => Document.setSettings(d, null as never)],
  ["Document.setCoreProperties(5)", "props", d => Document.setCoreProperties(d, 5 as never)],
  ["Document.addFont(name)", "font.name", d => Document.addFont(d, { name: null as never })],
  [
    "Document.addCustomProperty(name)",
    "name",
    d => Document.addCustomProperty(d, "", { type: "string", value: "v" })
  ],
  [
    "Document.addCustomProperty(number NaN)",
    "value.value",
    d => Document.addCustomProperty(d, "n", { type: "number", value: NaN })
  ],
  [
    "Document.addCustomProperty(type)",
    "value.type",
    d => Document.addCustomProperty(d, "n", { type: "blob", value: 1 } as never)
  ],
  // Build.*
  ["Build.text(null)", "content", () => Build.text(null as never)],
  ["Build.text(size NaN)", "properties.size", () => Build.text("x", { size: NaN })],
  ["Build.bold(kern)", "properties.kern", () => Build.bold("x", { kern: -1 })],
  ["Build.italic(5)", "content", () => Build.italic(5 as never)],
  ["Build.underline(style)", "style", () => Build.underline("x", "zig" as never)],
  [
    "Build.strikethrough(spacing)",
    "properties.spacing",
    () => Build.strikethrough("x", { spacing: 0.5 })
  ],
  ["Build.symbol(char)", "char", () => Build.symbol("Wingdings", null as never)],
  ["Build.field(instruction)", "instruction", () => Build.field(undefined as never)],
  [
    "Build.positionalTab(alignment)",
    "options.alignment",
    () => Build.positionalTab({ alignment: "middle" as never, relativeTo: "margin" })
  ],
  ["Build.paragraph(null)", "children", () => Build.paragraph(null as never)],
  ["Build.paragraph([null])", "children[0]", () => Build.paragraph([null as never])],
  [
    "Build.paragraph(alignment)",
    "properties.alignment",
    () => Build.paragraph([], { alignment: "justify" as never })
  ],
  ["Build.textParagraph(null)", "content", () => Build.textParagraph(null as never)],
  [
    "Build.textParagraph(run.color)",
    "properties.run.color",
    () => Build.textParagraph("x", { run: { color: "FFF" } })
  ],
  ["Build.heading(10)", "level", () => Build.heading("x", 10 as never)],
  ["Build.heading(content 5)", "content", () => Build.heading(5 as never)],
  ["Build.hyperlink(text)", "linkText", () => Build.hyperlink(null as never, { url: "u" })],
  ["Build.hyperlink(options)", "options", () => Build.hyperlink("x", null as never)],
  ["Build.hyperlink(url)", "options.url", () => Build.hyperlink("x", { url: 5 as never })],
  [
    "Build.hyperlink(properties)",
    "options.properties.size",
    () => Build.hyperlink("x", { url: "u", properties: { size: NaN } })
  ],
  ["Build.bookmarkStart(id)", "id", () => Build.bookmarkStart(NaN, "b")],
  ["Build.bookmarkStart(name)", "name", () => Build.bookmarkStart(1, null as never)],
  ["Build.bookmarkEnd(id)", "id", () => Build.bookmarkEnd(1.5)],
  ["Build.commentRangeStart(id)", "id", () => Build.commentRangeStart("1" as never)],
  ["Build.commentRangeEnd(id)", "id", () => Build.commentRangeEnd(Infinity)],
  ["Build.commentReference(id)", "id", () => Build.commentReference(x)],
  [
    "Build.insertedRun(revision)",
    "revision.author",
    () => Build.insertedRun(Build.text("x"), { id: 1, author: null as never })
  ],
  [
    "Build.deletedRun(revision id)",
    "revision.id",
    () => Build.deletedRun(Build.text("x"), { id: NaN, author: "a" })
  ],
  [
    "Build.movedFromRun(run)",
    "run",
    () => Build.movedFromRun(null as never, { id: 1, author: "a" })
  ],
  ["Build.moveToRangeStart(author)", "author", () => Build.moveToRangeStart(1, null as never)],
  ["Build.border(style)", "options.style", () => Build.border({ style: "solid" as never })],
  ["Build.border(size)", "options.size", () => Build.border({ size: NaN })],
  ["Build.border(color)", "options.color", () => Build.border({ color: "#000000" })],
  ["Build.border positional size", "size", () => Build.border("single", -1)],
  ["Build.gridBorders(size)", "options.size", () => Build.gridBorders({ size: 1.5 })],
  ["Build.gridBorders positional color", "color", () => Build.gridBorders(4, "black")],
  ["Build.cell(null)", "content", () => Build.cell(null as never)],
  ["Build.cell(gridSpan)", "properties.gridSpan", () => Build.cell("x", { gridSpan: 0 })],
  [
    "Build.cell(shading)",
    "properties.shading.fill",
    () => Build.cell("x", { shading: { fill: "grey" } })
  ],
  [
    "Build.cell(width type)",
    "properties.width.type",
    () => Build.cell("x", { width: { value: 1, type: "px" as never } })
  ],
  ["Build.row(null)", "cells", () => Build.row(null as never)],
  ["Build.row(height)", "properties.height.value", () => Build.row([], { height: { value: NaN } })],
  ["Build.table(null)", "rows", () => Build.table(null as never)],
  ["Build.table(layout)", "options.layout", () => Build.table([], { layout: "auto" as never })],
  ["Build.table(columnWidths)", "columnWidths[0]", () => Build.table([], undefined, [NaN])],
  ["Build.simpleTable(null)", "data", () => Build.simpleTable(null as never)],
  [
    "Build.floatingImage(rId)",
    "options.rId",
    () => Build.floatingImage({ rId: null as never, width: 1, height: 1 })
  ],
  [
    "Build.drawingShape(fillColor)",
    "fillColor",
    () => Build.drawingShape({ shapeType: "rect", width: 1, height: 1, fillColor: "auto" })
  ],
  [
    "Build.chart(series)",
    "options.series",
    () => Build.chart({ type: "bar", series: null as never })
  ],
  [
    "Build.shape(fill.color)",
    "options.fill.color",
    () =>
      Build.shape({
        shapeType: "rect",
        width: 1,
        height: 1,
        fill: { type: "solid", color: "#F00" }
      })
  ]
];

describe("every builder and editor refuses a value with no valid form", () => {
  it.each(INVALID)("%s → %s", (_, field, call) => {
    const doc = Document.create();
    Document.addParagraph(doc, "existing");
    const before = structuredClone(doc);
    let caught: unknown;
    try {
      call(doc);
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(DocxArgumentError);
    const err = caught as DocxArgumentError;
    expect(err.field).toBe(field);
    expect(err.message).toMatch(/^(Document|Build)\.\w+: /);
    expect(err.message).toContain(` ${field} must be `);
    expect(doc).toEqual(before);
  });

  it("a drawing error is an argument error too", () => {
    expect(() => Build.rect(NaN, 1)).toThrow(DocxArgumentError);
  });
});

describe("valid boundary values are still accepted and written", () => {
  it("accepts the edges of every range", async () => {
    const doc = Document.create();
    Document.addHeading(doc, "one", 1);
    Document.addHeading(doc, "nine", 9);
    Document.addHeading(doc, [Build.text("runs")]);
    Document.addParagraph(doc, "");
    Document.addParagraph(doc, "styled", {
      alignment: "both",
      spacing: { before: 0, after: 0, line: 240, lineRule: "auto" },
      indent: { left: -720, firstLine: 0 },
      outlineLevel: 9,
      numbering: { numId: 0, level: 8 },
      borders: { top: { style: "nil", size: 0, color: "auto" } },
      shading: { fill: "auto", pattern: "pct5", color: "ffffff" },
      tabs: [{ type: "decimal", position: -100, leader: "dot" }],
      run: {
        size: 0,
        kern: 0,
        position: -4,
        spacing: -20,
        scale: 600,
        color: { val: "00FF00", themeTint: "99" },
        underline: { style: "wavyDouble", color: "auto" },
        highlight: "none",
        font: "Arial"
      }
    });
    Document.addParagraph(doc, "bool underline", { run: { underline: false } });
    Document.insertContent(doc, Document.getContentCount(doc), Build.textParagraph("end"));
    expect(Document.getContent(doc, -1)).toBeUndefined();
    expect(Document.getContent(doc, 999)).toBeUndefined();
    expect(Document.removeContent(doc, -1)).toBeUndefined();
    expect(Document.removeContent(doc, 999)).toBeUndefined();
    Document.addTable(doc, [[""]], { width: { type: "auto", value: 0 }, columnWidths: [0] });
    Document.addTable(doc, []);
    Document.addBulletList(doc, ["a", [Build.text("b")]], 8);
    Document.addNumberedList(doc, [], 0);
    Document.addFootnote(doc, []);
    Document.addComment(doc, "", "c");
    Document.addTextBox(doc, "t", { width: 0, height: 0 });
    Document.addListDefinition(doc, "numbered", { numId: 40 });
    Document.addSectionBreak(doc, {
      pageSize: { width: 12240, height: 15840, orientation: "landscape" },
      margins: { top: -1440, right: 0, bottom: 0, left: 0, gutter: 0 },
      columns: { count: 1, space: 0 }
    });
    Document.setWatermark(doc, { type: "text", text: "D", rotation: 0, color: "C0C0C0" });
    Document.setBackground(doc, { color: "auto" });
    Document.addCustomProperty(doc, "when", { type: "date", value: new Date(0) });
    Document.addCustomProperty(doc, "n", { type: "number", value: -1.5 });
    Document.addStyle(doc, { type: "paragraph", styleId: "S", name: "", outlineLevel: 0 });
    Document.setDocDefaults(doc, { runProperties: { size: 2 } });
    Document.setHeader(doc, "default", { children: [] });
    Document.addContent(doc, Build.paragraph([Build.bookmarkStart(-1, ""), Build.bookmarkEnd(0)]));
    expect(Build.border("single", 0, "auto").size).toBe(0);
    expect(Build.gridBorders(undefined).top?.size).toBe(4);
    await Io.toBuffer(doc);
  });
});

// -----------------------------------------------------------------------------
// Content builders' deep fields
// -----------------------------------------------------------------------------

const m = [Build.mathRun("x")];
const run = Build.text("x");

const INVALID_CONTENT: [string, string, () => unknown][] = [
  ["Build.mathBlock(null)", "content", () => Build.mathBlock(null as never)],
  ["Build.mathRun(text)", "mathText", () => Build.mathRun(5 as never)],
  ["Build.mathRun(bold)", "properties.bold", () => Build.mathRun("x", { bold: "y" as never })],
  ["Build.mathFraction(type)", "fractionType", () => Build.mathFraction(m, m, "over" as never)],
  ["Build.mathFraction(denominator)", "denominator", () => Build.mathFraction(m, null as never)],
  ["Build.mathSqrt([null])", "content[0]", () => Build.mathSqrt([null as never])],
  ["Build.mathRoot(degree)", "degree", () => Build.mathRoot("2" as never, m)],
  ["Build.mathSum(options)", "options", () => Build.mathSum(m, 5 as never)],
  ["Build.mathIntegral(sup)", "sup", () => Build.mathIntegral(m, m, {} as never)],
  ["Build.mathProduct(content)", "content", () => Build.mathProduct(null as never)],
  ["Build.mathSuperscript(base)", "base", () => Build.mathSuperscript(null as never, m)],
  ["Build.mathSubscript(sub)", "subScript", () => Build.mathSubscript(m, null as never)],
  ["Build.mathSubSuperscript", "superScript", () => Build.mathSubSuperscript(m, m, null as never)],
  [
    "Build.mathPreSubSuperscript",
    "preSubScript",
    () => Build.mathPreSubSuperscript(m, null as never, m)
  ],
  ["Build.mathPhantom(show)", "options.show", () => Build.mathPhantom(m, { show: 1 as never })],
  ["Build.mathGroupChar(char)", "options.char", () => Build.mathGroupChar(m, { char: "ab" })],
  [
    "Build.mathGroupChar(position)",
    "options.position",
    () => Build.mathGroupChar(m, { position: "left" as never })
  ],
  [
    "Build.mathGroupChar(verticalAlign)",
    "options.verticalAlign",
    () => Build.mathGroupChar(m, { verticalAlign: "middle" as never })
  ],
  [
    "Build.mathBorderBox(hideTop)",
    "options.hideTop",
    () => Build.mathBorderBox(m, { hideTop: "yes" as never })
  ],
  ["Build.mathDelimiter(content)", "content[0]", () => Build.mathDelimiter([null as never])],
  [
    "Build.mathDelimiter(beginChar)",
    "options.beginChar",
    () => Build.mathDelimiter([m], { beginChar: "((" })
  ],
  ["Build.mathNary(char)", "char", () => Build.mathNary("", m)],
  ["Build.mathFunction(name)", "name", () => Build.mathFunction(null as never, m)],
  ["Build.mathLimit(type)", "limitType", () => Build.mathLimit(m, m, "side" as never)],
  ["Build.mathMatrix(rows)", "rows[0][0]", () => Build.mathMatrix([[null as never]])],
  ["Build.mathAccent(char)", "char", () => Build.mathAccent(m, "^^")],
  ["Build.mathBar(position)", "position", () => Build.mathBar(m, "middle" as never)],
  ["Build.mathBox(null)", "content", () => Build.mathBox(null as never)],
  ["Build.mathEquationArray(rows)", "rows[0]", () => Build.mathEquationArray([5 as never])],
  ["Build.ruby(base)", "baseText", () => Build.ruby(null as never, "r")],
  ["Build.ruby(text)", "rubyText[0]", () => Build.ruby("b", [null as never])],
  [
    "Build.ruby(align)",
    "properties.align",
    () => Build.ruby("b", "r", { align: "middle" as never })
  ],
  ["Build.ruby(fontSize)", "properties.fontSize", () => Build.ruby("b", "r", { fontSize: 9.5 })],
  ["Build.field(cachedValue)", "cachedValue", () => Build.field(" X ", 5 as never)],
  ["Build.pageNumberField(cached)", "cachedValue", () => Build.pageNumberField(1 as never)],
  ["Build.dateField(format)", "options.format", () => Build.dateField({ format: 5 as never })],
  ["Build.timeField(positional)", "format", () => Build.timeField(5 as never)],
  ["Build.sequenceField(id)", "identifier", () => Build.sequenceField("Figure 1")],
  [
    "Build.sequenceField(hide)",
    "options.hide",
    () => Build.sequenceField("F", { hide: 1 as never })
  ],
  ["Build.styleRefField(name)", "styleName", () => Build.styleRefField(null as never)],
  ["Build.refField(name)", "bookmarkName", () => Build.refField("")],
  [
    "Build.refField(hyperlink)",
    "options.hyperlink",
    () => Build.refField("b", { hyperlink: "y" as never })
  ],
  ["Build.pageRefField(name)", "bookmarkName", () => Build.pageRefField('a"b')],
  ["Build.noteRefField(name)", "bookmarkName", () => Build.noteRefField(undefined as never)],
  [
    "Build.hyperlinkField(tooltip)",
    "options.tooltip",
    () => Build.hyperlinkField("u", { tooltip: 1 as never })
  ],
  ["Build.quoteField(text)", "text", () => Build.quoteField(null as never)],
  [
    "Build.tocField(headingLevels)",
    "options.headingLevels",
    () => Build.tocField({ headingLevels: "3-1" })
  ],
  [
    "Build.tocField(tabLeader)",
    "options.tabLeader",
    () => Build.tocField({ tabLeader: "*" as never })
  ],
  ["Build.tocField(styles)", "options.styles[0]", () => Build.tocField({ styles: [1 as never] })],
  ["Build.tcField(level)", "options.level", () => Build.tcField("t", { level: 0 })],
  [
    "Build.indexEntryField(bold)",
    "options.bold",
    () => Build.indexEntryField("t", { bold: 1 as never })
  ],
  ["Build.indexField(columns)", "options.columns", () => Build.indexField({ columns: 0 })],
  ["Build.ifField(trueText)", "trueText", () => Build.ifField("a = b", null as never, "")],
  ["Build.includeTextField(path)", "filePath", () => Build.includeTextField(null as never)],
  ["Build.includePictureField(path)", "filePath", () => Build.includePictureField(5 as never)],
  [
    "Build.formTextField(maxLength)",
    "options.maxLength",
    () => Build.formTextField({ maxLength: -1 })
  ],
  [
    "Build.formTextField(name)",
    "options.name",
    () => Build.formTextField({ name: "n".repeat(66) })
  ],
  [
    "Build.formTextField(statusText)",
    "options.statusText",
    () => Build.formTextField({ statusText: "s".repeat(141) })
  ],
  ["Build.formCheckboxField(size)", "options.size", () => Build.formCheckboxField({ size: NaN })],
  [
    "Build.formCheckboxField(checked)",
    "options.checked",
    () => Build.formCheckboxField({ checked: "yes" as never })
  ],
  [
    "Build.formDropdownField(default)",
    "options.default",
    () => Build.formDropdownField({ entries: ["a", "b"], default: 2 })
  ],
  [
    "Build.formDropdownField(entries)",
    "options.entries[0]",
    () => Build.formDropdownField({ entries: [1 as never] })
  ],
  [
    "Build.structuredDocumentTag(content)",
    "content",
    () => Build.structuredDocumentTag(null as never)
  ],
  [
    "Build.structuredDocumentTag(id)",
    "properties.id",
    () => Build.structuredDocumentTag([], { id: 1.5 })
  ],
  [
    "Build.structuredDocumentTag(appearance)",
    "properties.appearance",
    () => Build.structuredDocumentTag([], { appearance: "box" as never })
  ],
  [
    "Build.structuredDocumentTag(dropdownList)",
    "properties.dropdownList[0].value",
    () => Build.structuredDocumentTag([], { dropdownList: [{ value: 1 as never }] })
  ],
  [
    "Build.structuredDocumentTag(date)",
    "properties.date.storeMappedDataAs",
    () => Build.structuredDocumentTag([], { date: { storeMappedDataAs: "iso" as never } })
  ],
  [
    "Build.structuredDocumentTag(checkbox)",
    "properties.checkbox.checkedChar",
    () => Build.structuredDocumentTag([], { checkbox: { checkedChar: "" } })
  ],
  [
    "Build.structuredDocumentTag(dataBinding)",
    "properties.dataBinding.xpath",
    () =>
      Build.structuredDocumentTag([], { dataBinding: { xpath: null as never, storeItemId: "x" } })
  ],
  ["Build.checkbox(checked)", "options.checked", () => Build.checkbox({ checked: 1 as never })],
  [
    "Build.checkbox(state)",
    "options.checkedState.value",
    () => Build.checkbox({ checkedState: { value: null as never } })
  ],
  [
    "Document.addTableOfContents(range)",
    "options.headingStyleRange",
    () => Document.addTableOfContents(Document.create(), { headingStyleRange: "1-10" })
  ],
  [
    "Document.addTableOfContents(leader)",
    "options.leader",
    () => Document.addTableOfContents(Document.create(), { leader: "stars" as never })
  ],
  [
    "Document.addTableOfContents(stylesWithLevels)",
    "options.stylesWithLevels[0].level",
    () =>
      Document.addTableOfContents(Document.create(), {
        stylesWithLevels: [{ styleName: "Title", level: 0 }]
      })
  ],
  [
    "Document.addTextBox(stroke)",
    "options.stroke",
    () => Document.addTextBox(Document.create(), "t", { stroke: "1" as never })
  ],
  [
    "Document.setSectionProperties(footnotes)",
    "props.footnoteProperties.numFmt",
    () =>
      Document.setSectionProperties(Document.create(), {
        footnoteProperties: { numFmt: "roman" as never }
      })
  ],
  [
    "Document.addSectionBreak(endnotes)",
    "props.endnoteProperties.position",
    () =>
      Document.addSectionBreak(Document.create(), {
        endnoteProperties: { position: "pageBottom" as never }
      })
  ],
  ["Build.table(alignment both)", "options.alignment", () => Build.table([], { alignment: "both" })]
];

describe("content builders check their deep fields", () => {
  it.each(INVALID_CONTENT)("%s → %s", (_, field, call) => {
    const err = failure(call);
    expect(err.field).toBe(field);
    expect(err.message).toMatch(/^(Document|Build)\.\w+: /);
  });

  it("valid content builders still build and write", async () => {
    const doc = Document.create();
    Document.addMath(doc, [
      Build.mathFraction(m, m, "lin"),
      Build.mathDelimiter([m], { beginChar: "", endChar: "]" }),
      Build.mathNary("∑", m),
      Build.mathAccent(m, "\u{1D400}"),
      Build.mathMatrix([[m, m]]),
      Build.mathEquationArray([m]),
      Build.mathGroupChar(m, { char: "⏟", position: "bottom", verticalAlign: "top" }),
      Build.mathLimit(m, m, "upper"),
      Build.mathBar(m, "bottom")
    ]);
    Document.addParagraph(
      doc,
      Build.paragraph([
        Build.ruby("漢", [run], { align: "center", fontSize: 10, raise: -2 }),
        Build.sequenceField("Figure", { hide: true }),
        Build.refField("_Ref1", { hyperlink: true }),
        Build.tocField({ headingLevels: "1-9", tabLeader: "." }),
        Build.tcField("t", { level: 9 }),
        Build.formTextField({ name: "n".repeat(65), maxLength: 0 }),
        Build.formDropdownField({ entries: ["a", "b"], default: 1 }),
        Build.formCheckboxField({ size: 20, checked: true }),
        Build.dateField("d MMMM yyyy"),
        Build.pageNumberField()
      ])
    );
    Document.addContent(doc, Build.checkbox({ checkedState: { value: "☒" } }));
    Document.addContent(
      doc,
      Build.structuredDocumentTag([Build.textParagraph("x")], {
        id: -5,
        appearance: "hidden",
        checkbox: { checked: true, checkedChar: "☒", uncheckedChar: "2610" }
      })
    );
    Document.addTableOfContents(doc, { headingStyleRange: "2-2", leader: "dot" });
    const xml = await partText(await Io.toBuffer(doc), "word/document.xml");
    // A glyph is written as its code point, which `w14:val` requires; a code point as it is.
    expect(xml).toContain('<w14:checkedState w14:val="2612"');
    expect(xml).toContain('<w14:uncheckedState w14:val="2610"');
    expect(xml).not.toContain('w14:val="☒"');
  });
});

// -----------------------------------------------------------------------------
// Other namespaces
// -----------------------------------------------------------------------------

// Begins with the TrueType sfnt version, so `Font.*` takes it as a font.
const BYTES = new Uint8Array([0, 1, 0, 0, 1, 2, 3]);
const SERIES = [{ name: "s", categories: ["a"], values: [1] }];

/** [call, field, call] — a call may be async, and then must reject. */
const INVALID_OPTIONS: [string, string, (doc: Doc) => unknown][] = [
  [
    "Security.protect(edit)",
    "options.edit",
    d => Security.protect(d, { edit: "everything" as never })
  ],
  [
    "Security.protect(spinCount)",
    "options.spinCount",
    d => Security.protect(d, { edit: "readOnly", spinCount: NaN })
  ],
  [
    "Security.protect(hashAlgorithm)",
    "options.hashAlgorithm",
    d => Security.protect(d, { edit: "readOnly", hashAlgorithm: "MD5" as never })
  ],
  ["Security.protect(null)", "options", d => Security.protect(d, null as never)],
  ["Security.encrypt(password)", "password", () => Security.encrypt(BYTES, null as never)],
  [
    "Security.encrypt(keyBits)",
    "options.keyBits",
    () => Security.encrypt(BYTES, "p", { keyBits: 512 as never })
  ],
  [
    "Security.encrypt(spinCount)",
    "options.spinCount",
    () => Security.encrypt(BYTES, "p", { spinCount: 10_000_001 })
  ],
  ["Font.embed(name)", "options.name", () => Font.embed({ name: "", data: BYTES })],
  ["Font.embed(data)", "options.data", () => Font.embed({ name: "F", data: null as never })],
  [
    "Font.embed(style)",
    "options.style",
    () => Font.embed({ name: "F", data: BYTES, style: "heavy" as never })
  ],
  [
    "Font.embed(charset)",
    "options.charset",
    () => Font.embed({ name: "F", data: BYTES, charset: 256 })
  ],
  [
    "Font.embed(panose1)",
    "options.panose1",
    () => Font.embed({ name: "F", data: BYTES, panose1: "0204" })
  ],
  [
    "Font.embedFamily(variant)",
    "variants.heavy",
    () => Font.embedFamily("F", { heavy: BYTES } as never)
  ],
  ["Ole.createEmbedding(progId)", "progId", () => Ole.createEmbedding(BYTES, "")],
  [
    "Ole.createEmbedding(fileName)",
    "options.fileName",
    () => Ole.createEmbedding(BYTES, "Excel.Sheet.12", { fileName: "../evil.bin" })
  ],
  [
    "Ole.createEmbedding(previewContentType)",
    "options.previewContentType",
    () => Ole.createEmbedding(BYTES, "Excel.Sheet.12", { previewImage: BYTES })
  ],
  [
    "Ole.add(widthPt)",
    "options.widthPt",
    d => Ole.add(d, Ole.createEmbedding(BYTES, "Excel.Sheet.12"), { widthPt: NaN })
  ],
  [
    "Ole.add(displayAs)",
    "options.displayAs",
    d => Ole.add(d, Ole.createEmbedding(BYTES, "P"), { displayAs: "link" as never })
  ],
  [
    "Template.fill(empty delimiter)",
    "options.delimiters",
    d => Template.fill(d, {}, { delimiters: ["", "}}"] })
  ],
  [
    "Template.fill(same delimiters)",
    "options.delimiters",
    d => Template.fill(d, {}, { delimiters: ["%", "%"] })
  ],
  ["Template.fill(data)", "data", d => Template.fill(d, null as never)],
  [
    "Template.fillTemplate(strict)",
    "options.strict",
    d => Template.fillTemplate(d, {}, { strict: "no" as never })
  ],
  [
    "Template.listTags(delimiters)",
    "options.delimiters",
    d => Template.listTags(d, { delimiters: ["<<"] as never })
  ],
  ["Styles.create(null)", "rules", () => Styles.create(null as never)],
  [
    "Styles.create(source)",
    "rules[0].source",
    () => Styles.create([{ source: "div" as never, conditions: [], target: { tagName: "p" } }])
  ],
  [
    "Styles.create(tagName)",
    "rules[0].target.tagName",
    () => Styles.create([{ source: "p", conditions: [], target: { tagName: "p onclick=x" } }])
  ],
  [
    "Styles.create(regex)",
    "rules[0].conditions[0].value",
    () =>
      Styles.create([
        {
          source: "p",
          conditions: [{ attribute: "style-name", value: "(", regex: true }],
          target: { tagName: "p" }
        }
      ])
  ],
  [
    "Styles.create(priority)",
    "rules[0].priority",
    () => Styles.create([{ source: "p", conditions: [], target: { tagName: "p" }, priority: NaN }])
  ],
  ["Styles.parse(dsl)", "dsl", () => Styles.parse(null as never)],
  ["Styles.merge(map)", "maps[1]", () => Styles.merge(Styles.DEFAULT, null as never)],
  [
    "Layout.compute(defaultFontSize)",
    "options.defaultFontSize",
    d => Layout.compute(d, { defaultFontSize: 0 })
  ],
  [
    "Layout.getPageMap(pageWidth)",
    "options.pageGeometry.pageWidth",
    d => Layout.getPageMap(d, { pageGeometry: { pageWidth: NaN } })
  ],
  [
    "Layout.compute(marginTop)",
    "options.pageGeometry.marginTop",
    d => Layout.compute(d, { pageGeometry: { marginTop: Infinity } })
  ],
  ["Layout.toSvg(scale)", "options.scale", d => Layout.toSvg(d, { scale: -1 })],
  ["Layout.pageToSvg(NaN)", "pageNumber", d => Layout.pageToSvg(d, NaN)],
  [
    "Layout.computedPageToSvg(1.5)",
    "pageNumber",
    d => Layout.computedPageToSvg(Layout.compute(d), 1.5)
  ],
  [
    "Transform.merge(sectionBreak)",
    "options.sectionBreak",
    d => Transform.merge([d, d], { sectionBreak: "page" as never })
  ],
  ["Transform.merge(null)", "documents[1]", d => Transform.merge([d, null as never])],
  ["Transform.split(by)", "options.by", d => Transform.split(d, { by: "chapter" as never })],
  [
    "Transform.split(headingLevel)",
    "options.headingLevel",
    d => Transform.split(d, { by: "heading", headingLevel: 0 })
  ],
  [
    "Transform.updateFields(updateToc)",
    "options.updateToc",
    d => Transform.updateFields(d, { updateToc: "yes" as never })
  ],
  [
    "Transform.updateFields(layoutOptions)",
    "options.layoutOptions.averageCharWidth",
    d => Transform.updateFields(d, { layoutOptions: { averageCharWidth: NaN } })
  ],
  [
    "Transform.fillFormFields(NaN)",
    'values["f"]',
    d => Transform.fillFormFields(d, new Map([["f", NaN]]))
  ],
  ["Transform.fillFormFields(number)", "values", d => Transform.fillFormFields(d, 42 as never)],
  ["WordHtml.toDocx(html)", "html", () => WordHtml.toDocx(null as never)],
  [
    "WordHtml.toDocxBody(size)",
    "options.defaultFontSize",
    () => WordHtml.toDocxBody("<p>x</p>", { defaultFontSize: 11.5 })
  ],
  [
    "WordHtml.toDocx(classStyles)",
    "options.classStyles.a",
    () => WordHtml.toDocx("<p>x</p>", { classStyles: { a: 1 as never } })
  ],
  ["WordMarkdown.toDocx(markdown)", "markdown", () => WordMarkdown.toDocx(5 as never)],
  [
    "WordMarkdown.toDocxBody(codeBlockFit)",
    "options.codeBlockFit",
    () => WordMarkdown.toDocxBody("x", { codeBlockFit: "scroll" as never })
  ],
  [
    "WordMarkdown.toDocx(contentWidth)",
    "options.contentWidth",
    () => WordMarkdown.toDocx("x", { contentWidth: NaN })
  ],
  [
    "WordExcel.toDocx(sheetHeadingLevel)",
    "options.sheetHeadingLevel",
    () => WordExcel.toDocx(Workbook.create(), { sheetHeadingLevel: 10 })
  ],
  [
    "WordExcel.toDocx(sheets)",
    "options.sheets[0]",
    () => WordExcel.toDocx(Workbook.create(), { sheets: [-1] })
  ],
  [
    "WordExcel.toChartExXml(type)",
    "options.type",
    () => WordExcel.toChartExXml({ type: "pie" as never, series: SERIES })
  ],
  [
    "WordExcel.toChartExXml(values)",
    "options.series[0].values[0]",
    () => WordExcel.toChartExXml({ type: "funnel", series: [{ name: "s", values: [NaN] }] })
  ],
  [
    "WordExcel.toChartXlsx(series)",
    "series[0].name",
    () => WordExcel.toChartXlsx([{ name: 1 as never, categories: [], values: [] }])
  ]
];

describe("other namespaces check the values they take", () => {
  it.each(INVALID_OPTIONS)("%s → %s", async (_, field, call) => {
    const doc = Document.create();
    Document.addParagraph(doc, "existing");
    const before = structuredClone(doc);
    let caught: unknown;
    try {
      await call(doc);
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(DocxArgumentError);
    const err = caught as DocxArgumentError;
    expect(err.field).toBe(field);
    expect(err.message).toMatch(
      /^(Security|Font|Ole|Template|Styles|Layout|Transform|WordHtml|WordMarkdown|WordExcel|WordCrypto)\.\w+: /
    );
    expect(doc).toEqual(before);
  });

  it("valid options still work, and an embedded font's charset is written in hex", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "x");
    await Security.protect(doc, { edit: "forms", spinCount: 0, hashAlgorithm: "SHA-512" });
    const font = Font.embed({ name: "F", data: BYTES, charset: 128, obfuscate: false });
    expect(font.fontDef.charset).toBe("80");
    expect(Font.embed({ name: "F", data: BYTES, charset: 0 }).fontDef.charset).toBe("00");
    Ole.add(doc, Ole.createEmbedding(BYTES, "P", { fileName: "o.bin" }), { widthPt: 0 });
    Template.fill(doc, {}, { delimiters: ["<<", ">>"] });
    Layout.compute(doc, { pageGeometry: { marginTop: -10, pageWidth: 600 } });
    expect(() => Layout.pageToSvg(doc, 99)).toThrow(RangeError);
    Transform.split(doc, { by: "heading", headingLevel: 9 });
    const formDoc = Document.create();
    Document.addParagraph(formDoc, Build.paragraph([Build.formTextField({ name: "n" })]));
    Transform.fillFormFields(formDoc, new Map<string, string | number | boolean>([["n", 1.5]]));
    WordHtml.toDocxBody("<p>x</p>", { defaultFontSize: 0 });
    await WordMarkdown.toDocxBody("x", { codeBlockFit: "wrap", contentWidth: 1 });
    const book = Workbook.create();
    Workbook.addWorksheet(book, "S");
    WordExcel.toDocx(book, { sheetHeadingLevel: 9, sheets: ["S", 0] });
    Styles.create([
      {
        source: "p",
        conditions: [{ attribute: "style-name", value: "^H", regex: true }],
        target: { tagName: "h1" }
      }
    ]);
  });
});

// -----------------------------------------------------------------------------
// Units
// -----------------------------------------------------------------------------

describe("Units checks its argument and keeps every finite result", () => {
  const names = Object.keys(Units) as (keyof typeof Units)[];

  it.each(names)("Units.%s", name => {
    const fn = Units[name] as (v: number) => number;
    const shared = (SharedUnits as unknown as Record<string, (v: number) => number>)[name]!;
    for (const v of [0, 1, -2.5, 1440, 0.3333]) {
      expect(fn(v)).toBe(shared(v));
    }
    for (const bad of [NaN, Infinity, -Infinity, "1", null, undefined]) {
      const err = failure(() => fn(bad as never));
      expect(err.message).toMatch(new RegExp(`^Units\\.${name}: \\w+ must be a finite number`));
    }
  });

  it("Units.ptToTwips(NaN) names the call", () => {
    expect(() => Units.ptToTwips(NaN)).toThrow(
      "Units.ptToTwips: pt must be a finite number, got NaN"
    );
  });
});

// -----------------------------------------------------------------------------
// Vocabulary at write
// -----------------------------------------------------------------------------

function handBuilt(apply: (doc: Writable) => void): Doc {
  const doc = Document.create() as Writable;
  apply(doc);
  return doc;
}
type Writable = { -readonly [K in keyof DocxDocument]: DocxDocument[K] };
const BAD = "bogus" as never;
const para = (properties: object, runProps?: object): BodyContent =>
  ({
    type: "paragraph",
    properties,
    children: [{ properties: runProps, content: [{ type: "text", text: "x" }] }]
  }) as never;
const tableWith = (t: object, r: object, c: object): BodyContent =>
  ({
    type: "table",
    properties: t,
    rows: [{ properties: r, cells: [{ properties: c, content: [para({})] }] }]
  }) as never;

const BAD_VOCABULARY: [string, (doc: Writable) => void][] = [
  ["body[0].children[0].properties.underline", d => (d.body = [para({}, { underline: BAD })])],
  [
    "body[0].children[0].properties.underline.style",
    d => (d.body = [para({}, { underline: { style: BAD } })])
  ],
  ["body[0].children[0].properties.highlight", d => (d.body = [para({}, { highlight: BAD })])],
  ["body[0].children[0].properties.vertAlign", d => (d.body = [para({}, { vertAlign: BAD })])],
  ["body[0].children[0].properties.effect", d => (d.body = [para({}, { effect: BAD })])],
  [
    "body[0].children[0].properties.emphasisMark",
    d => (d.body = [para({}, { emphasisMark: BAD })])
  ],
  [
    "body[0].children[0].properties.font.hint",
    d => (d.body = [para({}, { font: { ascii: "A", hint: BAD } })])
  ],
  [
    "body[0].children[0].properties.shading.pattern",
    d => (d.body = [para({}, { shading: { fill: "auto", pattern: BAD } })])
  ],
  ["body[0].properties.alignment", d => (d.body = [para({ alignment: BAD })])],
  ["body[0].properties.textAlignment", d => (d.body = [para({ textAlignment: BAD })])],
  ["body[0].properties.textDirection", d => (d.body = [para({ textDirection: BAD })])],
  [
    "body[0].properties.spacing.lineRule",
    d => (d.body = [para({ spacing: { line: 240, lineRule: BAD } })])
  ],
  [
    "body[0].properties.tabs[0].type",
    d => (d.body = [para({ tabs: [{ type: BAD, position: 1 }] })])
  ],
  [
    "body[0].properties.tabs[0].leader",
    d => (d.body = [para({ tabs: [{ type: "left", position: 1, leader: BAD }] })])
  ],
  ["body[0].properties.frame.wrap", d => (d.body = [para({ frame: { wrap: BAD } })])],
  ["body[0].properties.frame.hAnchor", d => (d.body = [para({ frame: { hAnchor: BAD } })])],
  ["body[0].properties.alignment", d => (d.body = [tableWith({ alignment: "both" }, {}, {})])],
  ["body[0].properties.layout", d => (d.body = [tableWith({ layout: BAD }, {}, {})])],
  [
    "body[0].properties.width.type",
    d => (d.body = [tableWith({ width: { value: 1, type: BAD } }, {}, {})])
  ],
  [
    "body[0].rows[0].properties.height.rule",
    d => (d.body = [tableWith({}, { height: { value: 1, rule: BAD } }, {})])
  ],
  [
    "body[0].rows[0].cells[0].properties.verticalAlign",
    d => (d.body = [tableWith({}, {}, { verticalAlign: BAD })])
  ],
  [
    "body[0].rows[0].cells[0].properties.textDirection",
    d => (d.body = [tableWith({}, {}, { textDirection: BAD })])
  ],
  ["sectionProperties.breakType", d => (d.sectionProperties = { breakType: BAD })],
  ["sectionProperties.verticalAlign", d => (d.sectionProperties = { verticalAlign: BAD })],
  ["sectionProperties.docGrid.type", d => (d.sectionProperties = { docGrid: { type: BAD } })],
  [
    "sectionProperties.lineNumbers.restart",
    d => (d.sectionProperties = { lineNumbers: { countBy: 1, restart: BAD } })
  ],
  [
    "sectionProperties.pageBorders.display",
    d => (d.sectionProperties = { pageBorders: { display: BAD } })
  ]
];

/**
 * A writer error names the exact path from the document to the value
 * (`body[0].children[0].properties.underline`).
 */

describe("a hand-built node with a word outside the vocabulary is refused at write", () => {
  it.each(BAD_VOCABULARY)("%s", async (field, apply) => {
    await expect(Io.toBuffer(handBuilt(apply))).rejects.toMatchObject({
      name: "DocxArgumentError",
      field
    });
  });

  it("writes the schema's whole vocabulary, wider than the builder's where it is", async () => {
    const doc = handBuilt(d => {
      d.body = [
        para({ alignment: "thaiDistribute", textDirection: "tbV" }, { underline: "wavyDouble" }),
        tableWith({ alignment: "end" }, {}, { verticalAlign: "both" })
      ];
    });
    const xml = await partText(await Io.toBuffer(doc), "word/document.xml");
    expect(xml).toContain('<w:vAlign w:val="both"/>');
    expect(xml).toContain('<w:textDirection w:val="tbV"/>');
  });
});

// -----------------------------------------------------------------------------
// Vocabulary at read
// -----------------------------------------------------------------------------

/** A package written from `doc`, with `edits` applied to `part` (`word/document.xml`) afterwards. */
async function editedPackage(
  doc: Doc,
  edits: [string, string][],
  part = "word/document.xml"
): Promise<Uint8Array> {
  const entries = new Map<string, Uint8Array>();
  for await (const entry of unzip(await Io.toBuffer(doc)).entries()) {
    entries.set(entry.path, await entry.bytes());
  }
  let xml = new TextDecoder().decode(entries.get(part)!);
  for (const [from, to] of edits) {
    expect(xml).toContain(from);
    xml = xml.replace(from, to);
  }
  entries.set(part, new TextEncoder().encode(xml));
  const archive = zip();
  for (const [path, data] of entries) {
    archive.add(path, data);
  }
  return archive.toBuffer();
}

describe("a file's word outside the vocabulary is kept as read, not dropped or refused", () => {
  const EDITS: [string, string][] = [
    ['<w:u w:val="single"/>', '<w:u w:val="evil" w:color="FF0000"/>'],
    ['<w:jc w:val="center"/>', '<w:jc w:val="bogus"/>'],
    ['<w:highlight w:val="yellow"/>', '<w:highlight w:val="mauve"/>'],
    ['<w:vertAlign w:val="superscript"/>', '<w:vertAlign w:val="up"/>'],
    ['<w:textDirection w:val="lrTb"/>', '<w:textDirection w:val="sideways"/>'],
    ['<w:jc w:val="right"/>', '<w:jc w:val="both"/>'],
    ['<w:tblLayout w:type="fixed"/>', '<w:tblLayout w:type="loose"/>'],
    ['<w:vAlign w:val="center"/>', '<w:vAlign w:val="middle"/>'],
    ['<w:type w:val="oddPage"/>', '<w:type w:val="sometimes"/>'],
    ['<w:tab w:val="right"', '<w:tab w:val="rightish"']
  ];

  async function source(): Promise<Uint8Array> {
    const doc = Document.create();
    Document.addParagraph(doc, "x", {
      alignment: "center",
      textDirection: "lrTb",
      tabs: [{ type: "right", position: 720 }],
      run: { underline: "single", highlight: "yellow", vertAlign: "superscript" }
    });
    Document.addTable(
      doc,
      Build.table([Build.row([Build.cell("c", { verticalAlign: "center" })])], {
        alignment: "right",
        layout: "fixed"
      })
    );
    Document.setSectionProperties(doc, { breakType: "oddPage" });
    return editedPackage(doc, EDITS);
  }

  it("reads each as an absent field, and writes the document back with every one of them", async () => {
    const doc = await Io.read(await source());
    const p = doc.body.find(b => b.type === "paragraph") as Paragraph;
    const r = p.children[0] as Run;
    expect(p.properties?.alignment).toBeUndefined();
    expect(p.properties?.textDirection).toBeUndefined();
    // A tab stop always has a type; it reads as the schema's default.
    expect(p.properties?.tabs?.[0]?.type).toBe("left");
    expect(r.properties?.underline).toBeUndefined();
    expect(r.properties?.highlight).toBeUndefined();
    expect(r.properties?.vertAlign).toBeUndefined();
    const t = doc.body.find(b => b.type === "table") as Table;
    expect(t.properties?.alignment).toBeUndefined();
    expect(t.properties?.layout).toBeUndefined();
    expect(t.rows[0]!.cells[0]!.properties?.verticalAlign).toBeUndefined();
    expect(doc.sectionProperties?.breakType).toBeUndefined();

    const xml = await partText(await Io.toBuffer(doc), "word/document.xml");
    for (const [, kept] of EDITS) {
      expect(xml).toContain(kept);
    }
  });

  it("reports each through Validation.validate, with the path to its field", async () => {
    const doc = await Io.read(await source());
    const kept = Validation.validate(doc).issues.filter(i => i.rule === "unrecognized-value");
    expect(kept.every(i => i.severity === "warning")).toBe(true);
    expect(kept.map(i => i.message.split(" ")[0]).sort()).toEqual(
      [
        '"bogus"',
        '"both"',
        '"evil"',
        '"loose"',
        '"mauve"',
        '"middle"',
        '"rightish"',
        '"sometimes"',
        '"sideways"',
        '"up"'
      ].sort()
    );
    expect(kept.map(i => i.path)).toContain("body[0].properties.alignment");
    expect(kept.map(i => i.path)).toContain("sectionProperties.breakType");
  });

  it("gives way to a value set afterwards, which is checked like any other", async () => {
    const doc = await Io.read(await source());
    const p = doc.body.find(b => b.type === "paragraph") as Paragraph;
    const r = p.children[0] as Run;
    Object.assign(p.properties!, { alignment: "end" });
    Object.assign(p.properties!.tabs![0]!, { type: "center" });
    Object.assign(r.properties!, { underline: "double" });
    const xml = await partText(await Io.toBuffer(doc), "word/document.xml");
    expect(xml).toContain('<w:jc w:val="end"/>');
    expect(xml).toContain('<w:tab w:val="center"');
    expect(xml).toContain('<w:u w:val="double"/>');
    expect(xml).not.toContain("bogus");
    expect(xml).not.toContain("rightish");
    expect(xml).not.toContain("evil");
    expect(Validation.validate(doc).issues.map(i => i.path)).not.toContain(
      "body[0].properties.alignment"
    );

    Object.assign(p.properties!, { alignment: "bogus" });
    await expect(Io.toBuffer(doc)).rejects.toMatchObject({
      name: "DocxArgumentError",
      field: expect.stringMatching(/^body\[\d+\]\.properties\.alignment$/)
    });
  });

  it("comes back when the field is cleared again, and goes with Document.discardUnrecognizedValues", async () => {
    const doc = await Io.read(await source());
    const p = doc.body.find(b => b.type === "paragraph") as Paragraph;
    Object.assign(p.properties!, { alignment: "end" });
    Object.assign(p.properties!, { alignment: undefined });
    expect(await partText(await Io.toBuffer(doc), "word/document.xml")).toContain(
      '<w:jc w:val="bogus"/>'
    );

    expect(Document.discardUnrecognizedValues(doc)).toBe(EDITS.length);
    expect(Validation.validate(doc).issues.filter(i => i.rule === "unrecognized-value")).toEqual(
      []
    );
    const xml = await partText(await Io.toBuffer(doc), "word/document.xml");
    for (const [, kept] of EDITS) {
      expect(xml).not.toContain(kept);
    }
    // A tab stop cannot be absent: it is written as what it reads as.
    expect(xml).toContain('<w:tab w:val="left"');
    expect(Document.discardUnrecognizedValues(doc)).toBe(0);
  });

  it("takes its options as the flat DiscardUnrecognizedValuesOptions type, as Excel does", async () => {
    const doc = await Io.read(await source());
    const options: DiscardUnrecognizedValuesOptions = { filter: () => false };
    expect(Document.discardUnrecognizedValues(doc, options)).toBe(0);
  });

  it("discards only the values a filter picks, by the path validate reports", async () => {
    const doc = await Io.read(await source());
    const kept = () =>
      Validation.validate(doc)
        .issues.filter(i => i.rule === "unrecognized-value")
        .map(i => i.path);
    const before = kept();
    expect(before).toContain("body[0].properties.alignment");
    expect(
      Document.discardUnrecognizedValues(doc, {
        filter: v => v.path === "body[0].properties.alignment" && v.value === "bogus"
      })
    ).toBe(1);
    expect(kept()).toEqual(before.filter(p => p !== "body[0].properties.alignment"));
    const xml = await partText(await Io.toBuffer(doc), "word/document.xml");
    expect(xml).not.toContain('w:val="bogus"');
    expect(xml).toContain('w:val="mauve"');
    expect(() => Document.discardUnrecognizedValues(doc, { filter: 1 as never })).toThrow(
      /options\.filter/
    );
  });

  it("a filtered discard also drops a value a field has taken over from, so it cannot come back", async () => {
    const doc = await Io.read(await source());
    const p = doc.body.find(b => b.type === "paragraph") as Paragraph;
    Object.assign(p.properties!, { alignment: "end" });
    expect(
      Document.discardUnrecognizedValues(doc, { filter: v => v.path.endsWith(".alignment") })
    ).toBe(1);
    Object.assign(p.properties!, { alignment: undefined });
    expect(await partText(await Io.toBuffer(doc), "word/document.xml")).not.toContain("bogus");
  });

  it("is found in a header too", async () => {
    const doc = Document.create();
    Document.setHeader(doc, "default", {
      children: [Build.textParagraph("h", { alignment: "center" })]
    });
    const read = await Io.read(
      await editedPackage(
        doc,
        [['<w:jc w:val="center"/>', '<w:jc w:val="bogus"/>']],
        "word/header1.xml"
      )
    );
    const kept = Validation.validate(read).issues.filter(i => i.rule === "unrecognized-value");
    // A read document keys its headers by relationship id.
    expect(kept.map(i => i.path)).toEqual([
      expect.stringMatching(/^headers\["rId\d+"\]\.content\.children\[0\]\.properties\.alignment$/)
    ]);
    expect(Document.discardUnrecognizedValues(read)).toBe(1);
    expect(await partText(await Io.toBuffer(read), "word/header1.xml")).not.toContain("bogus");
  });

  it("survives a copy of the document, and a spread of the properties it sits on", async () => {
    const doc = Document.clone(await Io.read(await source()));
    const p = doc.body.find(b => b.type === "paragraph") as Paragraph;
    const copy = { ...p, properties: { ...p.properties, keepNext: true } };
    (doc as { body: BodyContent[] }).body = [copy, ...doc.body.slice(1)];
    const xml = await partText(await Io.toBuffer(doc), "word/document.xml");
    expect(xml).toContain('<w:jc w:val="bogus"/>');
  });

  it("keeps every word the schema defines", async () => {
    const source = Document.create();
    Document.addParagraph(source, "x", {
      alignment: "thaiDistribute",
      run: { underline: "dashLongHeavy" }
    });
    const doc = await Io.read(await Io.toBuffer(source));
    const p = doc.body[0] as Paragraph;
    expect(p.properties?.alignment).toBe("thaiDistribute");
    expect((p.children[0] as Run).properties?.underline).toBe("dashLongHeavy");
  });

  it("the same words from a caller are still refused at write", async () => {
    for (const [node, field] of [
      [para({}, { underline: "evil" }), "body[0].children[0].properties.underline"],
      [para({ alignment: "bogus" }), "body[0].properties.alignment"]
    ] as const) {
      const doc = handBuilt(d => (d.body = [node]));
      await expect(Io.toBuffer(doc)).rejects.toMatchObject({
        name: "DocxArgumentError",
        field
      });
    }
  });
});
