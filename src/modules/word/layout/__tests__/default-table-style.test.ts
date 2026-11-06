import { Build, Document, Io, Validation } from "@word/index";
import { layoutDocumentFull } from "@word/layout/layout-full";
import type { LayoutTable } from "@word/layout/layout-model";
import type { DocxDocument, StyleDef } from "@word/types";
/**
 * A table that names no style takes the document's default table style — the
 * one marked `w:default="1"` — and its cell margins. Word's `Normal Table`
 * carries 0.08 inch left and right; without a default table style Word draws
 * the cells flush (Word for Mac 16.114 on N2: 20 pages with the default, 17
 * without). Layout assumed 5.4pt either way, and the writer emitted
 * `TableNormal` without `w:default`, so Word and layout never agreed.
 */
import { describe, expect, it } from "vitest";

function withTable(useDefaults: boolean): DocxDocument {
  const doc = Document.create();
  if (useDefaults) {
    Document.useDefaultStyles(doc);
  }
  Document.addTable(doc, Build.simpleTable([["cell"]], { borders: false }));
  return doc;
}

function firstCellInset(doc: DocxDocument): number {
  const page = layoutDocumentFull(doc).pages[0]!;
  const table = page.content.find((c): c is LayoutTable => c.type === "table")!;
  const cell = table.cells[0]!;
  return cell.content[0]!.rect.x - cell.rect.x;
}

describe("default table style", () => {
  it("writes Word's default styles with w:default, as Word does", async () => {
    const doc = withTable(true);
    const read = await Io.read(await Io.toBuffer(doc));
    const defaults = (read.styles ?? [])
      .filter(s => s.isDefault)
      .map(s => `${s.type}:${s.styleId}`)
      .sort();
    expect(defaults).toEqual([
      "character:DefaultParagraphFont",
      "numbering:NoList",
      "paragraph:Normal",
      "table:TableNormal"
    ]);
  });

  it("makes a referenced TableNormal the default table style", async () => {
    const doc = Document.create();
    Document.addTable(doc, Build.table([Build.row([Build.cell("x")])], { style: "TableGrid" }));
    const read = await Io.read(await Io.toBuffer(doc));
    expect(read.styles?.find(s => s.styleId === "TableNormal")?.isDefault).toBe(true);
  });

  it("keeps a document's own default table style the only one", async () => {
    const doc = Document.create();
    Document.addStyle(doc, { type: "table", styleId: "Mine", name: "Mine", isDefault: true });
    Document.addTable(doc, Build.table([Build.row([Build.cell("x")])], { style: "TableGrid" }));
    const read = await Io.read(await Io.toBuffer(doc));
    expect(
      (read.styles ?? []).filter(s => s.type === "table" && s.isDefault).map(s => s.styleId)
    ).toEqual(["Mine"]);
  });

  it("insets cells by the default table style's margins", () => {
    expect(firstCellInset(withTable(true))).toBeCloseTo(5.4, 6);
  });

  it("draws cells flush when the document has no default table style", () => {
    expect(firstCellInset(withTable(false))).toBeCloseTo(0, 6);
  });

  it("lets the table's and the cell's own margins override the style's", () => {
    const doc = withTable(true);
    const table = doc.body[0] as { properties?: object };
    table.properties = {
      ...table.properties,
      cellMargins: { left: { value: 40, type: "dxa" } }
    };
    expect(firstCellInset(doc)).toBeCloseTo(2, 6);
  });

  /** A document as read from disk: two table styles both marked `w:default`. */
  function twoDefaults(): DocxDocument {
    const doc = withTable(true);
    // Appended after the built-in `TableNormal`, flag intact — what a file
    // declaring two defaults reads as.
    (doc.styles as StyleDef[]).push({
      type: "table",
      styleId: "Wide",
      name: "Wide",
      isDefault: true,
      tableProperties: {
        cellMargins: { left: { value: 400, type: "dxa" }, right: { value: 400, type: "dxa" } }
      }
    });
    return doc;
  }

  it("uses the last of two default table styles, as Word does", () => {
    // Word for Mac 16.114 on final5 T1 (TableNormal and DefTbl both default):
    // the tables take DefTbl's pPr/rPr, 7 pages against 8 with TableNormal.
    expect(firstCellInset(twoDefaults())).toBeCloseTo(20, 6);
  });

  it("reports a second default style of one type", () => {
    const rules = Validation.validate(twoDefaults()).issues.map(i => i.rule);
    expect(rules).toContain("style-dup-default");
    expect(Validation.validate(withTable(true)).issues.map(i => i.rule)).not.toContain(
      "style-dup-default"
    );
  });

  it("makes a style added as default the only default of its type", () => {
    const doc = withTable(true);
    Document.addStyle(doc, {
      type: "table",
      styleId: "Wide",
      name: "Wide",
      isDefault: true,
      tableProperties: {
        cellMargins: { left: { value: 400, type: "dxa" }, right: { value: 400, type: "dxa" } }
      }
    });
    expect(
      (doc.styles ?? []).filter(s => s.type === "table" && s.isDefault).map(s => s.styleId)
    ).toEqual(["Wide"]);
    // Other types keep theirs.
    expect(doc.styles?.find(s => s.styleId === "Normal")?.isDefault).toBe(true);
    expect(firstCellInset(doc)).toBeCloseTo(20, 6);
  });
});
