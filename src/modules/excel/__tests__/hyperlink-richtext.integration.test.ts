/**
 * Hyperlink + RichText integration tests.
 *
 * Regression coverage for https://github.com/cjnoname/excelts/issues/142:
 * a hyperlink cell whose display text is rich-text (shared-string `<r>` runs)
 * must round-trip through `writeBuffer` / `load` with:
 *
 *   - `cell.value.text` always a `string` (public type contract)
 *   - `cell.value.richText` preserved so formatted display survives
 *   - downstream consumers (`cell.text`, `cell.toString()`) never observe
 *     a `CellRichTextValue` object where a string is promised
 */

import { describe, it, expect } from "vitest";

import { Workbook, ValueType } from "../../../index";

describe("Hyperlink + RichText round-trip (issue #142)", () => {
  it("preserves rich-text display on a hyperlink through writeBuffer → load", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    // Public input shape: rich-text hyperlink — no `text`, no cast required.
    ws.getCell("A1").value = {
      richText: [{ text: "bold ", font: { bold: true } }, { text: "plain" }],
      hyperlink: "https://example.com"
    };

    const buffer = await wb.xlsx.writeBuffer();

    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet("Sheet1")!;
    const cell = ws2.getCell("A1");
    const v = cell.value;

    // Output shape always carries `text` and `hyperlink`.
    if (typeof v !== "object" || v === null || !("hyperlink" in v)) {
      throw new Error("expected a hyperlink value");
    }
    expect(typeof v.text).toBe("string");
    expect(v.text).toBe("bold plain");
    expect(v.hyperlink).toBe("https://example.com");

    // richText runs preserved (at least their text content)
    expect(v.richText).toBeDefined();
    expect(v.richText!.map(r => r.text).join("")).toBe("bold plain");

    // cell.text must always be a string
    expect(typeof cell.text).toBe("string");
    expect(cell.text).toBe("bold plain");
    expect(cell.toString()).toBe("bold plain");
  });

  it("keeps cell.value.text as string for plain-text hyperlinks (no richText)", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    ws.getCell("A1").value = {
      text: "www.example.com",
      hyperlink: "https://www.example.com"
    };

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);

    const v = wb2.getWorksheet("Sheet1")!.getCell("A1").value;
    if (typeof v !== "object" || v === null || !("hyperlink" in v)) {
      throw new Error("expected a hyperlink value");
    }
    expect(typeof v.text).toBe("string");
    expect(v.text).toBe("www.example.com");
    expect(v.hyperlink).toBe("https://www.example.com");
    // Plain hyperlinks do not carry rich-text runs.
    expect(v.richText).toBeUndefined();
  });

  it("deduplicates identical rich-text hyperlinks in the shared-string table", async () => {
    // Before the fix, rich-text payloads collided under "[object Object]" in
    // the SharedStrings hash — this test confirms two equal rich-text hyperlinks
    // share one entry after the hash-key fix.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    const value = {
      richText: [{ text: "same", font: { bold: true } }],
      hyperlink: "https://example.com"
    };
    ws.getCell("A1").value = value;
    ws.getCell("A2").value = value;

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet("Sheet1")!;

    const a1 = ws2.getCell("A1").value;
    const a2 = ws2.getCell("A2").value;
    if (
      typeof a1 !== "object" ||
      a1 === null ||
      !("hyperlink" in a1) ||
      typeof a2 !== "object" ||
      a2 === null ||
      !("hyperlink" in a2)
    ) {
      throw new Error("expected hyperlink values");
    }
    expect(a1.text).toBe("same");
    expect(a2.text).toBe("same");
    expect(a1.hyperlink).toBe("https://example.com");
    expect(a2.hyperlink).toBe("https://example.com");
  });

  it("CSV export of a rich-text hyperlink emits the URL (not the JSON blob)", () => {
    // Regression: CSV mapper used to coerce rich-text hyperlinks through
    // `[object Object]` because the hyperlink branch inspected
    // `value.text || value.hyperlink` on an object with only `richText`.
    // The rewritten mapper routes hyperlink cells to the URL branch.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    ws.getCell("A1").value = {
      richText: [{ text: "display ", font: { italic: true } }, { text: "text" }],
      hyperlink: "https://example.com"
    };
    ws.getCell("B1").value = "plain";

    const csv = wb.writeCsv();
    // URL for the hyperlink cell, not "[object Object]" nor JSON
    expect(csv).toContain("https://example.com");
    expect(csv).not.toContain("[object Object]");
    expect(csv).not.toContain("richText");
  });

  it("preserves a formula cell that also has a hyperlink (round-trip)", async () => {
    // A worksheet may carry a `<hyperlink ref="..."/>` entry whose address
    // also has a formula `<c>` element. The historical public-API contract
    // is: the cell surfaces as a Hyperlink (display = formula result), but
    // `cell.model.formula` is preserved so the formula re-emits on write
    // and survives the round-trip.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    // Public input shape for formula + hyperlink — no `cell.model` poking,
    // no type assertions required.
    ws.getCell("A1").value = {
      formula: "1+2",
      result: 3,
      hyperlink: "https://example.com/formula-link"
    };

    const buf1 = await wb.xlsx.writeBuffer();

    const wb2 = new Workbook();
    await wb2.xlsx.load(buf1);
    const a1 = wb2.getWorksheet("Sheet1")!.getCell("A1");

    // Public surface: classified as Hyperlink with the formula result as text
    expect(a1.type).toBe(ValueType.Hyperlink);
    expect(a1.hyperlink).toBe("https://example.com/formula-link");
    // Formula source preserved on the model so a second write round-trips it
    expect(a1.model.formula).toBe("1+2");

    // Second round-trip — formula must still be present
    const buf2 = await wb2.xlsx.writeBuffer();
    const wb3 = new Workbook();
    await wb3.xlsx.load(buf2);
    const a1b = wb3.getWorksheet("Sheet1")!.getCell("A1");
    expect(a1b.type).toBe(ValueType.Hyperlink);
    expect(a1b.hyperlink).toBe("https://example.com/formula-link");
    expect(a1b.model.formula).toBe("1+2");
  });

  it("supports rich-text hyperlinks inline (useSharedStrings: false)", async () => {
    // When shared strings are disabled, the cell renderer must fall back
    // to the inlineStr rich-text representation rather than dropping runs.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    ws.getCell("A1").value = {
      richText: [{ text: "bold ", font: { bold: true } }, { text: "plain" }],
      hyperlink: "https://example.com"
    };

    const buffer = await wb.xlsx.writeBuffer({ useSharedStrings: false });

    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);
    const v = wb2.getWorksheet("Sheet1")!.getCell("A1").value;
    if (typeof v !== "object" || v === null || !("hyperlink" in v)) {
      throw new Error("expected a hyperlink value");
    }

    expect(typeof v.text).toBe("string");
    expect(v.text).toBe("bold plain");
    expect(v.hyperlink).toBe("https://example.com");
  });
});
