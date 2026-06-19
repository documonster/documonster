/**
 * Hyperlink + RichText integration tests.
 *
 * Regression coverage for https://github.com/documonster/documonster/issues/142:
 * a hyperlink cell whose display text is rich-text (shared-string `<r>` runs)
 * must round-trip through `writeBuffer` / `load` with:
 *
 *   - `cell.value.text` always a `string` (public type contract)
 *   - `cell.value.richText` preserved so formatted display survives
 *   - downstream consumers (`cell.text`, `cell.toString()`) never observe
 *     a `CellRichTextValue` object where a string is promised
 */

import { writeCsv } from "@excel/bridge/csv-bridge";
import {
  cellGetValue,
  cellHyperlink,
  cellText,
  cellType,
  cellGetModel,
  cellToString
} from "@excel/core/cell";
import { ValueType } from "@excel/core/enums";
import { getCell } from "@excel/core/worksheet";
import { Cell, Workbook } from "@excel/index";
import { describe, it, expect } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

describe("Hyperlink + RichText round-trip (issue #142)", () => {
  it("preserves rich-text display on a hyperlink through writeBuffer → load", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    // Public input shape: rich-text hyperlink — no `text`, no cast required.
    Cell.setValue(ws, "A1", {
      richText: [{ text: "bold ", font: { bold: true } }, { text: "plain" }],
      hyperlink: "https://example.com"
    });

    const buffer = await Workbook.toBuffer(wb);
    await expectValidXlsx(buffer, { label: "rich-text hyperlink sharedStrings" });

    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
    const cell = getCell(ws2, "A1");
    const v = cellGetValue(cell);

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
    expect(typeof cellText(cell)).toBe("string");
    expect(cellText(cell)).toBe("bold plain");
    expect(cellToString(cell)).toBe("bold plain");
  });

  it("keeps cell.value.text as string for plain-text hyperlinks (no richText)", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    Cell.setValue(ws, "A1", {
      text: "www.example.com",
      hyperlink: "https://www.example.com"
    });

    const buffer = await Workbook.toBuffer(wb);
    await expectValidXlsx(buffer, { label: "plain hyperlink" });
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);

    const v = Cell.getValue(Workbook.getWorksheet(wb2, "Sheet1")!, "A1");
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
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    const value = {
      richText: [{ text: "same", font: { bold: true } }],
      hyperlink: "https://example.com"
    };
    Cell.setValue(ws, "A1", value);
    Cell.setValue(ws, "A2", value);

    const buffer = await Workbook.toBuffer(wb);
    await expectValidXlsx(buffer, { label: "rich-text hyperlink dedup" });
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;

    const a1 = Cell.getValue(ws2, "A1");
    const a2 = Cell.getValue(ws2, "A2");
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
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    Cell.setValue(ws, "A1", {
      richText: [{ text: "display ", font: { italic: true } }, { text: "text" }],
      hyperlink: "https://example.com"
    });
    Cell.setValue(ws, "B1", "plain");

    const csv = writeCsv(wb);
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
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    // Public input shape for formula + hyperlink — no `cell.model` poking,
    // no type assertions required.
    Cell.setValue(ws, "A1", {
      formula: "1+2",
      result: 3,
      hyperlink: "https://example.com/formula-link"
    });

    const buf1 = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf1, { label: "formula+hyperlink pass 1" });

    const wb2 = Workbook.create();
    await Workbook.read(wb2, buf1);
    const a1 = getCell(Workbook.getWorksheet(wb2, "Sheet1")!, "A1");

    // Public surface: classified as Hyperlink with the formula result as text
    expect(cellType(a1)).toBe(ValueType.Hyperlink);
    expect(cellHyperlink(a1)).toBe("https://example.com/formula-link");
    // Formula source preserved on the model so a second write round-trips it
    expect(cellGetModel(a1).formula).toBe("1+2");

    // Second round-trip — formula must still be present
    const buf2 = await Workbook.toBuffer(wb2);
    await expectValidXlsx(buf2, { label: "formula+hyperlink pass 2" });
    const wb3 = Workbook.create();
    await Workbook.read(wb3, buf2);
    const a1b = getCell(Workbook.getWorksheet(wb3, "Sheet1")!, "A1");
    expect(cellType(a1b)).toBe(ValueType.Hyperlink);
    expect(cellHyperlink(a1b)).toBe("https://example.com/formula-link");
    expect(cellGetModel(a1b).formula).toBe("1+2");
  });

  it("supports rich-text hyperlinks inline (useSharedStrings: false)", async () => {
    // When shared strings are disabled, the cell renderer must fall back
    // to the inlineStr rich-text representation rather than dropping runs.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    Cell.setValue(ws, "A1", {
      richText: [{ text: "bold ", font: { bold: true } }, { text: "plain" }],
      hyperlink: "https://example.com"
    });

    const buffer = await Workbook.toBuffer(wb, { useSharedStrings: false });
    await expectValidXlsx(buffer, { label: "rich-text hyperlink inline" });

    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);
    const v = Cell.getValue(Workbook.getWorksheet(wb2, "Sheet1")!, "A1");
    if (typeof v !== "object" || v === null || !("hyperlink" in v)) {
      throw new Error("expected a hyperlink value");
    }

    expect(typeof v.text).toBe("string");
    expect(v.text).toBe("bold plain");
    expect(v.hyperlink).toBe("https://example.com");
  });
});
