import type { CellData } from "@excel/core/cell";
import {
  cellSetValue,
  cellEffectiveType,
  cellFont,
  cellFormula,
  cellGetValue,
  cellHyperlink,
  cellIsMerged,
  cellMaster,
  cellResult,
  cellSetFont,
  cellCreate,
  cellSetNote,
  cellSetNumFmt,
  cellText,
  cellType,
  cellMerge,
  cellUnmerge,
  cellIsMergedTo,
  _cellUpgradeToHyperlink,
  cellToCsvString,
  cellToString,
  cellAbsoluteAddress,
  cellDisplayText,
  cellFormulaType,
  cellFullAddress,
  cellHtml,
  cellIsHyperlink,
  cellGetModel,
  cellDestroy
} from "@excel/core/cell";
import type { ColumnData } from "@excel/core/column";
import { Enums } from "@excel/core/enums";
import type { RowData } from "@excel/core/row";
import { rowSetFont } from "@excel/core/row";
import { findRow, getCell, getRow, rowGetCell } from "@excel/core/worksheet";
import { findCell, getColumn } from "@excel/core/worksheet-core";
import { Cell, Workbook, Worksheet } from "@excel/index";
import type { CellHyperlinkValue } from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import { describe, it, expect, beforeEach } from "vitest";

// A thin facade over a real worksheet record so the existing test bodies
// (`sheetMock.getCell("A1")`, `.getColumn(n)`, `.getRow(n)`, `.reset()`) keep
// working but operate on genuine `CellData` / `RowData` / `ColumnData` handles.
interface SheetMock {
  ws: Worksheet.Handle;
  reset(): void;
  findRow(num: number): RowData | undefined;
  getRow(num: number): RowData;
  getColumn(num: number): ColumnData;
  createColumn(num: number): ColumnData;
  getCell(address: string): CellData;
  findCell(address: string): CellData | undefined;
}

const sheetMock: SheetMock = {
  ws: Workbook.addWorksheet(Workbook.create(), "mock"),
  reset() {
    this.ws = Workbook.addWorksheet(Workbook.create(), "mock");
  },
  findRow(num) {
    return findRow(this.ws, num);
  },
  getRow(num) {
    return getRow(this.ws, num);
  },
  getColumn(num) {
    return getColumn(this.ws, num);
  },
  createColumn(num) {
    return getColumn(this.ws, num);
  },
  getCell(address) {
    const fullAddress = colCache.decodeAddress(address);
    return getCell(this.ws, fullAddress.row, fullAddress.col);
  },
  findCell(address) {
    const fullAddress = colCache.decodeAddress(address);
    return findCell(this.ws, fullAddress.row, fullAddress.col);
  }
};

const fonts = {
  arialBlackUI14: {
    name: "Arial Black",
    family: 2,
    size: 14,
    underline: true,
    italic: true
  },
  comicSansUdB16: {
    name: "Comic Sans MS",
    family: 4,
    size: 16,
    underline: "double",
    bold: true
  },
  broadwayRedOutline20: {
    name: "Broadway",
    family: 5,
    size: 20,
    outline: true,
    color: { argb: "FFFF0000" }
  }
};

describe("Cell", () => {
  beforeEach(() => {
    sheetMock.reset();
  });
  it("stores values", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");

    expect(a1.address).toBe("A1");
    expect(cellAbsoluteAddress(a1)).toBe("$A$1");

    expect(cellType(a1)).toBe(Enums.ValueType.Null);

    cellSetValue(a1, 5);
    expect(cellGetValue(a1)).toBe(5);
    expect(cellType(a1)).toBe(Enums.ValueType.Number);

    const strValue = "Hello, World!";
    cellSetValue(a1, strValue);
    expect(cellGetValue(a1)).toBe(strValue);
    expect(cellType(a1)).toBe(Enums.ValueType.String);

    const dateValue = new Date();
    cellSetValue(a1, dateValue);
    expect(cellGetValue(a1)).toBe(dateValue);
    expect(cellType(a1)).toBe(Enums.ValueType.Date);

    const formulaValue = { formula: "A2", result: 5 };
    cellSetValue(a1, formulaValue);
    expect(cellGetValue(a1)).toEqual(formulaValue);
    expect(cellType(a1)).toBe(Enums.ValueType.Formula);

    const hyperlinkValue = {
      hyperlink: "http://www.link.com",
      text: "www.link.com"
    };
    cellSetValue(a1, hyperlinkValue);
    expect(cellGetValue(a1)).toEqual(hyperlinkValue);
    expect(cellType(a1)).toBe(Enums.ValueType.Hyperlink);

    cellSetValue(a1, null);
    expect(cellType(a1)).toBe(Enums.ValueType.Null);

    cellSetValue(a1, { json: "data" } as never);
    expect(cellType(a1)).toBe(Enums.ValueType.String);
  });
  it("validates options on construction", () => {
    const row = sheetMock.getRow(1);
    const column = sheetMock.getColumn(1);

    expect(() => {
      cellCreate(null as any, null as any, "A1");
    }).toThrow(Error);
    expect(() => {
      cellCreate(row, null as any, "A1");
    }).toThrow(Error);
    expect(() => {
      cellCreate(null as any, column, "A1");
    }).toThrow(Error);
  });
  it("merges", () => {
    const a1 = sheetMock.getCell("A1");
    const a2 = sheetMock.getCell("A2");

    cellSetValue(a1, 5);
    cellSetValue(a2, "Hello, World!");

    cellMerge(a2, a1);

    expect(cellGetValue(a2)).toBe(5);
    expect(cellType(a2)).toBe(Enums.ValueType.Merge);
    expect(a1._mergeCount).toBe(1);
    expect(cellIsMerged(a1)).toBeTruthy();
    expect(cellIsMerged(a2)).toBeTruthy();
    expect(cellIsMergedTo(a2, a1)).toBeTruthy();
    expect(cellMaster(a2)).toBe(a1);
    expect(cellMaster(a1)).toBe(a1);

    // assignment of slaves write to the master
    cellSetValue(a2, 7);
    expect(cellGetValue(a1)).toBe(7);

    // assignment of strings should add 1 ref
    const strValue = "Boo!";
    cellSetValue(a2, strValue);
    expect(cellGetValue(a1)).toBe(strValue);

    // unmerge should work also
    cellUnmerge(a2);
    expect(cellType(a2)).toBe(Enums.ValueType.Null);
    expect(a1._mergeCount).toBe(0);
    expect(cellIsMerged(a1)).not.toBeTruthy();
    expect(cellIsMerged(a2)).not.toBeTruthy();
    expect(cellIsMergedTo(a2, a1)).not.toBeTruthy();
    expect(cellMaster(a2)).toBe(a2);
    expect(cellMaster(a1)).toBe(a1);
  });

  it("upgrades from string to hyperlink", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");

    const strValue = "www.link.com";
    const linkValue = "http://www.link.com";

    cellSetValue(a1, strValue);

    _cellUpgradeToHyperlink(a1, linkValue);

    expect(cellType(a1)).toBe(Enums.ValueType.Hyperlink);
  });

  it("upgrades from rich text to hyperlink, preserving runs (issue #142)", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");
    const linkValue = "http://www.link.com";

    cellSetValue(a1, {
      richText: [{ text: "bold", font: { bold: true } }, { text: "-plain" }]
    });

    _cellUpgradeToHyperlink(a1, linkValue);

    expect(cellType(a1)).toBe(Enums.ValueType.Hyperlink);
    // CellHyperlinkValue.text contract: always a string
    const v = cellGetValue(a1) as CellHyperlinkValue;
    expect(typeof v.text).toBe("string");
    expect(v.text).toBe("bold-plain");
    expect(v.hyperlink).toBe(linkValue);
    expect(v.richText).toEqual([{ text: "bold", font: { bold: true } }, { text: "-plain" }]);
  });

  it("hyperlink value setter normalizes rich text into plain-text mirror (issue #142)", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");

    cellSetValue(a1, {
      richText: [{ text: "hello " }, { text: "world", font: { italic: true } }],
      hyperlink: "https://example.com",
      // Intentionally inconsistent: user-supplied text should be ignored
      // when richText is present, to keep the invariant text===flatten(richText).
      text: "ignored"
    });

    expect(cellType(a1)).toBe(Enums.ValueType.Hyperlink);
    const v = cellGetValue(a1) as CellHyperlinkValue;
    expect(v.text).toBe("hello world");
    expect(typeof v.text).toBe("string");
    expect(v.hyperlink).toBe("https://example.com");
    expect(v.richText).toEqual([{ text: "hello " }, { text: "world", font: { italic: true } }]);
  });

  it("hyperlink cell.text always returns a string, even for rich-text link (issue #142)", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");
    cellSetValue(a1, {
      richText: [{ text: "abc" }, { text: "def" }],
      hyperlink: "https://example.com"
    });

    expect(typeof cellText(a1)).toBe("string");
    expect(cellText(a1)).toBe("abcdef");
    expect(cellToString(a1)).toBe("abcdef");
  });

  it("does not classify { richText: [] } as a RichText cell", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");
    // An empty richText array carries no content. It must not produce a
    // RichText cell with no runs (which would later flatten to "").
    cellSetValue(a1, { richText: [] } as unknown as CellHyperlinkValue);
    expect(cellType(a1)).not.toBe(Enums.ValueType.RichText);
  });

  it("does not classify { richText: [], hyperlink } as a Hyperlink cell", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");
    // Empty richText and no plain text => not a valid hyperlink display.
    cellSetValue(a1, {
      richText: [],
      hyperlink: "https://example.com"
    } as unknown as CellHyperlinkValue);
    expect(cellType(a1)).not.toBe(Enums.ValueType.Hyperlink);
  });

  it("doesn't upgrade from non-string to hyperlink", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");

    const linkValue = "http://www.link.com";

    // null
    _cellUpgradeToHyperlink(a1, linkValue);
    expect(cellType(a1)).toBe(Enums.ValueType.Null);

    // number
    cellSetValue(a1, 5);
    _cellUpgradeToHyperlink(a1, linkValue);
    expect(cellType(a1)).toBe(Enums.ValueType.Number);

    // date
    cellSetValue(a1, new Date());
    _cellUpgradeToHyperlink(a1, linkValue);
    expect(cellType(a1)).toBe(Enums.ValueType.Date);

    // formula
    cellSetValue(a1, { formula: "A2" });
    _cellUpgradeToHyperlink(a1, linkValue);
    expect(cellType(a1)).toBe(Enums.ValueType.Formula);

    // hyperlink
    cellSetValue(a1, { hyperlink: "http://www.link2.com", text: "www.link2.com" });
    _cellUpgradeToHyperlink(a1, linkValue);
    expect(cellType(a1)).toEqual(Enums.ValueType.Hyperlink);

    // cleanup
    cellSetValue(a1, null);
  });

  it("inherits column styles", () => {
    sheetMock.getRow(1);
    const column = sheetMock.getColumn(1);

    column.style = {
      font: fonts.arialBlackUI14
    };

    const a1 = sheetMock.getCell("A1");
    expect(cellFont(a1)).toEqual(fonts.arialBlackUI14);
  });

  it("inherits row styles", () => {
    const row = sheetMock.getRow(1);
    sheetMock.getColumn(1);

    row.style = {
      font: fonts.broadwayRedOutline20
    };

    const a1 = sheetMock.getCell("A1");
    expect(cellFont(a1)).toEqual(fonts.broadwayRedOutline20);
  });

  it("has effective types", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");

    expect(cellType(a1)).toBe(Enums.ValueType.Null);
    expect(cellEffectiveType(a1)).toBe(Enums.ValueType.Null);

    cellSetValue(a1, 5);
    expect(cellType(a1)).toBe(Enums.ValueType.Number);
    expect(cellEffectiveType(a1)).toBe(Enums.ValueType.Number);

    cellSetValue(a1, "Hello, World!");
    expect(cellType(a1)).toBe(Enums.ValueType.String);
    expect(cellEffectiveType(a1)).toBe(Enums.ValueType.String);

    cellSetValue(a1, new Date());
    expect(cellType(a1)).toBe(Enums.ValueType.Date);
    expect(cellEffectiveType(a1)).toBe(Enums.ValueType.Date);

    cellSetValue(a1, { formula: "A2", result: 5 });
    expect(cellType(a1)).toEqual(Enums.ValueType.Formula);
    expect(cellEffectiveType(a1)).toBe(Enums.ValueType.Number);

    cellSetValue(a1, { formula: "A2", result: "Hello, World!" });
    expect(cellType(a1)).toEqual(Enums.ValueType.Formula);
    expect(cellEffectiveType(a1)).toBe(Enums.ValueType.String);

    cellSetValue(a1, { hyperlink: "http://www.link.com", text: "www.link.com" });
    expect(cellType(a1)).toEqual(Enums.ValueType.Hyperlink);
    expect(cellEffectiveType(a1)).toBe(Enums.ValueType.Hyperlink);
  });

  it("shares formulas", () => {
    const a1 = sheetMock.getCell("A1");
    const b1 = sheetMock.getCell("B1");
    const c1 = sheetMock.getCell("C1");

    cellSetValue(a1, 1);
    cellSetValue(b1, { formula: "A1+1", result: 2 });
    cellSetValue(c1, { sharedFormula: "B1", result: 3 });

    expect(cellType(b1)).toBe(Enums.ValueType.Formula);
    expect(cellFormulaType(b1)).toBe(Enums.FormulaType.Master);
    expect(cellType(c1)).toBe(Enums.ValueType.Formula);
    expect(cellFormulaType(c1)).toBe(Enums.FormulaType.Shared);
    expect(cellFormula(c1)).toBe("B1+1");
  });

  it("escapes dangerous html", () => {
    const a1 = sheetMock.getCell("A1");

    cellSetValue(a1, '<script>alert("yoohoo")</script>');

    expect(cellHtml(a1)).toBe("&lt;script&gt;alert(&quot;yoohoo&quot;)&lt;/script&gt;");
  });
  it("can set comment", () => {
    const a1 = sheetMock.getCell("A1");

    const comment = {
      texts: [
        {
          font: {
            size: 12,
            color: { theme: 0 },
            name: "Calibri",
            family: 2,
            scheme: "minor"
          },
          text: "This is "
        }
      ],
      margins: {
        insetmode: "auto",
        inset: [0.13, 0.13, 0.25, 0.25]
      },
      protection: {
        locked: "True",
        lockText: "True"
      },
      editAs: "twoCells"
    };

    cellSetNote(a1, comment as never);
    cellSetValue(a1, "test set value");

    expect(cellGetModel(a1).comment!.type).toBe("note");
    expect(cellGetModel(a1).comment!.note).toEqual(comment);
  });

  it("Cell comments supports setting margins, protection, and position properties", () => {
    const a1 = sheetMock.getCell("A1");

    const comment = {
      texts: [
        {
          font: {
            size: 12,
            color: { theme: 0 },
            name: "Calibri",
            family: 2,
            scheme: "minor"
          },
          text: "This is "
        }
      ],
      protection: {
        locked: "False",
        lockText: "True"
      }
    };

    cellSetNote(a1, comment as never);
    cellSetValue(a1, "test set value");

    expect(cellGetModel(a1).comment!.type).toBe("note");
    expect(cellGetModel(a1).comment!.note!.texts).toEqual(comment.texts);
    expect(cellGetModel(a1).comment!.note!.protection).toEqual(comment.protection);
    expect(cellGetModel(a1).comment!.note!.margins!.insetmode).toBe("auto");
    expect(cellGetModel(a1).comment!.note!.margins!.inset).toEqual([0.13, 0.13, 0.25, 0.25]);
    expect(cellGetModel(a1).comment!.note!.editAs).toBe("absolute");
  });

  // ===========================================================================
  // Boolean Values
  // ===========================================================================

  it("stores boolean values", () => {
    const a1 = sheetMock.getCell("A1");

    cellSetValue(a1, true);
    expect(cellGetValue(a1)).toBe(true);
    expect(cellType(a1)).toBe(Enums.ValueType.Boolean);

    cellSetValue(a1, false);
    expect(cellGetValue(a1)).toBe(false);
    expect(cellType(a1)).toBe(Enums.ValueType.Boolean);
  });

  // ===========================================================================
  // Error Values
  // ===========================================================================

  it("stores error values", () => {
    const a1 = sheetMock.getCell("A1");

    cellSetValue(a1, { error: "#DIV/0!" });
    expect(cellGetValue(a1)).toEqual({ error: "#DIV/0!" });
    expect(cellType(a1)).toBe(Enums.ValueType.Error);

    cellSetValue(a1, { error: "#VALUE!" });
    expect(cellGetValue(a1)).toEqual({ error: "#VALUE!" });
    expect(cellType(a1)).toBe(Enums.ValueType.Error);

    cellSetValue(a1, { error: "#REF!" });
    expect(cellGetValue(a1)).toEqual({ error: "#REF!" });
    expect(cellType(a1)).toBe(Enums.ValueType.Error);

    cellSetValue(a1, { error: "#NAME?" });
    expect(cellGetValue(a1)).toEqual({ error: "#NAME?" });
    expect(cellType(a1)).toBe(Enums.ValueType.Error);

    cellSetValue(a1, { error: "#N/A" });
    expect(cellGetValue(a1)).toEqual({ error: "#N/A" });
    expect(cellType(a1)).toBe(Enums.ValueType.Error);
  });

  // ===========================================================================
  // Rich Text Values
  // ===========================================================================

  it("stores rich text values", () => {
    const a1 = sheetMock.getCell("A1");

    const richText = {
      richText: [
        { font: { bold: true }, text: "Bold " },
        { font: { italic: true }, text: "Italic" }
      ]
    };

    cellSetValue(a1, richText);
    expect(cellType(a1)).toBe(Enums.ValueType.RichText);
    expect(cellText(a1)).toBe("Bold Italic");
  });

  // ===========================================================================
  // text Getter
  // ===========================================================================

  it("text getter returns appropriate string for each type", () => {
    const a1 = sheetMock.getCell("A1");

    // null
    cellSetValue(a1, null);
    expect(cellText(a1)).toBe("");

    // number
    cellSetValue(a1, 42);
    expect(cellText(a1)).toBe("42");

    // string
    cellSetValue(a1, "hello");
    expect(cellText(a1)).toBe("hello");

    // boolean
    cellSetValue(a1, true);
    expect(cellText(a1)).toBe("true");

    cellSetValue(a1, false);
    expect(cellText(a1)).toBe("false");

    // date
    cellSetValue(a1, new Date(2024, 0, 15));
    expect(cellText(a1)).toBeTypeOf("string");
    expect(cellText(a1).length).toBeGreaterThan(0);

    // hyperlink
    cellSetValue(a1, { text: "link text", hyperlink: "https://example.com" });
    expect(cellText(a1)).toBe("link text");

    // formula
    cellSetValue(a1, { formula: "A2+1", result: 99 });
    expect(cellText(a1)).toBe("99");
  });

  // ===========================================================================
  // toCsvString
  // ===========================================================================

  it("toCsvString returns correct CSV representation", () => {
    const a1 = sheetMock.getCell("A1");

    cellSetValue(a1, null);
    expect(cellToCsvString(a1)).toBe("");

    cellSetValue(a1, 42);
    expect(cellToCsvString(a1)).toBe("42");

    // Strings are quoted in CSV
    cellSetValue(a1, "hello");
    expect(cellToCsvString(a1)).toBe('"hello"');

    // Booleans are represented as 1/0 in CSV
    cellSetValue(a1, true);
    expect(cellToCsvString(a1)).toBe(1);

    cellSetValue(a1, false);
    expect(cellToCsvString(a1)).toBe(0);
  });

  // ===========================================================================
  // toString
  // ===========================================================================

  it("toString returns string representation", () => {
    const a1 = sheetMock.getCell("A1");

    cellSetValue(a1, null);
    expect(cellToString(a1)).toBe("");

    cellSetValue(a1, 42);
    expect(cellToString(a1)).toBe("42");

    cellSetValue(a1, "hello");
    expect(cellToString(a1)).toBe("hello");
  });

  // ===========================================================================
  // isHyperlink / hyperlink getters
  // ===========================================================================

  it("isHyperlink and hyperlink getters", () => {
    const a1 = sheetMock.getCell("A1");

    cellSetValue(a1, "plain text");
    expect(cellIsHyperlink(a1)).toBe(false);
    expect(cellHyperlink(a1)).toBeUndefined();

    cellSetValue(a1, { text: "link", hyperlink: "https://example.com" });
    expect(cellIsHyperlink(a1)).toBe(true);
    expect(cellHyperlink(a1)).toBe("https://example.com");

    cellSetValue(a1, 42);
    expect(cellIsHyperlink(a1)).toBe(false);
    expect(cellHyperlink(a1)).toBeUndefined();
  });

  // ===========================================================================
  // fullAddress
  // ===========================================================================

  it("fullAddress includes sheet, row, col, and address", () => {
    const a1 = sheetMock.getCell("A1");

    const full = cellFullAddress(a1);
    expect(full).toHaveProperty("row");
    expect(full).toHaveProperty("col");
    expect(full).toHaveProperty("address");
    expect(full.address).toBe("A1");
    expect(full.row).toBe(1);
    expect(full.col).toBe(1);
  });

  // ===========================================================================
  // $col$row
  // ===========================================================================

  it("$col$row returns absolute reference", () => {
    const b3 = sheetMock.getCell("B3");
    expect(cellAbsoluteAddress(b3)).toBe("$B$3");
  });

  // ===========================================================================
  // row and col getters
  // ===========================================================================

  it("row and col getters return numeric values", () => {
    const c5 = sheetMock.getCell("C5");
    expect(cellFullAddress(c5).row).toBe(5);
    expect(cellFullAddress(c5).col).toBe(3);
  });

  // ===========================================================================
  // formula / result getters
  // ===========================================================================

  it("formula and result getters work for formula cells", () => {
    const a1 = sheetMock.getCell("A1");

    // Non-formula
    cellSetValue(a1, 42);
    expect(cellFormula(a1)).toBeUndefined();
    expect(cellResult(a1)).toBeUndefined();

    // Formula with numeric result
    cellSetValue(a1, { formula: "B1+C1", result: 100 });
    expect(cellFormula(a1)).toBe("B1+C1");
    expect(cellResult(a1)).toBe(100);

    // Formula with string result
    cellSetValue(a1, { formula: 'CONCATENATE("a","b")', result: "ab" });
    expect(cellFormula(a1)).toBe('CONCATENATE("a","b")');
    expect(cellResult(a1)).toBe("ab");
  });

  // ===========================================================================
  // destroy
  // ===========================================================================

  it("destroy clears internal value reference", () => {
    const a1 = sheetMock.getCell("A1");

    cellSetValue(a1, "something");
    cellSetFont(a1, fonts.arialBlackUI14);
    expect(cellGetValue(a1)).toBe("something");

    cellDestroy(a1);
    // After destroy, accessing .type throws because _value is nullified.
    // This verifies destroy truly clears the cell's internal state.
    expect(() => cellType(a1)).toThrow();
  });

  it("cell inherits independent copy of row/column style at construction", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "test");

    // Set row style before creating cells
    rowSetFont(Worksheet.getRow(ws, 1), { bold: true, size: 14 });
    cellSetValue(rowGetCell(Worksheet.getRow(ws, 1), 1), "A1");

    // New cell in the same row should inherit font
    cellSetValue(rowGetCell(Worksheet.getRow(ws, 1), 2), "B1");
    expect(Cell.getStyle(ws, "B1").font).toEqual({ bold: true, size: 14 });

    // Mutating B1's font should not affect A1
    Cell.getStyle(ws, "B1").font!.bold = false;
    expect(Cell.getStyle(ws, "B1").font!.bold).toBe(false);
    expect(Cell.getStyle(ws, "A1").font!.bold).toBe(true);
  });

  // ===========================================================================
  // displayText — formatted value (applies numFmt)
  // ===========================================================================

  describe("displayText", () => {
    it("returns empty string for null / empty cell", () => {
      const a1 = sheetMock.getCell("A1");
      expect(cellDisplayText(a1)).toBe("");
    });

    it("returns the string value for string cells (no numFmt)", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, "hello");
      expect(cellDisplayText(a1)).toBe("hello");
    });

    it("applies numFmt to number cells", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, 1234.5);
      cellSetNumFmt(a1, "#,##0.00");
      expect(cellDisplayText(a1)).toBe("1,234.50");
    });

    it("formats number as percentage", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, 0.125);
      cellSetNumFmt(a1, "0.00%");
      expect(cellDisplayText(a1)).toBe("12.50%");
    });

    it("formats a Date using the cell's numFmt, not cellToString(Date)", () => {
      // Repro for issue #144: a date cell with numFmt "mm-dd-yy" should render
      // as "04-12-19", not the JS Date.cellToString(prototype) output that cell.text
      // would produce.
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, new Date(Date.UTC(2019, 3, 12)));
      cellSetNumFmt(a1, "mm-dd-yy");
      expect(cellDisplayText(a1)).toBe("04-12-19");

      // cell.text still returns the JS Date toString() for backwards compat.
      expect(cellText(a1)).not.toBe("04-12-19");
    });

    it("applies alternate date numFmt (dd.mm.yyyy)", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, new Date(Date.UTC(2019, 3, 12)));
      cellSetNumFmt(a1, "dd.mm.yyyy");
      expect(cellDisplayText(a1)).toBe("12.04.2019");
    });

    it("falls back to yyyy-mm-dd for a Date cell with no numFmt", () => {
      // Without a numFmt, Excel uses a locale-dependent short date; we emit
      // an ISO-like default instead of the raw serial number.
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, new Date(Date.UTC(2019, 3, 12)));
      expect(cellDisplayText(a1)).toBe("2019-04-12");
    });

    it("falls back to yyyy-mm-dd hh:mm:ss for a Date-with-time cell with no numFmt", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, new Date(Date.UTC(2019, 3, 12, 15, 30, 45)));
      expect(cellDisplayText(a1)).toBe("2019-04-12 15:30:45");
    });

    it("resolves month vs minute per-mm-occurrence in mixed date-time formats", () => {
      // Regression: previously a global `hasTimeContext` flag classified every
      // `mm` in "yyyy-mm-dd hh:mm:ss" as minutes, producing "2019-30-12 15:30:45".
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, new Date(Date.UTC(2019, 3, 12, 15, 30, 45)));
      cellSetNumFmt(a1, "yyyy-mm-dd hh:mm:ss");
      expect(cellDisplayText(a1)).toBe("2019-04-12 15:30:45");
    });

    it("uses the formula result for formula cells", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, { formula: "A2*2", result: 42.5 });
      cellSetNumFmt(a1, "0.0");
      expect(cellDisplayText(a1)).toBe("42.5");
    });

    it("uses the formula result Date with numFmt", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, { formula: "TODAY()", result: new Date(Date.UTC(2019, 3, 12)) });
      cellSetNumFmt(a1, "yyyy-mm-dd");
      expect(cellDisplayText(a1)).toBe("2019-04-12");
    });

    it("returns empty string for a formula without a result", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, { formula: "A2" });
      expect(cellDisplayText(a1)).toBe("");
    });

    it("falls back to cell.text for hyperlinks", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, { hyperlink: "http://www.link.com", text: "click me" });
      expect(cellDisplayText(a1)).toBe("click me");
    });

    it("accepts a NumFmt object (formatCode on style)", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, 1234);
      // NumFmt object form
      a1.style.numFmt = { id: 3, formatCode: "#,##0" };
      expect(cellDisplayText(a1)).toBe("1,234");
    });

    it("renders booleans as TRUE/FALSE", () => {
      const a1 = sheetMock.getCell("A1");
      cellSetValue(a1, true);
      expect(cellDisplayText(a1)).toBe("TRUE");
      cellSetValue(a1, false);
      expect(cellDisplayText(a1)).toBe("FALSE");
    });
  });
});
