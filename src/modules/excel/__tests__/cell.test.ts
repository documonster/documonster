import { Cell } from "@excel/cell";
import { Enums } from "@excel/enums";
import type { CellHyperlinkValue } from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import { describe, it, expect, beforeEach } from "vitest";

import { Workbook } from "../../../index";

interface SheetMock {
  rows: any[];
  columns: any[];
  reset(): void;
  findRow(num: number): any;
  getRow(num: number): any;
  findColumn(num: number): any;
  getColumn(num: number): any;
  createRow(num: number): any;
  createColumn(num: number): any;
  getCell(address: string): any;
  findCell(address: string): any;
}

const sheetMock: SheetMock = {
  rows: [],
  columns: [],
  reset() {
    this.rows = [];
    this.columns = [];
  },
  findRow(num) {
    return this.rows[num];
  },
  getRow(num) {
    return this.rows[num] || this.createRow(num);
  },
  findColumn(num) {
    return this.columns[num];
  },
  getColumn(num) {
    return this.columns[num] || this.createColumn(num);
  },
  createRow(num) {
    this.rows[num] = {
      cells: [],
      findCell(col: number) {
        return this.cells[col];
      },
      getCell(col: number) {
        return this.cells[col] || this.createCell(col);
      },
      createCell(col: number) {
        const address = colCache.encodeAddress(this.number, col);
        const column = sheetMock.getColumn(col);
        return (this.cells[col] = new Cell(this, column, address));
      },
      number: num,
      get worksheet() {
        return sheetMock;
      }
    };
    return this.rows[num];
  },
  createColumn(num) {
    this.columns[num] = {
      number: num,
      letter: colCache.n2l(num)
    };
    return this.columns[num];
  },
  getCell(address) {
    const fullAddress = colCache.decodeAddress(address);
    const row = this.getRow(fullAddress.row);
    return row.getCell(fullAddress.col);
  },
  findCell(address) {
    const fullAddress = colCache.decodeAddress(address);
    const row = this.getRow(fullAddress.row);
    return row && row.findCell(fullAddress.col);
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
    expect(a1.$col$row).toBe("$A$1");

    expect(a1.type).toBe(Enums.ValueType.Null);

    expect((a1.value = 5)).toBe(5);
    expect(a1.value).toBe(5);
    expect(a1.type).toBe(Enums.ValueType.Number);

    const strValue = "Hello, World!";
    expect((a1.value = strValue)).toBe(strValue);
    expect(a1.value).toBe(strValue);
    expect(a1.type).toBe(Enums.ValueType.String);

    const dateValue = new Date();
    expect((a1.value = dateValue)).toBe(dateValue);
    expect(a1.value).toBe(dateValue);
    expect(a1.type).toBe(Enums.ValueType.Date);

    const formulaValue = { formula: "A2", result: 5 };
    expect((a1.value = formulaValue)).toEqual(formulaValue);
    expect(a1.value).toEqual(formulaValue);
    expect(a1.type).toBe(Enums.ValueType.Formula);

    const hyperlinkValue = {
      hyperlink: "http://www.link.com",
      text: "www.link.com"
    };
    expect((a1.value = hyperlinkValue)).toEqual(hyperlinkValue);
    expect(a1.value).toEqual(hyperlinkValue);
    expect(a1.type).toBe(Enums.ValueType.Hyperlink);

    expect((a1.value = null)).toBeNull();
    expect(a1.type).toBe(Enums.ValueType.Null);

    expect((a1.value = { json: "data" })).toEqual({ json: "data" });
    expect(a1.type).toBe(Enums.ValueType.String);
  });
  it("validates options on construction", () => {
    const row = sheetMock.getRow(1);
    const column = sheetMock.getColumn(1);

    expect(() => {
      new Cell(null as any, null as any, "A1");
    }).toThrow(Error);
    expect(() => {
      new Cell(row, null as any, "A1");
    }).toThrow(Error);
    expect(() => {
      new Cell(null as any, column, "A1");
    }).toThrow(Error);
  });
  it("merges", () => {
    const a1 = sheetMock.getCell("A1");
    const a2 = sheetMock.getCell("A2");

    a1.value = 5;
    a2.value = "Hello, World!";

    a2.merge(a1);

    expect(a2.value).toBe(5);
    expect(a2.type).toBe(Enums.ValueType.Merge);
    expect(a1._mergeCount).toBe(1);
    expect(a1.isMerged).toBeTruthy();
    expect(a2.isMerged).toBeTruthy();
    expect(a2.isMergedTo(a1)).toBeTruthy();
    expect(a2.master).toBe(a1);
    expect(a1.master).toBe(a1);

    // assignment of slaves write to the master
    a2.value = 7;
    expect(a1.value).toBe(7);

    // assignment of strings should add 1 ref
    const strValue = "Boo!";
    a2.value = strValue;
    expect(a1.value).toBe(strValue);

    // unmerge should work also
    a2.unmerge();
    expect(a2.type).toBe(Enums.ValueType.Null);
    expect(a1._mergeCount).toBe(0);
    expect(a1.isMerged).not.toBeTruthy();
    expect(a2.isMerged).not.toBeTruthy();
    expect(a2.isMergedTo(a1)).not.toBeTruthy();
    expect(a2.master).toBe(a2);
    expect(a1.master).toBe(a1);
  });

  it("upgrades from string to hyperlink", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");

    const strValue = "www.link.com";
    const linkValue = "http://www.link.com";

    a1.value = strValue;

    a1._upgradeToHyperlink(linkValue);

    expect(a1.type).toBe(Enums.ValueType.Hyperlink);
  });

  it("upgrades from rich text to hyperlink, preserving runs (issue #142)", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");
    const linkValue = "http://www.link.com";

    a1.value = {
      richText: [{ text: "bold", font: { bold: true } }, { text: "-plain" }]
    };

    a1._upgradeToHyperlink(linkValue);

    expect(a1.type).toBe(Enums.ValueType.Hyperlink);
    // CellHyperlinkValue.text contract: always a string
    const v = a1.value as CellHyperlinkValue;
    expect(typeof v.text).toBe("string");
    expect(v.text).toBe("bold-plain");
    expect(v.hyperlink).toBe(linkValue);
    expect(v.richText).toEqual([{ text: "bold", font: { bold: true } }, { text: "-plain" }]);
  });

  it("hyperlink value setter normalizes rich text into plain-text mirror (issue #142)", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");

    a1.value = {
      richText: [{ text: "hello " }, { text: "world", font: { italic: true } }],
      hyperlink: "https://example.com",
      // Intentionally inconsistent: user-supplied text should be ignored
      // when richText is present, to keep the invariant text===flatten(richText).
      text: "ignored"
    };

    expect(a1.type).toBe(Enums.ValueType.Hyperlink);
    const v = a1.value as CellHyperlinkValue;
    expect(v.text).toBe("hello world");
    expect(typeof v.text).toBe("string");
    expect(v.hyperlink).toBe("https://example.com");
    expect(v.richText).toEqual([{ text: "hello " }, { text: "world", font: { italic: true } }]);
  });

  it("hyperlink cell.text always returns a string, even for rich-text link (issue #142)", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");
    a1.value = {
      richText: [{ text: "abc" }, { text: "def" }],
      hyperlink: "https://example.com"
    };

    expect(typeof a1.text).toBe("string");
    expect(a1.text).toBe("abcdef");
    expect(a1.toString()).toBe("abcdef");
  });

  it("does not classify { richText: [] } as a RichText cell", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");
    // An empty richText array carries no content. It must not produce a
    // RichText cell with no runs (which would later flatten to "").
    a1.value = { richText: [] } as unknown as CellHyperlinkValue;
    expect(a1.type).not.toBe(Enums.ValueType.RichText);
  });

  it("does not classify { richText: [], hyperlink } as a Hyperlink cell", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");
    // Empty richText and no plain text => not a valid hyperlink display.
    a1.value = {
      richText: [],
      hyperlink: "https://example.com"
    } as unknown as CellHyperlinkValue;
    expect(a1.type).not.toBe(Enums.ValueType.Hyperlink);
  });

  it("doesn't upgrade from non-string to hyperlink", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");

    const linkValue = "http://www.link.com";

    // null
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).toBe(Enums.ValueType.Null);

    // number
    a1.value = 5;
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).toBe(Enums.ValueType.Number);

    // date
    a1.value = new Date();
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).toBe(Enums.ValueType.Date);

    // formula
    a1.value = { formula: "A2" };
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).toBe(Enums.ValueType.Formula);

    // hyperlink
    a1.value = { hyperlink: "http://www.link2.com", text: "www.link2.com" };
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).toEqual(Enums.ValueType.Hyperlink);

    // cleanup
    a1.value = null;
  });

  it("inherits column styles", () => {
    sheetMock.getRow(1);
    const column = sheetMock.getColumn(1);

    column.style = {
      font: fonts.arialBlackUI14
    };

    const a1 = sheetMock.getCell("A1");
    expect(a1.font).toEqual(fonts.arialBlackUI14);
  });

  it("inherits row styles", () => {
    const row = sheetMock.getRow(1);
    sheetMock.getColumn(1);

    row.style = {
      font: fonts.broadwayRedOutline20
    };

    const a1 = sheetMock.getCell("A1");
    expect(a1.font).toEqual(fonts.broadwayRedOutline20);
  });

  it("has effective types", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell("A1");

    expect(a1.type).toBe(Enums.ValueType.Null);
    expect(a1.effectiveType).toBe(Enums.ValueType.Null);

    a1.value = 5;
    expect(a1.type).toBe(Enums.ValueType.Number);
    expect(a1.effectiveType).toBe(Enums.ValueType.Number);

    a1.value = "Hello, World!";
    expect(a1.type).toBe(Enums.ValueType.String);
    expect(a1.effectiveType).toBe(Enums.ValueType.String);

    a1.value = new Date();
    expect(a1.type).toBe(Enums.ValueType.Date);
    expect(a1.effectiveType).toBe(Enums.ValueType.Date);

    a1.value = { formula: "A2", result: 5 };
    expect(a1.type).toEqual(Enums.ValueType.Formula);
    expect(a1.effectiveType).toBe(Enums.ValueType.Number);

    a1.value = { formula: "A2", result: "Hello, World!" };
    expect(a1.type).toEqual(Enums.ValueType.Formula);
    expect(a1.effectiveType).toBe(Enums.ValueType.String);

    a1.value = { hyperlink: "http://www.link.com", text: "www.link.com" };
    expect(a1.type).toEqual(Enums.ValueType.Hyperlink);
    expect(a1.effectiveType).toBe(Enums.ValueType.Hyperlink);
  });

  it("shares formulas", () => {
    const a1 = sheetMock.getCell("A1");
    const b1 = sheetMock.getCell("B1");
    const c1 = sheetMock.getCell("C1");

    a1.value = 1;
    b1.value = { formula: "A1+1", result: 2 };
    c1.value = { sharedFormula: "B1", result: 3 };

    expect(b1.type).toBe(Enums.ValueType.Formula);
    expect(b1.formulaType).toBe(Enums.FormulaType.Master);
    expect(c1.type).toBe(Enums.ValueType.Formula);
    expect(c1.formulaType).toBe(Enums.FormulaType.Shared);
    expect(c1.formula).toBe("B1+1");
  });

  it("escapes dangerous html", () => {
    const a1 = sheetMock.getCell("A1");

    a1.value = '<script>alert("yoohoo")</script>';

    expect(a1.html).toBe("&lt;script&gt;alert(&quot;yoohoo&quot;)&lt;/script&gt;");
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

    a1.note = comment;
    a1.value = "test set value";

    expect(a1.model.comment.type).toBe("note");
    expect(a1.model.comment.note).toEqual(comment);
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

    a1.note = comment;
    a1.value = "test set value";

    expect(a1.model.comment.type).toBe("note");
    expect(a1.model.comment.note.texts).toEqual(comment.texts);
    expect(a1.model.comment.note.protection).toEqual(comment.protection);
    expect(a1.model.comment.note.margins.insetmode).toBe("auto");
    expect(a1.model.comment.note.margins.inset).toEqual([0.13, 0.13, 0.25, 0.25]);
    expect(a1.model.comment.note.editAs).toBe("absolute");
  });

  // ===========================================================================
  // Boolean Values
  // ===========================================================================

  it("stores boolean values", () => {
    const a1 = sheetMock.getCell("A1");

    a1.value = true;
    expect(a1.value).toBe(true);
    expect(a1.type).toBe(Enums.ValueType.Boolean);

    a1.value = false;
    expect(a1.value).toBe(false);
    expect(a1.type).toBe(Enums.ValueType.Boolean);
  });

  // ===========================================================================
  // Error Values
  // ===========================================================================

  it("stores error values", () => {
    const a1 = sheetMock.getCell("A1");

    a1.value = { error: "#DIV/0!" };
    expect(a1.value).toEqual({ error: "#DIV/0!" });
    expect(a1.type).toBe(Enums.ValueType.Error);

    a1.value = { error: "#VALUE!" };
    expect(a1.value).toEqual({ error: "#VALUE!" });
    expect(a1.type).toBe(Enums.ValueType.Error);

    a1.value = { error: "#REF!" };
    expect(a1.value).toEqual({ error: "#REF!" });
    expect(a1.type).toBe(Enums.ValueType.Error);

    a1.value = { error: "#NAME?" };
    expect(a1.value).toEqual({ error: "#NAME?" });
    expect(a1.type).toBe(Enums.ValueType.Error);

    a1.value = { error: "#N/A" };
    expect(a1.value).toEqual({ error: "#N/A" });
    expect(a1.type).toBe(Enums.ValueType.Error);
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

    a1.value = richText;
    expect(a1.type).toBe(Enums.ValueType.RichText);
    expect(a1.text).toBe("Bold Italic");
  });

  // ===========================================================================
  // text Getter
  // ===========================================================================

  it("text getter returns appropriate string for each type", () => {
    const a1 = sheetMock.getCell("A1");

    // null
    a1.value = null;
    expect(a1.text).toBe("");

    // number
    a1.value = 42;
    expect(a1.text).toBe("42");

    // string
    a1.value = "hello";
    expect(a1.text).toBe("hello");

    // boolean
    a1.value = true;
    expect(a1.text).toBe("true");

    a1.value = false;
    expect(a1.text).toBe("false");

    // date
    a1.value = new Date(2024, 0, 15);
    expect(a1.text).toBeTypeOf("string");
    expect(a1.text.length).toBeGreaterThan(0);

    // hyperlink
    a1.value = { text: "link text", hyperlink: "https://example.com" };
    expect(a1.text).toBe("link text");

    // formula
    a1.value = { formula: "A2+1", result: 99 };
    expect(a1.text).toBe("99");
  });

  // ===========================================================================
  // toCsvString
  // ===========================================================================

  it("toCsvString returns correct CSV representation", () => {
    const a1 = sheetMock.getCell("A1");

    a1.value = null;
    expect(a1.toCsvString()).toBe("");

    a1.value = 42;
    expect(a1.toCsvString()).toBe("42");

    // Strings are quoted in CSV
    a1.value = "hello";
    expect(a1.toCsvString()).toBe('"hello"');

    // Booleans are represented as 1/0 in CSV
    a1.value = true;
    expect(a1.toCsvString()).toBe(1);

    a1.value = false;
    expect(a1.toCsvString()).toBe(0);
  });

  // ===========================================================================
  // toString
  // ===========================================================================

  it("toString returns string representation", () => {
    const a1 = sheetMock.getCell("A1");

    a1.value = null;
    expect(a1.toString()).toBe("");

    a1.value = 42;
    expect(a1.toString()).toBe("42");

    a1.value = "hello";
    expect(a1.toString()).toBe("hello");
  });

  // ===========================================================================
  // isHyperlink / hyperlink getters
  // ===========================================================================

  it("isHyperlink and hyperlink getters", () => {
    const a1 = sheetMock.getCell("A1");

    a1.value = "plain text";
    expect(a1.isHyperlink).toBe(false);
    expect(a1.hyperlink).toBeUndefined();

    a1.value = { text: "link", hyperlink: "https://example.com" };
    expect(a1.isHyperlink).toBe(true);
    expect(a1.hyperlink).toBe("https://example.com");

    a1.value = 42;
    expect(a1.isHyperlink).toBe(false);
    expect(a1.hyperlink).toBeUndefined();
  });

  // ===========================================================================
  // fullAddress
  // ===========================================================================

  it("fullAddress includes sheet, row, col, and address", () => {
    const a1 = sheetMock.getCell("A1");

    const full = a1.fullAddress;
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
    expect(b3.$col$row).toBe("$B$3");
  });

  // ===========================================================================
  // row and col getters
  // ===========================================================================

  it("row and col getters return numeric values", () => {
    const c5 = sheetMock.getCell("C5");
    expect(c5.row).toBe(5);
    expect(c5.col).toBe(3);
  });

  // ===========================================================================
  // formula / result getters
  // ===========================================================================

  it("formula and result getters work for formula cells", () => {
    const a1 = sheetMock.getCell("A1");

    // Non-formula
    a1.value = 42;
    expect(a1.formula).toBeUndefined();
    expect(a1.result).toBeUndefined();

    // Formula with numeric result
    a1.value = { formula: "B1+C1", result: 100 };
    expect(a1.formula).toBe("B1+C1");
    expect(a1.result).toBe(100);

    // Formula with string result
    a1.value = { formula: 'CONCATENATE("a","b")', result: "ab" };
    expect(a1.formula).toBe('CONCATENATE("a","b")');
    expect(a1.result).toBe("ab");
  });

  // ===========================================================================
  // destroy
  // ===========================================================================

  it("destroy clears internal value reference", () => {
    const a1 = sheetMock.getCell("A1");

    a1.value = "something";
    a1.font = fonts.arialBlackUI14;
    expect(a1.value).toBe("something");

    a1.destroy();
    // After destroy, accessing .type throws because _value is nullified.
    // This verifies destroy truly clears the cell's internal state.
    expect(() => a1.type).toThrow();
  });

  it("cell inherits independent copy of row/column style at construction", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");

    // Set row style before creating cells
    ws.getRow(1).font = { bold: true, size: 14 };
    ws.getRow(1).getCell(1).value = "A1";

    // New cell in the same row should inherit font
    ws.getRow(1).getCell(2).value = "B1";
    expect(ws.getCell("B1").font).toEqual({ bold: true, size: 14 });

    // Mutating B1's font should not affect A1
    ws.getCell("B1").font!.bold = false;
    expect(ws.getCell("B1").font!.bold).toBe(false);
    expect(ws.getCell("A1").font!.bold).toBe(true);
  });

  // ===========================================================================
  // displayText — formatted value (applies numFmt)
  // ===========================================================================

  describe("displayText", () => {
    it("returns empty string for null / empty cell", () => {
      const a1 = sheetMock.getCell("A1");
      expect(a1.displayText).toBe("");
    });

    it("returns the string value for string cells (no numFmt)", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = "hello";
      expect(a1.displayText).toBe("hello");
    });

    it("applies numFmt to number cells", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = 1234.5;
      a1.numFmt = "#,##0.00";
      expect(a1.displayText).toBe("1,234.50");
    });

    it("formats number as percentage", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = 0.125;
      a1.numFmt = "0.00%";
      expect(a1.displayText).toBe("12.50%");
    });

    it("formats a Date using the cell's numFmt, not Date.toString()", () => {
      // Repro for issue #144: a date cell with numFmt "mm-dd-yy" should render
      // as "04-12-19", not the JS Date.prototype.toString() output that cell.text
      // would produce.
      const a1 = sheetMock.getCell("A1");
      a1.value = new Date(Date.UTC(2019, 3, 12));
      a1.numFmt = "mm-dd-yy";
      expect(a1.displayText).toBe("04-12-19");

      // cell.text still returns the JS Date toString() for backwards compat.
      expect(a1.text).not.toBe("04-12-19");
    });

    it("applies alternate date numFmt (dd.mm.yyyy)", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = new Date(Date.UTC(2019, 3, 12));
      a1.numFmt = "dd.mm.yyyy";
      expect(a1.displayText).toBe("12.04.2019");
    });

    it("falls back to yyyy-mm-dd for a Date cell with no numFmt", () => {
      // Without a numFmt, Excel uses a locale-dependent short date; we emit
      // an ISO-like default instead of the raw serial number.
      const a1 = sheetMock.getCell("A1");
      a1.value = new Date(Date.UTC(2019, 3, 12));
      expect(a1.displayText).toBe("2019-04-12");
    });

    it("falls back to yyyy-mm-dd hh:mm:ss for a Date-with-time cell with no numFmt", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = new Date(Date.UTC(2019, 3, 12, 15, 30, 45));
      expect(a1.displayText).toBe("2019-04-12 15:30:45");
    });

    it("resolves month vs minute per-mm-occurrence in mixed date-time formats", () => {
      // Regression: previously a global `hasTimeContext` flag classified every
      // `mm` in "yyyy-mm-dd hh:mm:ss" as minutes, producing "2019-30-12 15:30:45".
      const a1 = sheetMock.getCell("A1");
      a1.value = new Date(Date.UTC(2019, 3, 12, 15, 30, 45));
      a1.numFmt = "yyyy-mm-dd hh:mm:ss";
      expect(a1.displayText).toBe("2019-04-12 15:30:45");
    });

    it("uses the formula result for formula cells", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = { formula: "A2*2", result: 42.5 };
      a1.numFmt = "0.0";
      expect(a1.displayText).toBe("42.5");
    });

    it("uses the formula result Date with numFmt", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = { formula: "TODAY()", result: new Date(Date.UTC(2019, 3, 12)) };
      a1.numFmt = "yyyy-mm-dd";
      expect(a1.displayText).toBe("2019-04-12");
    });

    it("returns empty string for a formula without a result", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = { formula: "A2" };
      expect(a1.displayText).toBe("");
    });

    it("falls back to cell.text for hyperlinks", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = { hyperlink: "http://www.link.com", text: "click me" };
      expect(a1.displayText).toBe("click me");
    });

    it("accepts a NumFmt object (formatCode on style)", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = 1234;
      // NumFmt object form
      a1.style.numFmt = { id: 3, formatCode: "#,##0" };
      expect(a1.displayText).toBe("1,234");
    });

    it("renders booleans as TRUE/FALSE", () => {
      const a1 = sheetMock.getCell("A1");
      a1.value = true;
      expect(a1.displayText).toBe("TRUE");
      a1.value = false;
      expect(a1.displayText).toBe("FALSE");
    });
  });
});
