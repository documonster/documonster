import { testUtils } from "@excel/__tests__/shared";
import { describe, it, expect } from "vitest";

import { Workbook, ValueType } from "../../../index";

// =============================================================================
// Tests
// =============================================================================

describe("Workbook", () => {
  // ===========================================================================
  // Worksheet Access
  // ===========================================================================

  describe("worksheet access", () => {
    it("returns undefined for non-existent sheet by name", () => {
      const wb = new Workbook();
      wb.addWorksheet("first");
      expect(wb.getWorksheet("w00t")).toBeUndefined();
    });

    it("returns undefined for sheet 0", () => {
      const wb = new Workbook();
      wb.addWorksheet("first");
      expect(wb.getWorksheet(0)).toBeUndefined();
    });

    it("returns correct sheet by id after accessing worksheets or eachSheet", () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet("first");

      wb.eachSheet(() => {});
      const numSheets = wb.worksheets.length;

      expect(numSheets).toBe(1);
      expect(wb.getWorksheet(0)).toBeUndefined();
      expect(wb.getWorksheet(1) === sheet).toBe(true);
    });

    it("returns first worksheet when called with no arguments", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("first");
      wb.addWorksheet("second");

      expect(wb.getWorksheet()).toBe(ws1);
    });

    it("returns worksheet by name", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("target");
      wb.addWorksheet("other");

      expect(wb.getWorksheet("target")).toBe(ws);
    });

    it("returns worksheet by name case-insensitively", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("MySheet");

      expect(wb.getWorksheet("MySheet")).toBe(ws);
      expect(wb.getWorksheet("mysheet")).toBe(ws);
      expect(wb.getWorksheet("MYSHEET")).toBe(ws);
      expect(wb.getWorksheet("mySheet")).toBe(ws);
    });

    it("getWorksheet finds sheet that addWorksheet would reject as duplicate", () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet");

      // getWorksheet should find the existing sheet with different casing,
      // consistent with addWorksheet which would reject "sheet" as a duplicate
      const existing = wb.getWorksheet("sheet");
      expect(existing).toBeDefined();
      expect(existing!.name).toBe("Sheet");
    });

    it("returns worksheet by numeric id", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("first");
      const ws2 = wb.addWorksheet("second");

      expect(wb.getWorksheet(ws1.id)).toBe(ws1);
      expect(wb.getWorksheet(ws2.id)).toBe(ws2);
    });
  });

  // ===========================================================================
  // Worksheet Management
  // ===========================================================================

  describe("worksheet management", () => {
    it("removeWorksheet by id", () => {
      const wb = new Workbook();
      wb.addWorksheet("first");
      const ws2 = wb.addWorksheet("second");
      wb.addWorksheet("third");

      expect(wb.worksheets.length).toBe(3);

      wb.removeWorksheet(ws2.id);
      expect(wb.worksheets.length).toBe(2);
      expect(wb.getWorksheet("second")).toBeUndefined();
      expect(wb.getWorksheet("first")).toBeDefined();
      expect(wb.getWorksheet("third")).toBeDefined();
    });

    it("removeWorksheet by name (string id)", () => {
      const wb = new Workbook();
      wb.addWorksheet("alpha");
      wb.addWorksheet("beta");

      wb.removeWorksheet("alpha");
      expect(wb.worksheets.length).toBe(1);
      expect(wb.getWorksheet("alpha")).toBeUndefined();
      expect(wb.getWorksheet("beta")).toBeDefined();
    });

    it("removeWorksheet by name case-insensitively", () => {
      const wb = new Workbook();
      wb.addWorksheet("Alpha");
      wb.addWorksheet("Beta");

      wb.removeWorksheet("alpha");
      expect(wb.worksheets.length).toBe(1);
      expect(wb.getWorksheet("Alpha")).toBeUndefined();
      expect(wb.getWorksheet("Beta")).toBeDefined();
    });

    it("worksheets getter returns sheets in order", () => {
      const wb = new Workbook();
      wb.addWorksheet("A");
      wb.addWorksheet("B");
      wb.addWorksheet("C");

      const names = wb.worksheets.map(ws => ws.name);
      expect(names).toEqual(["A", "B", "C"]);
    });

    it("eachSheet iterates all worksheets", () => {
      const wb = new Workbook();
      wb.addWorksheet("one");
      wb.addWorksheet("two");
      wb.addWorksheet("three");

      const names: string[] = [];
      wb.eachSheet(ws => names.push(ws.name));
      expect(names).toEqual(["one", "two", "three"]);
    });

    it("addWorksheet rejects case-insensitive duplicate names", () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet");

      expect(() => wb.addWorksheet("sheet")).toThrow(/already exists/i);
      expect(() => wb.addWorksheet("SHEET")).toThrow(/already exists/i);
    });

    it("allows renaming a worksheet to a different casing of the same name", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet");

      // Renaming "Sheet" to "SHEET" should not throw -- it's the same sheet
      ws.name = "SHEET";
      expect(ws.name).toBe("SHEET");

      ws.name = "sheet";
      expect(ws.name).toBe("sheet");
    });

    it("renaming a worksheet still rejects duplicate names with other sheets", () => {
      const wb = new Workbook();
      wb.addWorksheet("Alpha");
      const ws2 = wb.addWorksheet("Beta");

      // Renaming Beta to "alpha" (case-insensitive match with Alpha) should throw
      expect(() => {
        ws2.name = "alpha";
      }).toThrow(/already exists/i);
    });
  });

  // ===========================================================================
  // Cell Types & Values
  // ===========================================================================

  describe("cell types", () => {
    it("stores shared string values properly", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      ws.getCell("A1").value = "Hello, World!";
      ws.getCell("A2").value = "Hello";
      ws.getCell("B2").value = "World";
      ws.getCell("C2").value = {
        formula: 'CONCATENATE(A2, ", ", B2, "!")',
        result: "Hello, World!"
      };
      ws.getCell("A3").value = `${["Hello", "World"].join(", ")}!`;

      // A1 and A3 should reference the same string object
      expect(ws.getCell("A1").value).toBe(ws.getCell("A3").value);
      // A1 and C2 result should share the same string
      expect(ws.getCell("A1").value).toBe(ws.getCell("C2").result);
    });

    it("assigns cell types properly", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      ws.getCell("A1").value = 7;
      ws.getCell("B1").value = "Hello, World!";
      ws.getCell("C1").value = 3.14;
      ws.getCell("D1").value = new Date();
      ws.getCell("E1").value = {
        text: "www.google.com",
        hyperlink: "http://www.google.com"
      };
      ws.getCell("A2").value = { formula: "A1", result: 7 };
      ws.getCell("B2").value = {
        formula: 'CONCATENATE("Hello", ", ", "World!")',
        result: "Hello, World!"
      };
      ws.getCell("C2").value = { formula: "D1", result: new Date() };

      expect(ws.getCell("A1").type).toBe(ValueType.Number);
      expect(ws.getCell("B1").type).toBe(ValueType.String);
      expect(ws.getCell("C1").type).toBe(ValueType.Number);
      expect(ws.getCell("D1").type).toBe(ValueType.Date);
      expect(ws.getCell("E1").type).toBe(ValueType.Hyperlink);
      expect(ws.getCell("A2").type).toBe(ValueType.Formula);
      expect(ws.getCell("B2").type).toBe(ValueType.Formula);
      expect(ws.getCell("C2").type).toBe(ValueType.Formula);
    });

    it("assigns rich text", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      ws.getCell("A1").value = {
        richText: [
          {
            font: { size: 12, color: { theme: 0 }, name: "Calibri", family: 2, scheme: "minor" },
            text: "This is "
          },
          {
            font: { italic: true, size: 12, color: { theme: 0 }, name: "Calibri", scheme: "minor" },
            text: "a"
          },
          {
            font: { size: 12, color: { theme: 1 }, name: "Calibri", family: 2, scheme: "minor" },
            text: " "
          },
          {
            font: { size: 12, color: { argb: "FFFF6600" }, name: "Calibri", scheme: "minor" },
            text: "colorful"
          },
          {
            font: { size: 12, color: { theme: 1 }, name: "Calibri", family: 2, scheme: "minor" },
            text: " text "
          },
          {
            font: { size: 12, color: { argb: "FFCCFFCC" }, name: "Calibri", scheme: "minor" },
            text: "with"
          },
          {
            font: { size: 12, color: { theme: 1 }, name: "Calibri", family: 2, scheme: "minor" },
            text: " in-cell "
          },
          {
            font: {
              bold: true,
              size: 12,
              color: { theme: 1 },
              name: "Calibri",
              family: 2,
              scheme: "minor"
            },
            text: "format"
          }
        ]
      };

      expect(ws.getCell("A1").text).toBe("This is a colorful text with in-cell format");
      expect(ws.getCell("A1").type).toBe(ValueType.RichText);
    });
  });

  // ===========================================================================
  // Images
  // ===========================================================================

  describe("images", () => {
    it("addImage and getImage round-trip", () => {
      const wb = new Workbook();
      const imageBuffer = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);

      const id = wb.addImage({
        buffer: imageBuffer,
        extension: "png"
      });

      expect(typeof id).toBe("number");
      const img = wb.getImage(id);
      expect(img).toBeDefined();
      expect(img!.extension).toBe("png");
    });

    it("getImage returns undefined for invalid id", () => {
      const wb = new Workbook();
      expect(wb.getImage(999)).toBeUndefined();
    });
  });

  // ===========================================================================
  // Metadata
  // ===========================================================================

  describe("metadata", () => {
    it("creator and dates can be set and read", () => {
      const wb = new Workbook();
      wb.creator = "Test Author";
      wb.created = new Date(2024, 0, 1);
      wb.modified = new Date(2024, 5, 15);

      expect(wb.creator).toBe("Test Author");
      expect(wb.created).toEqual(new Date(2024, 0, 1));
      expect(wb.modified).toEqual(new Date(2024, 5, 15));
    });

    it("properties can be set", () => {
      const wb = new Workbook();
      wb.title = "My Workbook";
      wb.subject = "Testing";

      expect(wb.title).toBe("My Workbook");
      expect(wb.subject).toBe("Testing");
    });
  });

  // ===========================================================================
  // Defined Names
  // ===========================================================================

  describe("defined names", () => {
    it("definedNames is accessible", () => {
      const wb = new Workbook();
      expect(wb.definedNames).toBeDefined();
    });
  });

  // ===========================================================================
  // Views
  // ===========================================================================

  describe("views", () => {
    it("views can be set and read", () => {
      const wb = new Workbook();
      wb.views = [
        {
          x: 0,
          y: 0,
          width: 10000,
          height: 20000,
          firstSheet: 0,
          activeTab: 0,
          visibility: "visible"
        }
      ];
      expect(wb.views.length).toBe(1);
      expect(wb.views[0].activeTab).toBe(0);
    });
  });

  // ===========================================================================
  // Duplicate Rows
  // ===========================================================================

  describe("duplicateRows", () => {
    it("inserts duplicates", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      ws.getCell("A1").value = "1.1";
      ws.getCell("A1").font = testUtils.styles.fonts.arialBlackUI14;
      ws.getCell("B1").value = "1.2";
      ws.getCell("B1").font = testUtils.styles.fonts.comicSansUdB16;
      ws.getCell("C1").value = "1.3";
      ws.getCell("C1").fill = testUtils.styles.fills.redDarkVertical;
      ws.getRow(1).numFmt = testUtils.styles.numFmts.numFmt1;

      ws.getCell("A2").value = "2.1";
      ws.getCell("A2").alignment = testUtils.styles.namedAlignments.topLeft;
      ws.getCell("B2").value = "2.2";
      ws.getCell("B2").alignment = testUtils.styles.namedAlignments.middleCentre;
      ws.getCell("C2").value = "2.3";
      ws.getCell("C2").alignment = testUtils.styles.namedAlignments.bottomRight;
      ws.getRow(2).numFmt = testUtils.styles.numFmts.numFmt2;

      ws.duplicateRow(1, 2, true);
      expect(ws.getRow(1).values).toEqual([, "1.1", "1.2", "1.3"]);
      expect(ws.getRow(2).values).toEqual([, "1.1", "1.2", "1.3"]);
      expect(ws.getRow(3).values).toEqual([, "1.1", "1.2", "1.3"]);
      expect(ws.getRow(4).values).toEqual([, "2.1", "2.2", "2.3"]);

      for (let i = 1; i <= 3; i++) {
        expect(ws.getCell(`A${i}`).font).toEqual(testUtils.styles.fonts.arialBlackUI14);
        expect(ws.getCell(`B${i}`).font).toEqual(testUtils.styles.fonts.comicSansUdB16);
        expect(ws.getCell(`C${i}`).fill).toEqual(testUtils.styles.fills.redDarkVertical);
      }
      expect(ws.getCell("A4").alignment).toEqual(testUtils.styles.namedAlignments.topLeft);
      expect(ws.getCell("B4").alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell("C4").alignment).toEqual(testUtils.styles.namedAlignments.bottomRight);

      expect(ws.getRow(1).numFmt).toBe(testUtils.styles.numFmts.numFmt1);
      expect(ws.getRow(2).numFmt).toBe(testUtils.styles.numFmts.numFmt1);
      expect(ws.getRow(3).numFmt).toBe(testUtils.styles.numFmts.numFmt1);
      expect(ws.getRow(4).numFmt).toBe(testUtils.styles.numFmts.numFmt2);
    });

    it("overwrites with duplicates", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      ws.getCell("A1").value = "1.1";
      ws.getCell("A1").font = testUtils.styles.fonts.arialBlackUI14;
      ws.getCell("B1").value = "1.2";
      ws.getCell("B1").font = testUtils.styles.fonts.comicSansUdB16;
      ws.getCell("C1").value = "1.3";
      ws.getCell("C1").fill = testUtils.styles.fills.redDarkVertical;
      ws.getRow(1).numFmt = testUtils.styles.numFmts.numFmt1;

      ws.getCell("A2").value = "2.1";
      ws.getCell("A2").alignment = testUtils.styles.namedAlignments.topLeft;
      ws.getCell("B2").value = "2.2";
      ws.getCell("B2").alignment = testUtils.styles.namedAlignments.middleCentre;
      ws.getCell("C2").value = "2.3";
      ws.getCell("C2").alignment = testUtils.styles.namedAlignments.bottomRight;
      ws.getRow(2).numFmt = testUtils.styles.numFmts.numFmt2;

      ws.getCell("A3").value = "3.1";
      ws.getCell("A3").fill = testUtils.styles.fills.redGreenDarkTrellis;
      ws.getCell("B3").value = "3.2";
      ws.getCell("B3").fill = testUtils.styles.fills.blueWhiteHGrad;
      ws.getCell("C3").value = "3.3";
      ws.getCell("C3").fill = testUtils.styles.fills.rgbPathGrad;
      ws.getRow(3).font = testUtils.styles.fonts.broadwayRedOutline20;

      ws.duplicateRow(1, 1, false);
      expect(ws.getRow(1).values).toEqual([, "1.1", "1.2", "1.3"]);
      expect(ws.getRow(2).values).toEqual([, "1.1", "1.2", "1.3"]);
      expect(ws.getRow(3).values).toEqual([, "3.1", "3.2", "3.3"]);

      for (let i = 1; i <= 2; i++) {
        expect(ws.getCell(`A${i}`).font).toEqual(testUtils.styles.fonts.arialBlackUI14);
        expect(ws.getCell(`A${i}`).alignment).toBeUndefined();
        expect(ws.getCell(`B${i}`).font).toEqual(testUtils.styles.fonts.comicSansUdB16);
        expect(ws.getCell(`B${i}`).alignment).toBeUndefined();
        expect(ws.getCell(`C${i}`).fill).toEqual(testUtils.styles.fills.redDarkVertical);
        expect(ws.getCell(`C${i}`).alignment).toBeUndefined();
      }

      expect(ws.getRow(1).numFmt).toBe(testUtils.styles.numFmts.numFmt1);
      expect(ws.getRow(2).numFmt).toBe(testUtils.styles.numFmts.numFmt1);
      expect(ws.getRow(3).numFmt).toBeUndefined();
      expect(ws.getRow(3).font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
    });
  });

  // ===========================================================================
  // Themes
  // ===========================================================================

  describe("themes", () => {
    it("clearThemes removes internal themes", () => {
      const wb = new Workbook();
      // Themes exist by default as undefined, but after loading an XLSX they'd be set.
      // clearThemes() should set _themes to undefined without throwing.
      wb.clearThemes();
      expect((wb as any)._themes).toBeUndefined();
    });
  });

  // ===========================================================================
  // Workbook Protection
  // ===========================================================================

  describe("workbook protection", () => {
    it("protect() sets lockStructure by default", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");
      await wb.protect();

      expect(wb.protection).toBeDefined();
      expect(wb.protection!.lockStructure).toBe(true);
      // No password → no hash fields
      expect(wb.protection!.algorithmName).toBeUndefined();
      expect(wb.protection!.hashValue).toBeUndefined();
    });

    it("protect() with password generates hash fields", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");
      await wb.protect("secret");

      expect(wb.protection).toBeDefined();
      expect(wb.protection!.lockStructure).toBe(true);
      expect(wb.protection!.algorithmName).toBe("SHA-512");
      expect(wb.protection!.hashValue).toBeTruthy();
      expect(wb.protection!.saltValue).toBeTruthy();
      expect(wb.protection!.spinCount).toBe(100000);
    });

    it("protect() with options overrides defaults", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");
      await wb.protect("pass", {
        lockStructure: false,
        lockWindows: true,
        spinCount: 50000
      });

      expect(wb.protection!.lockStructure).toBe(false);
      expect(wb.protection!.lockWindows).toBe(true);
      expect(wb.protection!.spinCount).toBe(50000);
    });

    it("protect() normalizes spinCount edge cases", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");

      // undefined → default 100000
      await wb.protect("a", { lockStructure: true });
      expect(wb.protection!.spinCount).toBe(100000);

      // negative → 0
      await wb.protect("a", { spinCount: -1 });
      expect(wb.protection!.spinCount).toBe(0);

      // fractional → rounded
      await wb.protect("a", { spinCount: 1.8 });
      expect(wb.protection!.spinCount).toBe(2);
    });

    it("unprotect() removes protection", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");
      await wb.protect("secret");
      expect(wb.protection).toBeDefined();

      wb.unprotect();
      expect(wb.protection).toBeUndefined();
    });

    it("protection survives model round-trip", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");
      await wb.protect("secret", { lockWindows: true });

      const model = wb.model;
      expect(model.protection).toBeDefined();
      expect(model.protection!.lockStructure).toBe(true);
      expect(model.protection!.lockWindows).toBe(true);

      const wb2 = new Workbook();
      wb2.addWorksheet("Sheet1");
      wb2.model = model;
      expect(wb2.protection).toBeDefined();
      expect(wb2.protection!.lockStructure).toBe(true);
      expect(wb2.protection!.lockWindows).toBe(true);
      expect(wb2.protection!.algorithmName).toBe("SHA-512");
    });

    it("protection round-trips through XLSX write/load", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");
      await wb.protect("test123", { lockStructure: true });

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      expect(wb2.protection).toBeDefined();
      expect(wb2.protection!.lockStructure).toBe(true);
      expect(wb2.protection!.algorithmName).toBe("SHA-512");
      expect(wb2.protection!.hashValue).toBe(wb.protection!.hashValue);
      expect(wb2.protection!.saltValue).toBe(wb.protection!.saltValue);
      expect(wb2.protection!.spinCount).toBe(100000);
    });

    it("unprotected workbook has no protection in XLSX", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      expect(wb2.protection).toBeUndefined();
    });
  });

  // ===========================================================================
  // Default Font
  // ===========================================================================

  describe("defaultFont", () => {
    it("is undefined by default", () => {
      const wb = new Workbook();
      expect(wb.defaultFont).toBeUndefined();
    });

    it("can be set and read back", () => {
      const wb = new Workbook();
      wb.defaultFont = { name: "Arial", size: 12 };

      expect(wb.defaultFont).toEqual({ name: "Arial", size: 12 });
    });

    it("can be cleared by setting to undefined", () => {
      const wb = new Workbook();
      wb.defaultFont = { name: "Arial", size: 12 };
      wb.defaultFont = undefined;

      expect(wb.defaultFont).toBeUndefined();
    });

    it("survives model round-trip", () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");
      wb.defaultFont = { name: "Times New Roman", size: 14 };

      const model = wb.model;
      expect(model.defaultFont).toEqual({ name: "Times New Roman", size: 14 });

      const wb2 = new Workbook();
      wb2.addWorksheet("Sheet1");
      wb2.model = model;
      expect(wb2.defaultFont).toEqual({ name: "Times New Roman", size: 14 });
    });

    it("round-trips through XLSX write/load", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "test";
      wb.defaultFont = { name: "Arial", size: 12, family: 2 };

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      // After round-trip, the default font should be preserved
      expect(wb2.defaultFont).toBeDefined();
      expect(wb2.defaultFont!.name).toBe("Arial");
      expect(wb2.defaultFont!.size).toBe(12);
    });

    it("writes defaultFont as fontId=0 in styles.xml", async () => {
      const { extractAll } = await import("@archive/unzip/extract");

      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "test";
      wb.defaultFont = { name: "Arial", size: 12 };

      const buffer = await wb.xlsx.writeBuffer();
      const entries = await extractAll(buffer as Uint8Array);
      const stylesXml = new TextDecoder().decode(entries.get("xl/styles.xml")!.data);

      // The first <font> in styles.xml should be Arial 12, not Calibri 11
      const firstFontMatch = stylesXml.match(/<font>([\s\S]*?)<\/font>/);
      expect(firstFontMatch).toBeTruthy();
      const firstFont = firstFontMatch![1];
      expect(firstFont).toContain('val="Arial"');
      expect(firstFont).toContain('val="12"');
      expect(firstFont).not.toContain('val="Calibri"');
    });
  });

  // ===========================================================================
  // calcProperties XLSX Round-Trip
  // ===========================================================================

  describe("calcProperties", () => {
    it("round-trips iterate/iterateCount/iterateDelta through XLSX write/load", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");
      wb.calcProperties = {
        fullCalcOnLoad: true,
        iterate: true,
        iterateCount: 200,
        iterateDelta: 0.01
      };

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      expect(wb2.calcProperties).toBeDefined();
      expect(wb2.calcProperties.fullCalcOnLoad).toBe(true);
      expect(wb2.calcProperties.iterate).toBe(true);
      expect(wb2.calcProperties.iterateCount).toBe(200);
      expect(wb2.calcProperties.iterateDelta).toBe(0.01);
    });

    it("preserves default calcProperties when not explicitly set", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1");

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      // Default: fullCalcOnLoad should be false, iterate fields undefined
      expect(wb2.calcProperties).toBeDefined();
      expect(wb2.calcProperties.fullCalcOnLoad).toBe(false);
      expect(wb2.calcProperties.iterate).toBeUndefined();
      expect(wb2.calcProperties.iterateCount).toBeUndefined();
      expect(wb2.calcProperties.iterateDelta).toBeUndefined();
    });
  });

  // ===========================================================================
  // Formula-based Defined Names XLSX Round-Trip
  // ===========================================================================

  describe("formula-based defined names", () => {
    it("round-trips formula-based defined names through XLSX write/load", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 1;

      // Register a formula-based defined name
      wb.definedNames.addFormula("MyFormula", "OFFSET(Sheet1!$A$1,0,0,3,1)");

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      // The formula-based name should survive round-trip
      const { ranges } = wb2.definedNames.getRanges("MyFormula");
      expect(ranges).toHaveLength(1);
      expect(ranges[0]).toBe("OFFSET(Sheet1!$A$1,0,0,3,1)");
    });

    it("preserves addFormula names alongside cell-reference names", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 1;

      wb.definedNames.add("Sheet1!$A$1:$A$3", "CellRange");
      wb.definedNames.addFormula("FormulaName", "SUM(Sheet1!$A$1:$A$3)");

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      // Cell-reference name
      const cellRange = wb2.definedNames.getRanges("CellRange");
      expect(cellRange.ranges).toHaveLength(1);
      expect(cellRange.ranges[0]).toContain("$A$1");

      // Formula name
      const formulaName = wb2.definedNames.getRanges("FormulaName");
      expect(formulaName.ranges).toHaveLength(1);
      expect(formulaName.ranges[0]).toBe("SUM(Sheet1!$A$1:$A$3)");
    });

    it("does not misclassify sheet names with parentheses as formulas", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Budget (2024)");
      ws.getCell("A1").value = 100;

      wb.definedNames.add("'Budget (2024)'!$A$1", "MyCell");

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const result = wb2.definedNames.getRanges("MyCell");
      expect(result.ranges).toHaveLength(1);
      expect(result.ranges[0]).toContain("$A$1");
      // Must NOT be treated as a formula expression
      expect(result.formulaExpression).toBeUndefined();
    });
  });
});
